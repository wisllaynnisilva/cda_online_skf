# -*- coding: utf-8 -*-
"""
cda_online_skf — coleta da plataforma SKF online (UBU e GERMANO) e carga no
Google Sheets.

Etapas (cada uma isolada; se uma falhar as independentes continuam):
  hierarquia, machines, submachines, parts, points (+ alarms e
  last measurements, numa única chamada por máquina), trends (D-1), notes.
"""

from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, time as dtime, timezone
import threading

import pandas as pd
import urllib3

from cda_utils import (
    Pipeline, anexar_novos, dia_referencia, env_json, gravar_aba, http, log,
)

# O servidor SKF usa certificado próprio: desliga só o aviso (verify=False).
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ---------------------------------------------------------------------------
# 1. CONFIGURAÇÃO
# ---------------------------------------------------------------------------

CREDENTIALS = env_json("SKF_CREDENTIALS")
BASE_URLS = env_json("SKF_BASE")          # {"ubu": "http://...", "germano": "..."}
MAPA_SITE = {"ubu": "UBU", "germano": "GERMANO"}
WORKERS = 6                               # requisições simultâneas por origem

PLANILHAS = {
    "hierarquia": "1w3hVCO0DgMKd9EkeZCRlDy_UXPSMQgNaPJCAJtmgYjI",
    "machines": "1gHImL_Hbr6teYb-sV4hzfoHExVFmTwlrjx93jgAsji0",
    "submachines": "115Ilr5gw8jkqVaHJmotdo5gVxwfLpx07tCTdMwQLLzc",
    "parts": "1ZOW0VeSOqSjCNZVvO48TxnAt6FdKhBSuCwFLq23Rd7k",
    "points": "1ax7LyNRt06naxNwPXXb9kOGLUkmSbZFw9-IjhV4-_Gk",
    "alarms": "172uMRb6-j8yitUVbCYp-iHWcSQGYpi6oqSX1gan2DXU",
    "lastmeasurements": "1GmfGrRxJsOUAHG8yJ6Wfn67SSAhR5dgPQEZXplAFnJI",
    "trends": "1uekxFKio9llwhP9CljTappvO3XaXrrV7g_WT2V3buIo",
    "notes": "1A70P76NH1Lxt3h-Hg0V0NJBaPRHufbjC80m_qtkxfac",
}

# Tabelas históricas: ao passar deste nº de linhas, o arquivo é arquivado como
# '<nome> revNNN' e a planilha ativa recomeça vazia (mesmo ID/link).
LIMITE_LINHAS = {"trends": 1_000_000}

# ---------------------------------------------------------------------------
# 2. CLIENTE DA API (token único por origem, renovado em 401)
# ---------------------------------------------------------------------------

_tokens = {}
_lock = threading.Lock()


def _novo_token(origem):
    resp = http("POST", f"{BASE_URLS[origem]}/token", data=CREDENTIALS, verify=False)
    token = resp.json()["access_token"]
    with _lock:
        _tokens[origem] = token
    return token


def _headers(origem):
    token = _tokens.get(origem) or _novo_token(origem)
    return {"Authorization": f"Bearer {token}", "Accept": "application/json"}


def api_get(origem, caminho, params=None):
    """GET autenticado. 404 devolve None (máquina/ponto sem dados)."""
    url = f"{BASE_URLS[origem]}{caminho}"
    try:
        resp = http(
            "GET", url,
            headers=_headers(origem),
            params=params,
            verify=False,
            ao_401=lambda: (_novo_token(origem), _headers(origem))[1],
        )
    except Exception as exc:  # noqa: BLE001
        resposta = getattr(exc, "response", None)
        if resposta is not None and resposta.status_code == 404:
            return None
        raise
    return resp.json()


def em_paralelo(func, itens, tolerancia=0.05):
    """
    Aplica func em paralelo. Itens que falham mesmo após os retries viram None
    e são registrados no log; se falharem mais de `tolerancia` (5%) dos itens,
    a etapa é abortada para não gravar dados incompletos na planilha.
    """
    itens = list(itens)
    falhas = []

    def seguro(item):
        try:
            return func(item)
        except Exception as exc:  # noqa: BLE001
            falhas.append(item)
            log.error("Falha no item %s: %s", item, exc)
            return None

    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        resultados = list(pool.map(seguro, itens))

    if itens and len(falhas) / len(itens) > tolerancia:
        raise RuntimeError(f"{len(falhas)}/{len(itens)} requisições falharam — etapa abortada")
    if falhas:
        log.warning("%d/%d itens falharam (dentro da tolerância)", len(falhas), len(itens))
    return resultados


# ---------------------------------------------------------------------------
# 3. FUNÇÕES AUXILIARES DE TRATAMENTO
# ---------------------------------------------------------------------------

def separar_path(df, coluna="path"):
    """Cria path_0, path_1, ... a partir do caminho 'A\\B\\C'."""
    df = df.copy()
    df[coluna] = df[coluna].fillna("").astype(str)
    partes = df[coluna].str.split(r"\\", expand=True)
    partes.columns = [f"path_{i}" for i in range(partes.shape[1])]
    df = pd.concat([df.drop(columns=[c for c in partes.columns if c in df]), partes], axis=1)
    if "path_1" not in df.columns:
        df["path_1"] = ""
    df["path_1"] = df["path_1"].fillna("").astype(str).str.strip().str.upper()
    return df


def filtrar_site(df):
    """Mantém apenas linhas cujo path_1 bate com a origem (UBU/GERMANO)."""
    df = separar_path(df)
    origem = df["origem"].fillna("").astype(str).str.strip().str.lower()
    return df[df["path_1"] == origem.map(MAPA_SITE)]


def para_float(valor):
    try:
        return float(str(valor).split()[0].replace(",", "."))
    except (ValueError, IndexError):
        return None


def ids_por_origem(df, coluna, origem):
    return df.loc[df["origem"] == origem, coluna].dropna().tolist()


# ---------------------------------------------------------------------------
# 4. HIERARQUIA (baixada uma vez por origem; usada em 2 etapas)
# ---------------------------------------------------------------------------

def flatten_hierarchy(data, parent_id=None, parent_name=None, rows=None):
    rows = [] if rows is None else rows
    for item in data or []:
        rows.append({
            "id": item.get("id"),
            "name": item.get("name"),
            "description": item.get("description"),
            "path": item.get("path"),
            "tag": item.get("tag"),
            "active": item.get("active"),
            "typeName": item.get("typeName"),
            "typeId": item.get("typeId"),
            "status": item.get("status"),
            "src": item.get("src"),
            "parent_id": parent_id,
            "parent_name": parent_name,
        })
        if item.get("children"):
            flatten_hierarchy(item["children"], item.get("id"), item.get("name"), rows)
    return rows


def etapa_hierarquia():
    dfs = []
    for origem in BASE_URLS:
        df = pd.DataFrame(flatten_hierarchy(api_get(origem, "/v1/hierarchy")))
        df["origem"] = origem
        dfs.append(df)
    df = pd.concat(dfs, ignore_index=True)

    df = separar_path(df)
    df["origem"] = df["origem"].fillna("").str.strip().str.lower()
    df["path_valido"] = df["path_1"] == df["origem"].map(MAPA_SITE)

    df = df[[
        "id", "name", "description", "path", "path_1", "path_valido", "tag",
        "active", "typeName", "typeId", "status", "src", "parent_id",
        "parent_name", "origem",
    ]].copy()
    gravar_aba(PLANILHAS["hierarquia"], df, "hierarquia")
    return df


def etapa_submachines(df_hier):
    df = df_hier[df_hier["typeName"] == "SubMachine"].rename(columns={"parent_name": "parent"})
    df = filtrar_site(df)
    df = df[["origem", "id", "name", "description", "parent", "path", "status", "active"]]
    gravar_aba(PLANILHAS["submachines"], df, "submachines")
    return df


# ---------------------------------------------------------------------------
# 5. MACHINES
# ---------------------------------------------------------------------------

def etapa_machines():
    dfs = []
    for origem in BASE_URLS:
        df = pd.DataFrame(api_get(origem, "/v1/machines") or [])
        df["origem"] = origem
        dfs.append(df)
    df = pd.concat(dfs, ignore_index=True)

    df = filtrar_site(df)[["origem", "id", "name", "path"]].copy()
    log.info("%d ativos encontrados", len(df))
    if df.empty:
        raise RuntimeError("Nenhuma máquina retornada — abortando etapas dependentes")
    gravar_aba(PLANILHAS["machines"], df, "machines")
    return df


# ---------------------------------------------------------------------------
# 6. PARTS
# ---------------------------------------------------------------------------

def etapa_parts(df_machines):
    registros = []
    for origem in BASE_URLS:
        ids = ids_por_origem(df_machines, "id", origem)
        respostas = em_paralelo(lambda mid: api_get(origem, f"/v1/machines/{mid}/parts"), ids)
        for mid, parts in zip(ids, respostas):
            for part in parts or []:
                base = {
                    "MachineId": mid,
                    "PartID": part.get("id"),
                    "PartName": part.get("name"),
                    "Type": part.get("type"),
                    "Ratio": part.get("ratio"),
                    "Brand": part.get("brand"),
                    "Typeno": part.get("typeno"),
                    "SpeedPointId": part.get("speedPointId"),
                    "origem": origem,
                }
                faults = part.get("faultFrequencies") or [{}]
                for f in faults:
                    registros.append({**base, "Name": f.get("name"), "Multiple": f.get("multiple")})

    df = pd.DataFrame(registros)
    if df.empty:
        log.warning("Nenhuma part retornada")
        return df
    df = df.merge(
        df_machines[["origem", "id", "path"]],
        left_on=["origem", "MachineId"], right_on=["origem", "id"], how="left",
    )
    df = filtrar_site(df)
    df = df[[
        "origem", "path", "MachineId", "PartID", "PartName", "Type", "Ratio",
        "Brand", "Typeno", "SpeedPointId", "Name", "Multiple",
    ]].drop_duplicates()
    gravar_aba(PLANILHAS["parts"], df, "parts")
    return df


# ---------------------------------------------------------------------------
# 7. POINTS + ALARMS + LAST MEASUREMENTS
#    Antes eram 3 varreduras idênticas em /machines/{id}/points (3x mais
#    chamadas). Agora: 1 chamada por máquina com IncludeLastMeasurement=true.
# ---------------------------------------------------------------------------

def _alarme(origem, mid, ponto):
    reg = {
        "origem": origem, "MachineId": mid, "ID": ponto.get("ID"),
        "HighAlarm": None, "HighWarning": None,
        "Freq_AlarmLevel": None, "Freq_WarningLevel": None,
    }
    summary = (ponto.get("OverallAlarm") or {}).get("Summary") or ""
    for parte in summary.lower().replace(" / ", "/").split("/"):
        valor = para_float(parte.split()[-1]) if parte.split() else None
        if "high alarm" in parte:
            reg["HighAlarm"] = valor
        elif "high warning" in parte:
            reg["HighWarning"] = valor

    overall = next(
        (f for f in ponto.get("Frequencies") or [] if f.get("Frequency") == "Overall"), {}
    )
    reg["Freq_AlarmLevel"] = para_float(overall.get("AlarmLevel", ""))
    reg["Freq_WarningLevel"] = para_float(overall.get("WarningLevel", ""))
    return reg


def etapa_points(df_machines):
    # chave (origem, id): IDs de máquina se repetem entre UBU e GERMANO
    mapa_path = {(o, i): pth for o, i, pth in zip(df_machines["origem"], df_machines["id"], df_machines["path"])}
    pontos, alarmes, medicoes = [], [], []

    for origem in BASE_URLS:
        ids = ids_por_origem(df_machines, "id", origem)
        respostas = em_paralelo(
            lambda mid: api_get(
                origem, f"/v1/machines/{mid}/points", {"IncludeLastMeasurement": "true"}
            ),
            ids,
        )
        for mid, lista in zip(ids, respostas):
            for p in lista or []:
                if not isinstance(p, dict):
                    continue
                pontos.append({
                    "origem": origem,
                    "path": mapa_path.get((origem, mid)),
                    "MachineID": mid,
                    "SubmachineID": p.get("ParentID"),
                    "ID": p.get("ID"),
                    "Name": p.get("Name"),
                    "Description": p.get("Description"),
                    "NodeTypeName": p.get("NodeTypeName"),
                    "EU": p.get("EU"),
                    "DetectionName": p.get("DetectionName"),
                })
                alarmes.append(_alarme(origem, mid, p))

                last = p.get("LastMeasurement") or {}
                for m in last.get("Measurements") or [None]:
                    m = m or {}
                    medicoes.append({
                        "ReadingTimeUTC": last.get("ReadingTimeUTC"),
                        "PointID": p.get("ID"),
                        "Speed": last.get("Speed"),
                        "SpeedUnits": last.get("SpeedUnits"),
                        "Direction": m.get("Direction"),
                        "ChannelName": m.get("ChannelName"),
                        "Level": m.get("Level"),
                        "Units": m.get("Units"),
                        "origem": origem,
                    })

    df_points = pd.DataFrame(pontos)
    log.info("%d pontos encontrados", len(df_points))
    if df_points.empty:
        raise RuntimeError("Nenhum ponto retornado")
    gravar_aba(PLANILHAS["points"], df_points, "points")

    df_alarms = pd.DataFrame(alarmes)
    if not df_alarms.empty:
        cols = ["HighAlarm", "HighWarning", "Freq_AlarmLevel", "Freq_WarningLevel"]
        df_alarms[cols] = df_alarms[cols].apply(pd.to_numeric, errors="coerce")
        df_alarms["ID"] = pd.to_numeric(df_alarms["ID"], errors="coerce").astype("Int64")
    gravar_aba(PLANILHAS["alarms"], df_alarms, "alarms")

    df_last = pd.DataFrame(medicoes)
    if not df_last.empty:
        df_last["ReadingTimeUTC"] = pd.to_datetime(df_last["ReadingTimeUTC"], errors="coerce")
        df_last = df_last[df_last["ReadingTimeUTC"].notna()]
    gravar_aba(PLANILHAS["lastmeasurements"], df_last, "lastmeasurements")

    return df_points


# ---------------------------------------------------------------------------
# 8. TRENDS (medições do dia de referência — padrão D-1) — carga incremental
# ---------------------------------------------------------------------------

def etapa_trends(df_points):
    dia = dia_referencia()
    inicio = datetime.combine(dia, dtime(0, 0, 0), tzinfo=timezone.utc)
    fim = datetime.combine(dia, dtime(23, 59, 59), tzinfo=timezone.utc)
    params = {
        "fromDateUTC": inicio.strftime("%Y-%m-%dT%H:%M:%SZ"),
        "toDateUTC": fim.strftime("%Y-%m-%dT%H:%M:%SZ"),
    }
    log.info("Trends de %s", dia)

    resultados = []
    for origem in BASE_URLS:
        ids = (
            pd.to_numeric(df_points.loc[df_points["origem"] == origem, "ID"], errors="coerce")
            .dropna().astype(int).unique().tolist()
        )
        respostas = em_paralelo(
            lambda pid: api_get(origem, f"/v1/points/{pid}/trendMeasurements", params), ids
        )
        for registros in respostas:
            for rec in registros or []:
                for m in rec.get("Measurements") or []:
                    resultados.append({
                        "ReadingTimeUTC": rec.get("ReadingTimeUTC"),
                        "PointID": rec.get("PointID"),
                        "ChannelName": m.get("ChannelName"),
                        "Direction": m.get("Direction"),
                        "Level": m.get("Level"),
                        "origem": origem,
                    })

    df = pd.DataFrame(resultados)
    if df.empty:
        log.warning("Nenhuma medição de trend em %s", dia)
        return df
    df = df[df["ChannelName"].isin(["Valor global", "Overall"]) & (df["Direction"] == "X")]
    df = df[["ReadingTimeUTC", "PointID", "Level", "origem"]].drop_duplicates()
    log.info("%d medições finais", len(df))

    anexar_novos(
        PLANILHAS["trends"], df,
        chaves=["ReadingTimeUTC", "PointID", "Level", "origem"],
        colunas_data=["ReadingTimeUTC"],
        nome="trends",
        limite_linhas=LIMITE_LINHAS["trends"],
    )
    return df


# ---------------------------------------------------------------------------
# 9. NOTES
# ---------------------------------------------------------------------------

def etapa_notes(df_points):
    registros = []
    for origem in BASE_URLS:
        for note in api_get(origem, "/v1/notes") or []:
            registros.append({
                "NoteID": note.get("idNote"),
                "PointID": note.get("idNode"),
                "NodeName": note.get("nodeName"),
                "CreatedAt": note.get("noteDateUTC"),
                "Title": note.get("title"),
                "Author": note.get("signature"),
                "Text": note.get("noteComment"),
                "Priority": note.get("priority"),
                "origem": origem,
            })
    df = pd.DataFrame(registros)
    if df.empty:
        log.warning("Nenhuma nota retornada")
        return df

    pontos = df_points[["ID", "path", "origem"]].drop_duplicates(["ID", "origem"])
    df = df.merge(pontos, left_on=["PointID", "origem"], right_on=["ID", "origem"], how="left")
    df = filtrar_site(df)
    df = df[[
        "NoteID", "path", "PointID", "NodeName", "Author", "CreatedAt",
        "Title", "Text", "Priority", "origem",
    ]].drop_duplicates()
    gravar_aba(PLANILHAS["notes"], df, "notes")
    return df


# ---------------------------------------------------------------------------
# MAIN
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    p = Pipeline("cda_online_skf")

    df_hier = p.etapa("hierarquia", etapa_hierarquia)
    p.etapa("submachines", lambda: etapa_submachines(df_hier), depende_de=["hierarquia"])

    df_machines = p.etapa("machines", etapa_machines)
    p.etapa("parts", lambda: etapa_parts(df_machines), depende_de=["machines"])
    df_points = p.etapa("points/alarms/lastmeasurements", lambda: etapa_points(df_machines), depende_de=["machines"])
    p.etapa("trends", lambda: etapa_trends(df_points), depende_de=["points/alarms/lastmeasurements"])
    p.etapa("notes", lambda: etapa_notes(df_points), depende_de=["points/alarms/lastmeasurements"])

    p.finalizar()
