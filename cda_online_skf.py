# -*- coding: utf-8 -*-
"""cda_online_skf.ipynb

# **1. BIBLIOTECAS**
"""

from datetime import datetime, timedelta, timezone
import pandas as pd
import requests
import urllib3
import json
import jwt
import sys
import os

import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe, set_with_dataframe

"""# **2. DADOS DE ACESSO**

## **2.1. Credenciais**
"""

CREDENTIALS = json.loads(os.environ["SKF_CREDENTIALS"])

"""## **2.2. URL's**

### **2.2.1. Base**
"""

BASE_URLS = json.loads(os.environ["SKF_BASE"])

"""### **2.2.2. Endpoint**"""

def build_urls(base_url):
    return {
        "token": f"{base_url}/token",
        "machines": f"{base_url}/v1/machines",
        "parts": f"{base_url}/v1/parts",
        "submachines": f"{base_url}/v1/hierarchy",
        "points": f"{base_url}/v1/points",
        "notes": f"{base_url}/v1/notes"
    }

"""## **2.3. Funções**

### **2.3.1. Token**
"""

def obter_token(base_url):
    url = f"{base_url}/token"

    response = requests.post(url, data=CREDENTIALS, verify=False)
    response.raise_for_status()

    return response.json()["access_token"]

"""### **2.3.2. Hierarquia**"""

def obter_arvore_hierarquia(token, base_url):
    url = f"{base_url}/v1/hierarchy"

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()

    return response.json()

"""## **2.4. Sheets**

### **2.4.1. Autenticação no Sheets**
"""

# Lê a variável de ambiente com o conteúdo do JSON da conta de serviço
service_account_info = json.loads(os.environ["GOOGLE_SERVICE_ACCOUNT"])

# Define os escopos de acesso (Google Sheets)
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

# Cria as credenciais usando o conteúdo do secret
creds = Credentials.from_service_account_info(service_account_info, scopes=SCOPES)

# Autentica no Google Sheets
gc = gspread.authorize(creds)

"""# **3. REQUISIÇÃO: MACHINES**

## **3.1. Execução**
"""

def get_machines(token, base_url, origem):
    url = f"{base_url}/v1/machines"

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()

    df = pd.DataFrame(response.json())
    df["origem"] = origem

    return df


def get_all_machines():
    tokens = {
        origem: obter_token(base_url)
        for origem, base_url in BASE_URLS.items()
    }

    return pd.concat(
        [
            get_machines(tokens[origem], base_url, origem)
            for origem, base_url in BASE_URLS.items()
        ],
        ignore_index=True
    )


df_machines = get_all_machines()

"""## **3.2. Estrutura e organização**"""

# Garantir string
df_machines["path"] = df_machines["path"].fillna("").astype(str)

# Split
df_path = df_machines["path"].str.split(r"\\", expand=True)
df_path.columns = [f"path_{i}" for i in range(df_path.shape[1])]

# Estrutura
df_machines = pd.concat(
    [df_machines[["origem", "id", "name", "path"]], df_path],
    axis=1
)

# Normalizar
df_machines["path_1"] = df_machines["path_1"].str.strip().str.upper()

# Filtro
mapa = {"ubu": "UBU", "germano": "GERMANO"}
df_machines = df_machines[
    df_machines["path_1"] == df_machines["origem"].map(mapa)
]

"""## **3.3. DataFrame**"""

# Selecionar colunas
df_machine = df_machines[['origem', 'id', 'name', 'path']].copy()

"""## **3.4. Lista de ID's**"""

# Extrair lista de IDs
machine_ids = df_machine['id'].tolist()

# Exibir
print(f"{len(machine_ids)} ativos encontrados")

"""## **3.5. Carga no Sheets**"""

# Nome da planilha
planilha_id = "1gHImL_Hbr6teYb-sV4hzfoHExVFmTwlrjx93jgAsji0"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Limpa a aba antes de escrever os dados
aba.clear()

# Envia o DataFrame para a aba
set_with_dataframe(aba, df_machine)

print("Dados enviados com sucesso para o Google Sheets!")

"""# **4. REQUISIÇÃO: SUBMACHINES**

## **4.1. Execução**
"""

def get_submachines(data, parent_name=None, submachines=None):
    if submachines is None:
        submachines = []

    for node in data:
        nome = node.get("name")

        if node.get("typeName") == "SubMachine":
            submachines.append({
                "id": node.get("id"),
                "name": nome,
                "description": node.get("description"),
                "parent": parent_name,
                "path": node.get("path"),
                "active": node.get("active"),
                "status": node.get("status")
            })

        if node.get("children"):
            get_submachines(node["children"], parent_name=nome, submachines=submachines)

    return submachines

dfs = []

for origem, base_url in BASE_URLS.items():
    token = obter_token(base_url)

    data_raw = obter_arvore_hierarquia(token, base_url)
    submachines = get_submachines(data_raw)

    df = pd.DataFrame(submachines)
    df["origem"] = origem

    dfs.append(df)

df_submachines = pd.concat(dfs, ignore_index=True)

"""## **4.2. Estrutura e organização**"""

# Garantir string
df_submachines["path"] = df_submachines["path"].fillna("").astype(str)

# Split
df_path = df_submachines["path"].str.split(r"\\", expand=True)
df_path.columns = [f"path_{i}" for i in range(df_path.shape[1])]

# Estrutura
df_submachines = pd.concat(
    [df_submachines[['origem', 'id', 'name', 'description', 'parent', 'path', 'active', 'status']], df_path],
    axis=1
)

# Normalizar
df_submachines["path_1"] = df_submachines["path_1"].str.strip().str.upper()

# Filtro
mapa = {"ubu": "UBU", "germano": "GERMANO"}
df_submachines = df_submachines[
    df_submachines["path_1"] == df_submachines["origem"].map(mapa)
]

"""## **4.3. DataFrame**"""

df_submachines = df_submachines[['origem', 'id', 'name', 'description', 'parent', 'path', 'status', 'active']]

"""## **4.4. Carga no Sheets**"""

# Nome da planilha
planilha_id = "115Ilr5gw8jkqVaHJmotdo5gVxwfLpx07tCTdMwQLLzc"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Limpa a aba antes de escrever os dados
aba.clear()

# Envia o DataFrame para a aba
set_with_dataframe(aba, df_submachines)

print("Dados enviados com sucesso para o Google Sheets!")

"""# **5. REQUISIÇÃO: MACHINE PART**

## **5.1. Execução**
"""

def get_machine_parts(token, base_url, machine_ids, origem):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    registros = []

    for mid in machine_ids:
        url = f"{base_url}/v1/machines/{mid}/parts"
        resp = requests.get(url, headers=headers, verify=False)

        if resp.status_code == 200:
            data = resp.json()

            for part in data:
                fault_list = part.get("faultFrequencies", [])

                if not fault_list:
                    registros.append({
                        "MachineId": mid,
                        "PartID": part.get("id"),
                        "PartName": part.get("name"),
                        "Type": part.get("type"),
                        "Ratio": part.get("ratio"),
                        "Brand": part.get("brand"),
                        "Typeno": part.get("typeno"),
                        "SpeedPointId": part.get("speedPointId"),
                        "Name": None,
                        "Multiple": None,
                        "origem": origem
                    })
                else:
                    for f in fault_list:
                        registros.append({
                            "MachineId": mid,
                            "PartID": part.get("id"),
                            "PartName": part.get("name"),
                            "Type": part.get("type"),
                            "Ratio": part.get("ratio"),
                            "Brand": part.get("brand"),
                            "Typeno": part.get("typeno"),
                            "SpeedPointId": part.get("speedPointId"),
                            "Name": f.get("name"),
                            "Multiple": f.get("multiple"),
                            "origem": origem
                        })

    return pd.DataFrame(registros)

dfs = []

for origem, base_url in BASE_URLS.items():
    token = obter_token(base_url)

    machine_ids_origem = df_machines[
        df_machines["origem"] == origem
    ]["id"].tolist()

    df = get_machine_parts(token, base_url, machine_ids_origem, origem)
    dfs.append(df)

df_parts = pd.concat(dfs, ignore_index=True)

"""## **5.2. Estrutura e organização**"""

# merge
df_parts = df_parts.merge(
    df_machines[["id", "path"]],
    left_on="MachineId",
    right_on="id",
    how="left"
)

# remover coluna duplicada do merge
df_parts = df_parts.drop(columns=["id"], errors="ignore")

# garantir string
df_parts["path"] = df_parts["path"].fillna("").astype(str)

# split
df_path = df_parts["path"].str.split(r"\\", expand=True)
df_path.columns = [f"path_{i}" for i in range(df_path.shape[1])]

df_parts = pd.concat([df_parts, df_path], axis=1)

# garantir path_1 válido
df_parts["path_1"] = df_parts.get("path_1", "").astype(str).str.strip().str.upper()

# filtro
mapa = {"ubu": "UBU", "germano": "GERMANO"}

df_parts = df_parts[
    df_parts["path_1"] == df_parts["origem"].map(mapa)
]

"""## **5.3. DataFrame**"""

colunas = [
    "origem", "path", "MachineId", "PartID", "PartName", "Type", "Ratio",
    "Brand", "Typeno", "SpeedPointId",
    "Name", "Multiple"
]

df_parts = df_parts[colunas].drop_duplicates()

"""## **5.4. Carga no Sheets**"""

# Nome da planilha
planilha_id = "1ZOW0VeSOqSjCNZVvO48TxnAt6FdKhBSuCwFLq23Rd7k"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Limpa a aba antes de escrever os dados (opcional)
aba.clear()

# Envia o DataFrame para a aba
set_with_dataframe(aba, df_parts)

print("Dados enviados com sucesso para o Google Sheets!")

"""# **6. REQUISIÇÃO: POINTS**

## **6.1. Execução**
"""

def get_points(token, base_url, machine_ids, origem):
    headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
    registros = []

    mapa_path = df_machines.set_index("id")["path"].to_dict()

    for mid in machine_ids:
        url = f"{base_url}/v1/machines/{mid}/points"
        resp = requests.get(url, headers=headers, verify=False)

        if resp.status_code != 200:
            continue

        registros.append({
            "path": mapa_path.get(mid),
            "MachineId": mid,
            "origem": origem,
            "data": resp.json()
        })

    return pd.DataFrame(registros)

dfs = []

for origem, base_url in BASE_URLS.items():
    token = obter_token(base_url)

    machine_ids_origem = df_machines[
        df_machines["origem"] == origem
    ]["id"].tolist()

    df_temp = get_points(token, base_url, machine_ids_origem, origem)
    dfs.append(df_temp)

df_points = pd.concat(dfs, ignore_index=True)

"""## **6.2. Estrutura e organização**"""

def explode_points(df_raw):
    registros = []

    for _, row in df_raw.iterrows():
        for p in row["data"]:
            registros.append({
                "origem": row["origem"],
                "path": row["path"],
                "MachineID": row["MachineId"],
                "SubmachineID": p.get("ParentID"),
                "ID": p.get("ID"),
                "Name": p.get("Name"),
                "Description": p.get("Description"),
                "NodeTypeName": p.get("NodeTypeName"),
                "EU": p.get("EU"),
                "DetectionName": p.get("DetectionName")
            })

    return pd.DataFrame(registros)

df_points = explode_points(df_points)

"""## **6.4. Lista**"""

# Extrai lista com os IDs dos ativos filtrados
point_ids = df_points['ID'].tolist()

# Exibe a lista
print(f"{len(point_ids)} pontos encontrados")

"""## **6.5. Carga no Sheets**"""

# Nome da planilha
planilha_id = "1ax7LyNRt06naxNwPXXb9kOGLUkmSbZFw9-IjhV4-_Gk"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Limpa a aba antes de escrever os dados (opcional)
aba.clear()

# Envia o DataFrame para a aba
set_with_dataframe(aba, df_points)

print("Dados enviados com sucesso para o Google Sheets!")

"""# **7. REQUISIÇÃO: ALARMS**

## **7.1. Execução**
"""

def get_alarms(token, base_url, machine_ids, origem):

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    registros = []

    for mid in machine_ids:
        url = f"{base_url}/v1/machines/{mid}/points"
        resp = requests.get(url, headers=headers, verify=False)

        if resp.status_code != 200:
            continue

        for ponto in resp.json():

            if not isinstance(ponto, dict):
                continue

            registro = {
                "origem": origem,
                "MachineId": mid,
                "ID": ponto.get("ID"),
                "HighAlarm": None,
                "HighWarning": None,
                "Freq_AlarmLevel": None,
                "Freq_WarningLevel": None
            }

            overall_alarm = ponto.get("OverallAlarm") or {}
            summary = overall_alarm.get("Summary", "")

            if summary:
                parts = summary.lower().replace(" / ", "/").split("/")

                for part in parts:
                    try:
                        val = float(part.split()[-1].replace(",", "."))
                    except:
                        val = None

                    if "high alarm" in part:
                        registro["HighAlarm"] = val
                    elif "high warning" in part:
                        registro["HighWarning"] = val

            freq_list = ponto.get("Frequencies", [])
            freq_overall = next(
                (f for f in freq_list if f.get("Frequency") == "Overall"),
                {}
            )

            try:
                registro["Freq_AlarmLevel"] = float(
                    str(freq_overall.get("AlarmLevel", "")).split()[0].replace(",", ".")
                )
            except:
                pass

            try:
                registro["Freq_WarningLevel"] = float(
                    str(freq_overall.get("WarningLevel", "")).split()[0].replace(",", ".")
                )
            except:
                pass

            registros.append(registro)

    df = pd.DataFrame(registros)

    cols = ["HighAlarm", "HighWarning", "Freq_AlarmLevel", "Freq_WarningLevel"]
    df[cols] = df[cols].apply(pd.to_numeric, errors="coerce")

    df["ID"] = pd.to_numeric(df["ID"], errors="coerce").astype("Int64")

    return df

dfs = []

for origem, base_url in BASE_URLS.items():
    token = obter_token(base_url)

    machine_ids_origem = df_machines[
        df_machines["origem"] == origem
    ]["id"].tolist()

    df_temp = get_alarms(token, base_url, machine_ids_origem, origem)
    dfs.append(df_temp)

df_alarms = pd.concat(dfs, ignore_index=True)

"""## **7.2. DataFrame**"""

print(f"Total de registros: {len(df_alarms)}")

"""## **7.3. Carga de Sheets**"""

# Nome da planilha
planilha_id = "172uMRb6-j8yitUVbCYp-iHWcSQGYpi6oqSX1gan2DXU"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Limpa a aba antes de escrever os dados (opcional)
aba.clear()

# Envia o DataFrame para a aba
set_with_dataframe(aba, df_alarms)

print("Dados enviados com sucesso para o Google Sheets!")

"""# **8. REQUISIÇÃO: LAST MEASUREMENTS**

## **8.1. Execução**
"""

def get_measurements(token, base_url, machine_ids, origem):

    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    registros = []

    for mid in machine_ids:
        url = f"{base_url}/v1/machines/{mid}/points?IncludeLastMeasurement=true"
        resp = requests.get(url, headers=headers, verify=False)

        if resp.status_code != 200:
            continue

        for ponto in resp.json():

            last = ponto.get("LastMeasurement") or {}

            measurements = last.get("Measurements") or [None]

            for m in measurements:
                registros.append({
                    "ReadingTimeUTC": last.get("ReadingTimeUTC"),
                    "PointID": ponto.get("ID"),
                    "Speed": last.get("Speed"),
                    "SpeedUnits": last.get("SpeedUnits"),
                    "Direction": m.get("Direction") if m else None,
                    "ChannelName": m.get("ChannelName") if m else None,
                    "Level": m.get("Level") if m else None,
                    "Units": m.get("Units") if m else None,
                    "origem": origem
                })

    return pd.DataFrame(registros)

dfs = []

for origem, base_url in BASE_URLS.items():
    token = obter_token(base_url)

    machine_ids_origem = df_machines[
        df_machines["origem"] == origem
    ]["id"].tolist()

    df_temp = get_measurements(token, base_url, machine_ids_origem, origem)
    dfs.append(df_temp)

df_lastmeasurements = pd.concat(dfs, ignore_index=True)

"""## **8.2. DataFrame**"""

df_lastmeasurements["ReadingTimeUTC"] = pd.to_datetime(
    df_lastmeasurements["ReadingTimeUTC"],
    errors="coerce"
)

df_lastmeasurements = df_lastmeasurements[
    df_lastmeasurements["ReadingTimeUTC"].notna()
]

"""## **8.3. Carga de Sheets**"""

# Nome da planilha
planilha_id = "1GmfGrRxJsOUAHG8yJ6Wfn67SSAhR5dgPQEZXplAFnJI"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Limpa a aba antes de escrever os dados (opcional)
aba.clear()

# Envia o DataFrame para a aba
set_with_dataframe(aba, df_lastmeasurements)

print("Dados enviados com sucesso para o Google Sheets!")

"""# **9. REQUISIÇÃO: MEASUREMENTS**

## **9.1. Execução**
"""

def consultar_trends(point_ids, token, base_url, origem):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    now_utc = datetime.now(timezone.utc)

    ontem_inicio = (now_utc - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
    ontem_fim = (now_utc - timedelta(days=1)).replace(hour=23, minute=59, second=59, microsecond=0)

    params = {
        "fromDateUTC": ontem_inicio.isoformat().replace("+00:00", "Z"),
        "toDateUTC": ontem_fim.isoformat().replace("+00:00", "Z")
    }

    resultados = []

    for pid in point_ids:
        url = f"{base_url}/v1/points/{int(pid)}/trendMeasurements"
        response = requests.get(url, headers=headers, params=params, verify=False)

        if response.status_code == 200:
            data = response.json()

            for record in data:
                base = {
                    "ReadingTimeUTC": record.get("ReadingTimeUTC"),
                    "PointID": record.get("PointID"),
                    "origem": origem
                }

                for m in record.get("Measurements", []):
                    resultados.append({
                        **base,
                        "ChannelName": m.get("ChannelName"),
                        "Direction": m.get("Direction"),
                        "Level": m.get("Level"),
                        "Units": m.get("Units")
                    })

    return pd.DataFrame(resultados)

dfs = []

for origem, base_url in BASE_URLS.items():
    token = obter_token(base_url)

    point_ids_origem = df_points[
        df_points["origem"] == origem
    ]["ID"].dropna().astype(int).tolist()

    df = consultar_trends(point_ids_origem, token, base_url, origem)
    dfs.append(df)

df_trends = pd.concat(dfs, ignore_index=True)

"""## **9.2. DataFrame**"""

df_trends = df_trends[
    (df_trends['ChannelName'].isin(['Valor global', 'Overall'])) &
    (df_trends['Direction'] == "X")
]

colunas = ['ReadingTimeUTC', 'PointID', 'Level', 'origem']

df_trendMeasurements = df_trends[colunas].drop_duplicates()

print(f"{len(df_trendMeasurements)} medições finais")

"""## **9.3. Carga de Sheets**"""

# Nome da planilha e aba
planilha_id = "1uekxFKio9llwhP9CljTappvO3XaXrrV7g_WT2V3buIo"
nome_da_aba = "Sheet1"

# Abre a planilha e aba
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Lê os dados atuais da aba
df_existente = get_as_dataframe(aba, evaluate_formulas=True).dropna(how="all")

# Colunas obrigatórias
colunas_chave = ['ReadingTimeUTC', 'PointID', 'Level']

# Se a aba está vazia ou não tem as colunas necessárias → cria cabeçalho
if df_existente.empty or not all(col in df_existente.columns for col in colunas_chave):
    aba.clear()
    set_with_dataframe(aba, pd.DataFrame(columns=colunas_chave), row=1, col=1, include_column_header=True)
    df_existente = pd.DataFrame(columns=colunas_chave)

# Garante que colunas estão no mesmo formato e ordem
df_existente = df_existente[colunas_chave].dropna()

# Remove duplicados e encontra apenas as linhas novas
df_novos = df_trendMeasurements[~df_trendMeasurements.isin(df_existente.to_dict(orient='list')).all(axis=1)]

# Se houver novos registros, adiciona abaixo
if not df_novos.empty:
    # Número de linhas já existentes (para inserir a partir da próxima linha vazia)
    ultima_linha = len(df_existente) + 2  # +1 para header, +1 para próxima
    set_with_dataframe(aba, df_novos, row=ultima_linha, col=1, include_column_header=False)
    print(f"{len(df_novos)} novas medições adicionadas à planilha!")
else:
    print("Nenhuma medição nova para inserir — tudo já está na planilha.")

"""## **9.4. Tratamento de duplicatas**"""

# Nome da planilha e aba
planilha_id = "1UK6AatDxCdqg8NZxgL8ZThyCSNXOh03_AY4jtGaumXc"
nome_da_aba = "Sheet1"

# Abre a planilha e aba
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Lê os dados da aba
df = get_as_dataframe(aba, evaluate_formulas=True).dropna(how="all")

# Remove duplicatas com base nas colunas chave
colunas_chave = ['ReadingTimeUTC', 'PointID', 'Level']
df_limpo = df.drop_duplicates(subset=colunas_chave, keep='first')

# Limpa aba (opcional, mas garante que não fica lixo antigo abaixo)
aba.clear()

# Reescreve os dados limpos na planilha (com cabeçalho)
set_with_dataframe(aba, df_limpo, include_column_header=True)

print(f"Removidas {len(df) - len(df_limpo)} duplicatas. Planilha atualizada com {len(df_limpo)} registros únicos.")

"""# **10. REQUISIÇÃO: NOTES**

## **10.1. Execução**
"""

def get_notes(token, base_url, origem):
    headers = {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json"
    }

    url = f"{base_url}/v1/notes"
    response = requests.get(url, headers=headers, verify=False)
    response.raise_for_status()

    data = response.json()

    registros = []
    for note in data:
        registros.append({
            "NoteID": note.get("idNote"),
            "PointID": note.get("idNode"),
            "NodeName": note.get("nodeName"),
            "CreatedAt": note.get("noteDateUTC"),
            "Title": note.get("title"),
            "Author": note.get("signature"),
            "Text": note.get("noteComment"),
            "Priority": note.get("priority"),
            "origem": origem
        })

    return pd.DataFrame(registros)

dfs = []

for origem, base_url in BASE_URLS.items():
    token = obter_token(base_url)

    df = get_notes(token, base_url, origem)
    dfs.append(df)

df_notes = pd.concat(dfs, ignore_index=True)

"""## **10.2. Estrutura e organização**"""

df_notes = df_notes.merge(
    df_points[["ID", "path"]],
    left_on="PointID",
    right_on="ID",
    how="left"
)

# Garantir string
df_notes["path"] = df_notes["path"].fillna("").astype(str)

# Split
df_path = df_notes["path"].str.split(r"\\", expand=True)
df_path.columns = [f"path_{i}" for i in range(df_path.shape[1])]

df_notes = pd.concat([df_notes, df_path], axis=1)

# Normalizar
df_notes["path_1"] = df_notes["path_1"].str.strip().str.upper()

# Filtro
mapa = {"ubu": "UBU", "germano": "GERMANO"}

df_notes = df_notes[
    df_notes["path_1"] == df_notes["origem"].map(mapa)
]

"""## **10.3. DataFrame**"""

colunas = [
    "NoteID", "path", "PointID", "NodeName", "Author",
    "CreatedAt", "Title", "Text",
    "Priority", "origem"
]

df_notes = df_notes[colunas].drop_duplicates()

"""## **10.4. Carga no Sheets**"""

# Nome da planilha
planilha_id = "1A70P76NH1Lxt3h-Hg0V0NJBaPRHufbjC80m_qtkxfac"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open_by_key(planilha_id)
aba = planilha.worksheet(nome_da_aba)

# Limpa a aba antes de escrever os dados (opcional)
aba.clear()

# Envia o DataFrame para a aba
set_with_dataframe(aba, df_notes)

print("Dados enviados com sucesso para o Google Sheets!")
