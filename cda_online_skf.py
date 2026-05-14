# -*- coding: utf-8 -*-
"""cda_online_skf.ipynb

# **1. BIBLIOTECAS**
"""

from datetime import datetime, timedelta, timezone
from IPython.display import display
import pandas as pd
import requests
import urllib3
import json
import jwt
import sys

import gspread
from google.oauth2.service_account import Credentials
from gspread_dataframe import get_as_dataframe, set_with_dataframe

"""# **2. DADOS DE ACESSO**

## **2.1. Credenciais**
"""

CREDENTIALS = {
    "username": os.getenv("SKF_USERNAME"),
    "password": os.getenv("SKF_PASSWORD"),
    "grant_type": "password"
}

"""## **2.2. URL's**

### **2.2.1. Base**
"""

BASE_URLS = {
    "ubu": os.getenv("BASE_UBU"),
    "germano": os.getenv("BASE_GERMANO")
}

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

# Exibir
display(df_machines.head(100))

"""## **3.3. DataFrame**"""

# Selecionar colunas
df_machine = df_machines[['origem', 'id', 'name', 'path']].copy()

# Visualizar
display(df_machine.head(5))

# download excel
df_machine.to_excel('cda_online_skf_machines.xlsx', index=False)

"""## **3.4. Lista de ID's**"""

# Extrair lista de IDs
machine_ids = df_machine['id'].tolist()

# Exibir
print(f"{len(machine_ids)} ativos encontrados")
print(machine_ids)

"""## **3.5. Carga no Sheets**"""

# Nome da planilha
nome_da_planilha = "cda_online_skf_machines"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open(nome_da_planilha)
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

# Exibir
display(df_submachines.head(1000))

"""## **4.3. DataFrame**"""

df_submachines = df_submachines[['origem', 'id', 'name', 'description', 'parent', 'path', 'status', 'active']]

display(df_submachines.head(5))

# download excel
df_submachines.to_excel('cda_online_skf_submachines.xlsx', index=False)

"""## **4.4. Carga no Sheets**"""

# Nome da planilha
nome_da_planilha = "cda_online_skf_submachines"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open(nome_da_planilha)
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

df_parts.head(10000)

# download excel
df_parts.to_excel('cda_online_skf_parts.xlsx', index=False)

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

# exibir
df_parts.head(10)

"""## **5.3. DataFrame**"""

colunas = [
    "origem", "path", "MachineId", "PartID", "PartName", "Type", "Ratio",
    "Brand", "Typeno", "SpeedPointId",
    "Name", "Multiple"
]

df_parts = df_parts[colunas].drop_duplicates()

display(df_parts.head(5))

# download excel
df_parts.to_excel('cda_online_skf_machineparts.xlsx', index=False)

"""## **5.4. Carga no Sheets**"""

# Nome da planilha
nome_da_planilha = "cda_online_skf_machineparts"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open(nome_da_planilha)
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

"""## **6.3. DataFrame**"""

display(df_points.head(5))

# download excel
df_points.to_excel('cda_online_skf_points.xlsx', index=False)

"""## **6.4. Lista**"""

# Extrai lista com os IDs dos ativos filtrados
point_ids = df_points['ID'].tolist()

# Exibe a lista
print(f"{len(point_ids)} pontos encontrados")
print(point_ids)

"""## **6.5. Carga no Sheets**"""

# Nome da planilha
nome_da_planilha = "cda_online_skf_points"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open(nome_da_planilha)
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

display(df_alarms.head(5))

# download excel
df_alarms.to_excel('cda_online_skf_alarms.xlsx', index=False)

"""## **7.3. Carga de Sheets**"""

# Nome da planilha
nome_da_planilha = "cda_online_skf_alarms"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open(nome_da_planilha)
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

df_lastmeasurements.head(5)

# download excel
df_lastmeasurements.to_excel('cda_online_skf_lastmeasurements.xlsx', index=False)

"""## **8.3. Carga de Sheets**"""

# Nome da planilha
nome_da_planilha = "cda_online_skf_lastmeasurements"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open(nome_da_planilha)
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
display(df_trendMeasurements.head(5))

# download excel
df_trendMeasurements.to_excel('cda_online_skf_measurements.xlsx', index=False)

"""## **9.3. Carga de Sheets**"""

# Nome da planilha e aba
nome_da_planilha = "cda_online_skf_measurements"
nome_da_aba = "Sheet1"

# Abre a planilha e aba
planilha = gc.open(nome_da_planilha)
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
nome_da_planilha = "cda_online_skf_measurements"
nome_da_aba = "Sheet1"

# Abre a planilha e aba
planilha = gc.open(nome_da_planilha)
aba = planilha.worksheet(nome_da_aba)

# Lê os dados da aba
df = get_as_dataframe(aba, evaluate_formulas=True).dropna(how="all")

# Remove duplicatas com base nas colunas chave
colunas_chave = ['ReadingTimeUTC', 'PointID', 'Level', 'Units']
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

df_notes.head(10)

"""## **10.3. DataFrame**"""

colunas = [
    "NoteID", "path", "PointID", "NodeName", "Author",
    "CreatedAt", "Title", "Text",
    "Priority", "origem"
]

df_notes = df_notes[colunas].drop_duplicates()

display(df_notes.head(5))

# download excel
df_notes.to_excel('cda_online_skf_notes.xlsx', index=False)

"""## **10.4. Carga no Sheets**"""

# Nome da planilha
nome_da_planilha = "cda_online_skf_notes"
nome_da_aba = "Sheet1"

# Abre a planilha
planilha = gc.open(nome_da_planilha)
aba = planilha.worksheet(nome_da_aba)

# Limpa a aba antes de escrever os dados (opcional)
aba.clear()

# Envia o DataFrame para a aba
set_with_dataframe(aba, df_notes)

print("Dados enviados com sucesso para o Google Sheets!")

"""# **11. PSEUDOCÓDIGO**

# CDA Online SKF

---

## 1. CONFIGURAÇÕES INICIAIS

```
DEFINIR CREDENTIALS = { username, password, grant_type }

DEFINIR BASE_URLS = {
    "ubu"     : "http://services.repcenter.skf.com:22011",
    "germano" : "http://services.repcenter.skf.com:20446"
}

FUNÇÃO build_urls(base_url):
    RETORNAR dicionário com endpoints:
        token, machines, parts, submachines, points, notes

AUTENTICAR no Google Sheets via OAuth (gspread + google.auth)
```

---

## 2. FUNÇÕES UTILITÁRIAS

### 2.1. obter_token(base_url)
```
FUNÇÃO obter_token(base_url):
    POST para base_url/token com CREDENTIALS
    SE erro: lançar exceção
    RETORNAR access_token da resposta
```

### 2.2. obter_arvore_hierarquia(token, base_url)
```
FUNÇÃO obter_arvore_hierarquia(token, base_url):
    GET para base_url/v1/hierarchy com Bearer token
    SE erro: lançar exceção
    RETORNAR JSON com a árvore hierárquica completa
```

---

## 3. MACHINES

### Busca e construção
```
FUNÇÃO get_machines(token, base_url, origem):
    GET para base_url/v1/machines com Bearer token
    df = DataFrame(resposta.json())
    adicionar coluna "origem"
    RETORNAR df

FUNÇÃO get_all_machines():
    PARA CADA (origem, base_url) EM BASE_URLS:
        token = obter_token(base_url)
        executar get_machines(token, base_url, origem)
    RETORNAR concatenação de todos os DataFrames
```

### Estrutura e organização
```
df_machines = get_all_machines()

converter coluna "path" para string
DIVIDIR "path" por "\" → colunas path_0, path_1, path_2, ...
normalizar path_1: remover espaços, converter para MAIÚSCULO

FILTRAR df_machines:
    manter apenas linhas onde path_1 == mapa[origem]
    (ex: "ubu" → "UBU", "germano" → "GERMANO")

df_machine = selecionar colunas [origem, id, name, path]
machine_ids = lista de IDs das machines filtradas

SALVAR Excel "cda_online_skf_machines.xlsx"
ENVIAR df_machine para Google Sheets "cda_online_skf_machines" (limpar + escrever)
```

---

## 4. SUBMACHINES

### Busca recursiva na hierarquia
```
FUNÇÃO get_submachines(data, parent_name, submachines=[]):
    PARA CADA nó EM data:
        SE nó.typeName == "SubMachine":
            ADICIONAR ao submachines:
                { id, name, description, parent, path, active, status }

        SE nó possui filhos (children):
            chamar get_submachines(nó.children, parent_name=nó.name, submachines)

    RETORNAR submachines
```

### Execução por origem
```
PARA CADA (origem, base_url) EM BASE_URLS:
    token = obter_token(base_url)
    data_raw = obter_arvore_hierarquia(token, base_url)
    submachines = get_submachines(data_raw)
    df = DataFrame(submachines) + coluna "origem"
    acumular em dfs

df_submachines = concatenar todos os dfs

converter "path" para string
DIVIDIR "path" por "\" → colunas path_0, path_1, ...
normalizar path_1
FILTRAR por origem (mesmo critério de machines)

selecionar colunas [origem, id, name, description, parent, path, status, active]

SALVAR Excel "cda_online_skf_submachines.xlsx"
ENVIAR para Google Sheets "cda_online_skf_submachines"
```

---

## 5. MACHINE PARTS

### Busca por machine ID
```
FUNÇÃO get_machine_parts(token, base_url, machine_ids, origem):
    PARA CADA machine_id EM machine_ids:
        GET para base_url/v1/machines/{machine_id}/parts

        SE status == 200:
            PARA CADA part na resposta:
                SE part NÃO tem faultFrequencies:
                    adicionar registro com campos da peça + Name=None, Multiple=None
                SENÃO:
                    PARA CADA frequência de falha:
                        adicionar registro com campos da peça + Name e Multiple da frequência

    RETORNAR DataFrame de registros
```

### Execução por origem
```
PARA CADA (origem, base_url) EM BASE_URLS:
    token = obter_token(base_url)
    machine_ids_origem = IDs de df_machines filtrados por origem
    df = get_machine_parts(token, base_url, machine_ids_origem, origem)
    acumular em dfs

df_parts = concatenar todos os dfs

JUNTAR df_parts com df_machines pelo MachineId → trazer coluna "path"
DIVIDIR "path" por "\" → colunas path_0, path_1, ...
normalizar e FILTRAR por origem

selecionar colunas [origem, path, MachineId, PartID, PartName, Type,
                    Ratio, Brand, Typeno, SpeedPointId, Name, Multiple]
remover duplicatas

SALVAR Excel "cda_online_skf_machineparts.xlsx"
ENVIAR para Google Sheets "cda_online_skf_machineparts"
```

---

## 6. POINTS

### Busca por machine ID
```
FUNÇÃO get_points(token, base_url, machine_ids, origem):
    mapa_path = dicionário { machine_id: path } a partir de df_machines

    PARA CADA machine_id EM machine_ids:
        GET para base_url/v1/machines/{machine_id}/points

        SE status == 200:
            adicionar registro: { path, MachineId, origem, data: resposta.json() }

    RETORNAR DataFrame com uma linha por machine (data ainda aninhada)
```

### Expansão dos pontos (explode)
```
FUNÇÃO explode_points(df_raw):
    PARA CADA linha do df_raw:
        PARA CADA ponto em linha.data:
            extrair: { origem, path, MachineID, SubmachineID, ID,
                       Name, Description, NodeTypeName, EU, DetectionName }
    RETORNAR DataFrame flat

df_points = explode_points(df_points)
point_ids = lista de IDs dos pontos

SALVAR Excel "cda_online_skf_points.xlsx"
ENVIAR para Google Sheets "cda_online_skf_points"
```

---

## 7. ALARMS

### Busca por machine ID (extrai alarmes dos pontos)
```
FUNÇÃO get_alarms(token, base_url, machine_ids, origem):
    PARA CADA machine_id EM machine_ids:
        GET para base_url/v1/machines/{machine_id}/points

        SE status != 200: pular

        PARA CADA ponto na resposta:
            registro = { origem, MachineId, ID, HighAlarm=None,
                         HighWarning=None, Freq_AlarmLevel=None, Freq_WarningLevel=None }

            // Extrair OverallAlarm
            summary = ponto.OverallAlarm.Summary
            SE summary não vazio:
                DIVIDIR summary por "/"
                PARA CADA parte:
                    SE "high alarm" → extrair valor → HighAlarm
                    SE "high warning" → extrair valor → HighWarning

            // Extrair frequência Overall
            freq_overall = primeiro item de Frequencies onde Frequency == "Overall"
            extrair Freq_AlarmLevel e Freq_WarningLevel como float

            adicionar registro

    converter colunas numéricas
    RETORNAR DataFrame

PARA CADA (origem, base_url) EM BASE_URLS:
    token = obter_token(base_url)
    machine_ids_origem = IDs filtrados por origem
    df_temp = get_alarms(...)
    acumular

df_alarms = concatenar

SALVAR Excel "cda_online_skf_alarms.xlsx"
ENVIAR para Google Sheets "cda_online_skf_alarms"
```

---

## 8. LAST MEASUREMENTS

### Busca (inclui última medição por ponto)
```
FUNÇÃO get_measurements(token, base_url, machine_ids, origem):
    PARA CADA machine_id EM machine_ids:
        GET para base_url/v1/machines/{machine_id}/points?IncludeLastMeasurement=true

        SE status != 200: pular

        PARA CADA ponto na resposta:
            last = ponto.LastMeasurement
            measurements = last.Measurements (ou [None] se vazio)

            PARA CADA medição em measurements:
                adicionar registro:
                    { ReadingTimeUTC, PointID, Speed, SpeedUnits,
                      Direction, ChannelName, Level, Units, origem }

    RETORNAR DataFrame

PARA CADA (origem, base_url) EM BASE_URLS:
    token = obter_token(base_url)
    machine_ids_origem = IDs filtrados por origem
    df_temp = get_measurements(...)
    acumular

df_lastmeasurements = concatenar

converter ReadingTimeUTC para datetime
FILTRAR: remover linhas com ReadingTimeUTC inválido (NaT)

SALVAR Excel "cda_online_skf_lastmeasurements.xlsx"
ENVIAR para Google Sheets "cda_online_skf_lastmeasurements"
```

---

## 9. MEASUREMENTS (Trend — por período D-1)

### Busca por point ID
```
FUNÇÃO consultar_trends(point_ids, token, base_url, origem):
    start = ontem 00:00:00 UTC
    end   = ontem 23:59:59 UTC
    params = { fromDateUTC: start, toDateUTC: end }

    PARA CADA point_id EM point_ids:
        GET para base_url/v1/points/{point_id}/trendMeasurements com params

        SE status == 200:
            PARA CADA record na resposta:
                base = { ReadingTimeUTC, PointID, origem }
                PARA CADA medição em record.Measurements:
                    adicionar { base + ChannelName, Direction, Level, Units }

    RETORNAR DataFrame

PARA CADA (origem, base_url) EM BASE_URLS:
    token = obter_token(base_url)
    point_ids_origem = IDs de df_points filtrados por origem
    df = consultar_trends(point_ids_origem, ...)
    acumular

df_trends = concatenar
```

### Filtro e carga incremental
```
FILTRAR df_trends:
    manter apenas linhas onde ChannelName EM ['Valor global', 'Overall']
    E Direction == "X"

df_trendMeasurements = selecionar [ReadingTimeUTC, PointID, Level, origem]
                        remover duplicatas

SALVAR Excel "cda_online_skf_measurements.xlsx"

// Carga incremental no Sheets (sem duplicar):
df_existente = ler dados atuais do Sheets
SE aba vazia ou sem colunas-chave:
    inicializar aba com cabeçalho
df_novos = df_trendMeasurements - df_existente (apenas registros novos)
SE df_novos não vazio:
    INSERIR df_novos abaixo da última linha existente
SENÃO:
    IMPRIMIR "Nenhuma medição nova"

// Tratamento de duplicatas (limpeza pós-carga):
ler aba completa
remover duplicatas por [ReadingTimeUTC, PointID, Level, Units]
limpar aba e reescrever dados limpos
```

---

## 10. NOTES

### Busca por origem
```
FUNÇÃO get_notes(token, base_url, origem):
    GET para base_url/v1/notes com Bearer token
    SE erro: lançar exceção

    PARA CADA note na resposta:
        extrair: { NoteID, PointID, NodeName, CreatedAt,
                   Title, Author, Text, Priority, origem }

    RETORNAR DataFrame
```

### Estrutura e organização
```
PARA CADA (origem, base_url) EM BASE_URLS:
    token = obter_token(base_url)
    df = get_notes(token, base_url, origem)
    acumular

df_notes = concatenar

JUNTAR df_notes com df_points pelo PointID → trazer coluna "path"
converter "path" para string
DIVIDIR "path" por "\" → colunas path_0, path_1, ...
normalizar path_1
FILTRAR por origem (mesmo critério das demais entidades)

selecionar colunas [NoteID, path, PointID, NodeName, Author,
                    CreatedAt, Title, Text, Priority, origem]
remover duplicatas

SALVAR Excel "cda_online_skf_notes.xlsx"
ENVIAR para Google Sheets "cda_online_skf_notes" (limpar + escrever)
```

---

## FLUXO GERAL (visão macro)

```
INÍCIO
│
├─ Configurar credenciais e URLs por origem (ubu / germano)
├─ Autenticar no Google Sheets
│
├─ [3]  Buscar MACHINES
│        └─ autenticar por origem → GET /machines
│        └─ filtrar por path, montar df_machines
│        └─ Excel + Sheets
│
├─ [4]  Buscar SUBMACHINES
│        └─ buscar hierarquia completa → percorrer árvore recursivamente
│        └─ filtrar por path, montar df_submachines
│        └─ Excel + Sheets
│
├─ [5]  Buscar MACHINE PARTS
│        └─ GET /machines/{id}/parts → expandir faultFrequencies
│        └─ juntar com df_machines → filtrar por path
│        └─ Excel + Sheets
│
├─ [6]  Buscar POINTS
│        └─ GET /machines/{id}/points → expandir pontos (explode)
│        └─ montar df_points com IDs para uso posterior
│        └─ Excel + Sheets
│
├─ [7]  Buscar ALARMS
│        └─ GET /machines/{id}/points → extrair OverallAlarm + Frequencies
│        └─ parsear texto do summary para valores numéricos
│        └─ Excel + Sheets
│
├─ [8]  Buscar LAST MEASUREMENTS
│        └─ GET /machines/{id}/points?IncludeLastMeasurement=true
│        └─ extrair última medição por ponto
│        └─ Excel + Sheets
│
├─ [9]  Buscar MEASUREMENTS (Trend D-1)
│        └─ GET /points/{id}/trendMeasurements com range de datas (ontem)
│        └─ filtrar: ChannelName "Overall" + Direction "X"
│        └─ carga incremental no Sheets (apenas registros novos)
│        └─ limpeza de duplicatas na aba
│
└─ [10] Buscar NOTES
         └─ GET /notes por origem
         └─ juntar com df_points → filtrar por path
         └─ Excel + Sheets

FIM
```

---

## PADRÃO DE FILTRAGEM POR ORIGEM (comum a todas as pesquisas)

```
// Aplicado após qualquer busca que retorne dados hierárquicos com "path"
1. converter coluna "path" para string
2. DIVIDIR "path" por "\" → gerar colunas path_0, path_1, path_2, ...
3. normalizar path_1: strip + upper
4. FILTRAR: manter apenas linhas onde
       path_1 == mapa[origem]
       mapa = { "ubu": "UBU", "germano": "GERMANO" }
```
"""
