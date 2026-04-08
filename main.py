import os, re, numpy as np, pandas as pd
from datetime import datetime
os.system("cls" if os.name == "nt" else "clear")

df = pd.read_excel("tareas.xlsx", sheet_name="Sheet1")

TEAM = {
  "Kenny Marcial Rodríguez García": True,
  "SMARTLIFE": True,
  "Martin Bou Mansour Bakhos": True,
  "Karlanis Tariffa del Chiaro": False,
  "IVENETI": True,
  "INVERSIONES PENALVA 2022": False,
  "TELECOMUNICACIONES K. SUAREZ": False,
  "Marcial Venancio Rodríguez González": False,
  "LATIN TELECOM": False,
  "CG SERVICIOS": False,
  "Moisés Enmanuel Urbina Villarreal": False,
  "Argenis José Acosta Veliz": False,
  "Jonayker Daniel Mendoza Pérez": False,
  "Jorge Luis Rodríguez Parra": True,
  "Socrates Antonio Sequera Sanchez": True,
  "Óscar Eduardo Henriquez Zambrano": False,
  "Jeryco Salvador Ortega Palacios": False,
  "SAT SERVICES": False,
  "LFM Consultor": False,
  "Ruben Dario Sanchez Avila": False,
  "Dainee Yanibeth Zambrano Torres": False,
  "GRUPO ARLO SYSTEM": False,
  "TRS 2048": False
}

def find_header_value(header_name):
  result = df[df["Etapa"].str.contains(f"{header_name} \(")]

  if result.empty: return 0
  return re.search(r"\d+", result["Etapa"].iloc[0]).group()

header_index_list = df[df["Etapa"].str.contains("\(", na=False)].index.to_list()
status = {
  "asignadas": find_header_value("Asignado"),
  "en_progreso": find_header_value("En Progreso"),
  "por_facturar": find_header_value("Por facturar"),
  "sections": {},
  "asigned_persons": {}
}

for i in range(len(header_index_list) - 1):
  start = header_index_list[i]
  end = header_index_list[i+1]
  
  section_name = re.search(r"(.*)(?:\s*\(\d+\))", df.loc[start, "Etapa"]).groups()[0].strip()
  content = df.iloc[start + 1:end].copy()  
  
  if section_name == "Asignado":
    mask_vacias = content["Etapa"].isna() & content["Etiquetas"].notna()
    content.loc[mask_vacias.shift(-1, fill_value=False), 'Etiquetas_Extra'] = content['Etiquetas'].shift(-1)
    content['Etiquetas'] = content['Etiquetas'].astype(str) + ", " + content['Etiquetas_Extra'].fillna('')
    content = content[content['Etapa'].notna()].drop(columns=['Etiquetas_Extra']).reset_index(drop=True)

  if section_name == "Hecho":
    actual_date = pd.to_datetime('2026-04-07')
    content['Última actualización de la etapa'] = pd.to_datetime(content['Última actualización de la etapa'], errors='coerce').dt.floor("D")
    content = content[content['Última actualización de la etapa'] == actual_date].copy()
    print(len(content))

  status["sections"][section_name] = content

for index, row in status["sections"]["Asignado"].iterrows():
  filter = re.search(r"s*\(User\)?", row["Personas asignadas"])
  etiqueta = str(row["Etiquetas"]).upper().replace(" ", "")
  type = "RF" if "RF" in etiqueta else "FTTH"

  asigned_person = row["Personas asignadas"][0:filter.start()].strip() if filter else row["Personas asignadas"].strip()

  if asigned_person not in status["asigned_persons"]:
    if TEAM[asigned_person]: 
      status["asigned_persons"][asigned_person] = {
        "FTTH": 0,
        "RF": 0
      }
    else: status["asigned_persons"][asigned_person] = 0

  if TEAM[asigned_person]: status["asigned_persons"][asigned_person][type] += 1
  else: status["asigned_persons"][asigned_person] += 1

def fetch_value(name, mix = False):
  alias = {
    "KENNY": "Kenny Marcial Rodríguez García",
    "SMARTLIFE": "SMARTLIFE",
    "MARTIN": "Martin Bou Mansour Bakhos",
    "KARLANIS": "Karlanis Tariffa del Chiaro",
    "IVENETI": "IVENETI",
    "PENALVA": "INVERSIONES PENALVA 2022",
    "KSUAREZ": "TELECOMUNICACIONES K. SUAREZ",
    "MARCIAL": "Marcial Venancio Rodríguez González",
    "LATIN": "LATIN TELECOM",
    "CG": "CG SERVICIOS",
    "MOISES": "Moisés Enmanuel Urbina Villarreal",
    "ARGENIS": "Argenis José Acosta Veliz",
    "JONAYKER": "Jonayker Daniel Mendoza Pérez",
    "TERASERVICES VALENCIA": "Jorge Luis Rodríguez Parra",
    "TERASERVICES PUERTO": "Socrates Antonio Sequera Sanchez",
    "OSCAR": "Óscar Eduardo Henriquez Zambrano",
    "JERYCO": "Jeryco Salvador Ortega Palacios",
    "SAT SERVICES": "SAT SERVICES",
    "LFM": "LFM Consultor",
    "PE": "Ruben Dario Sanchez Avila",
    "DAINEE": "Dainee Yanibeth Zambrano Torres",
    "ARLO": "GRUPO ARLO SYSTEM",
    "TRS 2048": "TRS 2048"
  }
  condition = status["asigned_persons"].get(alias[name])

  if mix: return f"({status["asigned_persons"][alias[name]]["RF"]})RF / ({status["asigned_persons"][alias[name]]["FTTH"]})FTTH" if condition else "(0)RF / (0)FTTH"
  return f"({status["asigned_persons"][alias[name]]})" if condition else "(0)"

print(f"""
STATUS DE LAS ASISTENCIAS

▪️ N° de Asistencias Asignadas: {status["asignadas"]}
▪️ N° Tickets de Asistencias en espera: 0

▪️ *ARGENIS ACOSTA:* {fetch_value("ARGENIS")}
▪️ *CG SERVICIOS:* {fetch_value("CG")}
▪️ *DAINEE ZAMBRANO:* {fetch_value("DAINEE")}
▪️ *GRUPO ARLO:* {fetch_value("ARLO")}
▪️ *INVERSIONES PEÑALVA:* {fetch_value("PENALVA")}
▪️ *IVENETI:* {fetch_value("IVENETI", True)}
▪️ *JERYCO ORTEGA:* {fetch_value("JERYCO")}
▪️ *JONAYKER MENDOZA:* {fetch_value("JONAYKER")}
▪️ *K. SUAREZ:* {fetch_value("KSUAREZ")}
▪️ *KARLANIS TARIFFA:* {fetch_value("KARLANIS")}
▪️ *KENNY RODRIGUEZ:* {fetch_value("KENNY", True)}
▪️ *LATIN TELECOM:* {fetch_value("LATIN")}
▪️ *LFM CONSULTOR:* {fetch_value("LFM")}
▪️ *MARCIAL RODRIGUEZ:* {fetch_value("MARCIAL")}
▪️ *MARTIN BOU:* {fetch_value("MARTIN", True)}
▪️ *MOISES URBINA:* {fetch_value("MOISES")}
▪️ *OSCAR HENRIQUEZ:* {fetch_value("OSCAR")}
▪️ *PE:* (0)AOC / {fetch_value("PE")}
▪️ *SMARTLIFE:* {fetch_value("SMARTLIFE", True)}
▪️ *TERASERVICES VALENCIA:* {fetch_value("TERASERVICES VALENCIA", True)}
▪️ *TERASERVICES PUERTO:* {fetch_value("TERASERVICES PUERTO", True)}
▪️ TRS 2048: ({fetch_value("TRS 2048", True)})

▪️ Asistencias en progreso: {status["en_progreso"]}
▪️ Asistencias por facturar: {status["por_facturar"]}
▪️ Clientes atendidos por asistencia hoy: 0
""")