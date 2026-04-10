import os, re, numpy as np, pandas as pd
os.system("cls" if os.name == "nt" else "clear")

GROUP = {
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

df = pd.read_excel("tareas.xlsx", sheet_name="Sheet1")
date = input("Introduzca el día a verificar (AAAA-MM-DD): ")

def find_header_value(header_name):
  result = df[df["Etapa"].str.contains(f"{header_name} \(")]

  if result.empty: return 0
  return re.search(r"\d+", result["Etapa"].iloc[0]).group()

status = {
  "asignadas": find_header_value("Asignado"),
  "en_progreso": find_header_value("En Progreso"),
  "por_facturar": find_header_value("Por facturar"),
  "sections": {},
  "asigned_persons": {}
}

header_index_list = df[df["Etapa"].str.contains("\(", na=False)].index.to_list()
for i in range(len(header_index_list) - 1):
  start = header_index_list[i]
  end = header_index_list[i+1]
  
  section_name = re.search(r"(.*)(?:\s*\(\d+\))", df.loc[start, "Etapa"]).groups()[0].strip()
  if section_name != "Asignado" and section_name != "Hecho" and section_name != "Por facturar": continue

  content = df.iloc[start + 1:end].copy()  
  
  if section_name == "Asignado":
    content['temp_group'] = content['Etapa'].notna().cumsum()
    content = content.groupby('temp_group').agg({
      'Etapa': 'first',
      'Personas asignadas': 'first',
      'Fecha límite': 'first',
      'Etiquetas': lambda x: ', '.join(x.dropna().astype(str))
    }).reset_index(drop=True)    
    content = content[['Etapa', 'Personas asignadas', 'Etiquetas', 'Fecha límite']]

  if section_name == "Hecho" or section_name == "Por facturar":
    actual_date = pd.to_datetime(date)
    content["Fecha límite"] = pd.to_datetime(content["Fecha límite"], errors="coerce").dt.floor("D")
    content = content[content["Fecha límite"] == actual_date].copy()

  status["sections"][section_name] = content

for index, row in status["sections"]["Asignado"].iterrows():
  filter = re.search(r"s*\(User\)?", row["Personas asignadas"])
  etiqueta = str(row["Etiquetas"]).upper().replace(" ", "")
  type = "RF" if "RF" in etiqueta else "FTTH"

  asigned_person = row["Personas asignadas"][0:filter.start()].strip() if filter else row["Personas asignadas"].strip()

  if asigned_person not in status["asigned_persons"]:
    if GROUP[asigned_person]: 
      status["asigned_persons"][asigned_person] = {
        "FTTH": 0,
        "RF": 0
      }
    else: status["asigned_persons"][asigned_person] = 0

  if GROUP[asigned_person]: status["asigned_persons"][asigned_person][type] += 1
  else: status["asigned_persons"][asigned_person] += 1

def fetch_value(name):
  condition = status["asigned_persons"].get(name)

  if condition and GROUP[name]: return f"({status["asigned_persons"][name]["RF"]})RF / ({status["asigned_persons"][name]["FTTH"]})FTTH"
  return f"({status["asigned_persons"][name]})" if condition else "(0)"

print(f"""
**STATUS DE LAS ASISTENCIAS**

▪️ _N° de Asistencias Asignadas:_ {status["asignadas"]}
▪️ _N° Tickets de Asistencias en espera:_ 0

▪️ *ARGENIS ACOSTA:* {fetch_value("Argenis José Acosta Veliz")}
▪️ *CG SERVICIOS:* {fetch_value("CG SERVICIOS")}
▪️ *DAINEE ZAMBRANO:* {fetch_value("Dainee Yanibeth Zambrano Torres")}
▪️ *GRUPO ARLO:* {fetch_value("GRUPO ARLO SYSTEM")}
▪️ *INVERSIONES PEÑALVA:* {fetch_value("INVERSIONES PENALVA 2022")}
▪️ *IVENETI:* {fetch_value("IVENETI")}
▪️ *JERYCO ORTEGA:* {fetch_value("Jeryco Salvador Ortega Palacios")}
▪️ *JONAYKER MENDOZA:* {fetch_value("Jonayker Daniel Mendoza Pérez")}
▪️ *K. SUAREZ:* {fetch_value("TELECOMUNICACIONES K. SUAREZ")}
▪️ *KARLANIS TARIFFA:* {fetch_value("Karlanis Tariffa del Chiaro")}
▪️ *KENNY RODRIGUEZ:* {fetch_value("Kenny Marcial Rodríguez García")}
▪️ *LATIN TELECOM:* {fetch_value("LATIN TELECOM")}
▪️ *LFM CONSULTOR:* {fetch_value("LFM")}
▪️ *MARCIAL RODRIGUEZ:* {fetch_value("Marcial Venancio Rodríguez González")}
▪️ *MARTIN BOU:* {fetch_value("Martin Bou Mansour Bakhos")}
▪️ *MOISES URBINA:* {fetch_value("Moisés Enmanuel Urbina Villarreal")}
▪️ *OSCAR HENRIQUEZ:* {fetch_value("Óscar Eduardo Henriquez Zambrano")}
▪️ *PE:* (0)AOC / {fetch_value("Ruben Dario Sanchez Avila")}
▪️ *SMARTLIFE:* {fetch_value("SMARTLIFE")}
▪️ *TERASERVICES ARAGUA:* {fetch_value("TRS 2048")}
▪️ *TERASERVICES VALENCIA:* {fetch_value("Jorge Luis Rodríguez Parra")}
▪️ *TERASERVICES PUERTO:* {fetch_value("Socrates Antonio Sequera Sanchez")}

▪️ _Asistencias en progreso:_ {status["en_progreso"]}
▪️ _Asistencias por facturar:_ {status["por_facturar"]}
▪️ _Clientes atendidos por asistencia hoy:_ {len(status['sections']["Hecho"]) + len(status['sections']["Por facturar"])}
""")