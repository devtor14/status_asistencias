import os, re, pandas as pd
os.system("cls" if os.name == 'nt' else "clear")

df = pd.read_excel('tareas.xlsx', sheet_name='Sheet1')
df["Categoria"] = df["Etapa"].where(df["Etapa"].str.contains(r"\(", na=False))
df["Categoria"] = df["Categoria"].ffill()

is_nan = df['Personas asignadas'].isna() & df['Etiquetas'].notna()
df['Etiquetas'] = df['Etiquetas'].shift(-1).where(is_nan.shift(-1), df['Etiquetas'])

df_limpio = df.dropna(subset=['Personas asignadas']).copy()
df_final = df_limpio[df_limpio['Categoria'].str.contains('Asignado', na=False)].copy()
reporte = df_final[['Personas asignadas', 'Etiquetas', 'Última actualización de la etapa']]

TEAM = {
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

MIX_TEAM = [
  TEAM["IVENETI"],
  TEAM["KENNY"],
  TEAM["MARTIN"],
  TEAM["SMARTLIFE"],
  TEAM["TERASERVICES VALENCIA"],
  TEAM["TERASERVICES PUERTO"],
  # TEAM["TERASERVICES ARAGUA"]
]

status = {
  "Asignado": 0
}

for index, row in reporte.iterrows():
  filter = re.search(r"s*\(User\)?", row['Personas asignadas'])
  etiqueta = str(row['Etiquetas']).upper()
  type = "RF" if "RF" == etiqueta else "FTTH"

  person_name = row["Personas asignadas"][0:filter.start()].strip() if filter else row["Personas asignadas"].strip()

  if person_name not in status:
    if person_name in MIX_TEAM: 
      status[person_name] = {
        "FTTH": 0,
        "RF": 0,
        "Total": 0
      }
    else: status[person_name] = 0

  if person_name in MIX_TEAM:
    status[person_name][type] += 1
    status[person_name]["Total"] += 1
  else: status[person_name] += 1

def exist_in_status(name, type = None):
  if type: return status[TEAM[name]][type] if TEAM[name] in status else 0 
  return status[TEAM[name]] if TEAM[name] in status else 0

print(f"""
STATUS DE LAS ASISTENCIAS

▪️ N° de Asistencias Asignadas: {status["Asignado"]}
▪️ N° Tickets de Asistencias en espera: 0

▪️ *ARGENIS ACOSTA:* ({exist_in_status("ARGENIS")})
▪️ *CG SERVICIOS:* ({exist_in_status("CG")})
▪️ *DAINEE ZAMBRANO: *({exist_in_status("DAINEE")})
▪️ *GRUPO ARLO:* ({exist_in_status("ARLO")})
▪️ *INVERSIONES PEÑALVA:* ({exist_in_status("PENALVA")})
▪️ *IVENETI:* ({exist_in_status("IVENETI", "RF")})RF / ({exist_in_status("IVENETI", "FTTH")})FTTH
▪️ *JERYCO ORTEGA:* ({exist_in_status("JERYCO")})
▪️ *JONAYKER MENDOZA:* ({exist_in_status("JONAYKER")})
▪️ *K. SUAREZ:* ({exist_in_status("KSUAREZ")})
▪️ *KARLANIS TARIFFA:* ({exist_in_status("KARLANIS")})
▪️ *KENNY RODRIGUEZ:* ({exist_in_status("KENNY", "RF")})RF / ({exist_in_status("KENNY", "FTTH")})FTTH
▪️ *LATIN TELECOM:* ({exist_in_status("LATIN")})
▪️ *LFM CONSULTOR:* ({exist_in_status("LFM")})
▪️ *MARCIAL RODRIGUEZ:* ({exist_in_status("MARCIAL")})
▪️ *MARTIN BOU:* ({exist_in_status("MARTIN", "RF")})RF / ({exist_in_status("MARTIN", "FTTH")})FTTH
▪️ *MOISES URBINA:* ({exist_in_status("MOISES")})
▪️ *OSCAR HENRIQUEZ:* ({exist_in_status("OSCAR")})
▪️ *PE:* (0)AOC / ({exist_in_status("PE")})FTTH
▪️ *SMARTLIFE:* ({exist_in_status("SMARTLIFE", "RF")})RF / ({exist_in_status("SMARTLIFE", "FTTH")})FTTH
▪️ *TERASERVICES VALENCIA:* ({exist_in_status("TERASERVICES VALENCIA", "RF")})RF / ({exist_in_status("TERASERVICES VALENCIA", "FTTH")})FTTH
▪️ *TERASERVICES PUERTO:* ({exist_in_status("TERASERVICES PUERTO", "RF")})RF / ({exist_in_status("TERASERVICES PUERTO", "FTTH")})FTTH
▪️ *TERASERVICES ARAGUA:* (0)

▪️ Asistencias en progreso: 0
▪️ Asistencias por facturar: 0
▪️ Clientes atendidos por asistencia hoy: 0
""")