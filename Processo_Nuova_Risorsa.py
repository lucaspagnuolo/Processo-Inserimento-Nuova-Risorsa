import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# Utility
def formatta_data(data):
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

def genera_samaccountname(nome, cognome, secondo_nome="", secondo_cognome="", esterno=False):
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext" if esterno else ""
    limit = 16 if esterno else 20
    cand = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit: return cand + suffix
    cand = f"{(n[:1])}{(sn[:1])}.{c}{sc}"
    if len(cand) <= limit: return cand + suffix
    return (f"{n[:1]}{sn[:1]}.{c}" )[:limit] + suffix

def build_full_name(cognome, secondo_cognome, nome, secondo_nome, esterno=False):
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

HEADER = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName","Surname",
    "employeeNumber","employeeID","department","Description","passwordNeverExpired",
    "ExpireDate","userprincipalname","mail","mobile","RimozioneGruppo","InserimentoGruppo",
    "disable","moveToOU","telephoneNumber","company"
]

st.title("1.1 Nuova Risorsa Interna")

# Form
nome            = st.text_input("Nome").strip().capitalize()
secondo_nome    = st.text_input("Secondo Nome").strip().capitalize()
cognome         = st.text_input("Cognome").strip().capitalize()
secondo_cognome = st.text_input("Secondo Cognome").strip().capitalize()
numero_telefono = st.text_input("Numero di Telefono", "").replace(" ", "")
description     = st.text_input("Description (lascia vuoto per <PC>)", "<PC>").strip()
codice_fiscale  = st.text_input("Codice Fiscale", "").strip()

ou                 = st.selectbox("OU", ["Utenti standard","Utenti VIP"])
employee_id        = st.text_input("Employee ID", "").strip()
department         = st.text_input("Dipartimento", "").strip()
inserimento_gruppo = "consip_vpn;dipendenti_wifi;mobile_wifi;GEDOGA-P-DOCGAR;GRPFreeDeskUser"
telephone_number   = "+39 06 854491"
company            = "Consip"

if st.button("Genera CSV Interna"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn  = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    row = [
        sAM, "SI", ou, cn.replace(" (esterno)",""), cn, cn,
        " ".join([nome, secondo_nome]).strip(),
        " ".join([cognome, secondo_cognome]).strip(),
        codice_fiscale, employee_id, department,
        description or "<PC>", "No", "",
        f"{sAM}@consip.it",
        f"{sAM}@consip.it",
        f"+39 {numero_telefono}" if numero_telefono else "",
        "", inserimento_gruppo, "", "",
        telephone_number, company
    ]
    buf = io.StringIO()
    csv.writer(buf).writerow(HEADER)
    csv.writer(buf).writerow(row)
    buf.seek(0)
    df = pd.DataFrame([row], columns=HEADER)
    st.dataframe(df)
    st.download_button("📥 Scarica CSV", buf.getvalue(),
                       file_name=f"{cognome}_{nome[:1]}_interno.csv",
                       mime="text/csv")
    st.success(f"✅ Creato {sAM}")

