import streamlit as st
import csv
import pandas as pd
from datetime import datetime, timedelta
import io

# ------------------------------------------------------------
# Caricamento configurazione da Excel caricato dall'utente
# ------------------------------------------------------------
def load_config_from_bytes(data: bytes):
    cfg = pd.read_excel(io.BytesIO(data), sheet_name="Risorsa Interna")
    # Sezione OU
    ou_df = cfg[cfg["Section"] == "OU"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "label"}
    )
    ou_options = dict(zip(ou_df["key"], ou_df["label"]))
    # Sezione InserimentoGruppi
    grp_df = cfg[cfg["Section"] == "InserimentoGruppi"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "app", "Label/Gruppi/Value": "gruppi"}
    )
    gruppi = dict(zip(grp_df["app"], grp_df["gruppi"]))
    # Sezione Defaults
    def_df = cfg[cfg["Section"] == "Defaults"][["Key/App", "Label/Gruppi/Value"]].rename(
        columns={"Key/App": "key", "Label/Gruppi/Value": "value"}
    )
    defaults = dict(zip(def_df["key"], def_df["value"]))
    return ou_options, gruppi, defaults

# ------------------------------------------------------------
# App 1.1: Nuova Risorsa Interna
# ------------------------------------------------------------
st.set_page_config(page_title="1.1 Nuova Risorsa Interna")
st.title("1.1 Nuova Risorsa Interna")

config_file = st.file_uploader(
    "Carica il file di configurazione (config_corrected.xlsx)",
    type=["xlsx"],
    help="Deve contenere il foglio “Risorsa Interna” con campo Section"
)
if not config_file:
    st.warning("Per favore carica il file di configurazione per continuare.")
    st.stop()

ou_options, gruppi, defaults = load_config_from_bytes(config_file.read())

# ------------------------------------------------------------
# Utility functions
# ------------------------------------------------------------
def formatta_data(data: str) -> str:
    for sep in ["-", "/"]:
        try:
            g, m, a = map(int, data.split(sep))
            dt = datetime(a, m, g) + timedelta(days=1)
            return dt.strftime("%m/%d/%Y 00:00")
        except:
            continue
    return data

def genera_samaccountname(nome: str, cognome: str,
                          secondo_nome: str = "", secondo_cognome: str = "",
                          esterno: bool = False) -> str:
    n, sn = nome.strip().lower(), secondo_nome.strip().lower()
    c, sc = cognome.strip().lower(), secondo_cognome.strip().lower()
    suffix = ".ext" if esterno else ""
    limit  = 16 if esterno else 20
    cand   = f"{n}{sn}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    cand = f"{n[:1]}{sn[:1]}.{c}{sc}"
    if len(cand) <= limit:
        return cand + suffix
    return (f"{n[:1]}{sn[:1]}.{c}")[:limit] + suffix

def build_full_name(cognome: str, secondo_cognome: str,
                    nome: str, secondo_nome: str,
                    esterno: bool = False) -> str:
    parts = [p for p in [cognome, secondo_cognome, nome, secondo_nome] if p]
    full  = " ".join(parts)
    return full + (" (esterno)" if esterno else "")

HEADER = [
    "sAMAccountName","Creation","OU","Name","DisplayName","cn","GivenName","Surname",
    "employeeNumber","employeeID","department","Description","passwordNeverExpired",
    "ExpireDate","userprincipalname","mail","mobile","RimozioneGruppo","InserimentoGruppo",
    "disable","moveToOU","telephoneNumber","company"
]

# ------------------------------------------------------------
# Form di input
# ------------------------------------------------------------
st.subheader("Modulo Inserimento Nuova Risorsa Interna")

employee_id        = st.text_input("Matricola", defaults.get("employee_id_default", "")).strip()
cognome            = st.text_input("Cognome").strip().capitalize()
secondo_cognome    = st.text_input("Secondo Cognome").strip().capitalize()
nome               = st.text_input("Nome").strip().capitalize()
secondo_nome       = st.text_input("Secondo Nome").strip().capitalize()
codice_fiscale     = st.text_input("Codice Fiscale", "").strip()
department         = st.text_input("Sigla Divisione-Area", defaults.get("department_default", "")).strip()
numero_telefono    = st.text_input("Mobile", "").replace(" ", "")
description        = st.text_input("PC (lascia vuoto per <PC>)", "<PC>").strip()

ou_labels    = list(ou_options.values())
default_ou   = defaults.get("ou_default", ou_labels[0] if ou_labels else "")
label_ou     = st.selectbox("Tipologia Utente", options=ou_labels,
                             index=ou_labels.index(default_ou) if default_ou in ou_labels else 0)
selected_ou_key = list(ou_options.keys())[ou_labels.index(label_ou)]
ou_value     = ou_options[selected_ou_key]

inserimento_gruppo = gruppi.get("interna", "")
telephone_number   = defaults.get("telephone_interna", "")
company            = defaults.get("company_interna", "")

# ------------------------------------------------------------
# Generazione CSV
# ------------------------------------------------------------
if st.button("Genera CSV Interna"):
    sAM = genera_samaccountname(nome, cognome, secondo_nome, secondo_cognome, False)
    cn  = build_full_name(cognome, secondo_cognome, nome, secondo_nome, False)
    given = f"{nome} {secondo_nome}".strip()
    surn  = f"{cognome} {secondo_cognome}".strip()
    mobile = f"+39 {numero_telefono}" if numero_telefono else ""
    telnum = telephone_number

    row = [
        sAM, "SI", ou_value,                      # index 0,1,2
        cn.replace(" (esterno)", ""),             # 3 → Name
        cn,                                       # 4 → DisplayName
        cn,                                       # 5 → cn
        given,                                    # 6 → GivenName
        surn,                                     # 7 → Surname
        codice_fiscale, employee_id, department,  # 8,9,10
        description or "<PC>", "No", "",          # 11,12,13
        f"{sAM}@consip.it", f"{sAM}@consip.it",  # 14,15
        mobile,                                  # 16 → mobile
        "", inserimento_gruppo, "", "",          # 17,18,19,20
        telnum,                                  # 21 → telephoneNumber
        company                                  # 22
    ]

    # Avvolgi tra virgolette i campi richiesti
    # OU (idx 2), Name (3), DisplayName (4), cn (5)
    for i in (2, 3, 4, 5):
        row[i] = f"\"{row[i]}\""
    # GivenName (6) solo se secondo_nome non vuoto
    if secondo_nome:
        row[6] = f"\"{row[6]}\""
    # Surname (7) solo se secondo_cognome non vuoto
    if secondo_cognome:
        row[7] = f"\"{row[7]}\""
    # mobile (16) e telephoneNumber (21)
    for i in (16, 21):
        row[i] = f"\"{row[i]}\""

    buf = io.StringIO()
    writer = csv.writer(buf, quoting=csv.QUOTE_NONE, escapechar="\\")
    writer.writerow(HEADER)
    writer.writerow(row)
    buf.seek(0)

    st.dataframe(pd.DataFrame([row], columns=HEADER))
    st.download_button(
        label="📥 Scarica CSV",
        data=buf.getvalue(),
        file_name=f"{cognome}_{nome[:1]}_interno.csv",
        mime="text/csv"
    )
    st.success(f"✅ File CSV generato per '{sAM}'")
