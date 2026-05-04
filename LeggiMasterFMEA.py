import pandas as pd
import os


def format_masters_db(file_path):
    """Formatta il file Excel, rimuove duplicati e scompatta le celle in modo intelligente."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Il file {file_path} non esiste!")

    # 1. Legge il foglio 'MASTERS'
    try:
        df = pd.read_excel(file_path, sheet_name='MASTERS')
    except ValueError:
        df = pd.read_excel(file_path, sheet_name=0)

    df.dropna(how='all', inplace=True)

    if df.empty:
        return 0

    colonna_fasi = df.columns[0]

    # 2. SCOMPOSIZIONE CELLE UNITE INTELLIGENTE
    # A. Prima riempiamo le celle vuote solo nella prima colonna (Nome Fase)
    df[colonna_fasi] = df[colonna_fasi].ffill()

    # B. Poi riempiamo il resto delle colonne, ma RAGGRUPPATE PER FASE!
    # In questo modo, una riga nuova (tutta vuota tranne il nome) non prenderà MAI i difetti della fase sopra.
    colonne_dati = df.columns[1:]
    df[colonne_dati] = df.groupby(colonna_fasi, sort=False)[colonne_dati].ffill()

    # 3. FORMATTAZIONE DEL TESTO (Solo iniziale maiuscola)
    def sistema_testo(valore):
        if isinstance(valore, str) and valore.strip():
            return valore.strip().capitalize()
        return valore

    for col in df.columns:
        df[col] = df[col].apply(sistema_testo)

    df.columns = [sistema_testo(col) for col in df.columns]

    # 4. ELIMINAZIONE DUPLICATI PERFETTI
    righe_iniziali = len(df)
    df.drop_duplicates(inplace=True)
    duplicati_rimossi = righe_iniziali - len(df)

    # 5. SALVATAGGIO NEL FOGLIO "database_funzionale"
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='database_funzionale', index=False)

    return duplicati_rimossi


if __name__ == "__main__":
    file_master = "MASTERS-FMEA.xlsx"
    try:
        print(f"🔄 Avvio pulizia e formattazione di {file_master}...")
        rimossi = format_masters_db(file_master)
        print(f"✅ Operazione completata! Rimossi {rimossi} duplicati.")
    except Exception as e:
        print(f"❌ Errore: {e}")