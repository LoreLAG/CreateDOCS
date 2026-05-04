import pandas as pd
import os


def format_master_cp_db(file_path):
    """
    Formatta il file Excel del Control Plan, unendo le due righe di intestazione,
    rimuove duplicati, formatta il testo e scompatta le celle unite.
    """
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Il file {file_path} non esiste!")

    # 1. Ignora le prime 8 righe di intestazione aziendale e usa le righe 9 e 10 di Excel (indici 8 e 9) come nomi delle colonne
    df = pd.read_excel(file_path, sheet_name=0, header=[8, 9])

    # ELIMINA le colonne e le righe fantasma (completamente vuote) per evitare crash
    df.dropna(axis=1, how='all', inplace=True)
    df.dropna(axis=0, how='all', inplace=True)

    # Appiattisce l'intestazione su due righe in una sola stringa (es. "Fase_Macchina")
    # Ignorando le colonne vuote e DEDUPLICANDO i nomi uguali
    nuove_colonne = []
    seen = {}
    for i, col in enumerate(df.columns.values):
        nome = ' '.join([str(c) for c in col if "Unnamed" not in str(c) and pd.notna(c)]).strip()

        # Se il nome risultante è vuoto, assegna un nome fittizio
        if not nome:
            nome = f"Colonna_SenzaNome_{i}"

        # Sistema Anti-Crash per colonne con lo stesso identico nome
        if nome in seen:
            seen[nome] += 1
            nome = f"{nome}_{seen[nome]}"
        else:
            seen[nome] = 1

        nuove_colonne.append(nome)

    df.columns = nuove_colonne

    if df.empty:
        return 0

    colonna_fasi = df.columns[0]
    colonne_dati = df.columns[1:]

    # 2. SCOMPOSIZIONE CELLE UNITE INTELLIGENTE
    # A. Riempiamo le celle vuote della Fase (colonna 0)
    df[colonna_fasi] = df[colonna_fasi].ffill()

    # B. Riempiamo il resto delle colonne, raggruppate per fase (eseguito colonna per colonna per evitare disallineamenti)
    for col in colonne_dati:
        df[col] = df.groupby(colonna_fasi, sort=False)[col].ffill()

    # 3. FORMATTAZIONE DEL TESTO
    def sistema_testo_maiuscolo(valore):
        """Forza tutto in MAIUSCOLO (usato per la prima colonna)."""
        if isinstance(valore, str) and valore.strip():
            return valore.strip().upper()
        return valore

    def sistema_testo_normale(valore):
        """Maiuscola la prima lettera, lascia intatto il resto (per salvare le sigle come SPC, CMM)."""
        if isinstance(valore, str) and valore.strip():
            testo = valore.strip()
            return testo[0].upper() + testo[1:]
        return valore

    # Applica TUTTO IN MAIUSCOLO alla prima colonna
    df[colonna_fasi] = df[colonna_fasi].apply(sistema_testo_maiuscolo)

    # Applica "prima lettera maiuscola, salva sigle" al resto delle colonne
    for col in colonne_dati:
        df[col] = df[col].apply(sistema_testo_normale)

    # Applica la formattazione "normale" anche ai titoli delle colonne appena appiattite
    df.columns = [sistema_testo_normale(col) for col in df.columns]

    # 4. ELIMINAZIONE DUPLICATI PERFETTI
    righe_iniziali = len(df)
    df.drop_duplicates(inplace=True)
    duplicati_rimossi = righe_iniziali - len(df)

    # 5. SALVATAGGIO NEL FOGLIO "database_funzionale"
    with pd.ExcelWriter(file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, sheet_name='database_funzionale', index=False)

    return duplicati_rimossi


if __name__ == "__main__":
    file_master_cp = "MASTERS-CP.xlsx"

    try:
        print(f"🔄 Avvio pulizia e formattazione di {file_master_cp}...")
        rimossi = format_master_cp_db(file_master_cp)
        print(f"✅ Operazione completata! Rimossi {rimossi} duplicati.")
        print(f"📁 Foglio 'database_funzionale' generato/aggiornato con successo.")
    except Exception as e:
        print(f"❌ Errore: {e}")