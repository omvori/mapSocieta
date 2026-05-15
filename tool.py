import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

OUTPUT_SHEET  = "008"
OUTPUT_COL    = "IDS"
OUTPUT_COL_CONFERENTE = "IDS_CONFERENTE"

#caricamento delle sheet via calamine
def _load_sheet(path: str, sheet: str, engine: str = "calamine") -> pd.DataFrame:
    '''
    carica solamente tutte le sheet
    '''
    all_sheets = pd.read_excel(path, sheet_name=None, engine=engine, dtype=str)
    if sheet in all_sheets:
        return all_sheets[sheet]
    try:
        idx = int(sheet)
        return list(all_sheets.values())[idx]
    except (ValueError, IndexError):
        available = list(all_sheets.keys())
        raise KeyError(f"foglio non trovato, disponibili: {available}")


def scrivi_colonne_xlsx(path: str, sheet_name: str, colonne_valori: dict):
    '''
    Scrive multiple colonne in un file Excel usando openpyxl
    colonne_valori = {'IDS': [...], 'IDS_CONFERENTE': [...]}
    '''
    
    wb = load_workbook(path)
    
    #Cerca per nome esatto
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        raise ValueError(f"Foglio '{sheet_name}' non trovato. Disponibili: {wb.sheetnames}")
    
    #Leggi gli header esistenti
    headers = {}
    max_col = ws.max_column
    for col in range(1, max_col + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value:
            headers[str(cell_value).strip()] = col
    
    #Per ogni colonna da scrivere
    for col_name, valori in colonne_valori.items():
        if col_name in headers:
            col_idx = headers[col_name]
        else:
            col_idx = max_col + 1
            ws.cell(row=1, column=col_idx, value=col_name)
            max_col += 1
        
        for i, val in enumerate(valori):
            ws.cell(row=i + 2, column=col_idx, value=val if val else None)
    
    wb.save(path)
    print(f"Scritte {len(colonne_valori)} colonne nel foglio '{ws.title}'")


def processa_cella_societa(cella: str, societa_to_value: dict, societa_non_trovate: set) -> list:
    '''
    Processa una singola cella che contiene una o più società separate da |
    Formato società: "0800-01_Eni S.p.A. - CORPORATE"
    Match esatto con la Key del dizionario
    '''
    if not cella:
        return []
    
    societa_list = [s.strip() for s in cella.split("|") if s.strip()]
    values = []
    
    for societa in societa_list:
        value = societa_to_value.get(societa)
        
        if value:
            values.append(value)
        else:
            societa_non_trovate.add(societa)
    
    return values


def map_societa(
    path_importtemplate: str,
    path_sip: str,
    output_sheet: str = OUTPUT_SHEET,
    output_col: str = OUTPUT_COL,
    output_col_conferente: str = OUTPUT_COL_CONFERENTE,
    limit: int = None,
) -> None:
    
    
    print(f"\nCaricamento ImportTemplate foglio '{output_sheet}'...")
    df_output = _load_sheet(path_importtemplate, output_sheet)
    
    print(f"Caricamento file sip_societa")
    df_sip = _load_sheet(path_sip, "Sheet1")
    
    if limit is not None:
        df_output = df_output.head(limit)
        print(f"Limite impostato a {limit}")
    
    df_output.columns = df_output.columns.str.strip()
    df_sip.columns = df_sip.columns.str.strip()
    
    societa_to_value = {}
    for _, row in df_sip.iterrows():
        key = str(row.get("Key", "")).strip()
        value = str(row.get("Value", "")).strip()
        
        if key and value:
            societa_to_value[key] = value
    
    print(f"Trovate {len(societa_to_value)} corrispondenze Key to Value")
    
    ids_values = []
    ids_conferente_values = []
    societa_non_trovate = set()
    
    for idx, row in df_output.iterrows():
        societa_cell = str(row.get("societa", "")) if not pd.isna(row.get("societa")) else ""
        societa_conferente_cell = str(row.get("societa_conferente", "")) if not pd.isna(row.get("societa_conferente")) else ""
        
        values_societa = processa_cella_societa(societa_cell, societa_to_value, societa_non_trovate)
        values_conferente = processa_cella_societa(societa_conferente_cell, societa_to_value, societa_non_trovate)
        
        #popolazione colonna IDS
        if values_societa:
            parte1 = "[" + ",".join([f'"{id_val}"' for id_val in values_societa]) + "]"
            ids_values.append(parte1)
        else:
            ids_values.append("")
        
        #popolazione colonna IDS_CONFERENTE
        if values_conferente:
            parte2 = "[" + ",".join([f'"{id_val}"' for id_val in values_conferente]) + "]"
            ids_conferente_values.append(parte2)
        else:
            ids_conferente_values.append("")

    if societa_non_trovate:
        print(f"\nSocietà NON trovate nel file sip ({len(societa_non_trovate)}):")
        for s in sorted(societa_non_trovate)[:20]:
            print(f"  - '{s}'")
        if len(societa_non_trovate) > 20:
            print(f"e altre {len(societa_non_trovate) - 20}")
    
    print(f"Scrittura colonne nel foglio '{output_sheet}'...")
    scrivi_colonne_xlsx(path_importtemplate, output_sheet, {
        output_col: ids_values,
        output_col_conferente: ids_conferente_values
    })
    print("Scrittura completata")


if __name__ == "__main__":
    print("tool mapping\n")
    
    importtemplate = input("Path ImportTemplate: ").strip().strip('"')
    sip = input("Path sip_societa: ").strip().strip('"')

    limit_raw = input("Limite righe per test (invio = none): ").strip()
    limit = None if not limit_raw else int(limit_raw)

    output_sheet = input(f"Foglio in ImportTemplate (invio = {OUTPUT_SHEET}): ").strip() or OUTPUT_SHEET
    output_col = input(f"Colonna IDS (invio = {OUTPUT_COL}): ").strip() or OUTPUT_COL
    output_col_conferente = input(f"Colonna IDS_CONFERENTE (invio = {OUTPUT_COL_CONFERENTE}): ").strip() or OUTPUT_COL_CONFERENTE
    
    map_societa(
        path_importtemplate=importtemplate,
        path_sip=sip,
        output_sheet=output_sheet,
        output_col=output_col,
        output_col_conferente=output_col_conferente,
        limit=limit,
    )
    
    print("\nCompletato")