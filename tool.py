import pandas as pd
from odf.opendocument import load
from odf.table import Table, TableRow, TableCell
from odf.text import P
import re

F9_SHEET      = "9"
FCODICI_SHEET = "Sheet1"
OUTPUT_SHEET  = "18"
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


def writeOnODS(path: str, sheet_name: str, colonna: str, valori: list):
    '''
    funzione che scrive sul foglio in formato ODS
    ricerca la colonna tramite il nome e ne evita il duplicato
    ritorna il documento salvato
    '''
    
    doc = load(path)
    
    target_table = None
    table_index = 0
    for table in doc.getElementsByType(Table):
        if table.getAttribute("name") == sheet_name:
            target_table = table
            break
        table_index += 1
    
    if target_table is None:
        try:
            idx = int(sheet_name)
            tables = list(doc.getElementsByType(Table))
            if 0 <= idx < len(tables):
                target_table = tables[idx]
        except ValueError:
            pass
    
    if target_table is None:
        raise ValueError(f"foglio {sheet_name} non trovato nel file ODS")
    
    rows = list(target_table.getElementsByType(TableRow))
    if not rows:
        return
    
    header_row = rows[0] #A1/B1
    header_cells = list(header_row.getElementsByType(TableCell))
    
    #ricerca della colonna nel param
    col_index = None
    current_col = 0
    for i, cell in enumerate(header_cells):
        repeat_cols = cell.getAttribute("numbercolumnsrepeated")
        if repeat_cols:
            repeat_cols = int(repeat_cols)
        else:
            repeat_cols = 1
        
        text_content = ""
        for p in cell.getElementsByType(P):
            if p.firstChild and hasattr(p.firstChild, 'data'):
                text_content += p.firstChild.data
        
        if text_content.strip() == colonna:
            col_index = current_col
            break
        current_col += repeat_cols
    
    if col_index is None:
        col_index = current_col
        new_cell = TableCell()
        p = P()
        p.addText(colonna)
        new_cell.addElement(p)
        header_row.addElement(new_cell)
    
    for row_index in range(1, len(valori) + 1):
        while len(rows) <= row_index:
            new_row = TableRow()
            target_table.addElement(new_row)
            rows = list(target_table.getElementsByType(TableRow))
        
        row = rows[row_index]
        cells = list(row.getElementsByType(TableCell))
        
        actual_col_count = 0
        for cell in cells:
            repeat_cols = cell.getAttribute("numbercolumnsrepeated")
            if repeat_cols:
                actual_col_count += int(repeat_cols)
            else:
                actual_col_count += 1
        
        while actual_col_count <= col_index:
            new_cell = TableCell()
            row.addElement(new_cell)
            cells = list(row.getElementsByType(TableCell))
            actual_col_count += 1
        
        target_cell = None
        current_col = 0

        #fix per colonne ripetute
        for cell in cells:
            repeat_cols = cell.getAttribute("numbercolumnsrepeated")
            if repeat_cols:
                repeat_cols = int(repeat_cols)
            else:
                repeat_cols = 1
            
            if current_col <= col_index < current_col + repeat_cols:
                if repeat_cols > 1 and col_index > current_col:
                    cell.setAttribute("numbercolumnsrepeated", str(col_index - current_col))
                    target_cell = TableCell()
                    parent = cell.parentNode
                    siblings = list(parent.childNodes)
                    cell_index = siblings.index(cell)
                    parent.insertBefore(target_cell, siblings[cell_index + 1] if cell_index + 1 < len(siblings) else None)
                    
                    remaining = repeat_cols - (col_index - current_col) - 1
                    if remaining > 0:
                        remaining_cell = TableCell()
                        remaining_cell.setAttribute("numbercolumnsrepeated", str(remaining))
                        parent.insertBefore(remaining_cell, target_cell.nextSibling)
                else:
                    target_cell = cell
                    if repeat_cols > 1:
                        cell.removeAttribute("numbercolumnsrepeated")
                break
            
            current_col += repeat_cols
        
        if target_cell is None:
            target_cell = TableCell()
            row.addElement(target_cell)
        
        for child in list(target_cell.childNodes):
            target_cell.removeChild(child)
        
        valore = valori[row_index - 1]
        if valore:
            p = P()
            p.addText(str(valore))
            target_cell.addElement(p)
    
    doc.save(path)


def parsef9(df_f9: pd.DataFrame) -> tuple[dict, dict]:
    '''
    funzione che copia ogni cella della colonna #ID dal foglio 9
    return: torna due dizionari che contengono: 
    nome_to_codici = {società1:0000-00}
    codici_to_nome = {0000-00:societa1}
    
    '''
    nome_to_codici = {}
    codice_to_nome = {}
    #ricerca della colonna id nel f9
    id_column = None
    for col in df_f9.columns:
        if '#ID' in col or 'ID' in col or 'id' in col.lower():
            id_column = col
            break
    
    if id_column is None:
        print(f"Colonna #ID non trovata. Colonne disponibili: {list(df_f9.columns)}")
        return nome_to_codici, codice_to_nome
    
    for _, row in df_f9.iterrows():
        cell_value = str(row.get(id_column, "")) if not pd.isna(row.get(id_column)) else ""
        if not cell_value:
            continue
        #rimozione del trattino nella colonna #ID
        if '_' in cell_value:
            parts = cell_value.split('_', 1)
            codice = parts[0].strip()
            nome = parts[1].strip()
            
            
            if codice and nome:
                if nome not in nome_to_codici:
                    nome_to_codici[nome] = []
                if codice not in nome_to_codici[nome]:
                    nome_to_codici[nome].append(codice)
                codice_to_nome[codice] = nome
        else: #da rimuovere
            match = re.match(r'^([0-9\-]+)\s+(.+)$', cell_value)
            if match:
                codice = match.group(1).strip()
                nome = match.group(2).strip()
                if codice and nome:
                    if nome not in nome_to_codici:
                        nome_to_codici[nome] = []
                    if codice not in nome_to_codici[nome]:
                        nome_to_codici[nome].append(codice)
                    codice_to_nome[codice] = nome
    
    return nome_to_codici, codice_to_nome 


def processa_cella_societa(cella: str, nome_to_codici: dict, codice_nome_to_sipid: dict, 
                           codice_to_sipid: dict, colonna_nome_convert: str,
                           societa_non_trovate: set, codici_non_trovati: set) -> list:
    '''
    funzione che lavora sulla colonna societa/societa_conferente e processa una singola cella che contiene una o più società separate da |
    se società = true, cerca nel foglio 9 la corrispondenza col codice
    restituisce una lista di sip_id trovati
    '''
    if not cella:
        return []
    
    societa_list = [s.strip() for s in cella.split("|") if s.strip()]
    sip_ids = []
    
    for societa in societa_list:
        codici = nome_to_codici.get(societa, [])
        if not codici:
            societa_non_trovate.add(societa)
            continue
        
        sip_id = None
        for codice in codici:
            if colonna_nome_convert:
                sip_id = codice_nome_to_sipid.get((codice, societa))
                if sip_id:
                    break
            
            sip_id = codice_to_sipid.get(codice)
            if sip_id:
                break
        
        if sip_id:
            sip_ids.append(sip_id)
        else:
            codici_non_trovati.update(codici)
    
    return sip_ids


def map_societa(
    path_f9: str,
    path_fcodici: str,
    f9_sheet: str = F9_SHEET,
    fcodici_sheet: str = FCODICI_SHEET,
    output_sheet: str = OUTPUT_SHEET,
    output_col: str = OUTPUT_COL,
    output_col_conferente: str = OUTPUT_COL_CONFERENTE,
    limit: int = None,
) -> None:
    '''
    3 fogli su cui lavora
    f9 = foglio con codici e nomi società
    fcodici = foglio convert con sip_ids
    output = il foglio dei procuratori

    codice_nome_to_sipid = mappa il nome e il codice per cercare un match dentro il fogli
    o convert con sipid
    codice_to_sipid = mappa solo codice e sipid
    '''
    
    df_f9 = _load_sheet(path_f9, f9_sheet)
    df_output = _load_sheet(path_f9, output_sheet)
    df_fcodici = _load_sheet(path_fcodici, fcodici_sheet)
    
    if limit is not None:
        df_output = df_output.head(limit)
        print(f"Limite impostato a {limit}")
    
    df_f9.columns = df_f9.columns.str.strip()
    df_output.columns = df_output.columns.str.strip()
    df_fcodici.columns = df_fcodici.columns.str.strip()
    
    colonna_nome_convert = input("\nColonna dei nomi nel file convert :").strip()
    
    nome_to_codici, codice_to_nome = parsef9(df_f9)
    
    duplicati = {nome: codici for nome, codici in nome_to_codici.items() if len(codici) > 1}
    if duplicati:
        print(f"trovati {len(duplicati)} nomi con codici multipli nel foglio 9")
    
    codice_nome_to_sipid = {}
    codice_to_sipid = {}
    
    for _, row in df_fcodici.iterrows():
        codice = str(row.get("cr082_codice", "")).strip()
        sip_id = str(row.get("cr082_sip_societaid", "")).strip()
        
        if codice and sip_id:
            if colonna_nome_convert and colonna_nome_convert in df_fcodici.columns:
                nome_convert = str(row.get(colonna_nome_convert, "")).strip()
                if nome_convert:
                    codice_nome_to_sipid[(codice, nome_convert)] = sip_id
            
            if codice not in codice_to_sipid:
                codice_to_sipid[codice] = sip_id
    
    ids_values = []
    ids_conferente_values = []
    societa_non_trovate = set()
    codici_non_trovati = set()
    
    for idx, row in df_output.iterrows():
        
        societa_cell = str(row.get("societa", "")) if not pd.isna(row.get("societa")) else ""
        societa_conferente_cell = str(row.get("societa_conferente", "")) if not pd.isna(row.get("societa_conferente")) else ""
        
        sip_ids_societa = processa_cella_societa(societa_cell, nome_to_codici, codice_nome_to_sipid, 
                                                  codice_to_sipid, colonna_nome_convert, 
                                                  societa_non_trovate, codici_non_trovati)
        
        sip_ids_conferente = processa_cella_societa(societa_conferente_cell, nome_to_codici, codice_nome_to_sipid, 
                                                     codice_to_sipid, colonna_nome_convert, 
                                                     societa_non_trovate, codici_non_trovati)
        
        # Colonna IDS (solo società)
        if sip_ids_societa:
            parte1 = "[" + ",".join([f'"{id_val}"' for id_val in sip_ids_societa]) + "]"
            ids_values.append(parte1)
        else:
            ids_values.append("")
        
        # Colonna IDS_CONFERENTE (solo società_conferente)
        if sip_ids_conferente:
            parte2 = "[" + ",".join([f'"{id_val}"' for id_val in sip_ids_conferente]) + "]"
            ids_conferente_values.append(parte2)
        else:
            ids_conferente_values.append("")
    
    writeOnODS(path_f9, output_sheet, output_col, ids_values)
    writeOnODS(path_f9, output_sheet, output_col_conferente, ids_conferente_values)
    print("Scrittura completata")


if __name__ == "__main__":
    print("Mappa Societa\n")
    
    f9      = input("Path ImportTemplate: ").strip().strip('"')
    fcodici = input("Path convert:        ").strip().strip('"')

    limit_raw = input("Limite righe per test (invio = none): ").strip()
    limit = None if not limit_raw else int(limit_raw)

    f9_sheet      = input(f"Foglio 9(invio = {F9_SHEET}):      ").strip() or F9_SHEET
    fcodici_sheet = input(f"Foglio convert(invio = {FCODICI_SHEET}): ").strip() or FCODICI_SHEET
    output_sheet  = input(f"Foglio 18(invio = {OUTPUT_SHEET}): ").strip() or OUTPUT_SHEET
    output_col    = input(f"Colonna(invio = {OUTPUT_COL}):    ").strip() or OUTPUT_COL
    output_col_conferente = input(f"Colonna conferente(invio = {OUTPUT_COL_CONFERENTE}): ").strip() or OUTPUT_COL_CONFERENTE
    
    map_societa(
        path_f9=f9,
        path_fcodici=fcodici,
        f9_sheet=f9_sheet,
        fcodici_sheet=fcodici_sheet,
        output_sheet=output_sheet,
        output_col=output_col,
        output_col_conferente=output_col_conferente,
        limit=limit,
    )