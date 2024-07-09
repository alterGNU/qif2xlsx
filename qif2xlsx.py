import pandas as pd
import argparse

def read_qif(file):
    with open(file, 'r') as f:
        lines = f.readlines()
    
    data = []
    record = {}
    for line in lines:
        if line.startswith('^'):
            if record:
                data.append(record)
            record = {}
        else:
            if line.startswith('D'):
                record['DATE'] = line[1:].strip()
            elif line.startswith('T'):
                total = line[1:].strip()
                if (total[0] == "-"):
                    record['Debit'] = total[1:]
                    record['Credit'] = ""
                else:
                    record['Debit'] = ""
                    record['Credit'] = total
            elif line.startswith('P'):
                record['Libelle'] = line[1:].strip()
            elif line.startswith('L'):
                record['Reference'] = line[1:].strip()
    if record:
        data.append(record)

    return data

def create_rows(data):
    rows = []
    for i, entry in enumerate(data):
        row_num = i + 1
        row_1 = [f"512", entry.get('DATE', ''), entry.get('Reference', ''), entry.get('Libelle', ''), entry.get('Debit', ''), entry.get('Credit', '')]
        row_2 = [f"XXX", entry.get('DATE', ''), entry.get('Reference', ''), entry.get('Libelle', ''), entry.get('Credit', ''), entry.get('Debit', '')]
        rows.extend([row_1, row_2])
    return rows

def main():
    parser = argparse.ArgumentParser(description="Convertie le format QIF vers le format Excel.")
    parser.add_argument("qif_file", help="chemin vers le fichier QIF a convertir")
    parser.add_argument("excel_file", help="chemin et nom du fichier excel a creer")
    args = parser.parse_args()

    # Lire les données du fichier .qif
    data = read_qif(args.qif_file)

    # Créer des lignes pour le DataFrame
    rows = create_rows(data)

    # Colonnes du DataFrame
    columns = ['Comptes', 'DATE', 'Reference', 'Libelle', 'Debit', 'Credit']

    # Créer le DataFrame
    df = pd.DataFrame(rows, columns=columns)

    # Enregistrer en tant que fichier Excel
    df.to_excel(args.excel_file, index=False)
    print(f"Conversion réussie : {args.qif_file} -> {args.excel_file}")

if __name__ == "__main__":
    main()
