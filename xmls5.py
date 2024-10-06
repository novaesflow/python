import os
import xml.etree.ElementTree as ET

def read_existing_procods(file_path):
    with open(file_path, 'r') as file:
        existing_procods = {line.strip() for line in file}
    return existing_procods

def generate_insert_statements_from_xml(file_path, existing_procods):
    ns = {'n': 'http://www.portalfiscal.inf.br/nfe'}
    tree = ET.parse(file_path)
    root = tree.getroot()
    unique_product_inserts = set()
    unique_tax_inserts = set()

    cst_mapping = {
        '01': ('04', '08'),
        '06': ('01', '05'),
        '04': ('02', '06'),
        '05': ('03', '07')
    }

    for det in root.findall(".//n:det", namespaces=ns):
        prod = det.find("n:prod", namespaces=ns)
        cEAN = prod.find("n:cEAN", namespaces=ns).text
        cProd = prod.find("n:cProd", namespaces=ns).text.zfill(14)
        if cEAN == "SEM GTIN":
            cEAN = cProd
        if cEAN not in existing_procods:
            xProd = prod.find("n:xProd", namespaces=ns).text
            uTrib = prod.find("n:uTrib", namespaces=ns).text
            vUnTrib = prod.find("n:vUnTrib", namespaces=ns).text
            ncm = prod.find("n:NCM", namespaces=ns).text
            cest = prod.find("n:CEST", namespaces=ns).text if prod.find("n:CEST", namespaces=ns) is not None else '0000000'
            xProdShort = xProd[:20]
            gencode = ncm[:2]

            product_sql = f"INSERT INTO PRODUTO (PROCOD, PRODES, PRODESRDZ, SECCOD, PROUNID, PROUNDCMP, PROUNDFUN, PROPESVAR, PROPRC1, PROPRCVDAVAR, PROPRCCST, PROPRCCSTMED, PROFLGALT, PROTABA, FUNCOD, PROITEEMB, PROFORLIN, PROIAT, PROIPPT, PROFIN, PRONCM, PROINCZF, PROITEEMBVDA, PROCEST, PRODESON, GENCODIGO, PROENVIAPDV) VALUES ('{cEAN}', '{xProd}', '{xProdShort}', '00', '{uTrib}', '{uTrib}', '{uTrib}', 'N', '{vUnTrib}', '{vUnTrib}', '{vUnTrib}', '{vUnTrib}', 'T', '0', '000001', '1', 'S', 'T', 'T', '1', '{ncm}', 'N', '1', '{cest}', 'N', '{gencode}', 'N');"
            unique_product_inserts.add(product_sql.strip())

            pis_cofins = det.find("n:imposto", namespaces=ns)
            if pis_cofins is not None:
                for tax_type in ['PIS', 'COFINS']:
                    tax_detail = pis_cofins.find(f"n:{tax_type}/n:{tax_type}Aliq", namespaces=ns)
                    if tax_detail is not None:
                        cst = tax_detail.find("n:CST", namespaces=ns).text
                        if cst in cst_mapping:
                            for impfedsim in cst_mapping[cst]:
                                tax_sql = f"INSERT INTO IMPOSTOS_FEDERAIS_PRODUTO (IMPFEDSIM, PROCOD) VALUES ('{impfedsim}', '{cEAN}');"
                                unique_tax_inserts.add(tax_sql)

    return unique_product_inserts, unique_tax_inserts

def process_all_xmls_in_directory(directory_path, existing_procods_path):
    existing_procods = read_existing_procods(existing_procods_path)
    xml_files = [os.path.join(directory_path, f) for f in os.listdir(directory_path) if f.endswith('.xml')]
    all_product_inserts = set()
    all_tax_inserts = set()
    for xml_file in xml_files:
        product_inserts, tax_inserts = generate_insert_statements_from_xml(xml_file, existing_procods)
        all_product_inserts.update(product_inserts)
        all_tax_inserts.update(tax_inserts)

    product_file_path = os.path.join(directory_path, 'product_insert_commands.sql')
    tax_file_path = os.path.join(directory_path, 'tax_insert_commands.sql')

    with open(product_file_path, 'w') as f:
        for command in all_product_inserts:
            f.write(command + "\n")

    with open(tax_file_path, 'w') as f:
        for command in all_tax_inserts:
            f.write(command + "\n")

    return product_file_path, tax_file_path

# Usage example
directory_path = 'D:\\XMLS'
existing_procods_path = 'D:\\XMLS\\expdata.txt'
product_sql_file, tax_sql_file = process_all_xmls_in_directory(directory_path, existing_procods_path)
print(f"All product INSERT commands have been saved to {product_sql_file}")
print(f"All tax INSERT commands have been saved to {tax_sql_file}")
