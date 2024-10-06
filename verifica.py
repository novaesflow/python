import os
import xml.etree.ElementTree as ET

def read_existing_procods(file_path):
    """ Lê os PROCODs existentes de um arquivo de texto. """
    with open(file_path, 'r') as file:
        existing_procods = {line.strip().zfill(14) for line in file}
    return existing_procods

def generate_sql_commands(directory_path, existing_procods):
    """ Verifica os PROCODs nos arquivos XML e gera comandos SQL para novos PROCODs. """
    ns = {'n': 'http://www.portalfiscal.inf.br/nfe'}
    new_product_inserts = set()
    new_tax_inserts = set()
    cst_mapping = {
        '01': ('04', '08'),
        '06': ('01', '05'),
        '04': ('02', '06'),
        '05': ('03', '07')
    }

    xml_files = [os.path.join(directory_path, f) for f in os.listdir(directory_path) if f.endswith('.xml')]
    for xml_file in xml_files:
        tree = ET.parse(xml_file)
        root = tree.getroot()

        for det in root.findall(".//n:det", namespaces=ns):
            prod = det.find("n:prod", namespaces=ns)
            cProd = prod.find("n:cProd", namespaces=ns).text.zfill(14)
            if cProd not in existing_procods:
                xProd = prod.find("n:xProd", namespaces=ns).text
                uTrib = prod.find("n:uTrib", namespaces=ns).text
                vUnTrib = prod.find("n:vUnTrib", namespaces=ns).text
                ncm = prod.find("n:NCM", namespaces=ns).text
                cest = prod.find("n:CEST", namespaces=ns).text if prod.find("n:CEST", namespaces=ns) is not None else '0000000'
                xProdShort = xProd[:20]
                gencode = ncm[:2]

                product_sql = f"""
                INSERT INTO PRODUTO (
                    PROCOD, PRODES, PRODESRDZ, SECCOD, PROUNID, PROUNDCMP, PROUNDFUN, PROPESVAR,
                    PROPRC1, PROPRCVDAVAR, PROPRCCST, PROPRCCSTMED, PROFLGALT, PROTABA, FUNCOD, PROITEEMB,
                    PROFORLIN, PROIAT, PROIPPT, PROFIN, PRONCM, PROINCZF, PROITEEMBVDA, PROCEST, PRODESON, GENCODIGO, PROENVIAPDV
                ) VALUES (
                    '{cProd}', '{xProd}', '{xProdShort}', '00', '{uTrib}', '{uTrib}', '{uTrib}', 'N',
                    '{vUnTrib}', '{vUnTrib}', '{vUnTrib}', '{vUnTrib}', 'T', '0', '000001', '1',
                    'S', 'T', 'T', '1', '{ncm}', 'N', '1', '{cest}', 'N', '{gencode}', 'N'
                );
                """
                new_product_inserts.add(product_sql.strip())

                pis_cofins = det.find("n:imposto", namespaces=ns)
                if pis_cofins is not None:
                    for tax_type in ['PIS', 'COFINS']:
                        tax_detail = pis_cofins.find(f"n:{tax_type}/n:{tax_type}Aliq", namespaces=ns)
                        if tax_detail is not None:
                            cst = tax_detail.find("n:CST", namespaces=ns).text
                            if cst in cst_mapping:
                                for impfedsim in cst_mapping[cst]:
                                    tax_sql = f"INSERT INTO IMPOSTOS_FEDERAIS_PRODUTO (IMPFEDSIM, PROCOD) VALUES ('{impfedsim}', '{cProd}');"
                                    new_tax_inserts.add(tax_sql)

    return new_product_inserts, new_tax_inserts

def save_sql_commands(new_product_inserts, new_tax_inserts, directory_path):
    """ Salva comandos SQL em arquivos. """
    product_file_path = os.path.join(directory_path, 'new_products_insert_commands.sql')
    tax_file_path = os.path.join(directory_path, 'new_taxes_insert_commands.sql')

    with open(product_file_path, 'w') as f:
        for command in new_product_inserts:
            f.write(command + "\n")

    with open(tax_file_path, 'w') as f:
        for command in new_tax_inserts:
            f.write(command + "\n")

    return product_file_path, tax_file_path

def main():
    directory_path = 'D:\\XMLS'
    existing_procods_path = 'D:\\XMLS\\expdata.txt'
    existing_procods = read_existing_procods(existing_procods_path)
    new_product_inserts, new_tax_inserts = generate_sql_commands(directory_path, existing_procods)
    product_sql_file, tax_sql_file = save_sql_commands(new_product_inserts, new_tax_inserts, directory_path)

    print(f"All new product INSERT commands have been saved to {product_sql_file}")
    print(f"All new tax INSERT commands have been saved to {tax_sql_file}")

if __name__ == "__main__":
    main()
