import os
import xml.etree.ElementTree as ET

def generate_insert_statements_from_xml(file_path):
    ns = {'n': 'http://www.portalfiscal.inf.br/nfe'}
    tree = ET.parse(file_path)
    root = tree.getroot()
    unique_eans = set()
    insert_commands = []

    for det in root.findall(".//n:det/n:prod", namespaces=ns):
        cEAN = det.find("n:cEAN", namespaces=ns).text.zfill(14)
        if cEAN not in unique_eans:
            unique_eans.add(cEAN)
            xProd = det.find("n:xProd", namespaces=ns).text
            uTrib = det.find("n:uTrib", namespaces=ns).text
            vUnTrib = det.find("n:vUnTrib", namespaces=ns).text
            ncm = det.find("n:NCM", namespaces=ns).text
            cest = det.find("n:CEST", namespaces=ns).text if det.find("n:CEST", namespaces=ns) is not None else '0000000'
            xProdShort = xProd[:20]
            gencode = ncm[:2]

            sql = f"""
            INSERT INTO PRODUTO (
                PROCOD, PRODES, PRODESRDZ, SECCOD, PROUNID, PROUNDCMP, PROUNDFUN, PROPESVAR,
                PROPRC1, PROPRCVDAVAR, PROPRCCST, PROPRCCSTMED, PROFLGALT, PROTABA, FUNCOD, PROITEEMB,
                PROFORLIN, PROIAT, PROIPPT, PROFIN, PRONCM, PROINCZF, PROITEEMBVDA, PROCEST, PRODESON, GENCODIGO
            ) VALUES (
                '{cEAN}', '{xProd}', '{xProdShort}', '00', '{uTrib}', '{uTrib}', '{uTrib}', 'N',
                '{vUnTrib}', '{vUnTrib}', '{vUnTrib}', '{vUnTrib}', 'T', '0', '000001', '1',
                'N', 'T', 'T', '1', '{ncm}', 'N', '1', '{cest}', 'N', '{gencode}'
            );
            """
            insert_commands.append(sql.strip())

    return insert_commands

def process_all_xmls_in_directory(directory_path):
    xml_files = [os.path.join(directory_path, f) for f in os.listdir(directory_path) if f.endswith('.xml')]
    all_inserts = []
    for xml_file in xml_files:
        all_inserts.extend(generate_insert_statements_from_xml(xml_file))
    
    # Saving to a .sql file in the same directory
    output_file_path = os.path.join(directory_path, 'insert_commands.sql')
    with open(output_file_path, 'w') as f:
        for command in all_inserts:
            f.write(command + "\n")
    return output_file_path

# Usage
directory_path = 'D:\\XMLS'
sql_file_path = process_all_xmls_in_directory(directory_path)
print(f"All INSERT commands have been saved to {sql_file_path}")
