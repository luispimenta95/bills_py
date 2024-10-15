import pdfplumber
import pandas as pd
import os


def extract_all_tables_from_pdf(pdf_path):
    """
    Extrai todas as tabelas de um PDF e retorna uma lista de DataFrames.
    """
    all_tables = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages):
            print(f"Extraindo tabelas da página {page_number + 1}")

            tables = page.extract_tables()

            for table in tables:
                # Converta a tabela para um DataFrame do pandas
                df = pd.DataFrame(table[1:], columns=table[0])

                # Trate possíveis quebras de linha nos dados
                df = df.applymap(lambda x: x.replace('\n', ' ') if isinstance(x, str) else x)

                # Adiciona a tabela à lista de todas as tabelas
                all_tables.append(df)

    return all_tables


def combine_tables(tables):
    """
    Combina uma lista de DataFrames, preservando a primeira linha da segunda tabela
    como continuação da primeira, e combinando o restante das tabelas normalmente.
    """
    if len(tables) < 2:
        raise ValueError("A lista de tabelas deve conter pelo menos duas tabelas para combinar.")

    # DataFrame da primeira tabela
    df = tables[0]

    # DataFrame da segunda tabela
    df2 = tables[1]

    # Verifique se df2 tem cabeçalhos e remova-os se necessário
    df2.columns = df.columns

    # Preserve a primeira linha de df2 como uma continuação de df
    if not df.empty and not df2.empty:
        # Adiciona a primeira linha de df2 a df
        df = pd.concat([df, df2.iloc[[0]]], ignore_index=True)

        # Adiciona o restante das linhas de df2 a df
        df_combined = pd.concat([df, df2.iloc[1:]], ignore_index=True)
    else:
        # Se df2 estiver vazio, apenas use df
        df_combined = df

    # Se houver mais tabelas, combine-as também
    for table in tables[2:]:
        df_combined = pd.concat([df_combined, table], ignore_index=True)

    return df_combined


# Caminho para o seu PDF
pdf_path = 'doc.pdf'

# Extraia todas as tabelas
all_tables = extract_all_tables_from_pdf(pdf_path)

# Caminho para o arquivo Excel de saída
excel_path = 'teste.xlsx'

# Remover o arquivo existente, se necessário
if os.path.isfile(excel_path):
    os.remove(excel_path)

# Combine as tabelas extraídas
df_combined = combine_tables(all_tables)

# Salve o DataFrame combinado em um arquivo Excel
df_combined.to_excel(excel_path, index=False)
