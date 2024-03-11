import pandas as pd
import os
import glob

# Varíavel com caminho dos arquivos
folder_path = 'src\\data\\raw'

# Lista com todos os arquivos de Excel
excel_files = glob.glob(os.path.join(folder_path, '*.xlsx'))

if not excel_files:
    print('nenhum arquivo compátivel encontrado')
else:
    # Tabela na memória para guardar o conteúdo dos arquivos
    dfs = []
    for excel_file in excel_files:
        try:
            # leio o arquivo de Excel
            df_temp = pd.read_excel(excel_file)
            # pego o nome do arquivo
            file_name = os.path.basename(excel_file)
            # pego o nome do arquivo de origem dos dados
            df_temp['filename'] = file_name

            # criar uma nova coluna chamada 'location'
            if 'brasil' in file_name:
                df_temp['location'] = 'br'
            elif 'france' in file_name.lower():
                df_temp['location'] = 'fr'
            elif 'italian' in file_name.lower():
                df_temp['location'] = 'it'
            # criar uma nova coluna chamada 'campaign'
            df_temp['campaign'] = df_temp['utm_link'].str.extract(r'utm_campaign=(.*)')
            # guarda dados tratados dentro de uma dataframe comum
            dfs.append(df_temp)
            
        except Exception as e:
            print(f"Erro ao ler arquivo {excel_file} : {e}")

    if dfs:
        # junta todas as tabelas em "dfs" em uma única tabela
        result = pd.concat(dfs, ignore_index=True)
        # caminho de saída
        output_file = os.path.join('src', 'data', 'ready', 'clean.xlsx')
        # configura o motor de escrita
        writer = pd.ExcelWriter(output_file, engine = 'xlsxwriter')
        # leva os dados do resultado a serem escritos no motor de Excel configurado 
        result.to_excel(writer, index=False)

        # salva o arquivo de Excel
        writer._save()
    else:
        print('Nenhum dado para ser salvo')