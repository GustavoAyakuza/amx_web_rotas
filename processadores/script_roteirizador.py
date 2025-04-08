import pandas as pd

def merge_routes(control_file, route_file, updated_file):
    # Lê a planilha de controle geral (RTS com volumes)
    control_df = pd.read_excel(control_file)

    # Lê a planilha de roteirização (do Zeo)
    route_df = pd.read_excel(route_file)

    # Seleciona apenas as colunas necessárias da planilha do Zeo
    route_df = route_df[['Número de série', 'Endereço']]

    # Faz a correspondência dos endereços
    merged_df = control_df.merge(route_df, how='left', on='Endereço')

    if 'Número de série' in merged_df.columns:
        merged_df['NºRota'] = merged_df['Número de série']
        merged_df.drop(columns=['Número de série'], inplace=True)

    # Salva o resultado
    merged_df.to_excel(updated_file, index=False, engine='openpyxl')
    print(f"Planilha final salva em: {updated_file}")
