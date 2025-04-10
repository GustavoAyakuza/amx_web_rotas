import pandas as pd
import re
import os

def limpar_texto(texto):
    if pd.isna(texto):
        return ""
    return str(texto).strip()

def process_excel(input_path, output_saida_controle, output_rts):
    df = pd.read_excel(input_path, header=None, skiprows=1)

    registros = []
    romaneio = ''
    cliente = ''
    endereco = ''
    volume = ''
    aguardando_volume = False

    for i in range(len(df)):
        linha = df.iloc[i]
        texto_linha = linha.astype(str)

        for idx, celula in enumerate(texto_linha):
            celula = limpar_texto(celula)

            # Romaneio
            if "Romaneio:" in celula:
                match = re.search(r'Romaneio:\s*([0-9]{2}\.[0-9]{3})', celula)
                if match:
                    romaneio = match.group(1)

            # Endereço
            if "Endereço:" in celula:
                endereco = celula.split("Endereço:")[1].strip()
                endereco = endereco.replace("Sinal:", "").strip()

            # Cliente
            if "Cliente:" in celula:
                cliente = celula.split("Cliente:")[1].strip()

            # Quantidade → pega todos os valores abaixo, na mesma coluna
            if "Quantidade" in celula:
                volume_total = 0
                col_index = idx
                linha_check = i + 1

                while linha_check < len(df):
                    valor_celula = df.iloc[linha_check, col_index]
                    valor_str = limpar_texto(valor_celula).replace(".", "").replace(",", ".")

                    # Se for vazio ou texto irrelevante, parar
                    if valor_str == "" or not re.match(r"^\d+(\.\d+)?$", valor_str):
                        break

                    try:
                        volume_total += int(float(valor_str))
                    except:
                        break

                    linha_check += 1

                volume = volume_total

                # Salvar registro se completo
                if romaneio and cliente and endereco:
                    registros.append({
                        'NºRota': '',
                        'Resp.': '',
                        'NºRomaneio': romaneio,
                        'Cliente': cliente,
                        'Endereço': endereco,
                        'VOL': volume,
                        'OBS': ''
                    })

                    romaneio = ''
                    cliente = ''
                    endereco = ''
                    volume = ''
                continue

    # Saída Controle
    df_saida = pd.DataFrame(registros, columns=['NºRota', 'Resp.', 'NºRomaneio', 'Cliente', 'Endereço', 'VOL', 'OBS'])
    df_saida.to_excel(output_saida_controle, index=False, engine='openpyxl')

    # RTS com 14 colunas: Endereço (coluna A), Romaneio (coluna N)
    rts_data = []
    for reg in registros:
        linha = [''] * 14
        linha[0] = reg['Endereço']
        linha[13] = reg['NºRomaneio']
        rts_data.append(linha)

    df_rts = pd.DataFrame(rts_data)
    df_rts.to_excel(output_rts, index=False, header=False, engine='openpyxl')

    print("✅ Arquivos gerados com sucesso:")
    print(f"→ Saída Controle: {output_saida_controle}")
    print(f"→ RTS: {output_rts}")
