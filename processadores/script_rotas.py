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

        for celula in texto_linha:
            celula = limpar_texto(celula)

            # Romaneio: buscar padrão 72.857 etc.
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

            # Quantidade: ativa leitura da próxima célula como volume
            if "Quantidade" in celula:
                aguardando_volume = True
                continue

            # Volume (logo após "Quantidade")
            if aguardando_volume:
                volume_raw = celula.replace(".", "").replace(",", ".")  # trata 3,000 ou 4,000
                try:
                    volume = int(float(volume_raw))
                except:
                    volume = 0  # fallback caso haja erro
                aguardando_volume = False

                # Salvar registro quando todos os dados estiverem presentes
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

                    # Reset
                    romaneio = ''
                    cliente = ''
                    endereco = ''
                    volume = ''
                break

    # Saída Controle
    df_saida = pd.DataFrame(registros, columns=['NºRota', 'Resp.', 'NºRomaneio', 'Cliente', 'Endereço', 'VOL', 'OBS'])
    df_saida.to_excel(output_saida_controle, index=False, engine='openpyxl')

    # RTS com 14 colunas: Endereço (coluna A), Romaneio (coluna N)
    rts_data = []
    for reg in registros:
        linha = [''] * 14
        linha[0] = reg['Endereço']     # Coluna A
        linha[13] = reg['NºRomaneio']  # Coluna N
        rts_data.append(linha)

    df_rts = pd.DataFrame(rts_data)
    df_rts.to_excel(output_rts, index=False, header=False, engine='openpyxl')  # Sem cabeçalhos

    print("✅ Arquivos gerados com sucesso:")
    print(f"→ Saída Controle: {output_saida_controle}")
    print(f"→ RTS: {output_rts}")
