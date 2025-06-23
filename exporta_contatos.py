import pandas as pd
import re
import os

def extrair_contatos_vcf(caminho_vcf):
    nomes = []
    telefones = []

    try:
        with open(caminho_vcf, 'r', encoding='utf-8') as arquivo:
            nome_atual = None
            telefone_atual = None

            for linha in arquivo:
                linha = linha.strip()

                # Captura nome
                if linha.startswith('FN:'):
                    nome_atual = linha[3:]

                # Captura telefone (primeiro telefone encontrado)
                if linha.startswith('TEL'):
                    match = re.search(r'[:](.+)', linha)
                    if match:
                        telefone_atual = match.group(1)

                # Fim do contato
                if linha == 'END:VCARD':
                    if nome_atual and telefone_atual:
                        nomes.append(nome_atual)
                        telefones.append(telefone_atual)
                    nome_atual = None
                    telefone_atual = None

        # Monta DataFrame com nome e telefone
        df = pd.DataFrame({
            'Nome': nomes,
            'Telefone': telefones
        })

        return df

    except FileNotFoundError:
        print(f"Erro: Arquivo '{caminho_vcf}' não encontrado.")
        return None
    except Exception as e:
        print(f"Ocorreu um erro ao ler o arquivo VCF: {e}")
        return None

def salvar_em_excel(df, caminho_saida):
    try:
        if os.path.exists(caminho_saida):
            print(f"A planilha '{caminho_saida}' já existe. Não será sobrescrita.")
            return
        df.to_excel(caminho_saida, index=False)
        print(f"Planilha salva com sucesso em: {caminho_saida}")
    except Exception as e:
        print(f"Erro ao salvar o arquivo Excel: {e}")

def main():
    '''
    Define os caminhos de entrada e saída, verifica a existência do arquivo VCF,
    extrai os contatos e salva em uma planilha Excel (caso ela ainda não exista).
    '''
    caminho_vcf = 'Contatos.vcf' 
    caminho_saida = 'contatos.xlsx' 

    # Verifica se o arquivo VCF existe
    if not os.path.exists(caminho_vcf):
        print(f"Arquivo '{caminho_vcf}' não encontrado no diretório atual.")
        return

    # Extrai contatos
    df_contatos = extrair_contatos_vcf(caminho_vcf)

    if df_contatos is not None and not df_contatos.empty:
        salvar_em_excel(df_contatos, caminho_saida)
    else:
        print("Nenhum contato válido foi encontrado no arquivo.")

if __name__ == '__main__':
    main()
