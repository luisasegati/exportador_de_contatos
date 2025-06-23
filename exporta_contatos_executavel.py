import pandas as pd
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Conversor VCF para Excel")
        self.root.minsize(450, 350)
        self.root.resizable(True, True)

        self.caminho_vcf = None
        self.pasta_saida = None

        # Label e botão para selecionar arquivo VCF
        tk.Label(root, text="Clique para selecionar um arquivo .vcf:", font=("Arial", 12)).pack(pady=10, fill='x')
        tk.Button(root, text="Selecionar Arquivo VCF", command=self.selecionar_arquivo, width=25, height=2).pack()
        self.label_arquivo = tk.Label(root, text="", wraplength=420, justify="center")
        self.label_arquivo.pack(pady=10, fill='x')

        # Botão para selecionar pasta onde salvar
        tk.Label(root, text="Selecione onde salvar a planilha Excel:", font=("Arial", 12)).pack(pady=10, fill='x')
        tk.Button(root, text="Selecionar Pasta de Saída", command=self.selecionar_pasta, width=25, height=2).pack()
        self.label_pasta = tk.Label(root, text="", wraplength=420, justify="center")
        self.label_pasta.pack(pady=10, fill='x')

        # Botão para executar a conversão
        self.botao_converter = tk.Button(root, text="Converter e Salvar", command=self.converter, width=25, height=2)
        self.botao_converter.pack(pady=15)
        self.botao_converter.config(state="disabled")

    def extrair_contatos_vcf(self, caminho_vcf):
        nomes = []
        telefones = []

        try:
            with open(caminho_vcf, 'r', encoding='utf-8') as arquivo:
                nome_atual = None
                telefone_atual = None

                for linha in arquivo:
                    linha = linha.strip()

                    if linha.startswith('FN:'):
                        nome_atual = linha[3:]

                    if linha.startswith('TEL'):
                        match = re.search(r'[:](.+)', linha)
                        if match:
                            telefone_atual = match.group(1)

                    if linha == 'END:VCARD':
                        if nome_atual and telefone_atual:
                            nomes.append(nome_atual)
                            telefones.append(telefone_atual)
                        nome_atual = None
                        telefone_atual = None

            df = pd.DataFrame({
                'Nome': nomes,
                'Telefone': telefones
            })

            return df

        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro ao ler o arquivo VCF: {e}")
            return None

    def salvar_em_excel(self, df, caminho_saida):
        try:
            df.to_excel(caminho_saida, index=False)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo Excel: {e}")

    def selecionar_arquivo(self):
        caminho = filedialog.askopenfilename(title="Selecione o arquivo VCF", filetypes=[("VCF files", "*.vcf")])
        if caminho:
            self.caminho_vcf = caminho
            self.label_arquivo.config(text=f"Arquivo selecionado:\n{self.caminho_vcf}")
            self.atualizar_estado_botao()

    def selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta para salvar a planilha")
        if pasta:
            self.pasta_saida = pasta
            self.label_pasta.config(text=f"Pasta selecionada:\n{self.pasta_saida}")
            self.atualizar_estado_botao()

    def atualizar_estado_botao(self):
        if self.caminho_vcf and self.pasta_saida:
            self.botao_converter.config(state="normal")
        else:
            self.botao_converter.config(state="disabled")

    def converter(self):
        df_contatos = self.extrair_contatos_vcf(self.caminho_vcf)

        if df_contatos is not None and not df_contatos.empty:
            caminho_saida = os.path.join(self.pasta_saida, 'contatos.xlsx')
            self.salvar_em_excel(df_contatos, caminho_saida)
            messagebox.showinfo("Sucesso", f"Planilha salva em:\n{caminho_saida}")
        else:
            messagebox.showinfo("Aviso", "Nenhum contato válido encontrado no arquivo.")

def main():
    root = tk.Tk()
    app = App(root)
    root.mainloop()

if __name__ == "__main__":
    main()
