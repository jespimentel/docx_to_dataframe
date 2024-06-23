import pandas as pd
import glob
import os
from docx import Document
import re

class DocxParaTexto:
    def __init__(self, caminho_para_arquivo_docx):
        self.caminho_para_arquivo_docx = caminho_para_arquivo_docx
        self.texto = ''

    def extrai_texto_de_docx(self):
        """ Retorna o texto extraído de um arquivo Word (docx). """
        try:
            documento = Document(self.caminho_para_arquivo_docx)
            for paragrafo in documento.paragraphs:
                self.texto += paragrafo.text + "\n"
            return self.texto
        except Exception as e:
            texto = f"Error reading DOCX file {self.caminho_para_arquivo_docx}: {e}\n"
        return texto

    def limpa_texto(self):
        """ Limpa o texto, substituindo quebras de linhas por espaços e múltiplos espaços por um único espaço """
        if self.texto:
            self.texto = self.texto.replace('\n', ' ')
            self.texto = self.texto.replace('\t', ' ')
            self.texto = re.sub(r'\s{2,}', ' ', self.texto)
            return self.texto
        else:
            print("O método 'extrai_texto_de_docx()' ainda não foi executado")

def relaciona_arquivos_docx(diretorio):
    padrao = os.path.join(diretorio, '**', '*.docx')
    todos_arquivos_docx = glob.glob(padrao, recursive=True)
    return todos_arquivos_docx

if __name__ == '__main__':
    pasta = input('Entre com o caminho para a pasta: ')
    arquivos_docx = relaciona_arquivos_docx(pasta)

    caminho_e_conteudo = []

    for arquivo in arquivos_docx:
        extrator = DocxParaTexto(arquivo)
        extrator.extrai_texto_de_docx()
        conteudo = extrator.limpa_texto()

        caminho_e_conteudo.append((arquivo, conteudo))

    # Convertendo para DataFrame e salvando como CSV
    df = pd.DataFrame(caminho_e_conteudo, columns=['caminho', 'conteudo'])
    df.to_csv('caminho_e_conteudo.csv', index=False, encoding='utf-8')