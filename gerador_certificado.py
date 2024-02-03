#### Importar Bibliotecas
import openpyxl #biblioteca para manipular planilhas
from PIL import Image, ImageDraw, ImageFont #biblioteca manipular imagem


### Abrir planilha
planilha_alunos = openpyxl.load_workbook('nomes.xlsx')
aba_planilha = planilha_alunos['Sheet1']

# Percorrer cada linha da planilha, exceto cabecalho - linha 1.
for linha in aba_planilha.iter_rows(min_row=2):

    # Pegar nome e idade do aluno
    nome_aluno = linha[0].value
    idade_aluno = str(int(linha[1].value))
    
    # Importar fonte usada
    fonte_nome = ImageFont.truetype('./fonts/SHOWG.TTF', 80)

    # Importar imagem do certificado
    imagem_certificado = Image.open('./certificado.png')

    # Adicionar nomes em cima da imagem
    desenhar_imagem = ImageDraw.Draw(imagem_certificado)
    desenhar_imagem.text((520,550), nome_aluno +' - '+idade_aluno+' anos', fill='black', font=fonte_nome)

    # Salvar imagem
    imagem_certificado.save(f'./certificados_gerados/{nome_aluno}_certificado.png')