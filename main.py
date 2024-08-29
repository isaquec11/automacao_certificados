
# pegar os dados da planilha
import openpyxl
from PIL import Image, ImageDraw, ImageFont

planilha_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
pagina_alunos = planilha_alunos['Sheet1']

for indice, linha in enumerate(pagina_alunos.iter_rows(min_row=2)):
    # armazenar as informacoes de cada celula
    nome_curso = linha[0].value
    nome_participante = linha[1].value
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_termino = linha[4].value
    carga_horario = linha[5].value
    data_emissao = linha[6].value

    # transferir os dados da planilha para o certificado
    fonte_nome = ImageFont.truetype('./tahomabd.ttf', 90)
    fonte_data = ImageFont.truetype('./tahoma.ttf', 55)
    fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)

    # abrir a imagem do certificado
    imagem = Image.open('./certificado_padrao.jpg')
    desenhar = ImageDraw.Draw(imagem)

    # escrever as informações no certificado
    desenhar.text((1020,827), nome_participante, fill='black', font=fonte_nome)
    desenhar.text((1060,950), nome_curso, fill='black', font=fonte_geral)
    desenhar.text((1435,1065), tipo_participacao, fill='black', font=fonte_geral)
    desenhar.text((1480, 1182), f'{str(carga_horario)} horas', fill='black', font=fonte_geral)

    desenhar.text((750, 1770), data_inicio, fill='black', font=fonte_data)
    desenhar.text((750, 1930), data_termino, fill='black', font=fonte_data)

    desenhar.text((2220, 1930), data_emissao, fill='black', font=fonte_data)


    # transferir para a imagem do certificado   

    imagem.save(f'./certificados_prontos/ {nome_participante} - certificado.png')


