import openpyxl
from PIL import Image, ImageDraw, ImageFont

# Abrir a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Folha']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2, max_row = 2)):
    # cada célula que contém a info que precisamos
    nome_curso = linha[0].value # Nome do curso
    nome_participante = linha[1].value # nome do participante
    bilhete = linha[2].value # tipo de participação
    carga_horaria = linha[5].value # carga horária
    
    data_inicio = linha[3].value # data início
    data_final = linha[4].value # data início
    nome_professor = linha[6].value # data emissão
    #bilhete = linha[7].value
    media = linha[7].value
  
    # Transferir os dados da planilha para a imagem do certificado
    # Definindo a fonte a ser usada
    fonte_nome = ImageFont.truetype('./ITCBLKAD.TTF',85)
    fonte_geral = ImageFont.truetype('./BRUSHSCI.TTF',85)
    fonte_data = ImageFont.truetype('./CENTAUR.TTF',70)
    
    image = Image.open('./2.png')
    desenhar = ImageDraw.Draw(image)
    
    desenhar.text((600,760), nome_participante,fill='indigo',font=fonte_nome)
    desenhar.text((450,899),nome_curso, fill='indigo',font=fonte_geral)
    desenhar.text((850,846),bilhete, fill='indigo',font=fonte_data)
    desenhar.text((1820, 900),str(carga_horaria),fill='indigo',font=fonte_geral)
    desenhar.text((800, 975),str(media),fill='indigo',font=fonte_geral)
    
    desenhar.text((460, 1055),data_inicio,fill='indigo',font=fonte_data)
    desenhar.text((1560, 1055),data_final,fill='indigo',font=fonte_data)
    
    desenhar.text((1350, 1190),nome_professor,fill='indigo',font=fonte_data)

    image.save(f'./{indice} {nome_participante} certificado.png')