import datetime

mydate = datetime.datetime.now()
ano = mydate.strftime('%Y')
data = int(mydate.strftime('%m'))
dia = int(mydate.strftime('%d'))

if (dia >= 10): data = data + 1
if (data == 13): data = 1

def month():
  if (data == 1): mes = 'JANEIRO'
  elif (data == 2): mes = 'FEVEREIRO'
  elif (data == 3): mes = 'MARÇO'
  elif (data == 4): mes = 'ABRIL'
  elif (data == 5): mes = 'MAIO'
  elif (data == 6): mes = 'JUNHO'
  elif (data == 7): mes = 'JULHO'
  elif (data == 8): mes = 'AGOSTO'
  elif (data == 9): mes = 'SETEMBRO'
  elif (data == 10): mes = 'OUTUBRO'
  elif (data == 11): mes = 'NOVEMBRO'
  elif (data == 12): mes = 'DEZEMBRO'

  if (data >= 1 or data <= 6):
    semester = '2º Semestre ' + ano
  else:
    semester = '1º Semestre ' + ano

  return mes, ano, semester