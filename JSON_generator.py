import docx2txt, json

text = docx2txt.process("plantilla.docx")
myList = [item for item in text.split('\n')]

asignatura = myList[2] + "-" + myList[6]
print("Procesando: " + asignatura)

titles = ["inicio",
          "termino",
          "modalidad",
          "nivel",
          "horas",
          "descripcion",
          "proposito",
          "objetivos",
          "unidades",
          "calendario",
          "docentes",
          "evaluaciones",
          "RA",
          "link"]

messages = ["La fecha de inicio de la asignatura " + asignatura + " es el ",
            "La fecha de termino de la asignatura " + asignatura + "es el ",
            "La modalidad de la asignatura " + asignatura + " es ",
            "Nivel de la asignatura " + asignatura + ": ",
            "Cantidad de horas de la asignatura " + asignatura + ": ",
            "Descripcion de la asignatura " + asignatura + ": ",
            "Proposito de la asignatura " + asignatura + ": ",
            "Objetivos de la asignatura " + asignatura + ": ",
            "Unidades de la asignatura " + asignatura + ": ",
            "Calendario de la asignatura " + asignatura + ": ",
            "Docentes de la asignatura " + asignatura + ": ",
            "Evaluaciones de la asignatura " + asignatura + ": ",
            "Resultados de Aprendizaje de la asignatura " + asignatura + ": ",
            "Link de la Sala Virtual de la asignatura " + asignatura + ": "]

position = 0

for i in range(8,len(myList),4):
  print(myList[i])
  x = {
    "title": asignatura,
    "text": messages[position] + myList[i+2] 
  }
  y = json.dumps(x)
  filename = myList[2] + "_" + titles[position] + ".json"
  position = position + 1
  with open(filename, 'w') as f:
    f.write(y)
  f.close()