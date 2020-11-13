import docx2txt, json

text = docx2txt.process("plantilla.docx")
myList = [item for item in text.split('\n')]

asignatura = myList[2] + "-" + myList[6]
print("Procesando: " + asignatura)

sections = ["Fecha de inicio",
            "Fecha de termino",
            "Modalidad",
            "Nivel",
            "Horas",
            "Descripcion",
            "Proposito",
            "Objetivos",
            "Unidades",
            "Calendario de actividades",
            "Docentes",
            "Evaluaciones",
            "Resultados de aprendizaje",
            "Link de Classroom"]

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
            "La fecha de termino de la asignatura " + asignatura + " es el ",
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

print("Generando archivos JSON....")

j = 8

for i in range(14):
  print(myList[j] + " -> OK")
  content = ""
  if myList[j] != sections[-1]:
    j += 1
    while myList[j] != sections[position + 1]:
      if myList[j] != "":
        content += myList[j] + "\n"
      j += 1
  else:
    j += 1
    while j < len(myList):
      if myList[j] != "":
        content += myList[j] + "\n"
      j += 1
  content = content.rstrip("\n")
  x = {
    "title": asignatura,
    "text": messages[position] + content 
  }
  y = json.dumps(x)
  filename = myList[2] + "_" + titles[position] + ".json"
  position = position + 1
  with open(filename, 'w') as f:
    f.write(y)
  f.close()