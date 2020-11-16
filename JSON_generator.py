import docx2txt, json

print("Seleccione el tipo de archivo que desea procesar:")
print("1.- Programa de Asignatura")
print("2.- Programa de Diplomado")
file_type = input()

print("Ingrese el nombre del archivo:")
file_name = raw_input()
file_path = "../Documentos/" + file_name + ".docx"

text = docx2txt.process(file_path)
myList = [item for item in text.split('\n')]

if file_type == 1:
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
  
  j = 8
  rango = 14
  title = asignatura

elif file_type == 2:
  programa = myList[2]
  print("Procesando: " + programa)
  
  sections = ["Fecha de inicio",
              "Fecha de termino",
              "Duracion",
              "Idioma",
              "Modalidad",
              "Descripcion",
              "Objetivos",
              "Perfil del participante",
              "Metodologia",
              "Docentes",
              "Link Programa"]

  titles = ["inicio",
            "termino",
            "duracion",
            "idioma",
            "modalidad",
            "descripcion",
            "objetivos",
            "perfil",
            "metodologia",
            "docentes",
            "link"]

  messages = ["La fecha de inicio del programa " + programa + " es el ",
              "La fecha de termino del programa " + programa + " es el ",
              "La duracion del programa " + programa + " es ",
              "El idioma del programa " + programa + ": ",
              "La modalidad del programa " + programa + ": ",
              "Descripcion del programa " + programa + ": ",
              "Objetivos del programa " + programa + ": ",
              "Perfil del programa " + programa + ": ",
              "Metodologia del programa " + programa + ": ",
              "Docentes del programa " + programa + ": ",
              "Link del programa " + programa + ": "]

  j = 4
  rango = 11
  title = programa

position = 0

print("Generando archivos JSON....")
for i in range(rango):
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
    "title": title,
    "text": messages[position] + content 
  }
  y = json.dumps(x)
  json_name = "../Data Discovery/" + myList[2] + "_" + titles[position] + ".json"
  position = position + 1
  with open(json_name, 'w') as f:
    f.write(y)
  f.close()