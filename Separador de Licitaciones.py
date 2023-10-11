import win32com.client
from bs4 import BeautifulSoup
import re

# Crear una instancia de Outlook
outlook = win32com.client.Dispatch("Outlook.Application")

# Obtener la bandeja de entrada
inbox = outlook.GetNamespace("MAPI").GetDefaultFolder(6)  # 6 representa la carpeta de la bandeja de entrada

# Iterar a través de los correos electrónicos no leídos en la bandeja de entrada
for message in inbox.Items:
    if message.UnRead:
            
        # Verificar si el correo tiene el asunto deseado
        # Utilizar expresiones regulares para buscar una tabla con los campos deseados
        body = message.Body
  
        pattern = r'ID LICITACION\s+CONVOCANTE\s+TIPO LICITACION\s+PROYECTO[\s\S]*?(?:(?:\n+\t*){2,}|\n*$)'
        match = re.search(pattern, body, re.IGNORECASE)
        print(match)
        if match:
        
            # Obtener el texto de la tabla
            table_text = match.group()

            # Dividir el texto de la tabla en filas
            rows = table_text.split('\n')

            # Inicializar diccionarios para cada tipo de licitación
            licitacion_obra = []
            licitacion_equipos = []
            licitacion_consultoria = []

            # Iterar a través de las filas de la tabla
            for row in rows:
                cells = row.split('\t')
                if len(cells) == 4:
                    id_licitacion, convocante, tipo_licitacion, proyecto = cells
                  
                    print(tipo_licitacion)
                    # Determinar el tipo de licitación y agregar la fila al diccionario correspondiente
                    if "obras civiles" in tipo_licitacion:
                        licitacion_obra.append((id_licitacion, convocante, proyecto))
                    elif "suministro de equipos" in tipo_licitacion:
                        licitacion_equipos.append((id_licitacion, convocante, proyecto))
                    elif "servicio consultoría" in tipo_licitacion:
                        licitacion_consultoria.append((id_licitacion, convocante, proyecto))

            # Imprimir las filas por tipo de licitación
            if licitacion_obra:
                print("Licitaciones de Obras Civiles:")
                for row in licitacion_obra:
                    print(row)

            if licitacion_equipos:
                print("\nLicitaciones de Suministro de Equipos:")
                for row in licitacion_equipos:
                    print(row)

            if licitacion_consultoria:
                print("\nLicitaciones de Servicios de Consultoría:")
                for row in licitacion_consultoria:
                    print(row)

        # Marcar el correo como leído
        message.UnRead = False
        message.Save()