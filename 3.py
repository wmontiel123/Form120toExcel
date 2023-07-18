import re

def encontrar_decimales(expresion):
    # Expresión regular para encontrar números decimales
    patron = r'\d+\.\d+'
    
    # Buscar todos los números decimales en la expresión
    decimales = re.findall(patron, expresion)
    
    # Reemplazar los decimales por "DD" en la expresión
    expresion_modificada = re.sub(patron, "DD", expresion)
    
    return decimales, expresion_modificada

# Ejemplo de uso
expresion = "2.5 + 3 - 4 * 4.8 / 2.1"
decimales, expresion_modificada = encontrar_decimales(expresion)

print("Números decimales encontrados:", decimales)
print("Expresión modificada:", expresion_modificada)
