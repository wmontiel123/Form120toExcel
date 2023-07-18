""" def calculator(expression):
    try:
        result = eval(expression)
        return result
    except Exception:
        return "Invalid expression"


# Example usage
expression = input("Enter an expression: ")
result = calculator(expression)
print("Result:", result)
 """
import re
import openpyxl


def encontrar_decimales(expresion, decimales):
    # Reemplazar "DD" por los números decimales encontrados
    for decimal in decimales:
        expresion = expresion.replace("DD", decimal, 1)
    
    return expresion

def evaluate_expression(expression, dictionary):
    patron = r'\d+\.\d+'
    
    # Buscar todos los números decimales en la expresión
    decimales = re.findall(patron, expression)
    
    # Reemplazar los decimales por "DD" en la expresión
    expresion_modificada = re.sub(patron, "DD", expression)
    
    for key, value in dictionary.items():
        expresion_modificada = expresion_modificada.replace(key, str(value))
        
        
    expression = encontrar_decimales(expresion_modificada,decimales)
        

    return eval(expression)


# Abre el archivo de Excel
workbook = openpyxl.load_workbook("C:\\Users\\Walter Montiel\\Desktop\\GT CO\\Nuevos\\Automatizacion Kaika.xlsx")

# Obtiene las hojas de cálculo
sheet1 = workbook['Balance']
sheet2 = workbook['Renglon']
sheet3 = workbook['Formulas']

# Diccionario para almacenar los grupos y sus sumas
grupos = {}

# Calcula la suma de cada grupo en la pestaña 2
for row in sheet2.iter_rows(min_row=2,values_only=True):

    grupo = row[2]
    monto = row[3]
    
    
    
    if grupo in grupos:
        grupos[grupo] += monto
    else:
        grupos[grupo] = monto


print(grupos)

expression_str = input("Enter the expression: ")
value_dict = {
    '1': 100,
    '2': 5,
    '3': 7,
    '4': 3
}

result = evaluate_expression(expression_str, value_dict)
print("Result:", result)