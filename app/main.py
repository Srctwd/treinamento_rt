
# código principal

# Treinamento RT 06/02/2025

# Desenvolva uma aplicação em Python com os seguintes requisitos:

# Modos de operação: O usuário seleciona um dos dois modos através de argumentos do sistema.
# Modo 1: Converte um número binário para base 10.
# Modo 2: Converte um número binário para base 7.
# Estrutura modular: O código deve ser organizado em módulos.
# Divisão de tarefas: Cada participante deve implementar pelo menos uma funcionalidade.
# Registro de histórico: A aplicação deve armazenar um histórico das conversões realizadas e dos valores inseridos pelo usuário.
# Validação de entrada: Deve haver tratamento para garantir que o usuário insira valores binários válidos.
from utils.math import calc_base10

print(f"Select the entry mode ")
print(f"1 - Convert the binary number to base 10.")
print(f"2 - Convert the binary number to base 7.")

mode = input("Entry the mode: ")

if mode == "1":
    while True:
        value = input("insert the value to be converted: ")
        result = calc_base10(value)
        if (result is not False):
            break # quebra o loop
    print(f"Final Result: {result}")

if mode == "2":
    while True:
        value = input("insert the value to be converted: ")
        result = calc_base10(value)
        if (result is not False):
            result = calc_base7(result)
            break # quebra o loop
        
    print(f"Final Result: {result}")
else:
    print("Error, mode dont identified")
