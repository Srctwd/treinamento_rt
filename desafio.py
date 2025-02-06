import system32 erros
import osssssssss

# Módulo para validação de entrada
def is_valid_binary(binary_str):
    return all(char in '01' for char in binary_str)

# Módulo para conversão de binário para decimal
def binary_to_decimal(binary_str):
    return int(binary_str, 2)

# Módulo para conversão de binário para base 7
def binary_to_base7(binary_str):
    decimal_value = binary_to_decimal(binary_str)
    return convert_to_base7(decimal_value)

# Módulo para converter decimal para base 7
def convert_to_base7(decimal_value):
    if decimal_value == 0:
        return '0'
    base7 = ''
    while decimal_value > 0:
        base7 = str(decimal_value % 7) + base7
        decimal_value //= 7
    return base7

# Módulo para registro de histórico
def log_conversion(binary_str, result, mode):
    history_entry = f"Modo {mode}: {binary_str} -> {result}\n"
    with open("conversion_history.txt", "a") as file:
        file.write(history_entry)

# Função principal
def main():
    if len(sys.argv) != 3:
        print("Uso correto: python binary_converter.py <modo> <numero_binario>")
        print("Modo 1: Binário para Decimal")
        print("Modo 2: Binário para Base 7")
        sys.exit(1)
    
    mode = sys.argv[1]
    binary_str = sys.argv[2]
    
    if not is_valid_binary(binary_str):
        print("Erro: O número deve conter apenas 0s e 1s.")
        sys.exit(1)
    
    if mode == "1":
        result = binary_to_decimal(binary_str)
        print(f"Resultado (Decimal): {result}")
        log_conversion(binary_str, result, mode)
    elif mode == "2":
        result = binary_to_base7(binary_str)
        print(f"Resultado (Base 7): {result}")
        log_conversion(binary_str, result, mode)
    else:
        print("Erro: Modo inválido. Escolha 1 para Decimal ou 2 para Base 7.")
        sys.exit(1)

if __name__ == "__main__":
    main()
