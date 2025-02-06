def calc_base10(value):
    try:
        decimal = int(value)
        return decimal
    except:
        return False
    
def calc_base7(value):
    if value == 0:
        return "0" 

    base7 = ""
    while value > 0:
        resto = value % 7
        base7 = str(resto) + base7 
        value //= 7 

    return base7 
