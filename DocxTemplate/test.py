from num2words import num2words
import math
#
# num = float(input("Enter number: "))
# n = math.trunc(num)
# # print(n)
#
#
# def numwords():
#     return num2words(n, lang='ru')
#
#
# number = str(numwords())
# # print(number)
#
#
# if "." in str(num):
#     cop = str(num).split(".")[-1]
#     # print(nn)
#
#     print(f"{number} рублей {cop} копеек")


def numerWords():
    cop2 = 0
    number = float(input())
    n = math.trunc(number)
    number_text = str(num2words(n, lang='ru'))
    if "." in str(number):
        cop = str(number).split(".")[-1]
        if "0" in cop:
            cop2 = cop
        else:
            cop2 = cop + "0"
    result = f"{number_text} рублей {cop2} копеек"
    return result

print(numerWords())