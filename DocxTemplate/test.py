from num2words import num2words
import math

num = float(input("Enter number: "))
n = math.trunc(num)
# print(n)


def numwords():
    return num2words(n, lang='ru')


number = str(numwords())
# print(number)


if "." in str(num):
    cop = str(num).split(".")[-1]
    # print(nn)

    print(f"{number} рублей {cop} копеек")