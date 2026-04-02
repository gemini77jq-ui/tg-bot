"""
Числа прописью на русском языке (родительный падеж).
Используется для формулировки вида "15 (пятнадцати) штук".
"""

ONES_GENITIVE = [
    "", "одного", "двух", "трёх", "четырёх",
    "пяти", "шести", "семи", "восьми", "девяти",
]

TEENS_GENITIVE = [
    "десяти", "одиннадцати", "двенадцати", "тринадцати", "четырнадцати",
    "пятнадцати", "шестнадцати", "семнадцати", "восемнадцати", "девятнадцати",
]

TENS_GENITIVE = [
    "", "десяти", "двадцати", "тридцати", "сорока",
    "пятидесяти", "шестидесяти", "семидесяти", "восьмидесяти", "девяноста",
]

HUNDREDS_GENITIVE = [
    "", "ста", "двухсот", "трёхсот", "четырёхсот",
    "пятисот", "шестисот", "семисот", "восьмисот", "девятисот",
]


def number_to_genitive(n: int) -> str:
    """Возвращает число прописью в родительном падеже (1-999)."""
    if n <= 0 or n >= 1000:
        return str(n)

    parts = []
    hundreds = n // 100
    remainder = n % 100
    tens = remainder // 10
    ones = remainder % 10

    if hundreds:
        parts.append(HUNDREDS_GENITIVE[hundreds])

    if 10 <= remainder <= 19:
        parts.append(TEENS_GENITIVE[remainder - 10])
    else:
        if tens:
            parts.append(TENS_GENITIVE[tens])
        if ones:
            parts.append(ONES_GENITIVE[ones])

    return " ".join(parts)
