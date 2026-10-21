"""
Vietnamese Number to Words Converter for Cost Estimation.
Example: 24000 -> Hai mươi tư nghìn đồng.
"""

def num2words_vi(n: float) -> str:
    """Convert an amount in VND to Vietnamese words."""
    try:
        n = int(round(n))
    except:
        return ""
    if n == 0:
        return "Không đồng"
    if n < 0:
        return "Âm " + num2words_vi(-n).lower()

    units = ["không", "một", "hai", "ba", "bốn", "năm", "sáu", "bảy", "tám", "chín"]
    scales = ["", "nghìn", "triệu", "tỷ", "nghìn tỷ", "triệu tỷ"]

    def read_three(num: int, is_highest: bool = False) -> str:
        hundred = num // 100
        ten = (num % 100) // 10
        one = num % 10
        parts = []

        if hundred > 0:
            parts.append(units[hundred] + " trăm")
        elif not is_highest:
            parts.append("không trăm")

        if ten == 0:
            if one > 0:
                if not is_highest or hundred > 0:
                    parts.append("lẻ " + units[one])
                else:
                    parts.append(units[one])
        elif ten == 1:
            if one == 0:
                parts.append("mười")
            elif one == 5:
                parts.append("mười lăm")
            else:
                parts.append("mười " + units[one])
        else:
            if one == 0:
                parts.append(units[ten] + " mươi")
            elif one == 1:
                parts.append(units[ten] + " mươi mốt")
            elif one == 4:
                parts.append(units[ten] + " mươi tư")
            elif one == 5:
                parts.append(units[ten] + " mươi lăm")
            else:
                parts.append(units[ten] + " mươi " + units[one])

        return " ".join(parts).strip()

    chunks = []
    temp = n
    while temp > 0:
        chunks.append(temp % 1000)
        temp //= 1000

    result_parts = []
    for i in range(len(chunks) - 1, -1, -1):
        c = chunks[i]
        if c > 0:
            c_text = read_three(c, i == len(chunks) - 1)
            scale = scales[i]
            if scale:
                result_parts.append(c_text + " " + scale)
            else:
                result_parts.append(c_text)

    text = " ".join(result_parts).strip()
    # Capitalize first letter
    if text:
        text = text[0].upper() + text[1:] + " đồng./."
    return text

if __name__ == "__main__":
    for val in [24000, 18816, 23799, 125000000, 1543200]:
        print(f"{val:,} -> {num2words_vi(val)}")
