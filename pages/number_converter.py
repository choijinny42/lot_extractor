import streamlit as st
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation

st.set_page_config(page_title="Number Converter", page_icon="🔢")
st.title("🔢 Number Converter")

# paste all the number-to-words functions here directly
# (ONES, TEENS, TENS, SCALES, _0_to_999_hy, int_to_armenian, number_to_armenian)
ONES = {
    0: "զրո",
    1: "մեկ",
    2: "երկու",
    3: "երեք",
    4: "չորս",
    5: "հինգ",
    6: "վեց",
    7: "յոթ",
    8: "ութ",
    9: "ինը",
}

TEENS = {
    10: "տասը",
    11: "տասնմեկ",
    12: "տասներկու",
    13: "տասներեք",
    14: "տասնչորս",
    15: "տասնհինգ",
    16: "տասնվեց",
    17: "տասնյոթ",
    18: "տասնութ",
    19: "տասնինը",
}

TENS = {
    2: "քսան",
    3: "երեսուն",
    4: "քառասուն",
    5: "հիսուն",
    6: "վաթսուն",
    7: "յոթանասուն",
    8: "ութսուն",
    9: "իննսուն",
}

SCALES = [
    (10**12, "տրիլիոն"),
    (10**9,  "միլիարդ"),
    (10**6,  "միլիոն"),
    (10**3,  "հազար"),
]

def _0_to_999_hy(n: int) -> str:
    if not (0 <= n <= 999):
        raise ValueError("n must be between 0 and 999")

    if n < 10:
        return ONES[n]
    if 10 <= n <= 19:
        return TEENS[n]
    if n < 100:
        tens = n // 10
        ones = n % 10
        if ones == 0:
            return TENS[tens]
        # Common Armenian practice: concatenate 21..99 as one word
        return TENS[tens] + ONES[ones]

    hundreds = n // 100
    rem = n % 100

    head = "հարյուր" if hundreds == 1 else f"{ONES[hundreds]} հարյուր"
    if rem == 0:
        return head
    return f"{head} {_0_to_999_hy(rem)}"

def int_to_armenian(n: int) -> str:
    if n == 0:
        return ONES[0]
    if n < 0:
        return "մինուս " + int_to_armenian(-n)

    parts = []
    x = n

    for scale_val, scale_name in SCALES:
        if x >= scale_val:
            group = x // scale_val
            x %= scale_val

            if scale_val == 10**3 and group == 1:
                parts.append("հազար")  # 1000 = հազար
            else:
                parts.append(f"{int_to_armenian(group)} {scale_name}")

    if x > 0:
        parts.append(_0_to_999_hy(x))

    return " ".join(parts)

def number_to_armenian(x, *, as_currency=False, major="դրամ", minor="լումա") -> str:
    s = str(x).strip().replace(" ", "").replace(",", ".")
    try:
        d = Decimal(s)
    except InvalidOperation:
        raise ValueError("Սխալ թիվ է։ Օրինակ՝ 123, 123.45 կամ -10.5")

    if as_currency:
        d = d.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        sign = "մինուս " if d.is_signed() else ""
        d = abs(d)

        whole = int(d)
        frac = int((d - Decimal(whole)) * 100)  # 0..99

        whole_words = int_to_armenian(whole) + f" {major}"
        if frac == 0:
            return sign + whole_words

        frac_words = _0_to_999_hy(frac) + f" {minor}"
        return sign + whole_words + " և " + frac_words

    # Non-currency: read digits after "կետ"
    sign = "մինուս " if d.is_signed() else ""
    d = abs(d)

    s2 = format(d, "f").rstrip("0").rstrip(".")
    if "." not in s2:
        return sign + int_to_armenian(int(s2))

    whole_str, frac_str = s2.split(".")
    whole = int(whole_str) if whole_str else 0
    frac_words = " ".join(ONES[int(ch)] for ch in frac_str) if frac_str else ONES[0]
    return sign + int_to_armenian(whole) + " կետ " + frac_words


mode = st.radio("Mode", ["Plain number", "Currency (դրամ)"])
as_currency = mode == "Currency (դրամ)"

tab1, tab2 = st.tabs(["Single value", "Bulk / paste from Excel"])

with tab1:
    value = st.text_input("Enter a number", placeholder="e.g. 123456")
    if value:
        try:
            result = number_to_armenian(value, as_currency=as_currency)
            st.success(result)
        except Exception as e:
            st.error(str(e))

with tab2:
    pasted = st.text_area("Paste a column from Excel (one number per line)")
    if pasted:
        lines = [l.strip() for l in pasted.splitlines() if l.strip()]
        results = []
        for v in lines:
            try:
                results.append(number_to_armenian(v, as_currency=as_currency))
            except Exception as e:
                results.append(f"[ՍԽԱԼ] {v} → {e}")
        st.text_area("Results", value="\n".join(results), height=300)
        st.download_button(
            "⬇️ Download results",
            data="\n".join(results).encode("utf-8-sig"),
            file_name="armenian_numbers.txt",
            mime="text/plain"
        )
