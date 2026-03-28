import re


def normalize_for_matching(text: str) -> str:
    if not text or not isinstance(text, str):
        return ""
    text = text.strip()
    text = re.sub(r'\s*#\s*', '#', text)
    text = re.sub(r'\s+', ' ', text)
    text = text.replace('（', '(').replace('）', ')')
    return text


def normalize_strict(text: str) -> str:
    if not text or not isinstance(text, str):
        return ""
    text = re.sub(r'[\s\-_·・\(\)（）]', '', text)
    return text.lower()


def split_product_option(product_name: str) -> tuple[str, str | None]:
    if not product_name or not isinstance(product_name, str):
        return (product_name or "", None)

    if '#' in product_name:
        parts = product_name.split('#', 1)
        base = parts[0].strip()
        option = '#' + parts[1].strip()
        return (base, option)

    return (product_name.strip(), None)
