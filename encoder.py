import json
import re
import os
import sys


def get_resource_path(filename):
    """
    Trả về đường dẫn tuyệt đối đến file tài nguyên (như rules.json),
    hoạt động đúng cả khi chạy trong môi trường PyInstaller --onefile.
    """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.abspath("."), filename)


def load_rules(path='rules.json'):
    """
    Load danh sách rule từ file JSON, chỉ lấy những rule đang bật (enabled=True).
    """
    path = get_resource_path(path)
    with open(path, encoding='utf-8') as f:
        all_rules = json.load(f)
    rules = [r for r in all_rules if r.get("enabled", True)]
    validate_rules(rules)
    return rules


def validate_rules(rules):
    """
    Kiểm tra các rule không bị trùng lặp hoặc sai logic.
    """
    all_from = set()
    all_to = set()
    for rule in rules:
        to_val = rule["to"]
        for from_val in rule["from"]:
            if from_val in all_from:
                raise ValueError(f"Từ thay thế trùng lặp: '{from_val}'")
            if from_val == to_val:
                raise ValueError(f"Từ gốc và mã hoá trùng nhau: '{from_val}'")
            all_from.add(from_val)
        if to_val in all_from or to_val in all_to:
            raise ValueError(f"Mã hoá trùng hoặc không hợp lệ: '{to_val}'")
        all_to.add(to_val)


def build_replacement_maps(rules, strict_boundary=False):
    """
    Tạo 2 dictionary:
    - encode_map: từ gốc → mã hoá
    - decode_map: mã hoá → từ gốc

    Nếu strict_boundary=True thì chỉ thay thế từ trùng chính xác (dùng \\b).
    Nếu False thì thay thế bất kỳ chuỗi con nào khớp (không giới hạn từ nguyên).
    """
    encode_map = {}
    decode_map = {}

    for rule in rules:
        to = rule["to"]
        for word in rule["from"]:
            # Nếu strict thì thay thế đúng từ (\\bWORD\\b), nếu không thì bất kỳ vị trí nào
            if strict_boundary:
                pattern = r'\b' + re.escape(word) + r'\b'
                reverse_pattern = r'\b' + re.escape(to) + r'\b'
            else:
                pattern = re.escape(word)
                reverse_pattern = re.escape(to)

            encode_map[pattern] = to
            decode_map[reverse_pattern] = word

    return encode_map, decode_map


def replace_text(text, replace_map):
    """
    Thay thế nội dung text theo mapping từ replace_map.
    """
    for pattern, replacement in replace_map.items():
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    return text
