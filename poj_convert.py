"""台文漢字 -> 白話字 (Pe̍h-ōe-jī, POJ) 轉換。

以 taibun 斷詞與查音為底，再用 POJ_OVERRIDES 校正本書常見的讀錯詞。
校正表以「taibun 斷出來的詞」為單位比對（單字只在被斷成單字時才套用，
不會拆散 taibun 已認得的複合詞，例如「成長」不會被「長→tn̂g」影響）。
PHRASE_OVERRIDES 可跨越 taibun 的斷詞邊界（必須剛好落在斷詞邊界上）。

用法：
    from poj_convert import to_poj
    to_poj("愛有耐性。")  # -> "Ài ū nāi-sèng."

未安裝 taibun 時 to_poj() 回傳空字串，不中斷產生流程。
"""
import re

try:
    from taibun import Converter, Tokeniser
    _conv = Converter(system="POJ", format="mark", punctuation="none",
                      sandhi="none", dialect="south")
    _tok = Tokeniser()
    AVAILABLE = True
except Exception:  # pragma: no cover
    AVAILABLE = False

# 單一斷詞的讀音校正（本調、教育部辭典優先音）
POJ_OVERRIDES = {
    # 虛詞、常用口語字
    "莫": "mài", "攏": "lóng", "才": "chiah", "個": "in", "上": "siōng",
    "濟": "chē", "工": "kang", "前": "chêng", "為": "ūi", "提": "the̍h",
    "相": "siong", "較": "khah", "仔": "á", "抑": "iah", "抑是": "iah-sī",
    "啥物": "siáⁿ-mih", "啥人": "siáng", "大的": "tōa ê", "做會": "chò ē",
    "袂": "bē", "會": "ē", "閣": "koh", "嘛": "mā", "矣": "--ah",
    # 本書高頻詞
    "工課": "khang-khòe", "做工課": "chò khang-khòe", "核心": "ha̍t-sim",
    "上尾": "siōng-bóe", "上好": "siōng-hó", "上要緊": "siōng iàu-kín",
    "條": "tiâu", "程": "thêng", "現": "hiān", "成": "sêng", "量": "liōng",
    "實": "si̍t", "放": "pàng", "跨": "khòa", "假": "ká", "關": "koan",
    "落": "lo̍h", "長": "tn̂g", "情": "chêng", "將": "chiong", "圈": "khoân",
    "推": "thui", "養": "ióng", "韌": "jūn", "盤": "phoân", "批": "phoe",
    "職": "chit", "場": "tiûⁿ", "鍵": "kiān", "曆": "le̍k", "與": "ú",
    "開": "khui",
}

# 跨斷詞邊界的片語校正
PHRASE_OVERRIDES = {
    "職場": "chit-tiûⁿ", "關鍵": "koan-kiān", "相關": "siong-koan",
    "關係": "koan-hē", "關掉": "koaiⁿ-tiāu", "關機": "koaiⁿ-ki",
    "關心": "koan-sim", "行程": "hêng-thêng", "流程": "liû-thêng",
    "議程": "gī-thêng", "時程": "sî-thêng", "過程": "kòe-thêng",
    "課程": "khò-thêng", "工程師": "kang-thêng-su", "工程": "kang-thêng",
    "落實": "lo̍k-si̍t", "段落": "toāⁿ-lo̍k", "個人": "kò-jîn",
    "個月": "kò-goe̍h", "個性": "kò-sèng", "幾若个": "kúi-nā ê",
    "能量": "lêng-liōng", "數量": "sò͘-liōng", "量化": "liōng-hòa",
    "培養": "pôe-ióng", "營養": "êng-ióng", "養成": "ióng-sêng",
    "盤點": "phoân-tiám", "推前": "thui-chêng", "推理": "thui-lí",
    "推進": "thui-chìn", "表現": "piáu-hiān", "發現": "hoat-hiān",
    "呈現": "thêng-hiān", "展現": "tián-hiān", "現實": "hiān-si̍t",
    "實現": "si̍t-hiān", "請假": "chhiáⁿ-ká", "假期": "ká-kî",
    "假設": "ká-siat", "放假": "pàng-ká", "行事曆": "hêng-sū-le̍k",
    "舒適圈": "su-sek-khoân", "能力圈": "lêng-le̍k-khoân",
    "開喙": "khui-chhùi", "開始": "khai-sí", "公開": "kong-khai",
    "開發": "khai-hoat", "開會": "khui-hōe", "開講": "khai-káng",
    "韌性": "jūn-sèng", "心情": "sim-chêng", "情緒": "chêng-sū",
    "提升": "thê-seng", "提早": "thê-chá", "提醒": "thê-chhéⁿ",
    "提出": "thê-chhut", "提供": "thê-kiong", "提問": "thê-būn",
    "成長": "sêng-tióng", "長期": "tn̂g-kî", "互相": "hō͘-siong",
    "相信": "siong-sìn", "鬥相共": "tàu-sio-kāng", "共識": "kiōng-sek",
    "共通": "kiōng-thong", "共贏": "kiōng-iâⁿ", "上班": "siōng-pan",
    "下班": "hā-pan", "落班": "lo̍h-pan", "頂懸": "téng-koân",
    "學會": "o̍h-ē", "沉澱": "tîm-tiān", "蜥蜴腦": "sek-ia̍h-náu",
    "步驟": "pō͘-chō͘", "嗜好": "sī-hòⁿ", "惦惦": "tiām-tiām",
    "恬恬": "tiām-tiām", "憨憨": "khám-khám", "憋氣": "pih-khùi",
    "汲取": "kip-chhú", "隔閡": "keh-gāi", "癒": "jú", "療癒": "liâu-jú",
    "門檻": "mn̂g-khám", "棘手": "kek-chhiú", "獲得": "he̍k-tit",
    "收穫": "siu-he̍k", "演繹": "ián-e̍k", "範疇": "hoān-tiû",
    "蒐集": "so͘-chi̍p", "越俎代庖": "oa̍t-chó͘-tāi-pâu",
    "个人": "ê lâng", "崗位": "kòng-ūi", "里程碑": "lí-thêng-pi",
    "睬": "chhap", "洽": "kap", "紅菜頭": "âng-chhài-thâu", "紅蘿蔔": "âng-lô-po̍k", "洽談": "kap-tâm",
    "這馬": "chit-má", "逐家": "ta̍k-ke", "毋過": "m̄-koh",
}

_HAN = re.compile(r"[㐀-鿿\U00020000-\U0002ffff]")
_PUNCT_NO_SPACE_BEFORE = set("，。、：；！？）」』》…,.:;!?)")
_PUNCT_NO_SPACE_AFTER = set("（「『《(")
_FW = {"，": ", ", "。": ". ", "、": ", ", "：": ": ", "；": "; ", "！": "! ",
       "？": "? ", "（": " (", "）": ") ", "「": " “", "」": "” ", "『": " ‘",
       "』": "’ ", "《": " «", "》": "» ", "～": "~", "—": "—", "……": "… "}
_MAX_PHRASE = max(len(k) for k in PHRASE_OVERRIDES)


def _read(token):
    if token in POJ_OVERRIDES:
        return POJ_OVERRIDES[token]
    if not _HAN.search(token):
        return token
    return _conv.get(token)


def _tokens(text):
    toks = _tok.tokenise(text)
    # 片語校正：從斷詞邊界開始、到斷詞邊界結束的最長片語
    out, i = [], 0
    while i < len(toks):
        hit = None
        acc = ""
        for j in range(i, len(toks)):
            acc += toks[j]
            if len(acc) > _MAX_PHRASE:
                break
            if acc in PHRASE_OVERRIDES:
                hit = (j, PHRASE_OVERRIDES[acc])
        if hit:
            out.append(hit[1]); i = hit[0] + 1
        else:
            out.append(_read(toks[i])); i += 1
    return out


def _cap_sentences(s):
    # 句首大寫（POJ 慣例）
    def up(m):
        return m.group(1) + m.group(2).upper()
    s = re.sub(r"(^|[.!?]\s+)([^\W\d_])", up, s)
    return s


def to_poj(text):
    if not AVAILABLE or not text:
        return ""
    pieces = _tokens(text)
    out = ""
    for p in pieces:
        if p in _FW:
            out = out.rstrip() + _FW[p]
            continue
        p = p.strip()
        if not p:
            continue
        if out and not out.endswith((" ", "(", "“", "‘", "«")) and not p[0] in _PUNCT_NO_SPACE_BEFORE:
            out += " "
        out += p
    out = re.sub(r"\s+", " ", out).strip()
    # 「上」統一用教育部優先音 siōng（taibun 南部腔會出 siāng）
    out = out.replace("siāng", "siōng").replace("Siāng", "Siōng")
    out = re.sub(r"\s+([,.;:!?)”’»])", r"\1", out)
    return _cap_sentences(out)


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    for s in sys.argv[1:] or ["面笑笑，加一个朋友、少一个冤家。毋管拄著啥物場面，攏愛沉會牢、忍會住，千萬莫衝碰。"]:
        print(s); print(to_poj(s)); print()
