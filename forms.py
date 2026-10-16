"""Чтение исходных форм Судебного департамента (data/raw/*.xls) и расчёт показателей для приложения.

Исходные файлы лежат в репозитории без изменений. Здесь всё, что с ними делается:
  1. по тексту шапки находим нужные колонки (в разные годы колонки стоят по-разному);
  2. по первой цифре в подписи строки определяем статью УК (части и редакции статьи суммируем);
  3. считаем показатели и доли.
"""
import re
from collections import defaultdict
from pathlib import Path

import xlrd

RAW = Path(__file__).parent / "data" / "raw"
YEARS = list(range(2016, 2026))
LATIN_TO_CYRILLIC = str.maketrans("aceopxy", "асеорху")   # в шапке 2016 г. встречается латинская «c» в слове «cвыше»
ARTICLE_NO = re.compile(r"^\s*(\d+(?:\.\d+)?)")        # «264.1 ч. 2 (в старой редакции)» -> «264.1»

# ---------- 1. Какие колонки нужны: ключ -> слова, которые должны быть в тексте шапки ----------
# Форма 10.3 — исходы дел (шапка в строках 5–7 листа «Раздел 1»)
OUTCOME_COLUMNS = {
    "convicted":     ["всего осуждено"],
    "acquitted":     ["оправдано"],
    "term_rehab":    ["прекращено", "за отсутствием"],            # реабилитирующие основания
    "amnesty":       ["прекращено", "по амнистии"],               # с 2025 г. колонка не заполняется
    "reconciled":    ["прекращено", "примирением"],
    "repentance":    ["прекращено", "раскаянием"],
    "court_fine":    ["прекращено", "судебный штраф"],
    "other_grounds": ["прекращено", "по другим основаниям"],
    "medical":       ["принудительные меры"],
}
# Виды наказания (форма 10.3, графы 3–15 и «освобождено от наказания»): ключ -> с чего начинается подпись графы.
# Графа без подписи (в 2025 г. пустые «дисциплинарная часть» и «ограничение по военной службе») попадает в other_main.
PUNISHMENT_COLUMNS = {
    "life":        "пожизненное",
    "prison":      "лишение свободы",
    "cond_prison": "условное осуждение к лишению",
    "cond_other":  "условное осуждение к иным",
    "fine":        "штраф",
}
PUNISHMENT_KEYS = list(PUNISHMENT_COLUMNS) + ["other_main", "released"]
TERMS = {"до 1 года": "term_up_to_1y", "свыше 1 до 2": "term_1_2y", "свыше 2 до 3": "term_2_3y",
         "свыше 3 до 5": "term_3_5y", "свыше 5 до 8": "term_5_8y", "свыше 8 до 10": "term_8_10y",
         "свыше 10 до 15": "term_10_15y", "свыше 15 до 20": "term_15_20y"}
FINES = {"до 5 тыс": "_up_to_5k", "свыше 5 тыс. до 25": "_5_25k", "свыше 25 тыс. до 100": "_25_100k",
         "свыше 100 тыс. до 300": "_100_300k", "свыше 300 тыс. до 500": "_300_500k",
         "свыше 500 тыс": "_500k_1m", "свыше 1 млн": "_over_1m"}
JFINES = {"до 5 тыс": "_up_to_5k", "свыше 5 тыс. до 25": "_5_25k", "свыше 25 тыс. до 100": "_25_100k",
          "свыше 100 тыс": "_over_100k"}


# подписи интервалов в приложении
TERM_BUCKETS = [("term_up_to_1y", "до 1 года"), ("term_1_2y", "1–2 года"), ("term_2_3y", "2–3 года"),
                ("term_3_5y", "3–5 лет"), ("term_5_8y", "5–8 лет"), ("term_8_10y", "8–10 лет"),
                ("term_10_15y", "10–15 лет"), ("term_15_20y", "15–20 лет")]
FINE_BUCKETS = [("_up_to_5k", "до 5 тыс."), ("_5_25k", "5–25 тыс."), ("_25_100k", "25–100 тыс."),
                ("_100_300k", "100–300 тыс."), ("_300_500k", "300–500 тыс."),
                ("_500k_1m", "500 тыс.–1 млн"), ("_over_1m", "свыше 1 млн")]
JFINE_BUCKETS = [("judicial_fine_up_to_5k", "до 5 тыс."), ("judicial_fine_5_25k", "5–25 тыс."),
                 ("judicial_fine_25_100k", "25–100 тыс."), ("judicial_fine_over_100k", "свыше 100 тыс.")]


def sentencing_key(group, sub):
    """Форма 10.3.1 — сроки и штрафы (шапка в строках 7–8): (группа, подпись) -> ключ или None."""
    pick = lambda table: next((v for k, v in table.items() if sub.startswith(k)), None)
    if group.startswith("лишение свободы: всего"):
        return "imprisonment_total"
    if group.startswith("наказание назначено ниже"):
        return "below_minimum_imprisonment"
    if group.startswith("срок лишения свободы") and pick(TERMS):
        return pick(TERMS)
    if group.startswith("штраф") and "гр. 12" in group and pick(FINES):      # основное наказание
        return "fine_main" + pick(FINES)
    if group.startswith("штраф") and "гр. 26" in group and pick(FINES):      # дополнительное
        return "fine_additional" + pick(FINES)
    if group.startswith("судебный штраф как мера") and pick(JFINES):
        return "judicial_fine" + pick(JFINES)
    if group.startswith("общая сумма штрафов") and "гр. 13" in group and sub.startswith("всего руб"):
        return "fine_main_total_rub"
    return None


# ---------- чтение листа ----------
def read_sheet(path):
    """Лист «Раздел 1» -> (строки с шапкой, строки с данными)."""
    sh = xlrd.open_workbook(path).sheet_by_name("Раздел 1")
    rows = [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]
    first = next(i for i, r in enumerate(rows) if r[2] == 1 and r[3] == 2) + 1     # строка с номерами граф «1, 2, 3 …»
    return rows[:first], rows[first:]


def header_texts(head, levels):
    """Текст шапки каждой колонки. Объединённые ячейки в .xls заполнены только в первой колонке,
    поэтому подпись «растягиваем» вправо до следующей непустой."""
    texts = [""] * len(head[0])
    for level in levels:
        cur = ""
        for c, v in enumerate(head[level]):
            v = " ".join(str(v).split()).lower().translate(LATIN_TO_CYRILLIC)
            if v:
                cur = v
            elif c == 0:
                cur = ""
            texts[c] += "|" + cur
    return texts


def punishment_columns(path):
    """Форма 10.3: {ключ: [номера колонок]} для видов наказания. Вместе с освобождёнными от наказания
    эти колонки раскладывают ВСЕХ осуждённых без остатка (проверено по всем годам, см. README)."""
    head, _ = read_sheet(path)
    group = header_texts(head, [4])                                     # «виды основного наказания» / «по приговору освобождено…»
    own = [" ".join(str(v).split()).lower().translate(LATIN_TO_CYRILLIC) for v in head[5]]   # подпись самой графы, без растягивания
    cols = {k: [] for k in PUNISHMENT_KEYS}
    for c in range(3, len(group)):
        if "виды основного наказания" in group[c]:
            key = next((k for k, start in PUNISHMENT_COLUMNS.items()
                        if own[c].startswith(start) and (k != "prison" or own[c] == start)), "other_main")
            cols[key].append(c)
        elif "по приговору освобождено" in group[c]:
            cols["released"].append(c)
    return cols


def find_columns(path, kind):
    head, _ = read_sheet(path)
    cols = {}
    if kind == "10.3":
        for c, text in enumerate(header_texts(head, [4, 5, 6])):
            for key, words in OUTCOME_COLUMNS.items():
                if all(w in text for w in words) and key not in cols:
                    cols[key] = c
    else:
        top, sub = header_texts(head, [6]), header_texts(head, [7])
        for c in range(2, len(top)):
            key = sentencing_key(top[c].strip("|"), sub[c].split("|")[-1])
            if key:
                cols[key] = c
    return cols


# ---------- 2–3. Суммы по статьям ----------
def number(v):
    return v if isinstance(v, (int, float)) else 0


def sum_by_article(year, kind):
    """{номер статьи: {показатель: сумма по всем её частям и редакциям}}"""
    path = RAW / f"form_{kind.replace('.', '_')}_{year}.xls"
    cols = find_columns(path, kind)
    punish = punishment_columns(path) if kind == "10.3" else {}
    _, rows = read_sheet(path)
    out = defaultdict(lambda: defaultdict(int))
    for r in rows:
        label = f"{r[0]:g}" if isinstance(r[0], float) else str(r[0])    # статья «106» хранится как число 106.0
        m = ARTICLE_NO.match(label)
        if m:                                     # строки без номера («Воинские преступления», итоги) пропускаем
            for key, c in cols.items():
                out[m.group(1)][key] += int(number(r[c]))
            for key, cs in punish.items():
                out[m.group(1)][key] += sum(int(number(r[c])) for c in cs)
    return out


def pct(a, b):
    return round(a / b * 100, 2) if b else None


def outcomes(c):
    """Показатели приложения из сумм по колонкам формы 10.3."""
    nonrehab = c["amnesty"] + c["reconciled"] + c["repentance"] + c["court_fine"] + c["other_grounds"]
    total = c["convicted"] + c["acquitted"] + c["term_rehab"] + nonrehab + c["medical"]
    sentence = {"sent_" + k: c[k] for k in PUNISHMENT_KEYS}
    # остаток = осуждено − всё распределённое по видам наказания; в форме он почти всегда 0
    # (единичные расхождения есть в самих таблицах 2016, 2017 и 2025 гг.) — показываем его явно, чтобы сумма сходилась
    sentence["sent_unallocated"] = c["convicted"] - sum(c[k] for k in PUNISHMENT_KEYS)
    custodial = c["life"] + c["prison"] + c["cond_prison"]            # все приговорённые к лишению свободы: реально и условно
    sentence["pct_cond_of_custodial"] = pct(c["cond_prison"], custodial)
    sentence["custodial_total"] = custodial
    return {**sentence, "convicted": c["convicted"], "acquitted": c["acquitted"], "term_rehab": c["term_rehab"],
            "term_nonrehab": nonrehab, "medical": c["medical"], "total": total,
            "pct_acquitted": pct(c["acquitted"], total), "pct_convicted": pct(c["convicted"], total),
            "pct_term_rehab": pct(c["term_rehab"], total), "pct_term_nonrehab": pct(nonrehab, total),
            "pct_medical": pct(c["medical"], total),
            "pct_of_verdicts": pct(c["acquitted"], c["convicted"] + c["acquitted"])}


def build(names):
    """names: {номер статьи: {"ru": ..., "en": ..., ...}} — статьи, которые показываем в приложении.
    Возвращает (исходы по статьям, сроки и штрафы по статьям)."""
    by_year = {y: sum_by_article(y, "10.3") for y in YEARS}
    sent_year = {y: sum_by_article(y, "10.3.1") for y in YEARS}
    articles, sentencing = [], {}
    last = YEARS[-1]
    for num in sorted(names, key=float):
        years, all_c, all_s = {}, defaultdict(int), defaultdict(int)
        # статья была в отчёте прошлого года, а в этом строки пусты -> данные не опубликованы
        unpublished = [last] if num in by_year[last - 1] and num not in by_year[last] else []
        sent = {}
        for y in YEARS:
            if y in unpublished:
                continue
            c = by_year[y].get(num, defaultdict(int))
            years[str(y)] = outcomes(c)
            for k in [*OUTCOME_COLUMNS, *PUNISHMENT_KEYS]:
                all_c[k] += c[k]
            s = {k: v for k, v in sent_year[y].get(num, {}).items() if v}
            sent[str(y)] = s
            for k, v in s.items():
                all_s[k] += v
        years["ALL"] = outcomes(all_c)
        sent["ALL"] = dict(all_s)
        sentencing[num] = sent
        a = dict(names[num], num=num, years=years)
        if unpublished:
            a["unpublished_years"] = unpublished
        articles.append(a)
    sentencing = {"term_buckets": TERM_BUCKETS, "fine_buckets": FINE_BUCKETS,
                  "jfine_buckets": JFINE_BUCKETS, "data": sentencing}
    return {"years": YEARS, "articles": articles}, sentencing
