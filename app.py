import csv
import io
import json
import textwrap
from pathlib import Path

import streamlit as st

import forms   # чтение исходных .xls и расчёт показателей (см. forms.py)


def html(s: str) -> str:
    """Dedent a triple-quoted HTML string so Streamlit's markdown parser
    never mistakes indented lines for a code block, then render it."""
    return textwrap.dedent(s).strip()


def render(s: str):
    st.markdown(html(s), unsafe_allow_html=True)


# ---------- Language state ----------
if "lang" not in st.session_state:
    st.session_state.lang = "ru"


def toggle_lang():
    st.session_state.lang = "en" if st.session_state.lang == "ru" else "ru"


lang = st.session_state.lang

# ---------- UI text dictionary ----------
T = {
    "ru": {
        "page_title": "УК РФ — % оправдательных приговоров",
        "intro": """
        Использованы формы 10.3 и 10.3.1 Судебного департамента при Верховном Суде РФ за 2016–2025 годы.
        Считаются люди, а не дела: каждый человек учитывается один раз — по основному (наиболее тяжкому)
        составу преступления, без учёта сложения. Если в приговоре несколько статей, он попадает в статистику
        только по основной.
        """,
        "badge": "УК РФ · форма 10.3",
        "h1": "Оправдательные приговоры по статьям УК РФ",
        "subtitle": "Выберите статью (номер или название — части объединены в статью целиком) и год.",
        "field_article": "Статья УК РФ",
        "field_year": "Год",
        "placeholder_article": "Введите номер или название статьи...",
        "all_years": "Все годы (2016–2025)",
        "year_phrase_all": "за все годы (2016–2025)",
        "not_found": "Статья не найдена.",
        "no_data": "Нет данных по этой статье {year_phrase} (лиц, учтённых в отчёте, не было).",
        "no_data_unpublished": "Данные по этой статье {year_phrase} не опубликованы: в отчёте Судебного департамента строки по статье пусты. Это не означает, что таких лиц не было.",
        "partial_all": "По этой статье за 2025 год данные не опубликованы, поэтому показатели «за все годы» посчитаны только за {years}.",
        "na_label": "н/д",
        "dl_title": "Исходные файлы и данные",
        "dl_note": "Формы Судебного департамента при Верховном Суде РФ в исходном виде (.xls, без изменений). Все расчёты приложения — в файле forms.py.",
        "dl_form_103": "Форма 10.3 — исходы дел",
        "dl_form_1031": "Форма 10.3.1 — сроки и штрафы",
        "dl_csv": "Скачать таблицу приложения (CSV): статья × год",
        "year_warning": """
        <b>Данные за 2025 год — с оговорками.</b><br>
        1) В отчёте 2025 г. не опубликованы строки по ст. 275–284.3 и 330.1 УК РФ (по ним показано «н/д»), а также строка «Воинские преступления». Это пропуск данных, а не отсутствие дел.<br>
        2) Падение общего числа осуждённых в 2025 г. (437,8 тыс. против 512,8 тыс. в 2024 г.) нельзя считать реальным сокращением преступности или репрессий: оно частично объясняется пропусками данных.
        """,
        "total_caption": "Всего рассмотрено лиц {year_phrase}:",
        "m_acquitted": "Оправдано",
        "m_convicted": "Осуждено",
        "m_rehab": "Прекращено, реабилитирующие",
        "m_nonrehab": "Прекращено, нереабилитирующие",
        "m_medical": "Принудительные меры мед. характера",
        "metric_note": "Проценты — доля от «всего рассмотрено» (в сумме по всем пяти показателям дают 100%).",
        "trend_title": "Динамика доли оправдательных по годам",
        "sentencing_title": "Размеры наказаний по статье",
        "kinds_title": "Какое наказание получили осуждённые",
        "cond_share": "Условное осуждение к лишению свободы — {pct}% от всех приговорённых к лишению свободы ({cond} из {custodial}; реальное, включая пожизненное, и условное вместе)",
        "k_prison": "Реальное лишение свободы на срок",
        "k_life": "Пожизненное лишение свободы",
        "k_cond_prison": "Условное осуждение к лишению свободы",
        "k_cond_other": "Условное осуждение к иным мерам",
        "k_fine": "Штраф (основное наказание)",
        "k_other_main": "Иные основные наказания (работы, ограничение свободы, арест, лишение права, воинские)",
        "k_released": "Освобождены от наказания / наказание не назначено",
        "k_unallocated": "Не распределено в форме (расхождение внутри отчёта)",
        "k_total": "Всего осуждено",
        "k_persons": "лиц",
        "k_share": "% от осуждённых",
        "terms_note": "В распределение сроков входит только реальное лишение свободы на срок; остальные осуждённые — в таблице выше.",
        "terms_gap": "Ещё {n} осуждены к лишению свободы, но в форме 10.3.1 не распределены по срокам.",
        "term_title": "Лишение свободы на срок — распределение сроков ({n} осуждено)",
        "below_min": "Из них ниже низшего предела: {n}",
        "fine_main_title": "Штраф (основное наказание) — {n} осуждено",
        "fine_add_title": "Штраф (дополнительное наказание) — {n} осуждено",
        "jfine_title": "Судебный штраф, ст. 104.4 УК РФ — {n} лиц (с 2018 г.)",
        "rub_total": "Общая сумма (с 2018 г.): {n} руб.",
        "rub_missing": "Точная сумма в рублях не собиралась до 2018 года — доступны только диапазоны.",
        "article_word": "Статья",
        "footer": """
        Источник: формы 10.3 и 10.3.1 Судебного департамента при Верховном Суде РФ, 2016–2025
        (данные за 2025 г. — по состоянию на 08.04.2026).
        Части статей объединены в показатель по статье целиком.<br>
        «Всего рассмотрено» = осуждено + оправдано + прекращено (реабилитирующие: отсутствие
        события/состава преступления, непричастность) + прекращено (нереабилитирующие: амнистия,
        примирение сторон, деятельное раскаяние, судебный штраф, иные основания) +
        принудительные меры медицинского характера к невменяемым. Проценты считаются от этой суммы.<br>
        Считаются люди (лица), а не дела или приговоры. Каждый человек учитывается один раз — по основному
        (наиболее тяжкому) составу преступления, без учёта сложения, поэтому показатели по статье — это не все
        приговоры, где статья упоминалась.<br>
        С 2025 г. в форме 10.3 не заполняется графа «прекращено по амнистии» (в 2024 г. — 11 человек по всей стране),
        поэтому «нереабилитирующие основания» за 2025 г. могут быть занижены на эту величину.<br>
        Доля условного осуждения = условно осуждённые к лишению свободы ÷ (все приговорённые к лишению свободы: реально, включая пожизненное, + условно).
        Доля от всех осуждённых показана в таблице видов наказания.<br>
        Показатели за 2021–2022 гг. по ряду статей учтены в сводных строках отчёта и отнесены к первой статье строки.<br><br>
        Названия статей на английском — авторский перевод (единого актуального официального
        перевода УК РФ на английский не существует).
        """,
    },
    "en": {
        "page_title": "Russian Criminal Code — Acquittal Rate",
        "intro": """
        Based on Forms 10.3 and 10.3.1 of the Judicial Department of the Supreme Court of the Russian Federation, 2016–2025.
        The figures count people, not cases: each person is counted once, under the main (most serious) offence,
        without aggregation of offences. If a verdict covers several articles, the person appears in the statistics
        under the main one only.
        """,
        "badge": "Russian CC · Form 10.3",
        "h1": "Acquittals by Article of the Russian Criminal Code",
        "subtitle": "Choose an article (by number or name — parts are combined into the article as a whole) and a year.",
        "field_article": "Article of the Criminal Code",
        "field_year": "Year",
        "placeholder_article": "Type an article number or name...",
        "all_years": "All years (2016–2025)",
        "year_phrase_all": "for all years (2016–2025)",
        "not_found": "Article not found.",
        "no_data": "No data for this article {year_phrase} (no persons recorded in the report).",
        "no_data_unpublished": "Data for this article {year_phrase} have not been published: the rows for the article in the Judicial Department's report are blank. This does not mean there were no such persons.",
        "partial_all": "Data for this article are not published for 2025, so the \"all years\" figures cover {years} only.",
        "na_label": "n/a",
        "dl_title": "Source files and data",
        "dl_note": "Forms of the Judicial Department of the Supreme Court of the Russian Federation as published (.xls, unmodified). All calculations of the app are in forms.py.",
        "dl_form_103": "Form 10.3 — case outcomes",
        "dl_form_1031": "Form 10.3.1 — terms and fines",
        "dl_csv": "Download the app's table (CSV): article × year",
        "year_warning": """
        <b>2025 data come with caveats.</b><br>
        1) The 2025 report does not publish rows for Arts. 275–284.3 and 330.1 of the Criminal Code (shown as "n/a") nor the "Military crimes" row. This is missing data, not an absence of persons.<br>
        2) The drop in the total number of convictions in 2025 (437.8k vs 512.8k in 2024) cannot be read as a real decline in crime or repression: it is partly explained by omissions in the data.
        """,
        "total_caption": "Total persons considered {year_phrase}:",
        "m_acquitted": "Acquitted",
        "m_convicted": "Convicted",
        "m_rehab": "Terminated, exculpatory grounds",
        "m_nonrehab": "Terminated, non-exculpatory grounds",
        "m_medical": "Compulsory medical measures",
        "metric_note": "Percentages are the share of \"total persons considered\" (all five figures sum to 100%).",
        "trend_title": "Acquittal rate trend by year",
        "sentencing_title": "Sentence sizes for this article",
        "kinds_title": "What the convicted received",
        "cond_share": "Suspended sentences of imprisonment — {pct}% of everyone sentenced to imprisonment ({cond} of {custodial}; actual, incl. life, plus suspended)",
        "k_prison": "Actual imprisonment for a term",
        "k_life": "Life imprisonment",
        "k_cond_prison": "Suspended sentence of imprisonment",
        "k_cond_other": "Suspended sentence of other punishments",
        "k_fine": "Fine (principal punishment)",
        "k_other_main": "Other principal punishments (labour, restriction of liberty, arrest, disqualification, military)",
        "k_released": "Released from punishment / no punishment imposed",
        "k_unallocated": "Not allocated in the form (inconsistency inside the report)",
        "k_total": "Total convicted",
        "k_persons": "persons",
        "k_share": "% of convicted",
        "terms_note": "The term distribution covers actual imprisonment for a term only; the other convicted persons are in the table above.",
        "terms_gap": "Another {n} were sentenced to imprisonment but are not broken down by term in form 10.3.1.",
        "term_title": "Imprisonment for a term — distribution of terms ({n} sentenced)",
        "below_min": "Of which below the statutory minimum: {n}",
        "fine_main_title": "Fine (principal punishment) — {n} sentenced",
        "fine_add_title": "Fine (additional punishment) — {n} sentenced",
        "jfine_title": "Judicial fine, Art. 104.4 of the Criminal Code — {n} persons (since 2018)",
        "rub_total": "Total amount (since 2018): {n} RUB",
        "rub_missing": "Exact rouble amounts were not collected before 2018 — only ranges are available.",
        "article_word": "Article",
        "footer": """
        Source: Forms 10.3 and 10.3.1 of the Judicial Department of the Supreme Court of the Russian Federation,
        2016–2025 (2025 data as of 08.04.2026). Article parts are combined into a single figure per article.<br>
        "Total considered" = convicted + acquitted + terminated (exculpatory grounds: absence of the event
        or elements of a crime, non-involvement) + terminated (non-exculpatory grounds: amnesty,
        reconciliation of the parties, active repentance, judicial fine, other grounds) +
        compulsory medical measures for persons found insane. Percentages are calculated from this total.<br>
        The figures count people (persons), not cases or sentences. Each person is counted once — under the main
        (most serious) offence, without aggregation of offences — so figures per article are not all sentences
        in which the article was mentioned.<br>
        From 2025 the form 10.3 column "terminated due to amnesty" is no longer filled in (11 people nationwide in 2024),
        so "non-exculpatory grounds" for 2025 may be understated by that amount.<br>
        Suspended-sentence share = persons given suspended imprisonment ÷ (all persons sentenced to imprisonment: actual, incl. life, + suspended).
        The share of all convicted persons is given in the table of punishment types.<br>
        For some articles, figures for 2021–2022 sit in combined rows of the report and are attributed to the first article of the row.<br><br>
        English article titles are an original translation by the author — no single up-to-date official
        English translation of the Russian Criminal Code exists.
        """,
    },
}
t = T[lang]

BUCKET_LABELS_EN = {
    "term_up_to_1y": "up to 1 year", "term_1_2y": "1–2 years", "term_2_3y": "2–3 years",
    "term_3_5y": "3–5 years", "term_5_8y": "5–8 years", "term_8_10y": "8–10 years",
    "term_10_15y": "10–15 years", "term_15_20y": "15–20 years",
    "_up_to_5k": "up to 5,000", "_5_25k": "5,000–25,000", "_25_100k": "25,000–100,000",
    "_100_300k": "100,000–300,000", "_300_500k": "300,000–500,000",
    "_500k_1m": "500,000–1,000,000", "_over_1m": "over 1,000,000",
    "judicial_fine_up_to_5k": "up to 5,000", "judicial_fine_5_25k": "5,000–25,000",
    "judicial_fine_25_100k": "25,000–100,000", "judicial_fine_over_100k": "over 100,000",
}


def bucket_label(key, ru_label):
    return BUCKET_LABELS_EN.get(key, ru_label) if lang == "en" else ru_label


# ---------- Page config ----------
st.set_page_config(
    page_title=t["page_title"],
    page_icon="⚖️",
    layout="centered",
)

# ---------- Load CSS ----------
css_path = Path(__file__).parent / "style.css"
if css_path.exists():
    st.markdown(f"<style>{css_path.read_text(encoding='utf-8')}</style>", unsafe_allow_html=True)

# ---------- Load data ----------
# Данные считаются на лету из исходных файлов Судебного департамента (data/raw/*.xls),
# названия статей берутся из article_names.json. Результат кэшируется.
@st.cache_data
def load_data():
    names = json.loads((Path(__file__).parent / "article_names.json").read_text(encoding="utf-8"))
    return forms.build(names)


data, sentencing_data = load_data()
articles = data["articles"]
years = data["years"]

by_num = {a["num"]: a for a in articles}
article_nums = [a["num"] for a in articles]


@st.cache_data
def read_bytes(path):
    return Path(path).read_bytes()


@st.cache_data
def table_csv():
    """Таблица, которую показывает приложение: статья × год (для скачивания)."""
    fields = ["convicted", "acquitted", "term_rehab", "term_nonrehab", "medical", "total",
              "pct_acquitted", "pct_convicted", "pct_term_rehab", "pct_term_nonrehab", "pct_medical", "pct_of_verdicts",
              "sent_prison", "sent_life", "sent_cond_prison", "sent_cond_other", "sent_fine", "sent_other_main",
              "sent_released", "sent_unallocated", "pct_cond_of_custodial"]
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(["article", "name", "year"] + fields)
    for a in articles:
        for y in years:
            yd = a["years"].get(str(y))            # None = данные за год не опубликованы
            w.writerow([a["num"], a["name"], y] + ([yd[k] for k in fields] if yd else [""] * len(fields)))
    return buf.getvalue().encode("utf-8-sig")


# ---------- Language toggle button + downloads of the source files ----------
top_l, top_r = st.columns([5, 1])
with top_l:
    with st.expander(t["dl_title"]):
        st.caption(t["dl_note"])
        for kind, label in (("10_3", t["dl_form_103"]), ("10_3_1", t["dl_form_1031"])):
            st.markdown(f"**{label}**")
            cols = st.columns(5)
            for i, y in enumerate(years):
                path = forms.RAW / f"form_{kind}_{y}.xls"
                cols[i % 5].download_button(str(y), data=read_bytes(path), file_name=path.name,
                                            mime="application/vnd.ms-excel", key=f"dl_{kind}_{y}")
        st.download_button(t["dl_csv"], data=table_csv(), file_name="outcomes_by_article_year.csv",
                           mime="text/csv", key="dl_csv")
with top_r:
    st.button(
        "EN" if lang == "ru" else "RU",
        on_click=toggle_lang,
        use_container_width=True,
        key="lang_toggle",
    )

# ---------- Header ----------
render(f"""
<div class="app-header">
<div class="app-header-badge">{t['badge']}</div>
<h1>{t['h1']}</h1>
<p class="app-subtitle">{t['subtitle']}</p>
<div class="app-intro">{html(t['intro'])}</div>
</div>
""")

# ---------- Controls ----------
col1, col2 = st.columns([2, 1])

with col1:
    render(f'<div class="field-label">{t["field_article"]}</div>')

    def format_article(num):
        a = by_num[num]
        return a["display"] if lang == "ru" else a["display_en"]

    selected_num = st.selectbox(
        t["field_article"],
        options=article_nums,
        index=0,
        format_func=format_article,
        label_visibility="collapsed",
        placeholder=t["placeholder_article"],
        key="article_select",
    )

with col2:
    render(f'<div class="field-label">{t["field_year"]}</div>')
    year_options = ["ALL"] + [str(y) for y in sorted(years, reverse=True)]

    def format_year(y):
        return t["all_years"] if y == "ALL" else y

    selected_year = st.selectbox(
        t["field_year"],
        options=year_options,
        index=0,
        format_func=format_year,
        label_visibility="collapsed",
        key="year_select",
    )

article = by_num.get(selected_num)
year_phrase = t["year_phrase_all"] if selected_year == "ALL" else (
    f"в {selected_year} году" if lang == "ru" else f"in {selected_year}"
)

render('<div class="divider"></div>')

# ---------- Result ----------
if article is None:
    st.warning(t["not_found"])
else:
    year_data = article["years"].get(selected_year)
    article_name = article["name"] if lang == "ru" else article["name_en"]
    unpublished_years = article.get("unpublished_years", [])
    year_unpublished = selected_year != "ALL" and int(selected_year) in unpublished_years

    render(f"""
    <div class="result-card">
    <div class="result-article">{t['article_word']} {article['num']}</div>
    <div class="result-name">{article_name}</div>
    """)

    if year_unpublished:
        render(f"""
        <div class="no-data no-data-unpublished">
        {t['no_data_unpublished'].format(year_phrase=year_phrase)}
        </div>
        </div>
        """)
    elif year_data is None or year_data["total"] == 0:
        render(f"""
        <div class="no-data">
        {t['no_data'].format(year_phrase=year_phrase)}
        </div>
        </div>
        """)
    else:
        def fmt(n):
            return f"{n:,}".replace(",", " ")

        total = year_data["total"]

        render(f"""
        <div class="total-caption">{t['total_caption'].format(year_phrase=year_phrase)}
        <span>{fmt(total)}</span></div>
        """)

        metrics = [
            (t["m_acquitted"], year_data["acquitted"], year_data["pct_acquitted"], "m-acquitted", True),
            (t["m_convicted"], year_data["convicted"], year_data["pct_convicted"], "m-convicted", False),
            (t["m_rehab"], year_data["term_rehab"], year_data["pct_term_rehab"], "m-rehab", False),
            (t["m_nonrehab"], year_data["term_nonrehab"], year_data["pct_term_nonrehab"], "m-nonrehab", False),
            (t["m_medical"], year_data["medical"], year_data["pct_medical"], "m-medical", False),
        ]

        cards = []
        for label, count, pct_val, css_class, featured in metrics:
            pct_val = pct_val or 0.0
            featured_class = "featured" if featured else ""
            cards.append(html(f"""
            <div class="metric-card {css_class} {featured_class}">
            <div class="metric-label">{label}</div>
            <div class="metric-row">
            <div class="metric-count">{fmt(count)}</div>
            <div class="metric-pct">{pct_val:.2f}%</div>
            </div>
            </div>
            """))

        cards_html = "".join(cards)
        render(f'<div class="metric-grid">{cards_html}</div>')
        partial_note = ""
        if selected_year == "ALL" and unpublished_years:
            covered = f"{min(years)}–{min(unpublished_years) - 1}"
            partial_note = f'<div class="metric-note metric-note-warn">{t["partial_all"].format(years=covered)}</div>'
        render(f"""
        <div class="metric-note">{t['metric_note']}</div>
        {partial_note}
        </div>
        """)

    # ---------- Trend across years ----------
    render(f'<div class="trend-title">{t["trend_title"]}</div>')

    max_pct = max(
        (v["pct_acquitted"] for v in article["years"].values() if v["pct_acquitted"] is not None),
        default=1,
    )
    max_pct = max(max_pct, 0.1)

    trend_rows = []
    for y in sorted(years):
        yd = article["years"].get(str(y))
        if int(y) in unpublished_years:
            bar_width = 0
            label = t["na_label"]
        elif yd is None or yd["pct_acquitted"] is None:
            bar_width = 0
            label = "—"
        else:
            bar_width = max(2, int(yd["pct_acquitted"] / max_pct * 100))
            label = f"{yd['pct_acquitted']:.2f}%"
        active_class = "active" if str(y) == selected_year else ""
        trend_rows.append(html(f"""
        <div class="trend-row {active_class}">
        <div class="trend-year">{y}</div>
        <div class="trend-bar-track">
        <div class="trend-bar-fill" style="width:{bar_width}%;"></div>
        </div>
        <div class="trend-pct">{label}</div>
        </div>
        """))

    render(f'<div class="trend-block">{"".join(trend_rows)}</div>')

    # ---------- Какое наказание получили осуждённые (форма 10.3) ----------
    if not year_unpublished and year_data and year_data["convicted"] > 0:
        def n_fmt(n):
            return f"{n:,}".replace(",", " ")

        conv = year_data["convicted"]
        kind_rows = []
        for key in ("prison", "life", "cond_prison", "cond_other", "fine", "other_main", "released", "unallocated"):
            n = year_data["sent_" + key]
            if n == 0 and key in ("other_main", "unallocated", "cond_other", "fine"):
                continue                       # пустые строки не показываем, кроме основных
            kind_rows.append(f"<tr><td>{t['k_' + key]}</td><td>{n_fmt(n)}</td><td>{n / conv * 100:.1f}%</td></tr>")
        cond_line = ""
        if year_data["custodial_total"] and year_data["sent_cond_prison"]:
            cond_line = '<div class="cond-share">' + t["cond_share"].format(
                pct=f"{year_data['pct_cond_of_custodial']:.1f}", cond=n_fmt(year_data["sent_cond_prison"]),
                custodial=n_fmt(year_data["custodial_total"])) + "</div>"
        render(f'<div class="trend-title">{t["kinds_title"]}</div>')
        render(f"""
        <div class="kinds-wrap">{cond_line}
        <table class="kinds-table">
        <thead><tr><th></th><th>{t['k_persons']}</th><th>{t['k_share']}</th></tr></thead>
        <tbody>{''.join(kind_rows)}</tbody>
        <tfoot><tr><td>{t['k_total']}</td><td>{n_fmt(conv)}</td><td>100%</td></tr></tfoot>
        </table>
        </div>
        """)

    # ---------- Sentencing sizes (form 10.3.1) ----------
    sd = sentencing_data.get("data", {}).get(selected_num) if sentencing_data else None
    year_sd = sd.get(selected_year) if sd else None

    if year_sd:
        def bucket_rows(buckets, denom):
            rows_ = []
            for key, ru_label in buckets:
                label = bucket_label(key, ru_label)
                count = year_sd.get(key, 0)
                pct_v = round(count / denom * 100, 1) if denom > 0 else 0
                bar_w = max(2, int(pct_v)) if count > 0 else 0
                rows_.append(html(f"""
                <div class="bucket-row">
                <div class="bucket-label">{label}</div>
                <div class="bucket-bar-track">
                <div class="bucket-bar-fill" style="width:{bar_w}%;"></div>
                </div>
                <div class="bucket-count">{count}</div>
                <div class="bucket-pct">{pct_v}%</div>
                </div>
                """))
            return "".join(rows_)

        term_buckets = sentencing_data["term_buckets"]
        fine_buckets = sentencing_data["fine_buckets"]
        jfine_buckets = sentencing_data["jfine_buckets"]

        sections = []

        # Imprisonment terms
        term_total = sum(year_sd.get(k, 0) for k, _ in term_buckets)
        if term_total > 0:
            below_min = year_sd.get("below_minimum_imprisonment", 0)
            rows_html = bucket_rows(term_buckets, term_total)
            below_note = f'<div class="sentencing-note">{t["below_min"].format(n=below_min)}</div>' if below_min else ""
            gap = year_sd.get("imprisonment_total", 0) - term_total          # л/с «всего» минус сумма интервалов
            gap_note = f'<div class="sentencing-note">{t["terms_gap"].format(n=gap)}</div>' if gap > 0 else ""
            below_note += gap_note + f'<div class="sentencing-note">{t["terms_note"]}</div>'
            sections.append(html(f"""
            <div class="sentencing-title">{t['term_title'].format(n=term_total)}</div>
            <div class="sentencing-block">{rows_html}</div>
            {below_note}
            """))

        # Main fine
        fine_main_total = sum(year_sd.get("fine_main" + k, 0) for k, _ in fine_buckets)
        if fine_main_total > 0:
            fine_rows = "".join(html(f"""
            <div class="bucket-row">
            <div class="bucket-label">{bucket_label(k, label)}</div>
            <div class="bucket-bar-track">
            <div class="bucket-bar-fill" style="width:{max(2, int(round(year_sd.get('fine_main'+k,0)/fine_main_total*100))) if year_sd.get('fine_main'+k,0) else 0}%;"></div>
            </div>
            <div class="bucket-count">{year_sd.get('fine_main'+k,0)}</div>
            <div class="bucket-pct">{round(year_sd.get('fine_main'+k,0)/fine_main_total*100,1)}%</div>
            </div>
            """) for k, label in fine_buckets)
            rub_note = ""
            total_rub = year_sd.get("fine_main_total_rub", 0)
            if total_rub:
                rub_note = f'<div class="sentencing-note">{t["rub_total"].format(n=f"{total_rub:,}".replace(",", " "))}</div>'
            elif selected_year != "ALL" and int(selected_year) < 2018:
                rub_note = f'<div class="sentencing-note">{t["rub_missing"]}</div>'
            sections.append(html(f"""
            <div class="sentencing-title">{t['fine_main_title'].format(n=fine_main_total)}</div>
            <div class="sentencing-block">{fine_rows}</div>
            {rub_note}
            """))

        # Additional fine
        fine_add_total = sum(year_sd.get("fine_additional" + k, 0) for k, _ in fine_buckets)
        if fine_add_total > 0:
            fine_add_rows = "".join(html(f"""
            <div class="bucket-row">
            <div class="bucket-label">{bucket_label(k, label)}</div>
            <div class="bucket-bar-track">
            <div class="bucket-bar-fill" style="width:{max(2, int(round(year_sd.get('fine_additional'+k,0)/fine_add_total*100))) if year_sd.get('fine_additional'+k,0) else 0}%;"></div>
            </div>
            <div class="bucket-count">{year_sd.get('fine_additional'+k,0)}</div>
            <div class="bucket-pct">{round(year_sd.get('fine_additional'+k,0)/fine_add_total*100,1)}%</div>
            </div>
            """) for k, label in fine_buckets)
            sections.append(html(f"""
            <div class="sentencing-title">{t['fine_add_title'].format(n=fine_add_total)}</div>
            <div class="sentencing-block">{fine_add_rows}</div>
            """))

        # Judicial fine (ст. 104.4)
        jfine_total = sum(year_sd.get(k, 0) for k, _ in jfine_buckets)
        if jfine_total > 0:
            jfine_rows = bucket_rows(jfine_buckets, jfine_total)
            sections.append(html(f"""
            <div class="sentencing-title">{t['jfine_title'].format(n=jfine_total)}</div>
            <div class="sentencing-block">{jfine_rows}</div>
            """))

        if sections:
            render(f'<div class="trend-title">{t["sentencing_title"]}</div>')
            render(f'<div class="sentencing-wrap">{"".join(sections)}</div>')

# ---------- Caveat for 2025 ----------
if selected_year in ("ALL", "2025"):
    render(f'<div class="year-warning">{html(t["year_warning"])}</div>')

# ---------- Footer ----------
footer_text = html(t['footer'])
render(f"""
<div class="footer-note">
{footer_text}
</div>
""")
