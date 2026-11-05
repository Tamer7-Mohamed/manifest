#!/usr/bin/env python3
"""
استخراج بيانات المانيفست من PDF (نص حقيقي) إلى Excel، صف لكل بوليصة،
مع تطبيق رولز Hapag-Lloyd Egypt.

التثبيت (مرة واحدة):
    pip install pdfplumber pandas openpyxl

الاستخدام:
    python manifest_to_excel_hapag_egypt.py file.pdf              -> file.xlsx
    python manifest_to_excel_hapag_egypt.py file.pdf out.xlsx
    python manifest_to_excel_hapag_egypt.py file.pdf --raw        -> يطبع النص الخام للتشخيص
"""
import json
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd
import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# ===== للتشغيل من VS Code: اكتب مسار الـ PDF هنا (بين علامتي التنصيص) واضغط Run =====
# مثال: PDF_PATH = r"C:\Users\FORMAT\Desktop\CAMN0001_605702373.pdf"
# لو سيبته فاضي هتفتحلك نافذة تختار منها الملف.
PDF_PATH = r""
# ===================================================================================

BL_RE = re.compile(r"^\s*([A-Z0-9]{8,})\s+(BL|SWB)\b")  # BL أو SWB (Sea Waybill)
STOP = re.compile(r"^\s*(SHIPPER|CONSIGNEE|NOTIFY|PLACE OF)")
CONT_RE = re.compile(r"([A-Z]{4})\s?(\d{7})\s*\((\w+)\)")
num = lambda s: float(s.replace(",", "") or 0)


COUNTRIES = {
    "ANDORRA": "AD", "PRINCIPALITY OF ANDORRA": "AD", "U.A.E": "AE", "UAE": "AE",
    "UNITED ARAB EMIRATES": "AE", "AFGHANISTAN": "AF", "ISLAMIC REPUBLIC OF AFGHANISTAN": "AF",
    "ANTIGUA AND BARBUDA": "AG", "ANGUILLA": "AI", "ALBANIA": "AL",
    "REPUBLIC OF ALBANIA": "AL", "ARMENIA": "AM", "REPUBLIC OF ARMENIA": "AM", "ANGOLA": "AO",
    "REPUBLIC OF ANGOLA": "AO", "ANTARCTICA": "AQ", "ARGENTINA": "AR",
    "ARGENTINE REPUBLIC": "AR", "AMERICAN SAMOA": "AS", "AUSTRIA": "AT",
    "REPUBLIC OF AUSTRIA": "AT", "AUSTRALIA": "AU", "ARUBA": "AW", "AZERBAIJAN": "AZ",
    "REPUBLIC OF AZERBAIJAN": "AZ", "BOSNIA AND HERZEGOVINA": "BA",
    "REPUBLIC OF BOSNIA AND HERZEGOVINA": "BA", "BARBADOS": "BB", "BANGLADESH": "BD",
    "PEOPLE'S REPUBLIC OF BANGLADESH": "BD", "BELGIUM": "BE", "KINGDOM OF BELGIUM": "BE",
    "BURKINA FASO": "BF", "BULGARIA": "BG", "REPUBLIC OF BULGARIA": "BG", "BAHRAIN": "BH",
    "KINGDOM OF BAHRAIN": "BH", "BURUNDI": "BI", "REPUBLIC OF BURUNDI": "BI", "BENIN": "BJ",
    "REPUBLIC OF BENIN": "BJ", "BERMUDA": "BM", "BRUNEI": "BN", "BRUNEI DARUSSALAM": "BN",
    "BOLIVIA": "BO", "PLURINATIONAL STATE OF BOLIVIA": "BO", "BRAZIL": "BR",
    "FEDERATIVE REPUBLIC OF BRAZIL": "BR", "BAHAMAS": "BS",
    "COMMONWEALTH OF THE BAHAMAS": "BS", "BHUTAN": "BT", "KINGDOM OF BHUTAN": "BT",
    "BOUVET ISLAND": "BV", "BOTSWANA": "BW", "REPUBLIC OF BOTSWANA": "BW", "BELARUS": "BY",
    "REPUBLIC OF BELARUS": "BY", "BELIZE": "BZ", "CANADA": "CA",
    "DEMOCRATIC REPUBLIC OF THE CONGO": "CD", "DR CONGO": "CD",
    "CENTRAL AFRICAN REPUBLIC": "CF", "CONGO": "CG", "REPUBLIC OF THE CONGO": "CG",
    "SWISS CONFEDERATION": "CH", "SWITZERLAND": "CH", "COTE D IVOIRE": "CI",
    "IVORY COAST": "CI", "COOK ISLANDS": "CK", "CHILE": "CL", "REPUBLIC OF CHILE": "CL",
    "CAMEROON": "CM", "REPUBLIC OF CAMEROON": "CM", "CHINA": "CN",
    "PEOPLE'S REPUBLIC OF CHINA": "CN", "COLOMBIA": "CO", "REPUBLIC OF COLOMBIA": "CO",
    "COSTA RICA": "CR", "REPUBLIC OF COSTA RICA": "CR", "CUBA": "CU", "REPUBLIC OF CUBA": "CU",
    "CABO VERDE": "CV", "CAPE VERDE": "CV", "REPUBLIC OF CABO VERDE": "CV",
    "CHRISTMAS ISLAND": "CX", "CYPRUS": "CY", "REPUBLIC OF CYPRUS": "CY",
    "CZECH REPUBLIC": "CZ", "CZECHIA": "CZ", "FEDERAL REPUBLIC OF GERMANY": "DE",
    "GERMANY": "DE", "DJIBOUTI": "DJ", "REPUBLIC OF DJIBOUTI": "DJ", "DENMARK": "DK",
    "KINGDOM OF DENMARK": "DK", "COMMONWEALTH OF DOMINICA": "DM", "DOMINICA": "DM",
    "DOMINICAN REPUBLIC": "DO", "ALGERIA": "DZ",
    "PEOPLE'S DEMOCRATIC REPUBLIC OF ALGERIA": "DZ", "ECUADOR": "EC",
    "REPUBLIC OF ECUADOR": "EC", "ESTONIA": "EE", "REPUBLIC OF ESTONIA": "EE",
    "ARAB REPUBLIC OF EGYPT": "EG", "EGYPT": "EG", "WESTERN SAHARA": "EH", "ERITREA": "ER",
    "THE STATE OF ERITREA": "ER", "KINGDOM OF SPAIN": "ES", "SPAIN": "ES", "ETHIOPIA": "ET",
    "FEDERAL DEMOCRATIC REPUBLIC OF ETHIOPIA": "ET", "FINLAND": "FI",
    "REPUBLIC OF FINLAND": "FI", "FIJI": "FJ", "REPUBLIC OF FIJI": "FJ",
    "FALKLAND ISLANDS": "FK", "FEDERATED STATES OF MICRONESIA": "FM", "FAROE ISLANDS": "FO",
    "FRANCE": "FR", "FRENCH REPUBLIC": "FR", "GABON": "GA", "GABONESE REPUBLIC": "GA",
    "ENGLAND": "GB", "GREAT BRITAIN": "GB", "SCOTLAND": "GB", "U.K": "GB", "UK": "GB",
    "UNITED KINGDOM": "GB", "UNITED KINGDOM OF GREAT BRITAIN AND NORTHERN IRELAND": "GB",
    "GRENADA": "GD", "GEORGIA": "GE", "FRENCH GUIANA": "GF", "GUERNSEY": "GG", "GHANA": "GH",
    "REPUBLIC OF GHANA": "GH", "GIBRALTAR": "GI", "GREENLAND": "GL", "GAMBIA": "GM",
    "REPUBLIC OF THE GAMBIA": "GM", "GUINEA": "GN", "REPUBLIC OF GUINEA": "GN",
    "GUADELOUPE": "GP", "EQUATORIAL GUINEA": "GQ", "REPUBLIC OF EQUATORIAL GUINEA": "GQ",
    "GREECE": "GR", "HELLENIC REPUBLIC": "GR",
    "SOUTH GEORGIA AND THE SOUTH SANDWICH ISLANDS": "GS", "GUATEMALA": "GT",
    "REPUBLIC OF GUATEMALA": "GT", "GUAM": "GU", "GUINEA-BISSAU": "GW",
    "REPUBLIC OF GUINEA-BISSAU": "GW", "GUYANA": "GY", "REPUBLIC OF GUYANA": "GY",
    "HONG KONG": "HK", "HONG KONG SPECIAL ADMINISTRATIVE REGION OF CHINA": "HK",
    "HEARD ISLAND AND MCDONALD ISLANDS": "HM", "HONDURAS": "HN", "REPUBLIC OF HONDURAS": "HN",
    "CROATIA": "HR", "REPUBLIC OF CROATIA": "HR", "HAITI": "HT", "REPUBLIC OF HAITI": "HT",
    "HUNGARY": "HU", "INDONESIA": "ID", "REPUBLIC OF INDONESIA": "ID", "IRELAND": "IE",
    "ISRAEL": "IL", "STATE OF ISRAEL": "IL", "ISLE OF MAN": "IM", "INDIA": "IN",
    "REPUBLIC OF INDIA": "IN", "BRITISH INDIAN OCEAN TERRITORY": "IO", "IRAQ": "IQ",
    "REPUBLIC OF IRAQ": "IQ", "IRAN": "IR", "ISLAMIC REPUBLIC OF IRAN": "IR", "ICELAND": "IS",
    "REPUBLIC OF ICELAND": "IS", "ITALIAN REPUBLIC": "IT", "ITALY": "IT", "JERSEY": "JE",
    "JAMAICA": "JM", "HASHEMITE KINGDOM OF JORDAN": "JO", "JORDAN": "JO", "JAPAN": "JP",
    "KENYA": "KE", "REPUBLIC OF KENYA": "KE", "KYRGYZ REPUBLIC": "KG", "KYRGYZSTAN": "KG",
    "CAMBODIA": "KH", "KINGDOM OF CAMBODIA": "KH", "KIRIBATI": "KI",
    "REPUBLIC OF KIRIBATI": "KI", "COMOROS": "KM", "UNION OF THE COMOROS": "KM",
    "SAINT KITTS AND NEVIS": "KN", "DEMOCRATIC PEOPLE'S REPUBLIC OF KOREA": "KP",
    "NORTH KOREA": "KP", "KOREA": "KR", "SOUTH KOREA": "KR", "KUWAIT": "KW",
    "STATE OF KUWAIT": "KW", "CAYMAN ISLANDS": "KY", "KAZAKHSTAN": "KZ",
    "REPUBLIC OF KAZAKHSTAN": "KZ", "LAO PEOPLE'S DEMOCRATIC REPUBLIC": "LA", "LAOS": "LA",
    "LEBANESE REPUBLIC": "LB", "LEBANON": "LB", "SAINT LUCIA": "LC", "LIECHTENSTEIN": "LI",
    "PRINCIPALITY OF LIECHTENSTEIN": "LI", "DEMOCRATIC SOCIALIST REPUBLIC OF SRI LANKA": "LK",
    "SRI LANKA": "LK", "LIBERIA": "LR", "REPUBLIC OF LIBERIA": "LR",
    "KINGDOM OF LESOTHO": "LS", "LESOTHO": "LS", "LITHUANIA": "LT",
    "REPUBLIC OF LITHUANIA": "LT", "GRAND DUCHY OF LUXEMBOURG": "LU", "LUXEMBOURG": "LU",
    "LATVIA": "LV", "REPUBLIC OF LATVIA": "LV", "LIBYA": "LY", "KINGDOM OF MOROCCO": "MA",
    "MOROCCO": "MA", "MONACO": "MC", "PRINCIPALITY OF MONACO": "MC", "MOLDOVA": "MD",
    "REPUBLIC OF MOLDOVA": "MD", "MONTENEGRO": "ME", "MADAGASCAR": "MG",
    "REPUBLIC OF MADAGASCAR": "MG", "MARSHALL ISLANDS": "MH",
    "REPUBLIC OF THE MARSHALL ISLANDS": "MH", "MACEDONIA": "MK", "NORTH MACEDONIA": "MK",
    "REPUBLIC OF NORTH MACEDONIA": "MK", "MALI": "ML", "REPUBLIC OF MALI": "ML", "BURMA": "MM",
    "MYANMAR": "MM", "REPUBLIC OF MYANMAR": "MM", "MONGOLIA": "MN", "MACAO": "MO",
    "MACAO SPECIAL ADMINISTRATIVE REGION OF CHINA": "MO", "MACAU": "MO",
    "COMMONWEALTH OF THE NORTHERN MARIANA ISLANDS": "MP", "NORTHERN MARIANA ISLANDS": "MP",
    "MARTINIQUE": "MQ", "ISLAMIC REPUBLIC OF MAURITANIA": "MR", "MAURITANIA": "MR",
    "MONTSERRAT": "MS", "MALTA": "MT", "REPUBLIC OF MALTA": "MT", "MAURITIUS": "MU",
    "REPUBLIC OF MAURITIUS": "MU", "MALDIVES": "MV", "REPUBLIC OF MALDIVES": "MV",
    "MALAWI": "MW", "REPUBLIC OF MALAWI": "MW", "MEXICO": "MX", "UNITED MEXICAN STATES": "MX",
    "MALAYSIA": "MY", "MOZAMBIQUE": "MZ", "REPUBLIC OF MOZAMBIQUE": "MZ", "NAMIBIA": "NA",
    "REPUBLIC OF NAMIBIA": "NA", "NEW CALEDONIA": "NC", "NIGER": "NE",
    "REPUBLIC OF THE NIGER": "NE", "NORFOLK ISLAND": "NF", "FEDERAL REPUBLIC OF NIGERIA": "NG",
    "NIGERIA": "NG", "NICARAGUA": "NI", "REPUBLIC OF NICARAGUA": "NI", "HOLLAND": "NL",
    "KINGDOM OF THE NETHERLANDS": "NL", "NETHERLANDS": "NL", "THE NETHERLANDS": "NL",
    "KINGDOM OF NORWAY": "NO", "NORWAY": "NO", "FEDERAL DEMOCRATIC REPUBLIC OF NEPAL": "NP",
    "NEPAL": "NP", "NAURU": "NR", "REPUBLIC OF NAURU": "NR", "NIUE": "NU", "NEW ZEALAND": "NZ",
    "OMAN": "OM", "SULTANATE OF OMAN": "OM", "PANAMA": "PA", "REPUBLIC OF PANAMA": "PA",
    "PERU": "PE", "REPUBLIC OF PERU": "PE", "FRENCH POLYNESIA": "PF",
    "INDEPENDENT STATE OF PAPUA NEW GUINEA": "PG", "PAPUA NEW GUINEA": "PG",
    "PHILIPPINES": "PH", "REPUBLIC OF THE PHILIPPINES": "PH",
    "ISLAMIC REPUBLIC OF PAKISTAN": "PK", "PAKISTAN": "PK", "POLAND": "PL",
    "REPUBLIC OF POLAND": "PL", "SAINT PIERRE AND MIQUELON": "PM", "PITCAIRN": "PN",
    "PUERTO RICO": "PR", "PALESTINE": "PS", "THE STATE OF PALESTINE": "PS", "PORTUGAL": "PT",
    "PORTUGUESE REPUBLIC": "PT", "PALAU": "PW", "REPUBLIC OF PALAU": "PW", "PARAGUAY": "PY",
    "REPUBLIC OF PARAGUAY": "PY", "QATAR": "QA", "STATE OF QATAR": "QA", "ROMANIA": "RO",
    "REPUBLIC OF SERBIA": "RS", "SERBIA": "RS", "RUSSIA": "RU", "RUSSIAN FEDERATION": "RU",
    "RWANDA": "RW", "RWANDESE REPUBLIC": "RW", "KINGDOM OF SAUDI ARABIA": "SA",
    "SAUDI ARABIA": "SA", "SOLOMON ISLANDS": "SB", "REPUBLIC OF SEYCHELLES": "SC",
    "SEYCHELLES": "SC", "REPUBLIC OF THE SUDAN": "SD", "SUDAN": "SD",
    "KINGDOM OF SWEDEN": "SE", "SWEDEN": "SE", "REPUBLIC OF SINGAPORE": "SG",
    "SINGAPORE": "SG", "REPUBLIC OF SLOVENIA": "SI", "SLOVENIA": "SI",
    "SVALBARD AND JAN MAYEN": "SJ", "SLOVAK REPUBLIC": "SK", "SLOVAKIA": "SK",
    "REPUBLIC OF SIERRA LEONE": "SL", "SIERRA LEONE": "SL", "REPUBLIC OF SAN MARINO": "SM",
    "SAN MARINO": "SM", "REPUBLIC OF SENEGAL": "SN", "SENEGAL": "SN",
    "FEDERAL REPUBLIC OF SOMALIA": "SO", "SOMALIA": "SO", "REPUBLIC OF SURINAME": "SR",
    "SURINAME": "SR", "REPUBLIC OF SOUTH SUDAN": "SS", "SOUTH SUDAN": "SS",
    "DEMOCRATIC REPUBLIC OF SAO TOME AND PRINCIPE": "ST", "SAO TOME AND PRINCIPE": "ST",
    "EL SALVADOR": "SV", "REPUBLIC OF EL SALVADOR": "SV", "SYRIA": "SY",
    "SYRIAN ARAB REPUBLIC": "SY", "ESWATINI": "SZ", "KINGDOM OF ESWATINI": "SZ",
    "SWAZILAND": "SZ", "TURKS AND CAICOS ISLANDS": "TC", "CHAD": "TD",
    "REPUBLIC OF CHAD": "TD", "FRENCH SOUTHERN TERRITORIES": "TF", "TOGO": "TG",
    "TOGOLESE REPUBLIC": "TG", "KINGDOM OF THAILAND": "TH", "THAILAND": "TH",
    "REPUBLIC OF TAJIKISTAN": "TJ", "TAJIKISTAN": "TJ", "TOKELAU": "TK",
    "DEMOCRATIC REPUBLIC OF TIMOR-LESTE": "TL", "EAST TIMOR": "TL", "TIMOR-LESTE": "TL",
    "TURKMENISTAN": "TM", "REPUBLIC OF TUNISIA": "TN", "TUNISIA": "TN",
    "KINGDOM OF TONGA": "TO", "TONGA": "TO", "TURKEY": "TR", "TURKIYE": "TR",
    "REPUBLIC OF TRINIDAD AND TOBAGO": "TT", "TRINIDAD AND TOBAGO": "TT", "TUVALU": "TV",
    "TAIWAN": "TW", "TANZANIA": "TZ", "UNITED REPUBLIC OF TANZANIA": "TZ", "UKRAINE": "UA",
    "REPUBLIC OF UGANDA": "UG", "UGANDA": "UG", "UNITED STATES MINOR OUTLYING ISLANDS": "UM",
    "U.S.": "US", "U.S.A": "US", "UNITED STATES": "US", "UNITED STATES OF AMERICA": "US",
    "USA": "US", "EASTERN REPUBLIC OF URUGUAY": "UY", "URUGUAY": "UY",
    "REPUBLIC OF UZBEKISTAN": "UZ", "UZBEKISTAN": "UZ",
    "SAINT VINCENT AND THE GRENADINES": "VC", "BOLIVARIAN REPUBLIC OF VENEZUELA": "VE",
    "VENEZUELA": "VE", "BRITISH VIRGIN ISLANDS": "VG",
    "VIRGIN ISLANDS OF THE UNITED STATES": "VI", "SOCIALIST REPUBLIC OF VIET NAM": "VN",
    "VIET NAM": "VN", "VIETNAM": "VN", "REPUBLIC OF VANUATU": "VU", "VANUATU": "VU",
    "WALLIS AND FUTUNA": "WF", "INDEPENDENT STATE OF SAMOA": "WS", "SAMOA": "WS",
    "KOSOVO": "XK", "REPUBLIC OF YEMEN": "YE", "YEMEN": "YE", "MAYOTTE": "YT",
    "REPUBLIC OF SOUTH AFRICA": "ZA", "SOUTH AFRICA": "ZA", "REPUBLIC OF ZAMBIA": "ZM",
    "ZAMBIA": "ZM", "REPUBLIC OF ZIMBABWE": "ZW", "ZIMBABWE": "ZW",
}

COUNTRY_NAMES = {
    "AD": "Andorra", "AE": "United Arab Emirates", "AF": "Afghanistan",
    "AG": "Antigua and Barbuda", "AI": "Anguilla", "AL": "Albania", "AM": "Armenia",
    "AO": "Angola", "AQ": "Antarctica", "AR": "Argentina", "AS": "American Samoa",
    "AT": "Austria", "AU": "Australia", "AW": "Aruba", "AX": "Åland Islands",
    "AZ": "Azerbaijan", "BA": "Bosnia and Herzegovina", "BB": "Barbados", "BD": "Bangladesh",
    "BE": "Belgium", "BF": "Burkina Faso", "BG": "Bulgaria", "BH": "Bahrain", "BI": "Burundi",
    "BJ": "Benin", "BL": "Saint Barthélemy", "BM": "Bermuda", "BN": "Brunei", "BO": "Bolivia",
    "BQ": "Bonaire, Sint Eustatius and Saba", "BR": "Brazil", "BS": "Bahamas", "BT": "Bhutan",
    "BV": "Bouvet Island", "BW": "Botswana", "BY": "Belarus", "BZ": "Belize", "CA": "Canada",
    "CC": "Cocos (Keeling) Islands", "CD": "Congo, The Democratic Republic of the",
    "CF": "Central African Republic", "CG": "Congo", "CH": "Switzerland",
    "CI": "Côte d'Ivoire", "CK": "Cook Islands", "CL": "Chile", "CM": "Cameroon",
    "CN": "China", "CO": "Colombia", "CR": "Costa Rica", "CU": "Cuba", "CV": "Cabo Verde",
    "CW": "Curaçao", "CX": "Christmas Island", "CY": "Cyprus", "CZ": "Czechia",
    "DE": "Germany", "DJ": "Djibouti", "DK": "Denmark", "DM": "Dominica",
    "DO": "Dominican Republic", "DZ": "Algeria", "EC": "Ecuador", "EE": "Estonia",
    "EG": "Egypt", "EH": "Western Sahara", "ER": "Eritrea", "ES": "Spain", "ET": "Ethiopia",
    "FI": "Finland", "FJ": "Fiji", "FK": "Falkland Islands (Malvinas)",
    "FM": "Micronesia, Federated States of", "FO": "Faroe Islands", "FR": "France",
    "GA": "Gabon", "GB": "United Kingdom", "GD": "Grenada", "GE": "Georgia",
    "GF": "French Guiana", "GG": "Guernsey", "GH": "Ghana", "GI": "Gibraltar",
    "GL": "Greenland", "GM": "Gambia", "GN": "Guinea", "GP": "Guadeloupe",
    "GQ": "Equatorial Guinea", "GR": "Greece",
    "GS": "South Georgia and the South Sandwich Islands", "GT": "Guatemala", "GU": "Guam",
    "GW": "Guinea-Bissau", "GY": "Guyana", "HK": "Hong Kong",
    "HM": "Heard Island and McDonald Islands", "HN": "Honduras", "HR": "Croatia",
    "HT": "Haiti", "HU": "Hungary", "ID": "Indonesia", "IE": "Ireland", "IL": "Israel",
    "IM": "Isle of Man", "IN": "India", "IO": "British Indian Ocean Territory", "IQ": "Iraq",
    "IR": "Iran", "IS": "Iceland", "IT": "Italy", "JE": "Jersey", "JM": "Jamaica",
    "JO": "Jordan", "JP": "Japan", "KE": "Kenya", "KG": "Kyrgyzstan", "KH": "Cambodia",
    "KI": "Kiribati", "KM": "Comoros", "KN": "Saint Kitts and Nevis", "KP": "North Korea",
    "KR": "South Korea", "KW": "Kuwait", "KY": "Cayman Islands", "KZ": "Kazakhstan",
    "LA": "Laos", "LB": "Lebanon", "LC": "Saint Lucia", "LI": "Liechtenstein",
    "LK": "Sri Lanka", "LR": "Liberia", "LS": "Lesotho", "LT": "Lithuania", "LU": "Luxembourg",
    "LV": "Latvia", "LY": "Libya", "MA": "Morocco", "MC": "Monaco", "MD": "Moldova",
    "ME": "Montenegro", "MF": "Saint Martin (French part)", "MG": "Madagascar",
    "MH": "Marshall Islands", "MK": "North Macedonia", "ML": "Mali", "MM": "Myanmar",
    "MN": "Mongolia", "MO": "Macao", "MP": "Northern Mariana Islands", "MQ": "Martinique",
    "MR": "Mauritania", "MS": "Montserrat", "MT": "Malta", "MU": "Mauritius", "MV": "Maldives",
    "MW": "Malawi", "MX": "Mexico", "MY": "Malaysia", "MZ": "Mozambique", "NA": "Namibia",
    "NC": "New Caledonia", "NE": "Niger", "NF": "Norfolk Island", "NG": "Nigeria",
    "NI": "Nicaragua", "NL": "Netherlands", "NO": "Norway", "NP": "Nepal", "NR": "Nauru",
    "NU": "Niue", "NZ": "New Zealand", "OM": "Oman", "PA": "Panama", "PE": "Peru",
    "PF": "French Polynesia", "PG": "Papua New Guinea", "PH": "Philippines", "PK": "Pakistan",
    "PL": "Poland", "PM": "Saint Pierre and Miquelon", "PN": "Pitcairn", "PR": "Puerto Rico",
    "PS": "Palestine", "PT": "Portugal", "PW": "Palau", "PY": "Paraguay", "QA": "Qatar",
    "RE": "Réunion", "RO": "Romania", "RS": "Serbia", "RU": "Russia", "RW": "Rwanda",
    "SA": "Saudi Arabia", "SB": "Solomon Islands", "SC": "Seychelles", "SD": "Sudan",
    "SE": "Sweden", "SG": "Singapore", "SH": "Saint Helena, Ascension and Tristan da Cunha",
    "SI": "Slovenia", "SJ": "Svalbard and Jan Mayen", "SK": "Slovakia", "SL": "Sierra Leone",
    "SM": "San Marino", "SN": "Senegal", "SO": "Somalia", "SR": "Suriname",
    "SS": "South Sudan", "ST": "Sao Tome and Principe", "SV": "El Salvador",
    "SX": "Sint Maarten (Dutch part)", "SY": "Syria", "SZ": "Eswatini",
    "TC": "Turks and Caicos Islands", "TD": "Chad", "TF": "French Southern Territories",
    "TG": "Togo", "TH": "Thailand", "TJ": "Tajikistan", "TK": "Tokelau", "TL": "Timor-Leste",
    "TM": "Turkmenistan", "TN": "Tunisia", "TO": "Tonga", "TR": "Turkey",
    "TT": "Trinidad and Tobago", "TV": "Tuvalu", "TW": "Taiwan", "TZ": "Tanzania",
    "UA": "Ukraine", "UG": "Uganda", "UM": "United States Minor Outlying Islands",
    "US": "United States", "UY": "Uruguay", "UZ": "Uzbekistan",
    "VA": "Holy See (Vatican City State)", "VC": "Saint Vincent and the Grenadines",
    "VE": "Venezuela", "VG": "Virgin Islands, British", "VI": "Virgin Islands, U.S.",
    "VN": "Vietnam", "VU": "Vanuatu", "WF": "Wallis and Futuna", "WS": "Samoa", "XK": "Kosovo",
    "YE": "Yemen", "YT": "Mayotte", "ZA": "South Africa", "ZM": "Zambia", "ZW": "Zimbabwe",
}

_CRE = re.compile(r"(?<![A-Z])(" + "|".join(sorted(map(re.escape, COUNTRIES), key=len, reverse=True)) + r")(?![A-Z])")


def country_code(shipper_lines, alltxt, place_of_receipt):
    """كود دولة الشاحن: 1) حقل COUNTRY  2) آخر دولة مذكورة في عنوانه  3) أول حرفين من كود ميناء الشحن."""
    m = re.search(r"EXPORTER REGISTRATION\s*COUNTRY\s*:\s*([A-Z. ]+)", alltxt)
    cand = [m.group(1).strip()] if m else []
    cand.append(" ".join(shipper_lines).upper())
    for t in cand:
        found = _CRE.findall(t)
        if found:
            return COUNTRIES[found[-1]]
    m = re.match(r"([A-Z]{2})[A-Z0-9]{3}\b", place_of_receipt or "")
    return m.group(1) if m else ""


TRANSIT_RE = re.compile(r"\bTRANSIT\b|\bIN[- ]TRANSIT\b|ترانزيت", re.I)
FREEZONE_RE = re.compile(
    r"\bFREE[\s-]*(?:TRADE[\s-]*)?(?:ZONE|AREA)S?\b|\bFREEZONE\b|\bFTZ\b|\bFZ\b|\bFZE\b|\bFZCO\b|\bFZC\b|\bFZ-?LLC\b|منطقة حرة|المنطقة الحرة",
    re.I)


def find_flags(full_lines):
    """يدوّر على ترانزيت / فري زون في بيانات البوليصة (بدون صفحات الشروط CLAUSES)."""
    end = next((i for i, l in enumerate(full_lines) if "CLAUSES" in l), len(full_lines))
    body = full_lines[:end]

    def hits(rx):
        out = []
        for l in body:
            for m in rx.finditer(l):
                a, b = max(0, m.start() - 25), min(len(l), m.end() + 25)
                t = re.sub(r"\s+", " ", l[a:b]).strip()
                if t not in out:
                    out.append(t)
        return " | ".join(out)
    return hits(TRANSIT_RE), hits(FREEZONE_RE)


def page_lines(page):
    """سطور الصفحة بنفس التخطيط (الخط monospace فالأعمدة بتتحفظ)."""
    txt = page.extract_text(layout=True) or ""
    return [l.rstrip() for l in txt.split("\n")]


def read_blocks(pdf_path, raw=False):
    """يقسم الملف لبلوكات، كل بلوك = بوليصة واحدة (ممكن تمتد على أكتر من صفحة)."""
    blocks, cur = [], None
    with pdfplumber.open(pdf_path) as pdf:
        for n, page in enumerate(pdf.pages, 1):
            lines = page_lines(page)
            if raw:
                print(f"===== PAGE {n} =====")
                print("\n".join(lines))
            hi = next((i for i, l in enumerate(lines)
                       if "MARKS" in l and "QUANTITY" in l), None)
            if hi is None:
                continue
            cut = lines[hi].index("QUANTITY")  # حدّ عمود الوصف
            for l in lines[hi + 1:]:
                if not l.strip() or re.match(r"^\s*-{10,}\s*$", l) \
                        or re.match(r"^\s*DOCUMENT\b", l) \
                        or "CONTINUED" in l:
                    continue
                left, right = l[:cut], l[cut:]
                if BL_RE.match(left):
                    cur = {"left": [], "right": []}
                    blocks.append(cur)
                if cur is not None:
                    cur["left"].append(left)
                    cur["right"].append(right)
    return blocks


def clean(lines):
    out = []
    for l in lines:
        l = re.sub(r"\s*BILL TOTALS.*", "", l).strip()
        if l and not re.fullmatch(r"[\d,]+", l) and not re.fullmatch(r"\(\d{3}\)", l):
            out.append(l)
    return out


def section(left, start_re, stop_re=STOP):
    """السطور من عنوان القسم لحد أول عنوان قسم تاني."""
    i = next((k for k, l in enumerate(left) if re.match(start_re, l)), None)
    if i is None:
        return []
    j = next((k for k in range(i + 1, len(left)) if stop_re.match(left[k])), len(left))
    first = re.sub(start_re, "", left[i], count=1).strip()
    return clean([first] + left[i + 1:j])


def find_id(lines):
    m = re.search(r"\b(?:TAX\s*ID|TAX\s*NO|VAT|REG(?:ISTRATION)?\s*NO|ID)\s*[:.]?\s*([A-Z0-9][A-Z0-9-]{4,})",
                  " ".join(lines), re.I)
    return m.group(1) if m else ""


def parse_block(b):
    left, right = b["left"], b["right"]
    ltxt, rtxt = "\n".join(left), "\n".join(right)
    alltxt = ltxt + "\n" + rtxt

    mbl = BL_RE.match(left[0])
    bl, doc_type = mbl.group(1), mbl.group(2)
    ref = (re.search(r"HL-REF:?\s*(\S+)", alltxt) or [None, ""])[1]

    shipper = section(left, r"^\s*SHIPPER:?")
    consignee = section(left, r"^\s*CONSIGNEE:?")
    notify = section(left, r"^\s*NOTIFY PARTY 1:?")
    first = lambda L: L[0] if L else ""
    por = first(section(left, r"^\s*PLACE OF RECEIPT.*?:"))
    pod = first(section(left, r"^\s*PLACE OF DELIVERY.*?:"))

    m = re.search(r"ACID\D{0,12}(\d{10,})", alltxt, re.I) or re.search(r"\b(\d{19})\b", alltxt)
    acid = m.group(1) if m else ""
    m = re.search(r"SHIPPER\s*ID\s*:\s*(\S+)", alltxt)
    shipper_id = m.group(1) if m else find_id(shipper)
    m = re.search(r"CONSIGNEE\s*(?:VAT|ID|TAX\s*ID)\s*:\s*(\S+)", alltxt)
    consignee_id = m.group(1) if m else find_id(consignee)

    # فصل سطور الأصناف عن سطور BILL TOTALS (عشان الأوزان ماتتحسبش مرتين)
    full = [a + b_ for a, b_ in zip(left, right)]
    ti = next((i for i, l in enumerate(full) if "BILL TOTALS" in l), None)
    items = full[:ti] if ti is not None else full
    itxt = "\n".join(items)
    ritems = "\n".join(right[:ti] if ti is not None else right)
    tot = "\n".join(full[ti:]) if ti is not None else ""

    UNIT = r"(?:UNITS?|CARTONS?|PALLETS?|PACKAGES?|PKGS?|BAGS?|DRUMS?|BALES?|ROLLS?|CRATES?|BOXES|PIECES|PCS)"
    cartons = [int(num(x)) for x in re.findall(r"(?<![\d.,])(\d[\d,]*)\s+" + UNIT + r"\b", itxt)]
    gross = [num(x) for x in re.findall(r"([\d,]+\.\d+)\s*KGM\s*GROSS", ritems)]
    vol = [num(x) for x in re.findall(r"([\d,]+\.\d+)\s*MTQ\s*GROSS", ritems)]
    hs = [x.strip() for x in re.findall(r"HS-CODE\s*:?\s*([\d ]+)", ritems)]
    comm = [x for x in re.findall(r"COMMODITY\s*:\s*(\d+)", ritems)]
    conts = CONT_RE.findall(ritems)
    seals = re.findall(r"SEAL NO\s*\d*:?\s*(\S+)", ritems)
    tare = re.findall(r"TARE WEIGHT:?\s*(\d[\d,]*)", ritems)

    tg = [num(x) for x in re.findall(r"([\d,]+\.\d+)\s*KGM\s*GROSS", tot)]
    tv = [num(x) for x in re.findall(r"([\d,]+\.\d+)\s*MTQ\s*GROSS", tot)]
    tc = re.search(r"BILL TOTALS\s+(\d[\d,]*)", tot)
    tot_cartons = int(num(tc.group(1))) if tc else sum(cartons)

    transit, freezone = find_flags(full)
    ccode = country_code(shipper, alltxt, por)
    pref = re.match(r"([A-Z]{2})(?![A-Z])", shipper_id.upper())
    id_ok = bool(ccode and pref and pref.group(1) == ccode)

    summary = {
        "BL": bl, "Doc Type": doc_type, "HL-REF": ref,
        "Shipper": " | ".join(shipper), "Shipper ID": shipper_id,
        "Shipper Country": COUNTRY_NAMES.get(ccode, ""),
        "Shipper Country Code": ccode,
        "Shipper ID Check": shipper_id if id_ok else "راجع المانيفست",
        "Consignee": " | ".join(consignee), "Consignee ID": consignee_id,
        "Notify": " | ".join(notify), "Notify ID": consignee_id if "SAME AS CONSIGNEE" in " ".join(notify) else find_id(notify),
        "ACID": acid, "Place of Receipt": por, "Place of Delivery": pod,
        "Transit": transit, "Free Zone": freezone, "Part of 2 Cargo": "YES" if _part_of_two(alltxt) else "",
        "Containers": " ; ".join(a + "  " + b_ for a, b_, _ in conts),
        "Seals": " ; ".join(seals),
        "HS Code": " ; ".join(dict.fromkeys(hs)),
        "Commodity Code": " ; ".join(dict.fromkeys(comm)),
        "Cartons": tot_cartons,
        "Gross KGM": round(tg[0] if tg else sum(gross), 3),
        "Volume MTQ": round(tv[0] if tv else sum(vol), 3),
    }
    g = lambda lst, i: lst[i] if i < len(lst) else ""
    details = [{
        "BL": bl, "Container": a + "  " + b_, "Type": t, "Seal": g(seals, i),
        "Tare": g(tare, i), "Cartons": g(cartons, i),
        "Gross KGM": g(gross, i), "Volume MTQ": g(vol, i), "HS Code": g(hs, i),
    } for i, (a, b_, t) in enumerate(conts)]
    return summary, details


# ===================== النسخة السريعة (بتتنادى من index.html) =====================
_pdf = None
_BAD = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


def open_pdf(path):
    global _pdf
    if _pdf is not None:
        try:
            _pdf.close()
        except Exception:
            pass
    _pdf = pdfplumber.open(path)
    return len(_pdf.pages)


def process(a, b):
    """يقرأ الصفحات [a, b) ويرجّع البلوكات JSON.
    أول بلوك ممكن يكون orphan = تكملة بوليصة بدأت في الدفعة اللي قبلها."""
    blocks, cur = [], None
    for n in range(a, b):
        page = _pdf.pages[n]
        lines = page_lines(page)
        page.flush_cache()
        hi = next((i for i, l in enumerate(lines) if "MARKS" in l and "QUANTITY" in l), None)
        if hi is None:
            continue
        cut = lines[hi].index("QUANTITY")
        for l in lines[hi + 1:]:
            if not l.strip() or re.match(r"^\s*-{10,}\s*$", l) \
                    or re.match(r"^\s*DOCUMENT\b", l) or "CONTINUED" in l:
                continue
            left, right = l[:cut], l[cut:]
            if BL_RE.match(left):
                cur = {"left": [], "right": []}
                blocks.append(cur)
            elif cur is None:
                cur = {"left": [], "right": [], "orphan": True}
                blocks.append(cur)
            cur["left"].append(left)
            cur["right"].append(right)
    return json.dumps(blocks)


# ===================== HAPAG-LLOYD EGYPT RULES =====================
MAX_20_KG = 25000       # حد وزن حاوية 20'
SHARED_20_KG = 24000    # حاوية 20' مشتركة بين بوليصتين: خطأ لو أي وزن >= الرقم ده
MAX_40_KG = 50000       # حد وزن حاوية 40' (غيّره هنا لو الحد الفعلي مختلف)

ERR, WARN = "خطأ", "تنبيه"

_EGYPT_RE = re.compile(
    r"\bEGYPT\b|\bEG\b|\bCAIRO\b|\bALEXANDRIA\b|\bGIZA\b|\bDAMIETTA\b|\bPORT\s*SAID\b|\bSUEZ\b"
    r"|\bSOKHNA\b|\bISMAILIA\b|\bTANTA\b|\bMANSOURA\b|\bOBOUR\b|\bHELIOPOLIS\b|\bNASR\s*CITY\b"
    r"|\bMAADI\b|\b10TH OF RAMADAN\b|\b6TH OF OCTOBER\b|مصر|القاهرة|الإسكندرية|الاسكندرية")
_BANK_RE = re.compile(r"\bBANK\b|\bBANQUE\b|\bBANCO\b|\bBANCA\b|\bBANKING\b|بنك|مصرف")


def _norm(v):
    return re.sub(r"\s+", " ", str(v or "")).strip().upper()


def _num(v):
    try:
        return float(str(v).replace(",", "")) if v not in ("", None) else None
    except Exception:
        return None


def _size(container_type):
    """يرجّع 20 أو 40 أو None (لو النوع مش معروف)."""
    t = _norm(container_type)
    if re.match(r"^20", t) or re.match(r"^2\d[A-Z]\d$", t) or re.search(r"\b20\b|20\s*'|20FT", t):
        return 20
    if re.match(r"^(40|45|48)", t) or re.match(r"^[4LM]\d[A-Z]\d$", t) or re.search(r"\b(40|45)\b|40FT", t):
        return 40
    return None


def _part_of_two(text):
    t = _norm(text)
    return bool(re.search(r"1\s+OF\s+2\s+PART\s+CARGO|1\s*/\s*2\s+PART\s+CARGO|PART\s+OF\s+2", t))


def _is_egypt_text(text):
    t = _norm(text)
    if _EGYPT_RE.search(t):
        return True
    return any(COUNTRIES[c] == "EG" for c in _CRE.findall(t))


def _issue(bl, rule, details, sev=ERR):
    return {"Severity": sev, "B/L": bl, "Rule": rule, "Details": details}


def build_rule_issues(summaries, details):
    issues = []
    sum_by_bl, det_by_bl = {}, {}
    for s in summaries:
        sum_by_bl.setdefault(_norm(s.get("BL")), s)
    for d in details:
        det_by_bl.setdefault(_norm(d.get("BL")), []).append(d)
    bl_counts = Counter(_norm(s.get("BL")) for s in summaries if _norm(s.get("BL")))
    acid_counts = Counter(_norm(s.get("ACID")) for s in summaries if _norm(s.get("ACID")))
    done = set()

    for s in summaries:
        bl = s.get("BL", "")
        nbl = _norm(bl)
        first = nbl not in done
        done.add(nbl)

        # ---- بيانات أساسية ----
        if first and bl_counts.get(nbl, 0) > 1:
            issues.append(_issue(bl, "Duplicate B/L", f"ظهر {bl_counts[nbl]} مرات في المانيفست"))
        acid = _norm(s.get("ACID"))
        if not acid:
            issues.append(_issue(bl, "Missing ACID", "لا يوجد رقم ACID مستخرج"))
        elif acid_counts.get(acid, 0) > 1:
            issues.append(_issue(bl, "Duplicate ACID", f"رقم ACID {acid} مكرر {acid_counts[acid]} مرات"))

        # ---- الشاحن لازم يكون برا مصر ----
        code = s.get("Shipper Country Code", "")
        if code == "EG":
            issues.append(_issue(bl, "Egyptian Shipper", "الشاحن مصري؛ المطلوب شاحن من خارج مصر"))
        elif not code:
            issues.append(_issue(bl, "Shipper country unknown", "تعذر تحديد دولة الشاحن، راجع المانيفست", WARN))
        elif _is_egypt_text(s.get("Shipper")):
            issues.append(_issue(bl, "Shipper address mentions Egypt",
                                 f"دولة الشاحن المستخرجة {code} لكن العنوان فيه إشارة لمصر: {s.get('Shipper')}", WARN))
        if not _norm(s.get("Shipper ID")):
            issues.append(_issue(bl, "Missing Shipper ID", "راجع بيانات الشاحن / Shipper ID"))
        if _norm(s.get("Shipper ID Check")) == "راجع المانيفست":
            issues.append(_issue(bl, "Shipper ID / Country mismatch", "كود الشاحن لا يتوافق مع دولة الشاحن المستخرجة"))

        # ---- المستلم لازم يكون مصري أو بنك ----
        cons = s.get("Consignee")
        if not _norm(cons):
            issues.append(_issue(bl, "Missing Consignee", "لا يوجد Consignee مستخرج"))
        elif not (_is_egypt_text(cons) or _BANK_RE.search(_norm(cons))):
            issues.append(_issue(bl, "Consignee not Egypt / not a bank", f"المستلم: {cons}"))

        # ---- الوجهة ----
        pod = _norm(s.get("Place of Delivery"))
        if pod and not re.search(r"\bEGDAM\b|DAMIETTA", pod):
            issues.append(_issue(bl, "Destination not EGDAM", f"مكان التسليم المستخرج: {pod}", WARN))

        # ---- ترانزيت / فري زون: تنبيه فقط ----
        if s.get("Transit"):
            issues.append(_issue(bl, "Transit", f"تنبيه: البوليصة فيها ترانزيت → {s.get('Transit')}", WARN))
        if s.get("Free Zone"):
            issues.append(_issue(bl, "Free Zone", f"تنبيه: البوليصة فيها فري زون → {s.get('Free Zone')}", WARN))

        # ---- فحوصات الحاويات (مرة واحدة لكل بوليصة) ----
        if not first:
            continue
        rows = det_by_bl.get(nbl, [])
        if not rows:
            issues.append(_issue(bl, "No containers found", "لم يتم استخراج أي حاوية لهذه البوليصة", WARN))
        part2 = bool(s.get("Part of 2 Cargo"))
        for d in rows:
            cno, w, size = d.get("Container", ""), _num(d.get("Gross KGM")), _size(d.get("Type"))
            if not _norm(d.get("Seal")):
                issues.append(_issue(bl, "Missing Seal", f"الحاوية {cno} لا يوجد لها Seal مستخرج"))
            if w is None:
                issues.append(_issue(bl, "Weight missing", f"تعذر قراءة الوزن للحاوية {cno}", WARN))
            if size is None:
                issues.append(_issue(bl, "Container size unknown",
                                     f"تعذر تحديد حجم الحاوية {cno} (النوع: {d.get('Type', '')})", WARN))
            if size == 20 and w is not None:
                if w > MAX_20_KG:
                    issues.append(_issue(bl, "20' overweight", f"{cno}: {w:,.0f} kg (الحد {MAX_20_KG:,})"))
                if part2 and w >= MAX_20_KG:
                    issues.append(_issue(bl, "20' Part-of-2 cargo at 25 tons", f"{cno}: {w:,.0f} kg؛ مراجعة خاصة"))
            if size == 40 and w is not None and w > MAX_40_KG:
                issues.append(_issue(bl, "40' overweight", f"{cno}: {w:,.0f} kg (الحد {MAX_40_KG:,})"))

    # ---- حاوية على أكتر من بوليصة ----
    by_cont = {}
    for d in details:
        c = _norm(d.get("Container"))
        if c:
            by_cont.setdefault(c, []).append(d)
    fields = [("Shipper", "الشاحن"), ("Shipper ID", "رقم الشاحن"), ("Consignee", "المستلم"),
              ("Consignee ID", "رقم المستلم"), ("Notify", "النوتيفاي"),
              ("Place of Receipt", "مكان الاستلام"), ("Place of Delivery", "مكان التسليم")]
    for c, rows in by_cont.items():
        bls = list(dict.fromkeys(d.get("BL", "") for d in rows))
        if len(bls) == 1:
            if len(rows) > 1:
                issues.append(_issue(bls[0], "Container repeated in same B/L", f"{c} ظهرت {len(rows)} مرات في نفس البوليصة", WARN))
            continue
        label = " / ".join(bls)
        weights = [_num(d.get("Gross KGM")) for d in rows]
        wtxt = " | ".join(f"{d.get('BL', '')}: {w:,.0f} kg" if w is not None else f"{d.get('BL', '')}: وزن غير مقروء"
                          for d, w in zip(rows, weights))
        diffs = []
        for key, name in fields:
            if len({_norm(sum_by_bl.get(_norm(b), {}).get(key)) for b in bls}) > 1:
                diffs.append(name)
        if len({_norm(d.get("Seal")) for d in rows}) > 1:
            diffs.append("الـ Seal")
        acids = {_norm(sum_by_bl.get(_norm(b), {}).get("ACID")) for b in bls}
        acid_txt = "نفس الـ ACID" if len(acids) == 1 else "ACID مختلف"
        if not diffs:
            issues.append(_issue(label, "Shared container – identical parties",
                                 f"{c} مشتركة بين {len(bls)} بوالص بنفس الشاحن والمستلم والنوتيفاي والموانئ والـ Seal ({acid_txt}) — {wtxt}", WARN))
        else:
            issues.append(_issue(label, "Shared container – different parties",
                                 f"{c} مشتركة بين {len(bls)} بوالص ومختلفة في: {'، '.join(diffs)} ({acid_txt}) — {wtxt}"))
        size = next((z for z in (_size(d.get("Type")) for d in rows) if z), None)
        if size == 20:
            if any(w is not None and w >= SHARED_20_KG for w in weights):
                issues.append(_issue(label, "20' shared container – weight >= 24t", f"{c}: {wtxt}"))
            if any(w is None for w in weights):
                issues.append(_issue(label, "20' shared container – weight missing", f"{c}: {wtxt}", WARN))

    issues.sort(key=lambda i: i["Severity"] != ERR)
    return issues


def finish(parts_json):
    """يجمّع الدفعات بالترتيب، يلزق البلوكات المقسومة، يعمل parse ويكتب /tmp/out.xlsx."""
    blocks = []
    for part in json.loads(parts_json):
        for b in json.loads(part):
            if b.get("orphan"):
                if blocks:
                    blocks[-1]["left"] += b["left"]
                    blocks[-1]["right"] += b["right"]
            else:
                blocks.append(b)
    if not blocks:
        raise ValueError("ما لقيتش بوالص في الملف. اتأكد إنه PDF بنص حقيقي.")

    summaries, details, errors = [], [], []
    for b in blocks:
        try:
            s, d = parse_block(b)
            summaries.append(s)
            details.extend(d)
        except Exception as e:
            errors.append({"First line": b["left"][0].strip()[:80], "Error": repr(e)})

    rule_issues = build_rule_issues(summaries, details)
    n_err = sum(1 for i in rule_issues if i["Severity"] == ERR)
    n_warn = len(rule_issues) - n_err

    summary_rows = [
        {"Item": "عدد البوالص", "Severity": "", "Count": len(summaries)},
        {"Item": "عدد الحاويات", "Severity": "", "Count": len(details)},
        {"Item": "بوالص فشلت قراءتها (شيت Errors)", "Severity": "", "Count": len(errors)},
        {"Item": "إجمالي الأخطاء", "Severity": ERR, "Count": n_err},
        {"Item": "إجمالي التنبيهات", "Severity": WARN, "Count": n_warn},
    ]
    cnt = Counter((i["Severity"], i["Rule"]) for i in rule_issues)
    for (sev, rule), n in sorted(cnt.items(), key=lambda x: (x[0][0] != ERR, -x[1])):
        summary_rows.append({"Item": rule, "Severity": sev, "Count": n})

    issues_rows = rule_issues or [{"Severity": "", "B/L": "", "Rule": "لا توجد ملاحظات", "Details": ""}]
    sheets = [("Summary", summary_rows), ("Manifest", summaries), ("Containers", details),
              ("Rule Issues", issues_rows)]
    if errors:
        sheets.append(("Errors", errors))

    red = PatternFill("solid", fgColor="F8CBAD")
    yellow = PatternFill("solid", fgColor="FFF2CC")
    wb = Workbook()
    for i, (name, rows) in enumerate(sheets):
        ws = wb.active if i == 0 else wb.create_sheet()
        ws.title = name
        if not rows:
            continue
        heads = list(rows[0].keys())
        widths = [len(h) for h in heads]
        ws.append(heads)
        for r in rows:
            vals = []
            for k, h in enumerate(heads):
                v = r.get(h, "")
                if isinstance(v, str):
                    v = _BAD.sub("", v)
                widths[k] = max(widths[k], len(str(v)))
                vals.append(v)
            ws.append(vals)
        for c in ws[1]:
            c.font = Font(bold=True)
        if name in ("Rule Issues", "Summary"):
            for row in ws.iter_rows(min_row=2):
                sev = row[0].value if name == "Rule Issues" else row[1].value
                if sev == ERR:
                    for c in row:
                        c.fill = red
                elif sev == WARN:
                    for c in row:
                        c.fill = yellow
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for k, w in enumerate(widths):
            ws.column_dimensions[ws.cell(row=1, column=k + 1).column_letter].width = min(w + 2, 60)
    wb.save("/tmp/out.xlsx")
    return json.dumps([len(summaries), len(details), len(errors)])
# =====================================================================================


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if not args:
        path = PDF_PATH.strip().strip('"')
        if not path:
            from tkinter import Tk, filedialog
            Tk().withdraw()
            path = filedialog.askopenfilename(title="اختار ملف المانيفست PDF",
                                              filetypes=[("PDF", "*.pdf")])
        if not path:
            sys.exit("ما اخترتش ملف.")
        args = [path]
    pdf_path = Path(args[0])
    raw = "--raw" in sys.argv
    blocks = read_blocks(pdf_path, raw=raw)
    if raw:
        return
    if not blocks:
        sys.exit("ما لقيتش بوالص. شغّل الأمر مع --raw وشوف شكل النص.")

    summaries, details = [], []
    for b in blocks:
        s, d = parse_block(b)
        summaries.append(s)
        details.extend(d)

    out = Path(args[1]) if len(args) > 1 else pdf_path.with_suffix(".xlsx")
    with pd.ExcelWriter(out, engine="openpyxl") as xw:
        pd.DataFrame(summaries).to_excel(xw, sheet_name="Manifest", index=False)
        pd.DataFrame(details).to_excel(xw, sheet_name="Containers", index=False)
        pd.DataFrame(build_rule_issues(summaries, details) or
                     [{"Severity": "", "B/L": "", "Rule": "لا توجد ملاحظات", "Details": ""}]
                     ).to_excel(xw, sheet_name="Rule Issues", index=False)
        for ws in xw.book.worksheets:  # عرض أعمدة مناسب
            for col in ws.columns:
                w = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(w + 2, 60)
    print(f"تم: {len(summaries)} بوليصة، {len(details)} حاوية -> {out}")


if __name__ == "__main__":
    main()
