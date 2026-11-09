# -*- coding: utf-8 -*-
"""
APAR komment kateqoriyalaşdırma skripti — yerli kompüterdə işə salmaq üçün.

İSTİFADƏ:
    python run_classification.py giris_faylı.xlsx

Giriş faylında "Comment" adlı sütun olmalıdır (birinci sətir başlıq).
Nəticə eyni qovluqda "*_kateqoriyalasdirilmis.xlsx" adı ilə yaranacaq.
"""
import sys
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from collections import Counter
from classify import classify, CODE_MAP, UNCAT_NAME, UNCAT_CODE

GROUPS = [
    ("🔧 Texniki Nasazlıq", [
        ("T01", "Əyləc / Tormoz"), ("T02", "Pedal"), ("T03", "Zəncir"), ("T04", "Oturacaq"),
        ("T05", "Sükan"), ("T06", "Təkər / Şin / Hava"), ("T07", "Fiziki Kilid"), ("T08", "Dayaq ayağı"),
        ("T09", "Motor / Mühərrik / Elektrik dartısı"), ("T10", "Batareya"), ("T11", "Səbət"),
        ("T12", "Bolt / Vint"), ("T13", "Ümumi Texniki Nasazlıq"), ("T14", "Təhlükəsizlik Problemi"),
        ("T15", "Ümumi Baxım / Servis ehtiyacı"), ("T16", "Gear / Sürət pilləsi / Ötürücü"),
        ("T17", "Sürət / Performans"), ("T18", "Telefon tutacağı"),
        ("T19", "Sükan–Təkər Uyğunsuzluğu / Alignment"), ("T20", "İşıq / Fara"), ("T21", "Qaz / Throttle"),
        ("T22", "Siqnal / Zəng"), ("T23", "Disk / Jant"), ("T24", "Amortizator / Asma sistemi"),
        ("T25", "Gediş zamanı özbaşına dayanma / kilidlənmə"),
        ("T26", "Ekran / Göstərici Paneli"), ("T27", "Əl Tutacağı Rezini / Qulpu"),
    ]),
    ("📱 Tətbiq və Rəqəmsal", [
        ("A01", "GPS / Xəritə / Location problemi"),
        ("A02", "App və Qeydiyyat / Login / QR / Session start-end"),
        ("A03", "Bluetooth bağlantısı"),
    ]),
    ("💳 Ödəniş və Qiymət", [
        ("P01", "Qiymət Narazılığı"), ("P02", "Ödəniş Xətaları"),
    ]),
    ("🏗️ İnfrastruktur", [
        ("I01", "Parking Çatışmazlığı"), ("I02", "Xidmət Zonası / Service Area"),
        ("I03", "Velosiped/Skuter Çatışmazlığı"), ("I04", "Velosiped Yolu / Yol İnfrastrukturu"),
    ]),
    ("💡 Təklif və İstəklər", [
        ("S01", "Yeni Funksiya və Təchizat Təklifi"),
    ]),
    ("🙂 Müsbət", [
        ("M01", "Məmnuniyyət / Müsbət rəy"),
    ]),
    ("📦 Digər", [
        ("O01", "Təsnif olunmamış"),
    ]),
]
ALL_CODES = [code for _, items in GROUPS for code, _ in items]


def main():
    if len(sys.argv) < 2:
        print("İstifadə: python run_classification.py giris_faylı.xlsx")
        sys.exit(1)

    src = sys.argv[1]
    base, _ = os.path.splitext(src)
    out = base + "_kateqoriyalasdirilmis.xlsx"

    wb_src = openpyxl.load_workbook(src, read_only=True)
    ws_src = wb_src.active  # birinci vərəqi götürür
    comments = [r[0] for r in ws_src.iter_rows(min_row=2, values_only=True) if r[0]]
    print(f"Oxundu: {len(comments)} komment")

    FONT = "Arial"
    header_font = Font(name=FONT, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    base_font = Font(name=FONT, size=10)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Kategoriyalar"
    headers = ["Comment"] + ALL_CODES + ["Matched_Codes", "Matched_Categories", "Category_Count"]
    ws.append(headers)
    for i in range(1, len(headers) + 1):
        c = ws.cell(row=1, column=i)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    counter_code = Counter()
    for comment in comments:
        hits = classify(comment)
        codes = [CODE_MAP[h] for h in hits] if hits else [UNCAT_CODE]
        names = hits if hits else [UNCAT_NAME]
        for code in codes:
            counter_code[code] += 1
        code_set = set(codes)
        row = [comment] + [1 if code in code_set else 0 for code in ALL_CODES]
        row += [", ".join(codes), ", ".join(names), len(codes)]
        ws.append(row)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.font = base_font
    ws.column_dimensions["A"].width = 55
    for i in range(2, 2 + len(ALL_CODES)):
        ws.column_dimensions[get_column_letter(i)].width = 5
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions

    ws2 = wb.create_sheet("Xulase")
    ws2.append(["Kod", "Subcategory", "Komment sayı", "Faiz"])
    for c in ws2[1]:
        c.font = header_font
        c.fill = header_fill
    total = len(comments)
    for _, items in GROUPS:
        for code, name in items:
            ws2.append([code, name, counter_code[code], round(100 * counter_code[code] / total, 1)])
    ws2.column_dimensions["A"].width = 8
    ws2.column_dimensions["B"].width = 45

    wb.save(out)
    print(f"Nəticə yadda saxlanıldı: {out}")


if __name__ == "__main__":
    main()
