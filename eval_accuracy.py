# -*- coding: utf-8 -*-
"""
Dəqiqlik hesablama skripti.

İSTİFADƏ:
    1. "deqiqlik_yoxlama_nümunəsi.xlsx" faylını açın
    2. Hər sətirdə "Duzgun_Kodlar" sütununa SİZCƏ doğru kodları yazın (vergüllə ayırıb, məs: T01, P01)
       - Skriptin tapdığı proqnoz düzgündürsə, sadəcə "Tehmin_Kodlari" dəyərini kopyalayın
       - Əgər komment heç bir kateqoriyaya aid deyilsə, "O01" yazın
    3. Faylı yadda saxlayın
    4. Terminal-da işə salın: python eval_accuracy.py deqiqlik_yoxlama_nümunəsi.xlsx
"""
import sys
import re
import openpyxl
from collections import defaultdict


def parse_codes(cell_value):
    if not cell_value:
        return set()
    return set(c.strip().upper() for c in re.split(r"[,;]", str(cell_value)) if c.strip())


def main():
    if len(sys.argv) < 2:
        print("İstifadə: python eval_accuracy.py deqiqlik_yoxlama_nümunəsi.xlsx")
        sys.exit(1)

    path = sys.argv[1]
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb["Yoxlama"]

    rows = list(ws.iter_rows(min_row=4, values_only=True))  # header 3-cü sətirdə idi

    total = 0
    exact_match = 0
    tp = defaultdict(int)
    fp = defaultdict(int)
    fn = defaultdict(int)
    jaccard_sum = 0.0
    skipped = 0

    for row in rows:
        if not row or not row[0]:
            continue
        comment, predicted_raw, correct_raw = row[0], row[1], row[2]
        if not correct_raw or str(correct_raw).strip() == "":
            skipped += 1
            continue  # hələ doldurulmayıb

        predicted = parse_codes(predicted_raw)
        correct = parse_codes(correct_raw)

        total += 1
        if predicted == correct:
            exact_match += 1

        union = predicted | correct
        inter = predicted & correct
        jaccard_sum += (len(inter) / len(union)) if union else 1.0

        for code in predicted:
            if code in correct:
                tp[code] += 1
            else:
                fp[code] += 1
        for code in correct:
            if code not in predicted:
                fn[code] += 1

    if total == 0:
        print("Heç bir sətir doldurulmayıb — 'Duzgun_Kodlar' sütununu doldurub yenidən yoxlayın.")
        print(f"({skipped} sətir hələ boşdur)")
        return

    print(f"Qiymətləndirilən sətir sayı: {total}  (boş buraxılan: {skipped})")
    print(f"Tam-uyğunluq dəqiqliyi (exact match accuracy): {100*exact_match/total:.1f}%")
    print(f"Orta Jaccard oxşarlığı (multi-label üçün): {100*jaccard_sum/total:.1f}%")
    print()

    all_codes = sorted(set(tp) | set(fp) | set(fn))
    print(f"{'Kod':6s} {'Precision':>10s} {'Recall':>10s} {'F1':>8s} {'TP':>5s} {'FP':>5s} {'FN':>5s}")
    print("-" * 55)
    macro_p, macro_r, macro_f1, n = 0, 0, 0, 0
    for code in all_codes:
        p = tp[code] / (tp[code] + fp[code]) if (tp[code] + fp[code]) else 0.0
        r = tp[code] / (tp[code] + fn[code]) if (tp[code] + fn[code]) else 0.0
        f1 = 2 * p * r / (p + r) if (p + r) else 0.0
        print(f"{code:6s} {p*100:>9.1f}% {r*100:>9.1f}% {f1*100:>7.1f}% {tp[code]:>5d} {fp[code]:>5d} {fn[code]:>5d}")
        macro_p += p
        macro_r += r
        macro_f1 += f1
        n += 1

    if n:
        print("-" * 55)
        print(f"{'MAKRO ORTA':6s} {100*macro_p/n:>9.1f}% {100*macro_r/n:>9.1f}% {100*macro_f1/n:>7.1f}%")

    print("""
QEYD (hesabat üçün):
- Exact match accuracy: nə qədər kommentin BÜTÜN kodları (nə çox, nə əskik) düz tapılıb
- Jaccard oxşarlığı: multi-label üçün daha ədalətli ölçü — qismən düzgün proqnozları da qismən mükafatlandırır
- Precision (kod üzrə): skript bu kodu verdikdə, neçə faiz halda həqiqətən doğru idi
- Recall (kod üzrə): bu koda aid olan kommentlərin neçə faizini skript tapa bildi
- Aşağı Precision → skript "həddindən artıq geniş" tutur (false positive çox)
- Aşağı Recall → skript "qaçırır" (false negative çox), lüğət genişləndirilməlidir
""")


if __name__ == "__main__":
    main()
