# -*- coding: utf-8 -*-
import re
import sys
import openpyxl
from collections import defaultdict, Counter
from classify import classify, CODE_MAP, UNCAT_NAME, UNCAT_CODE


# ---------------------------------------------------------------------
# Human label text -> our code. Keys are normalized (lowercased) category
# names as they appear AFTER the "|" in the human_label column.
# Sorted by length (desc) at lookup time so more specific phrases match
# before short generic ones.
# ---------------------------------------------------------------------
TEXT_TO_CODE = {
    "məmnuniyyət": "M01",
    "ümumi texniki nasazlıq": "T13",
    "təsnif olunmamış": "O01",
    "qiymət narazılığı": "P01",
    "əyləc (tormoz)": "T01",
    "əyləc": "T01",
    "pedal": "T02",
    "oturacaq": "T04",
    "sükan-təkər uyğunsuzluğu (alignment / razval)": "T19",
    "sükan": "T05",
    "kilid (fiziki / mexaniki)": "T07",
    "kilid": "T07",
    "yeni funksiya və təchizat təklifi": "S01",
    "app və qeydiyyat": "A02",
    "zəncir": "T03",
    "batareya": "T10",
    "təkər və hava": "T06",
    "motor": "T09",
    "gear / sürət pilləsi / ötürücü": "T16",
    "sürət ötürücüsü (gear)": "T16",
    "sürət / performans": "T17",
    "velosiped yolu / yol infrastrukturu": "I04",
    "ödəniş xətaları": "P02",
    "telefon tutacağı": "T18",
    "telefon tutacaği": "T18",
    "parking çatışmazlığı": "I01",
    "təhlükəsizlik problemi": "T14",
    "baxım / texniki xidmət": "T15",
    "səbət": "T11",
    "dayaq ayağı": "T08",
    "siqnal / zəng (horn/bell)": "T22",
    "zəng": "T22",
    "gps və yerləşmə": "A01",
    "velosiped/skuter çatışmazlığı (availability)": "I03",
    "işıqlandırma / stop işığı": "T20",
    "qaz / akselerator (throttle)": "T21",
    "boltlar / vintlər": "T12",
    "disk / jant": "T23",
    "xidmət zonası": "I02",
    "ümumi baxım / servis ehtiyacı": "T15",
}
# Sort keys longest-first so specific phrases are checked before generic substrings
SORTED_KEYS = sorted(TEXT_TO_CODE.keys(), key=len, reverse=True)


def parse_human_label(label):
    """Robustly extract our internal codes from a messy human_label string."""
    if not label or str(label).strip().lower() == "none":
        return None  # unlabeled -> skip
    text = str(label).lower()
    codes = set()
    for key in SORTED_KEYS:
        if key in text:
            codes.add(TEXT_TO_CODE[key])
    return codes


def main():
    if len(sys.argv) < 2:
        print("Usage: python eval_manual_labels.py manual_validation.xlsx")
        sys.exit(1)
    wb = openpyxl.load_workbook(sys.argv[1], read_only=True)
    ws = wb["Sheet1"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))

    total = 0
    skipped = 0
    exact_match = 0
    jaccard_sum = 0.0
    tp = defaultdict(int)
    fp = defaultdict(int)
    fn = defaultdict(int)
    unrecognized = []
    mismatches = []  # for manual inspection

    for vid, comment, human_label in rows:
        gold = parse_human_label(human_label)
        if gold is None:
            skipped += 1
            continue
        if not gold:
            # label text existed but matched nothing in our dictionary -> flag it
            unrecognized.append((vid, comment, human_label))
            skipped += 1
            continue

        hits = classify(comment)
        predicted = set(CODE_MAP[h] for h in hits) if hits else {UNCAT_CODE}

        total += 1
        if predicted == gold:
            exact_match += 1
        else:
            mismatches.append((vid, comment, sorted(gold), sorted(predicted)))

        union = predicted | gold
        inter = predicted & gold
        jaccard_sum += (len(inter) / len(union)) if union else 1.0

        for c in predicted:
            if c in gold:
                tp[c] += 1
            else:
                fp[c] += 1
        for c in gold:
            if c not in predicted:
                fn[c] += 1

    print(f"Cəmi sətir: {len(rows)}")
    print(f"Qiymətləndirildi: {total}")
    print(f"Buraxıldı (etiketsiz və ya tanınmayan): {skipped}")
    if unrecognized:
        print(f"  -> tanınmayan etiket mətni olan sətirlər: {len(unrecognized)}")
        for vid, c, lbl in unrecognized[:10]:
            print(f"     id={vid} label={lbl!r} comment={c!r}")
    print()
    print(f"TAM UYĞUNLUQ DƏQİQLİYİ (exact match accuracy): {100*exact_match/total:.1f}%")
    print(f"ORTA JACCARD OXŞARLIĞI (multi-label): {100*jaccard_sum/total:.1f}%")
    print()

    all_codes = sorted(set(tp) | set(fp) | set(fn))
    print(f"{'Kod':6s} {'Precision':>10s} {'Recall':>10s} {'F1':>8s} {'TP':>5s} {'FP':>5s} {'FN':>5s}")
    print("-" * 55)
    macro_p = macro_r = macro_f1 = 0.0
    n = 0
    for code in all_codes:
        p = tp[code] / (tp[code] + fp[code]) if (tp[code] + fp[code]) else 0.0
        r = tp[code] / (tp[code] + fn[code]) if (tp[code] + fn[code]) else 0.0
        f1 = 2 * p * r / (p + r) if (p + r) else 0.0
        print(f"{code:6s} {p*100:>9.1f}% {r*100:>9.1f}% {f1*100:>7.1f}% {tp[code]:>5d} {fp[code]:>5d} {fn[code]:>5d}")
        macro_p += p; macro_r += r; macro_f1 += f1; n += 1
    if n:
        print("-" * 55)
        print(f"{'MAKRO':6s} {100*macro_p/n:>9.1f}% {100*macro_r/n:>9.1f}% {100*macro_f1/n:>7.1f}%")

    # micro (aggregate) precision/recall
    TP = sum(tp.values()); FP = sum(fp.values()); FN = sum(fn.values())
    micro_p = TP / (TP + FP) if (TP + FP) else 0.0
    micro_r = TP / (TP + FN) if (TP + FN) else 0.0
    micro_f1 = 2 * micro_p * micro_r / (micro_p + micro_r) if (micro_p + micro_r) else 0.0
    print(f"{'MİKRO':6s} {100*micro_p:>9.1f}% {100*micro_r:>9.1f}% {100*micro_f1:>7.1f}%")

    print("\n\n=== ƏN ÇOX SƏHV EDİLƏN 25 NÜMUNƏ (manual baxış üçün) ===")
    for vid, c, gold, pred in mismatches[:25]:
        print(f"id={vid} | comment={c!r}")
        print(f"   İNSAN:   {gold}")
        print(f"   SKRIPT:  {pred}")

    return mismatches


if __name__ == "__main__":
    main()
