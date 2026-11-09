# -*- coding: utf-8 -*-
import re
import openpyxl
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Normalization: many AZ comments are typed without diacritics
# (islemir instead of işləmir). We strip AZ-specific letters down to their
# closest ASCII Latin letter before matching, and match RU (Cyrillic) / EN
# patterns unchanged.
# ---------------------------------------------------------------------------
_TRANSLIT = str.maketrans({
    "ə": "e", "ı": "i", "ö": "o", "ü": "u", "ç": "c", "ş": "s", "ğ": "g",
})

def normalize(text):
    return str(text).lower().translate(_TRANSLIT)

# ---------------------------------------------------------------------------
# Category definitions: name -> list of regex patterns (matched on the
# normalized/lowercased text).
# ---------------------------------------------------------------------------

CATS = {}

def add(name, patterns):
    CATS[name] = [re.compile(p, re.IGNORECASE | re.UNICODE) for p in patterns]

# ---------------- 1. TEXNIKI NASAZLIQ ----------------
# NOTE: AZ patterns are written WITHOUT diacritics (matched against
# normalize()'d text). RU (Cyrillic) and EN patterns need no such handling.

add("Əyləc / Tormoz", [
    r"\beylec\w*", r"\btorm[uo]z\w*", r"\bfren(le|i|ler|lemir|lemir)?\b",
    r"\bтормоз\w*", r"\bbrake\w*",
])

add("Pedal", [
    r"\bpedal\w*", r"\bпедал\w*",
])

add("Zəncir", [
    r"\bzencir\w*", r"\bsep(i|in|inde|inden|inne|len)?\b", r"\bцеп[ьи]\w*", r"\bchain\w*",
])

add("Oturacaq", [
    r"\boturaca[qg]\w*", r"\boturaq\b", r"\bsidenk\w*", r"\bseat\w*", r"\bsaddle\w*",
    r"\bсиден\w*", r"\bсидушк\w*", r"\bседушк\w*", r"\bкресл\w*",
])

add("Sükan", [
    r"\bsukan\w*", r"\brul[a-z]*\b", r"\brol(u|un|a)?\b",
    r"\bрул[ья]\w*", r"\bруль\b",
    r"\bhandle\s?bar\w*", r"\bsteering\w*",
])

add("Təkər / Şin / Hava", [
    r"\bteker\w*", r"\bsin(i|ler|de|den)?\b", r"\blastik\w*",
    r"\bkoles\w*", r"\bколес\w*", r"\bшин\w*", r"\bскат\w*",
    r"\bwheel\w*", r"\btire\w*", r"\btyre\w*", r"\bflat\s?tire\w*",
])

add("Fiziki Kilid", [
    r"\bkilid\w*", r"\bqifil\w*", r"\bacar\w*", r"\bacilm(ir|adi|adı)\w*", r"\backilmir\w*", r"\bklikle\w*",
    r"\bзамок\w*", r"\bзамк\w*", r"\bне\s?открыва\w*",
    r"\block\b(?!er)", r"\bunlock\w*", r"\blocked\b",
])

add("Dayaq ayağı", [
    r"\bstoyk\w*", r"\bdayaq\s?ayag\w*", r"\bdayag\w*", r"\bkikstend\w*", r"\bkickstend\w*", r"\bsoykemek\w*",
    r"\bподножк\w*", r"\bстойк\w*", r"\bkickstand\w*", r"\bкикстенд\w*",
])

add("Motor / Mühərrik / Elektrik dartısı", [
    r"\bmotor\w*", r"\bmuherrik\w*", r"\belektrik\w*",
    r"\bмотор\w*", r"\bдвигател\w*", r"\belectric\s?motor\w*", r"\bengine\w*",
])

add("Batareya", [
    r"\bbatare\w*", r"\bakkumulyator\w*", r"\bsarj\w*",
    r"\bдоз?ар[яе]жен\w*", r"\bбатаре\w*", r"\bаккумулятор\w*", r"\bзаряд\w*",
    r"\bbattery\w*", r"\bcharg\w*",
])

add("Səbət", [
    r"\bsebet\w*",
    r"\bкорзин\w*", r"\bbasket\w*",
])

add("Bolt / Vint", [
    r"\bbolt\w*", r"\bvint\w*", r"\bболт\w*", r"\bвинт\w*", r"\bscrew\w*",
])

add("Təhlükəsizlik Problemi", [
    r"\btehluke\w*", r"\bqeza\w*", r"\byix\w*",
    r"\bопасн\w*", r"\bавари\w*", r"\bупал\w*", r"\bупад\w*",
    r"\bdanger\w*", r"\bunsafe\w*", r"\baccident\w*", r"\bsafety\b",
])

add("Ümumi Baxım / Servis ehtiyacı", [
    r"\bbaxim\w*", r"\bservis\w*",
    r"\bобслужива\w*", r"\bремонт\w*", r"\bсервис\w*",
    r"\bmaintenance\w*", r"\brepair\w*", r"\bservice\b",
])

add("Gear / Sürət pilləsi / Ötürücü", [
    r"\bsuret\s?pille\w*", r"\boturucu\w*", r"\bskorus\w*", r"\bskoros\w*", r"\bskors\w*", r"\bsuret\w*\s?qutu\w*",
    r"\bgear\w*", r"\bпередач\w*", r"\bскорост\w*\s?переключ\w*", r"\bшестерн\w*",
])

add("Sürət / Performans", [
    r"\bsuretlen\w*", r"\byavas\s?ged\w*", r"\bsuret\s?az\w*", r"\bagir\s?ged\w*", r"\bzeif\s?ged\w*",
    r"\bagir\s?isley\w*", r"\bsuret\w*\s?yigmir\w*",
    r"\bперфоманс\w*", r"\bмедленн\w*\s?ед\w*", r"\bне\s?разгоняет\w*",
    r"\bскорост\w*\s?маленьк\w*", r"\bне\s?набирает\s?скорост\w*",
    r"\bslow\b", r"\bперформанс\w*",
])

add("Telefon tutacağı", [
    r"\btelefon\s?tutaca[qg]\w*", r"\btelefon\s?qulp\w*", r"\btelefon\s?saxlayan\w*", r"\btelefon\s?yer\w*",
    r"\btelefon\w*\s?dus(du|dü)\w*",
    r"\bдержател[ья]\s?телефон\w*", r"\bphone\s?holder\w*", r"\bphone\s?mount\w*",
])

add("Sükan–Təkər Uyğunsuzluğu / Alignment", [
    r"(?=.*\b(?:rul|sukan|rol)\w*)(?=.*\bteker\w*)(?=.*\beyri\w*)",
    r"\bduz\s?getmir\w*",
    r"\bрул[ья]?\s?крив\w*", r"\bедет\s?крив\w*",
    r"\bmisalign\w*", r"\bcrooked\w*",
])

add("İşıq / Fara", [
    r"\bisi[qg]\w*", r"\bfara\w*", r"\bciraq\w*",
    r"\bфар[аы]\w*", r"\bсвет\b", r"\bфонар\w*",
    r"\blight\b", r"\bheadlight\w*", r"\blamp\w*",
])

add("Qaz / Throttle", [
    r"\bqaz\s?tutacag\w*", r"\bqaz\s?el\w*", r"\bqaz\s?pedal\w*",
    r"\bгаз\b", r"\bручк\w*\s?газ\w*",
    r"\bthrottle\w*", r"\baccelerator\w*",
])

add("Siqnal / Zəng", [
    r"\bzeng\w*", r"\bsiqnal\w*", r"\bsignal\w*", r"\bqorna\w*",
    r"\bзвонок\w*", r"\bсигнал\w*", r"\bгудок\w*",
    r"\bbell\b", r"\bhorn\b",
])

add("Disk / Jant", [
    r"\bdisk\w*", r"\bjant\w*", r"\bдиск\w*", r"\bобод\w*", r"\brim\b",
])

add("Amortizator / Asma sistemi", [
    r"\bamortizator\w*", r"\bamortizasiya\w*", r"\bamortizasi\w*",
    r"\bамортизатор\w*", r"\bамортизаци\w*",
    r"\bsuspension\w*", r"\bshock\s?absorb\w*",
])

add("Gediş zamanı özbaşına dayanma / kilidlənmə", [
    r"\byolun\s?ortasinda\w*", r"\byari\s?yolda\s?qoy\w*", r"\bozu\s?ozune\s?(dayan|stop|sto?nub|kilidlen)\w*",
    r"\bozbasina\s?(dayan|sonub|kilidlen)\w*", r"\bbirden\s?(bire\s?)?dayan\w*", r"\byolda\s?sondu\w*",
])

add("Ekran / Göstərici Paneli", [
    r"\bekran\w*", r"\bэкран\w*", r"\bдисплей\w*", r"\bscreen\w*", r"\bdisplay\w*",
])

add("Əl Tutacağı Rezini / Qulpu", [
    r"\brezin\w*", r"\bрезин\w*", r"\bgrip\b", r"\bhand\s?grip\w*",
])

add("Ümumi Texniki Nasazlıq", [
    r"\bislemir\w*", r"\bislemedi\w*", r"\bishlemir\w*", r"\bislemeyib\w*", r"\biwlemir\w*", r"\biwlemedi\w*",
    r"\bxarab\w*", r"\bnasaz\w*", r"\bsinib\w*", r"\bsiniq\w*", r"\bkohne\w*", r"\blaxli\w*", r"\btarazliq\w*",
    r"\baktivles\w*", r"\bacilmir\w*", r"\bacilmadi\w*", r"\bavariy?a\w*", r"\btemir\w*\s?ehtiyac\w*",
    r"\bise\s?dusmur\w*", r"\bise\s?dusmedi\w*", r"\bgetmir\w*", r"\byerimir\w*", r"\bhereket\s?etm(ir|edi)\w*",
    r"\bbaslamir\w*", r"\bterpen\w*", r"\bgitmiyor\w*",
    r"\bses\s?(gelir|edir|cixir|verir)\w*", r"\bqirig\w*", r"\bqiriq\w*",
    r"\bне\s?работа\w*", r"\bсломан\w*", r"\bнеисправ\w*", r"\bполомк\w*", r"\bне\s?включ\w*",
    r"\bне\s?поеха\w*", r"\bне\s?едет\w*", r"\bне\s?ехал\w*", r"\bне\s?запустил\w*", r"\bне\s?завел\w*",
    r"\bскрип\w*", r"\bвизж\w*",
    r"\bnot\s?work\w*", r"\bdoesn.?t\s?work\w*", r"\bbroken\b", r"\bdidn.?t\s?(go|start|move)\w*",
    r"\bwon.?t\s?start\w*", r"\bhard\s?to\s?move\w*", r"\bnot\s?start\w*",
    r"\bmalfunction\w*", r"\bglitch\w*", r"\bnoisy\b", r"\bsqueak\w*",
])

# ---------------- 2. TƏTBİQ VƏ RƏQƏMSAL ----------------

add("GPS / Xəritə / Location problemi", [
    r"\bxerite\w*", r"\bgps\b", r"\blokasiya\w*",
    r"\bкарт[аеы]\w*", r"\bлокаци\w*", r"\bместоположени\w*", r"\bнавигаци\w*",
    r"\bmap\b", r"\blocation\w*", r"\bnavigation\w*",
])

add("App və Qeydiyyat / Login / QR / Session", [
    r"\btetbiq\w*", r"\baplikasiya\w*", r"\bqeydiyyat\w*", r"\bgiris\s?et\w*",
    r"\bqr\s?kod\w*", r"\bqr\b", r"\bsessiya\w*", r"\bhesab\w*", r"\bsifr\w*",
    r"\bsonlandir\w*", r"\bbitire\s?bilm\w*", r"\bbitirm\w*\s?bilm\w*", r"\bbitirmek\s?problem\w*",
    r"\baktivles\w*", r"\baktiv\s?olmad\w*", r"\baktiv\s?et\w*", r"\byeniden\s?aktiv\w*",
    r"\bxeta\s?\d+", r"\bошибк\w*\s?\d+",
    r"\bпри?ложени\w*", r"\bрегистраци\w*", r"\bвход\w*", r"\bсесси\w*", r"\bпарол\w*",
    r"\bзаверш\w*\s?(поездк|аренд)\w*",
    r"\bapp\b", r"\bapplication\w*", r"\blog\s?in\w*", r"\bregist\w*", r"\bsession\w*",
    r"\baccount\w*", r"\bpassword\w*", r"\bend\s?ride\w*", r"\bcan.?t\s?end\w*",
])

add("Bluetooth bağlantısı", [
    r"\bbluetooth\w*", r"\bblutuz\w*", r"\bблютуз\w*",
    r"\bbaglant\w*\s?(qosul|itir)\w*", r"\bсоединени\w*",
    r"\bconnect\w*\s?(fail|lost|issue)\w*",
])

# ---------------- 3. ÖDƏNİŞ VƏ QİYMƏT ----------------

add("Qiymət Narazılığı", [
    r"\bbaha\w*", r"\bqiymet\w*", r"\bucuz\w*",
    r"\bдорог\w*", r"\bцен[аы]\w*", r"\bдешевл\w*",
    r"\bexpensive\w*", r"\bprice\w*", r"\bcostly\w*", r"\bcheap\w*",
])

add("Ödəniş Xətaları", [
    r"\bodenis\w*", r"\bpul\s?cix\w*", r"\bkartdan\s?pul\w*", r"\bbalans\w*", r"\bikiqat\s?tut\w*", r"\bmanat\s?yaz\w*",
    r"\bgeri\s?qaytar\w*", r"\bpulu\w*\s?geri\w*", r"\bpul\w*\s?qayit\w*",
    r"\bоплат\w*", r"\bденьги\s?списал\w*", r"\bбаланс\w*", r"\bдважды\s?списал\w*", r"\bпо?списан\w*",
    r"\bверните\w*",
    r"\bpayment\w*", r"\bcharged\s?twice\w*", r"\bdeduct\w*", r"\brefund\w*",
])

# ---------------- 4. İNFRASTRUKTUR ----------------

add("Parking Çatışmazlığı", [
    r"\bparking\w*", r"\bparkovka\w*", r"\bdayanacaq(?!\s?eylec)\w*", r"\bpark\s?yer\w*", r"\bparklanma\w*",
    r"\bпарков\w*", r"\bпаркинг\w*", r"\bparking\s?spot\w*",
])

add("Xidmət Zonası / Service Area", [
    r"\bxidmet\s?zona\w*", r"\bzona\w*",
    r"\bзон[аеы]\w*", r"\bservice\s?area\w*", r"\bzone\b",
])

add("Velosiped/Skuter Çatışmazlığı", [
    r"\bvelosiped\w*\s?(azdir|yoxdur|tapilmir|catismir)\w*", r"\bskuter\w*\s?(azdir|yoxdur|tapilmir)\w*",
    r"\bсамокат\w*\s?(мало|нет)\w*", r"\bвелосипед\w*\s?(мало|нет)\w*",
    r"\bnot\s?enough\s?(bike|scooter)\w*", r"\bno\s?(bike|scooter)s?\s?available\w*", r"\bunavailable\b",
    r"\bmore\s?(bike|scooter)s?\b",
    r"\bне\s?возможно\s?найти\w*", r"\bневозможно\s?найти\w*",
])

add("Velosiped Yolu / Yol İnfrastrukturu", [
    r"\bvelosiped\s?yol\w*", r"\bvelosiped\s?xett\w*", r"\byol\s?infrastruktur\w*", r"\bzolaq\w*",
    r"\bвелодорожк\w*", r"\bдорог(?!о\b)\w*",
    r"\bbike\s?lane\w*", r"\broad\b", r"\bpavement\w*",
])

# ---------------- 5. TƏKLİF VƏ İSTƏKLƏR ----------------

add("Yeni Funksiya və Təchizat Təklifi", [
    r"\bteklif\s?ed\w*", r"\bolsa\s?yaxsi\s?olar\w*", r"\bolsa\s?gozel\s?olar\w*", r"\belave\s?edin\w*",
    r"\belave\w*\s?edil\w*", r"\belave\w*\s?ols\w*", r"\bgelsin\b", r"\bqoyun\b", r"\bqoysaniz\w*",
    r"\bartirsaniz\w*", r"\bwant\s?more\w*",
    r"\bпредлагаю\w*", r"\bдобавьте\w*", r"\bпоставьте\w*", r"\bхотелось\s?бы\w*", r"\bнужно\s?сделать\w*",
    r"\bsuggest\w*", r"\bplease\s?add\w*", r"\bwould\s?be\s?nice\w*", r"\bit\s?would\s?be\s?good\s?if\w*",
])

# ---------------- 6. MÜSBƏT ----------------

add("Müsbət Məmnuniyyət", [
    r"\bela\b", r"\bsuper\b", r"\btesekkur\w*", r"\bsagol\w*", r"\bsag\s?olun\w*", r"\bgozel\w*", r"\byaxsi\w*",
    r"\bmukemmel\w*", r"\bbomba\w*", r"\bproblemsiz\w*",
    r"\bсупер\b", r"\bотличн\w*", r"\bспасибо\w*", r"\bкласс\b", r"\bхорош\w*", r"\bпревосходн\w*",
    r"\bgreat\b", r"\bgood\b", r"\bthanks\w*", r"\bthank\s?you\w*", r"\bawesome\w*", r"\bexcellent\w*",
    r"\blove\s?it\w*", r"\bnice\b", r"\bcool\b", r"\bperfect\w*",
    r"👍", r"❤", r"😍", r"🙏", r"🎉", r"😘", r"💯", r"🥰", r"😊",
])

CATEGORY_ORDER = list(CATS.keys())

CODE_MAP = {
    "Əyləc / Tormoz": "T01",
    "Pedal": "T02",
    "Zəncir": "T03",
    "Oturacaq": "T04",
    "Sükan": "T05",
    "Təkər / Şin / Hava": "T06",
    "Fiziki Kilid": "T07",
    "Dayaq ayağı": "T08",
    "Motor / Mühərrik / Elektrik dartısı": "T09",
    "Batareya": "T10",
    "Səbət": "T11",
    "Bolt / Vint": "T12",
    "Ümumi Texniki Nasazlıq": "T13",
    "Təhlükəsizlik Problemi": "T14",
    "Ümumi Baxım / Servis ehtiyacı": "T15",
    "Gear / Sürət pilləsi / Ötürücü": "T16",
    "Sürət / Performans": "T17",
    "Telefon tutacağı": "T18",
    "Sükan–Təkər Uyğunsuzluğu / Alignment": "T19",
    "İşıq / Fara": "T20",
    "Qaz / Throttle": "T21",
    "Siqnal / Zəng": "T22",
    "Disk / Jant": "T23",
    "Amortizator / Asma sistemi": "T24",
    "Gediş zamanı özbaşına dayanma / kilidlənmə": "T25",
    "Ekran / Göstərici Paneli": "T26",
    "Əl Tutacağı Rezini / Qulpu": "T27",
    "GPS / Xəritə / Location problemi": "A01",
    "App və Qeydiyyat / Login / QR / Session": "A02",
    "Bluetooth bağlantısı": "A03",
    "Qiymət Narazılığı": "P01",
    "Ödəniş Xətaları": "P02",
    "Parking Çatışmazlığı": "I01",
    "Xidmət Zonası / Service Area": "I02",
    "Velosiped/Skuter Çatışmazlığı": "I03",
    "Velosiped Yolu / Yol İnfrastrukturu": "I04",
    "Yeni Funksiya və Təchizat Təklifi": "S01",
    "Müsbət Məmnuniyyət": "M01",
}
UNCAT_NAME = "Təsnif olunmamış"
UNCAT_CODE = "O01"

def classify(text):
    t = normalize(text)
    hits = []
    for name, patterns in CATS.items():
        for p in patterns:
            if p.search(t):
                hits.append(name)
                break
    # "Ümumi Texniki Nasazlıq" is a catch-all: if a more specific technical
    # (or app/payment/infra) category already explains the comment, don't
    # also tag the generic one. (Confirmed against 500 manually labeled
    # comments: humans do not double-tag the generic bucket in this case.)
    GENERIC_CATCH_ALL = "Ümumi Texniki Nasazlıq"
    POSITIVE = "Müsbət Məmnuniyyət"
    if GENERIC_CATCH_ALL in hits:
        others = [h for h in hits if h not in (GENERIC_CATCH_ALL, POSITIVE)]
        if others:
            hits = [h for h in hits if h != GENERIC_CATCH_ALL]

    # A strong imperative request ("qoyun", "gəlsin", "добавьте"...) means the
    # person is asking for something new, not describing today's shortage —
    # in that case keep only the suggestion tag, not the infrastructure one.
    STRONG_REQUEST = re.compile(
        r"\bgelsin\b|\bqoyun\b|\bqoysaniz\w*|\bartirsaniz\w*|\bдобавьте\w*|\bпоставьте\w*|"
        r"\bнужно\s?сделать\w*|\bwant\s?more\w*",
        re.IGNORECASE,
    )
    if "Yeni Funksiya və Təchizat Təklifi" in hits and STRONG_REQUEST.search(t):
        hits = [h for h in hits if h not in ("Parking Çatışmazlığı", "Velosiped/Skuter Çatışmazlığı")]

    return hits

if __name__ == "__main__":
    tests = [
        "sağ əyləc tutmur",
        "не смог отрегулировать кресло, сиденье зажато",
        "qiymət bahadı)",
        "high price",
        "ulduza parking gəlsin",
        "əla",
        "спасибо",
        "не работает ничего",
        "batareya 0 dadı",
        "GPS düzgün göstərmir",
    ]
    for t in tests:
        print(t, "->", classify(t))
