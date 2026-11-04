"""ประเทศ/ธงชาติสำหรับรายการนานาชาติ (ชิงแชมป์โลก)

ioc   = รหัสที่ใช้ในการแข่งขัน (FIPJP/IOC)
flag  = ชื่อไฟล์ธงใน static/flags/<flag>.svg (ธงจากชุด flag-icons, MIT)
names = ชื่อที่ยอมรับเวลาจับคู่สังกัด (อังกฤษ / ฝรั่งเศส / ไทย)
"""

NATIONS = [
    # ioc, flag, french, english, thai, extra aliases
    ("ALG", "dz", "Algérie", "Algeria", "แอลจีเรีย", ()),
    ("GER", "de", "Allemagne", "Germany", "เยอรมนี", ()),
    ("AND", "ad", "Andorre", "Andorra", "อันดอร์รา", ()),
    ("ENG", "gb-eng", "Angleterre", "England", "อังกฤษ", ()),
    ("ARG", "ar", "Argentine", "Argentina", "อาร์เจนตินา", ()),
    ("AUS", "au", "Australie", "Australia", "ออสเตรเลีย", ()),
    ("BEL", "be", "Belgique", "Belgium", "เบลเยียม", ()),
    ("BRU", "bn", "Brunei", "Brunei", "บรูไน", ("Brunei Darussalam",)),
    ("BFA", "bf", "Burkina Faso", "Burkina Faso", "บูร์กินาฟาโซ", ()),
    ("BEN", "bj", "Bénin", "Benin", "เบนิน", ()),
    ("CMR", "cm", "Cameroun", "Cameroon", "แคเมอรูน", ()),
    ("CAN", "ca", "Canada", "Canada", "แคนาดา", ()),
    ("CHN", "cn", "Chine", "China", "จีน", ()),
    ("TPE", "tpe", "Chinese Taipei", "Chinese Taipei", "จีนไทเป", ()),
    ("COM", "km", "Comores", "Comoros", "คอโมโรส", ()),
    ("CIV", "ci", "Côte d'Ivoire", "Ivory Coast", "โกตดิวัวร์", ("Côte d’Ivoire", "Cote d'Ivoire")),
    ("DEN", "dk", "Danemark", "Denmark", "เดนมาร์ก", ()),
    ("ESP", "es", "Espagne", "Spain", "สเปน", ()),
    ("EST", "ee", "Estonie", "Estonia", "เอสโตเนีย", ()),
    ("FIN", "fi", "Finlande", "Finland", "ฟินแลนด์", ()),
    ("FRA", "fr", "France", "France", "ฝรั่งเศส", ()),
    ("HUN", "hu", "Hongrie", "Hungary", "ฮังการี", ()),
    ("MRI", "mu", "Île Maurice", "Mauritius", "มอริเชียส", ("Ile Maurice", "Maurice")),
    ("INA", "id", "Indonésie", "Indonesia", "อินโดนีเซีย", ()),
    ("IRL", "ie", "Irlande", "Ireland", "ไอร์แลนด์", ()),
    ("ISR", "il", "Israël", "Israel", "อิสราเอล", ()),
    ("ITA", "it", "Italie", "Italy", "อิตาลี", ()),
    ("JPN", "jp", "Japon", "Japan", "ญี่ปุ่น", ()),
    ("LIB", "lb", "Liban", "Lebanon", "เลบานอน", ()),
    ("LUX", "lu", "Luxembourg", "Luxembourg", "ลักเซมเบิร์ก", ()),
    ("MAD", "mg", "Madagascar", "Madagascar", "มาดากัสการ์", ()),
    ("MAS", "my", "Malaisie", "Malaysia", "มาเลเซีย", ()),
    ("MLI", "ml", "Mali", "Mali", "มาลี", ()),
    ("MTN", "mr", "Mauritanie", "Mauritania", "มอริเตเนีย", ()),
    ("MON", "mc", "Monaco", "Monaco", "โมนาโก", ()),
    ("NIG", "ne", "Niger", "Niger", "ไนเจอร์", ()),
    ("NED", "nl", "Pays-Bas", "Netherlands", "เนเธอร์แลนด์", ("Holland",)),
    ("POL", "pl", "Pologne", "Poland", "โปแลนด์", ()),
    ("CZE", "cz", "Tchéquie", "Czech Republic", "เช็กเกีย", ("Czechia", "République tchèque")),
    ("SIN", "sg", "Singapour", "Singapore", "สิงคโปร์", ()),
    ("SVK", "sk", "Slovaquie", "Slovakia", "สโลวาเกีย", ()),
    ("SWE", "se", "Suède", "Sweden", "สวีเดน", ()),
    ("SUI", "ch", "Suisse", "Switzerland", "สวิตเซอร์แลนด์", ()),
    ("SEN", "sn", "Sénégal", "Senegal", "เซเนกัล", ()),
    ("PYF", "pf", "Tahiti", "Tahiti", "ตาฮีตี", ("French Polynesia", "Polynésie française")),
    ("THA", "th", "Thaïlande", "Thailand", "ไทย", ("ประเทศไทย",)),
    ("TUN", "tn", "Tunisie", "Tunisia", "ตูนิเซีย", ()),
    ("USA", "us", "États-Unis", "United States", "สหรัฐอเมริกา", ("USA", "United States of America", "Etats-Unis")),
    # ประเทศที่มักร่วมรายการเปตองอื่น ๆ (เผื่อรายการเอเชีย/ซีเกมส์)
    ("LAO", "la", "Laos", "Laos", "ลาว", ("Lao PDR",)),
    ("CAM", "kh", "Cambodge", "Cambodia", "กัมพูชา", ()),
    ("VIE", "vn", "Viêt Nam", "Vietnam", "เวียดนาม", ("Viet Nam",)),
    ("PHI", "ph", "Philippines", "Philippines", "ฟิลิปปินส์", ()),
    ("MYA", "mm", "Myanmar", "Myanmar", "เมียนมา", ()),
    ("TLS", "tl", "Timor-Leste", "Timor-Leste", "ติมอร์-เลสเต", ()),
    ("KOR", "kr", "Corée du Sud", "South Korea", "เกาหลีใต้", ("Korea",)),
    ("IND", "in", "Inde", "India", "อินเดีย", ()),
    ("HKG", "hk", "Hong Kong", "Hong Kong", "ฮ่องกง", ()),
    ("MAR", "ma", "Maroc", "Morocco", "โมร็อกโก", ()),
    ("POR", "pt", "Portugal", "Portugal", "โปรตุเกส", ()),
    ("SCO", "gb-sct", "Écosse", "Scotland", "สกอตแลนด์", ()),
    ("WAL", "gb-wls", "Pays de Galles", "Wales", "เวลส์", ()),
    ("NZL", "nz", "Nouvelle-Zélande", "New Zealand", "นิวซีแลนด์", ()),
    ("AUT", "at", "Autriche", "Austria", "ออสเตรีย", ()),
    ("NOR", "no", "Norvège", "Norway", "นอร์เวย์", ()),
    ("TUR", "tr", "Turquie", "Türkiye", "ตุรกี", ("Turkey", "TÜRKIYE")),
    ("SLO", "si", "Slovénie", "Slovenia", "สโลวีเนีย", ()),
    ("CRO", "hr", "Croatie", "Croatia", "โครเอเชีย", ()),
    ("BRA", "br", "Brésil", "Brazil", "บราซิล", ()),
    ("TOG", "tg", "Togo", "Togo", "โตโก", ()),
    ("GUI", "gn", "Guinée", "Guinea", "กินี", ()),
    ("DJI", "dj", "Djibouti", "Djibouti", "จิบูตี", ()),
    ("COD", "cd", "RD Congo", "DR Congo", "คองโก", ()),
]

FLAG_FILES = sorted({n[1] for n in NATIONS})


def _norm(text: str) -> str:
    import unicodedata
    text = unicodedata.normalize("NFKD", str(text or "")).encode("ascii", "ignore").decode() or str(text or "")
    return "".join(ch for ch in text.lower().replace("’", "'") if ch.isalnum())


def _build_index():
    idx = {}
    for ioc, flag, fr, en, th, extra in NATIONS:
        for name in (ioc, fr, en, th, *extra):
            key = _norm(name) or str(name).strip().lower()
            idx.setdefault(key, ioc)
            idx.setdefault(str(name).strip().lower(), ioc)
    return idx


_INDEX = _build_index()
BY_IOC = {n[0]: n for n in NATIONS}


def nation_code(text) -> str | None:
    """หา IOC code จากข้อความสังกัด: 'Belgique (BEL)', 'BEL', 'BELGIUM', 'เบลเยียม' → 'BEL'"""
    import re
    raw = str(text or "").strip()
    if not raw:
        return None
    m = re.search(r"\(([A-Za-z]{3})\)\s*$", raw)
    if m and m.group(1).upper() in BY_IOC:
        return m.group(1).upper()
    for key in (raw.lower(), _norm(raw)):
        if key and key in _INDEX:
            return _INDEX[key]
    # 'BEL - Belgique' หรือ 'BEL Belgique'
    m = re.match(r"^([A-Za-z]{3})\b", raw)
    if m and m.group(1).upper() in BY_IOC and len(raw) > 3 and not raw[3:4].isalpha():
        return m.group(1).upper()
    return None


def flag_file(text) -> str | None:
    code = nation_code(text)
    return BY_IOC[code][1] if code else None


def client_index() -> dict:
    """ตารางให้ฝั่งหน้าเว็บ: ชื่อที่ normalize แล้ว → ชื่อไฟล์ธง"""
    return {key: BY_IOC[ioc][1] for key, ioc in _INDEX.items()}
