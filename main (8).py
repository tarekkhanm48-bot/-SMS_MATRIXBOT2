import os, sys, re, sqlite3, datetime, time
import requests as http_requests
from threading import Thread

from flask import Flask, request as flask_req
from telethon import TelegramClient, events, Button
from telethon.tl.functions.channels import GetParticipantRequest
from telethon.tl.functions.bots import SetBotCommandsRequest
from telethon.tl.types import BotCommand, BotCommandScopeDefault
import openpyxl

# ════════════════════════════════════════════════════════════
#  CONFIG
# ════════════════════════════════════════════════════════════
API_ID      = 35358753
API_HASH    = '0df6c2d76a90d00bf022f5def552cbb9'
BOT_TOKEN   = '8610107064:AAHfEg5skZBUGj7rqedDyXhNnZLsCHuFYuI'
SUPER_ADMIN = 8048890397
OTP_GROUP   = -1003926163328   # numeric supergroup ID — always works
BOT_FILE    = os.path.abspath(__file__)

FORCE_SUBS = {
    "UnlimitedWSMethod0":  "📢 Channel 1",
    "QueenIraMoniBithiBD": "📢 Channel 2",
    "QueenIraMoniBithi_BD":"💬 Group 1",
    "unlimited_ws_method": "💬 Group 2",
    "SMS_MatrixBOT_OTP":   "🔥 OTP Group",
}

SVC = {
    # Messaging / Social
    "whatsapp":          "💬 WhatsApp",
    "whatsapp_business": "💼 WhatsApp Business",
    "telegram":          "🔹 Telegram",
    "facebook":          "🌐 Facebook",
    "instagram":         "📸 Instagram",
    "tiktok":            "🎵 TikTok",
    "twitter":           "🐦 Twitter / X",
    "snapchat":          "👻 Snapchat",
    "discord":           "🎮 Discord",
    "viber":             "📳 Viber",
    "line":              "💚 Line",
    "wechat":            "💚 WeChat",
    "linkedin":          "💼 LinkedIn",
    # Tech giants
    "google":            "🔍 Google",
    "apple":             "🍎 Apple",
    "microsoft":         "🪟 Microsoft",
    "yahoo":             "🟣 Yahoo",
    # Entertainment / Streaming
    "netflix":           "🎬 Netflix",
    "spotify":           "🎵 Spotify",
    "steam":             "🎮 Steam",
    # E-commerce / Finance
    "amazon":            "🛒 Amazon",
    "shopee":            "🛍 Shopee",
    "lazada":            "🛍 Lazada",
    "paypal":            "💳 PayPal",
    "binance":           "🪙 Binance",
    "bybit":             "🪙 ByBit",
    "okx":               "🪙 OKX",
    # Asian apps
    "naver":             "🟢 Naver",
    "rednote":           "📕 RedNote",
    "grab":              "🚗 Grab",
    "uber":              "🚖 Uber",
}

COUNTRIES = [
    ("Afghanistan","af","🇦🇫"),("Albania","al","🇦🇱"),("Algeria","dz","🇩🇿"),
    ("Angola","ao","🇦🇴"),("Argentina","ar","🇦🇷"),("Armenia","am","🇦🇲"),
    ("Australia","au","🇦🇺"),("Austria","at","🇦🇹"),("Azerbaijan","az","🇦🇿"),
    ("Bahrain","bh","🇧🇭"),("Bangladesh","bd","🇧🇩"),("Belarus","by","🇧🇾"),
    ("Belgium","be","🇧🇪"),("Bolivia","bo","🇧🇴"),("Brazil","br","🇧🇷"),
    ("Bulgaria","bg","🇧🇬"),("Burundi","bi","🇧🇮"),("Cambodia","kh","🇰🇭"),
    ("Cameroon","cm","🇨🇲"),("Canada","ca","🇨🇦"),("Chad","td","🇹🇩"),
    ("Chile","cl","🇨🇱"),("China","cn","🇨🇳"),("Colombia","co","🇨🇴"),
    ("Congo","cg","🇨🇬"),("Costa Rica","cr","🇨🇷"),("Croatia","hr","🇭🇷"),
    ("Cuba","cu","🇨🇺"),("Czech Republic","cz","🇨🇿"),("DR Congo","cd","🇨🇩"),
    ("Denmark","dk","🇩🇰"),("Dominican Rep","do","🇩🇴"),("Ecuador","ec","🇪🇨"),
    ("Egypt","eg","🇪🇬"),("El Salvador","sv","🇸🇻"),("Ethiopia","et","🇪🇹"),
    ("Finland","fi","🇫🇮"),("France","fr","🇫🇷"),("Germany","de","🇩🇪"),
    ("Ghana","gh","🇬🇭"),("Greece","gr","🇬🇷"),("Guatemala","gt","🇬🇹"),
    ("Guinea","gn","🇬🇳"),("Honduras","hn","🇭🇳"),("Hungary","hu","🇭🇺"),
    ("India","in","🇮🇳"),("Indonesia","id","🇮🇩"),("Iran","ir","🇮🇷"),
    ("Iraq","iq","🇮🇶"),("Ireland","ie","🇮🇪"),("Israel","il","🇮🇱"),
    ("Italy","it","🇮🇹"),("Jamaica","jm","🇯🇲"),("Japan","jp","🇯🇵"),
    ("Jordan","jo","🇯🇴"),("Kazakhstan","kz","🇰🇿"),("Kenya","ke","🇰🇪"),
    ("Kuwait","kw","🇰🇼"),("Kyrgyzstan","kg","🇰🇬"),("Laos","la","🇱🇦"),
    ("Lebanon","lb","🇱🇧"),("Libya","ly","🇱🇾"),("Madagascar","mg","🇲🇬"),
    ("Malawi","mw","🇲🇼"),("Malaysia","my","🇲🇾"),("Mali","ml","🇲🇱"),
    ("Mexico","mx","🇲🇽"),("Moldova","md","🇲🇩"),("Mongolia","mn","🇲🇳"),
    ("Morocco","ma","🇲🇦"),("Mozambique","mz","🇲🇿"),("Myanmar","mm","🇲🇲"),
    ("Nepal","np","🇳🇵"),("Netherlands","nl","🇳🇱"),("New Zealand","nz","🇳🇿"),
    ("Nicaragua","ni","🇳🇮"),("Niger","ne","🇳🇪"),("Nigeria","ng","🇳🇬"),
    ("Norway","no","🇳🇴"),("Oman","om","🇴🇲"),("Pakistan","pk","🇵🇰"),
    ("Palestine","ps","🇵🇸"),("Panama","pa","🇵🇦"),("Paraguay","py","🇵🇾"),
    ("Peru","pe","🇵🇪"),("Philippines","ph","🇵🇭"),("Poland","pl","🇵🇱"),
    ("Portugal","pt","🇵🇹"),("Qatar","qa","🇶🇦"),("Romania","ro","🇷🇴"),
    ("Russia","ru","🇷🇺"),("Rwanda","rw","🇷🇼"),("Saudi Arabia","sa","🇸🇦"),
    ("Senegal","sn","🇸🇳"),("Serbia","rs","🇷🇸"),("Sierra Leone","sl","🇸🇱"),
    ("Somalia","so","🇸🇴"),("South Africa","za","🇿🇦"),("South Korea","kr","🇰🇷"),
    ("Spain","es","🇪🇸"),("Sri Lanka","lk","🇱🇰"),("Sudan","sd","🇸🇩"),
    ("Sweden","se","🇸🇪"),("Switzerland","ch","🇨🇭"),("Syria","sy","🇸🇾"),
    ("Taiwan","tw","🇹🇼"),("Tajikistan","tj","🇹🇯"),("Tanzania","tz","🇹🇿"),
    ("Thailand","th","🇹🇭"),("Togo","tg","🇹🇬"),("Tunisia","tn","🇹🇳"),
    ("Turkey","tr","🇹🇷"),("Turkmenistan","tm","🇹🇲"),("Uganda","ug","🇺🇬"),
    ("Ukraine","ua","🇺🇦"),("UAE","ae","🇦🇪"),("UK","uk","🇬🇧"),
    ("USA","us","🇺🇸"),("Uruguay","uy","🇺🇾"),("Uzbekistan","uz","🇺🇿"),
    ("Venezuela","ve","🇻🇪"),("Vietnam","vn","🇻🇳"),("Yemen","ye","🇾🇪"),
    ("Zambia","zm","🇿🇲"),("Zimbabwe","zw","🇿🇼"),
]
COUNTRY_MAP = {s: (n, f) for n, s, f in COUNTRIES}

# International calling codes — used to filter uploaded numbers by country
DIALCODE: dict[str, list[str]] = {
    "af":["+93"],"al":["+355"],"dz":["+213"],"ao":["+244"],"ar":["+54"],
    "am":["+374"],"au":["+61"],"at":["+43"],"az":["+994"],"bh":["+973"],
    "bd":["+880"],"by":["+375"],"be":["+32"],"bo":["+591"],"br":["+55"],
    "bg":["+359"],"bi":["+257"],"kh":["+855"],"cm":["+237"],"ca":["+1"],
    "td":["+235"],"cl":["+56"],"cn":["+86"],"co":["+57"],"cg":["+242"],
    "cr":["+506"],"hr":["+385"],"cu":["+53"],"cz":["+420"],"cd":["+243"],
    "dk":["+45"],"do":["+1809","+1829","+1849","+1"],"ec":["+593"],"eg":["+20"],
    "sv":["+503"],"et":["+251"],"fi":["+358"],"fr":["+33"],"de":["+49"],
    "gh":["+233"],"gr":["+30"],"gt":["+502"],"gn":["+224"],"hn":["+504"],
    "hu":["+36"],"in":["+91"],"id":["+62"],"ir":["+98"],"iq":["+964"],
    "ie":["+353"],"il":["+972"],"it":["+39"],"jm":["+1876","+1"],"jp":["+81"],
    "jo":["+962"],"kz":["+77","+76","+7"],"ke":["+254"],"kw":["+965"],
    "kg":["+996"],"la":["+856"],"lb":["+961"],"ly":["+218"],"mg":["+261"],
    "mw":["+265"],"my":["+60"],"ml":["+223"],"mx":["+52"],"md":["+373"],
    "mn":["+976"],"ma":["+212"],"mz":["+258"],"mm":["+95"],"np":["+977"],
    "nl":["+31"],"nz":["+64"],"ni":["+505"],"ne":["+227"],"ng":["+234"],
    "no":["+47"],"om":["+968"],"pk":["+92"],"ps":["+970"],"pa":["+507"],
    "py":["+595"],"pe":["+51"],"ph":["+63"],"pl":["+48"],"pt":["+351"],
    "qa":["+974"],"ro":["+40"],"ru":["+7"],"rw":["+250"],"sa":["+966"],
    "sn":["+221"],"rs":["+381"],"sl":["+232"],"so":["+252"],"za":["+27"],
    "kr":["+82"],"es":["+34"],"lk":["+94"],"sd":["+249"],"se":["+46"],
    "ch":["+41"],"sy":["+963"],"tw":["+886"],"tj":["+992"],"tz":["+255"],
    "th":["+66"],"tg":["+228"],"tn":["+216"],"tr":["+90"],"tm":["+993"],
    "ug":["+256"],"ua":["+380"],"ae":["+971"],"uk":["+44"],"us":["+1"],
    "uy":["+598"],"uz":["+998"],"ve":["+58"],"vn":["+84"],"ye":["+967"],
    "zm":["+260"],"zw":["+263"],
}

def num_matches_country(phone: str, short: str) -> bool:
    """Return True if phone number starts with any calling code for the given country."""
    prefixes = DIALCODE.get(short, [])
    if not prefixes:
        return True  # unknown country → allow all
    norm = phone.strip()
    if not norm.startswith('+'):
        norm = '+' + norm
    return any(norm.startswith(p) for p in prefixes)

# ════════════════════════════════════════════════════════════
#  FLASK keep-alive + OTP webhook
# ════════════════════════════════════════════════════════════
flask_app = Flask('')

def mask_number(phone: str) -> str:
    """
    Mask middle digits with × symbols.
    Always adds + prefix.
    Format: +{first n-4 digits}××{last 2 digits}
    Example: +99271009959 → +9927100××59
             +99298585094 → +99298580××94
    """
    digits = re.sub(r'[^\d]', '', phone)
    n = len(digits)
    if n <= 5:
        return '+' + digits
    n_front = max(5, n - 4)
    return '+' + digits[:n_front] + '××' + digits[-2:]

@flask_app.route('/')
def home(): return "✅ SMS MATRIX BOT is running!"

@flask_app.route('/alive')
def alive(): return "OK", 200

import threading as _threading
_otp_lock = _threading.Lock()

def _detect_country_from_phone(norm_digits: str):
    """Fallback: detect country short/name/flag from phone prefix using DIALCODE."""
    # sort by longest prefix first to avoid false matches (+1 vs +1876)
    sorted_codes = sorted(
        ((code.lstrip('+'), short) for short, codes in DIALCODE.items() for code in codes),
        key=lambda x: -len(x[0])
    )
    for prefix_digits, short in sorted_codes:
        if norm_digits.startswith(prefix_digits):
            info = COUNTRY_MAP.get(short)
            if info:
                return short, info[0], info[1]
    return '', 'Unknown', '🌍'

def _detect_service_from_msg(msg: str) -> str:
    """Detect service key from OTP message text — WhatsApp Business first (more specific)."""
    ml = msg.lower()
    # WhatsApp Business — check before plain WhatsApp (more specific)
    if ('whatsapp business' in ml or 'wabiz' in ml or
            'wa business' in ml or 'whatsapp for business' in ml or
            'whatsapp_business' in ml or
            ('whatsapp' in ml and 'business' in ml)):
        return 'whatsapp_business'
    if 'whatsapp' in ml:                            return 'whatsapp'
    if 'telegram' in ml:                            return 'telegram'
    if 'facebook' in ml or ' fb ' in ml:            return 'facebook'
    if 'instagram' in ml:                           return 'instagram'
    if 'tiktok' in ml or 'tik tok' in ml or 'tik-tok' in ml: return 'tiktok'
    if 'twitter' in ml or 'x.com' in ml:            return 'twitter'
    if 'snapchat' in ml:                            return 'snapchat'
    if 'discord' in ml:                             return 'discord'
    if 'viber' in ml:                               return 'viber'
    if 'wechat' in ml or 'weixin' in ml:            return 'wechat'
    if 'line' == ml.strip() or 'line app' in ml or 'line ' in ml or ml.endswith('line'): return 'line'
    if 'linkedin' in ml:                            return 'linkedin'
    if 'google' in ml or 'gmail' in ml:             return 'google'
    if 'apple' in ml or 'icloud' in ml or 'apple id' in ml: return 'apple'
    if 'microsoft' in ml or 'outlook' in ml or 'hotmail' in ml: return 'microsoft'
    if 'yahoo' in ml:                               return 'yahoo'
    if 'netflix' in ml:                             return 'netflix'
    if 'spotify' in ml:                             return 'spotify'
    if 'steam' in ml:                               return 'steam'
    if 'amazon' in ml:                              return 'amazon'
    if 'shopee' in ml:                              return 'shopee'
    if 'lazada' in ml:                              return 'lazada'
    if 'paypal' in ml:                              return 'paypal'
    if 'binance' in ml:                             return 'binance'
    if 'bybit' in ml:                               return 'bybit'
    if 'okx' in ml or 'okex' in ml:                return 'okx'
    if 'naver' in ml:                               return 'naver'
    if 'rednote' in ml or 'red note' in ml or 'xiaohongshu' in ml: return 'rednote'
    if 'grab' in ml:                                return 'grab'
    if 'uber' in ml:                                return 'uber'
    return ''


def _detect_language(text: str) -> tuple:
    """Detect language of the OTP message from Unicode character ranges.
    Returns (language_name, emoji)."""
    if re.search(r'[\u0980-\u09FF]', text):           return ('বাংলা',    '🇧🇩')
    if re.search(r'[\u0600-\u06FF\u0750-\u077F]', text): return ('Arabic',  '🌙')
    if re.search(r'[\u4E00-\u9FFF\u3400-\u4DBF\uF900-\uFAFF]', text): return ('Chinese', '🇨🇳')
    if re.search(r'[\u3040-\u309F\u30A0-\u30FF]', text): return ('Japanese','🇯🇵')
    if re.search(r'[\uAC00-\uD7A3\u1100-\u11FF]', text): return ('Korean',  '🇰🇷')
    if re.search(r'[\u0400-\u04FF]', text):           return ('Russian',   '🇷🇺')
    if re.search(r'[\u0E00-\u0E7F]', text):           return ('Thai',      '🇹🇭')
    if re.search(r'[\u0900-\u097F]', text):           return ('Hindi',     '🇮🇳')
    if re.search(r'[\u0600-\u06FF]', text):           return ('Urdu',      '🇵🇰')
    if re.search(r'[\u0750-\u077F\u08A0-\u08FF]', text): return ('Arabic', '🌙')
    return ('English', '🔤')

def _send_otp_notifications(phone: str, raw_msg: str, sender: str,
                            service_hint: str = '', panel_name: str = ''):
    """
    Core OTP notification sender — called from both webhook (main thread)
    and panel pollers (thread executors).
    Uses its own DB connection to avoid cursor race conditions.
    panel_name: if provided, increments forwarded_stats for that panel today.
    """
    now  = str(datetime.datetime.now())
    norm = re.sub(r'[^\d]', '', phone)
    tg_url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"

    # ── Thread-safe DB block ─────────────────────────────────────
    uid_owner     = None
    country_short = ''
    service_key   = ''
    support_link  = 'https://t.me/SMS_MatrixBOT_OTP'

    with _otp_lock:
        try:
            _db2 = sqlite3.connect(DB, check_same_thread=False)
            _c2  = _db2.cursor()

            _c2.execute("INSERT INTO otp_log(phone,sender,received_at) VALUES(?,?,?)",
                        (phone, sender, now))

            # ── Owner lookup — try multiple normalizations ───────────
            # Strip +, -, spaces, parentheses from both sides
            norm_variants = [norm]
            if norm.startswith('0'):  norm_variants.append(norm[1:])
            if len(norm) > 10:        norm_variants.append(norm[-10:])  # last 10 digits fallback
            uid_owner = None
            for nv in norm_variants:
                _c2.execute(
                    "SELECT user_id FROM user_number_assignments "
                    "WHERE REPLACE(REPLACE(REPLACE(REPLACE(number,'+',''),'-',''),' ',''),'(','')=?",
                    (nv,))
                row = _c2.fetchone()
                if row:
                    uid_owner = row[0]; break

            _c2.execute("SELECT country, service FROM premium_stock "
                        "WHERE REPLACE(REPLACE(number,'+',''),'-','')=?", (norm,))
            srow = _c2.fetchone()
            country_short = srow[0] if srow else ''
            service_key   = srow[1] if srow else ''

            if uid_owner:
                td = str(datetime.date.today())
                _c2.execute("SELECT otp_count FROM history WHERE user_id=? AND date=?",
                            (uid_owner, td))
                h = _c2.fetchone()
                if h:
                    _c2.execute("UPDATE history SET otp_count=otp_count+1 "
                                "WHERE user_id=? AND date=?", (uid_owner, td))
                else:
                    _c2.execute("INSERT INTO history VALUES(?,?,1,0)", (uid_owner, td))

            _c2.execute("SELECT value FROM bot_links WHERE key='support_group'")
            lr = _c2.fetchone()
            if lr: support_link = lr[0]

            # ── Track forwarded_stats per panel per day ──────────
            if panel_name:
                _td = str(datetime.date.today())
                _c2.execute(
                    "INSERT INTO forwarded_stats(panel_name,date,forwarded_count) "
                    "VALUES(?,?,1) ON CONFLICT(panel_name,date) "
                    "DO UPDATE SET forwarded_count=forwarded_count+1",
                    (panel_name, _td))

            _db2.commit()
            _db2.close()
        except Exception as e:
            print(f"[OTP-DB] {e}")

    # ── Resolve country/service labels (with fallback detection) ──
    country_name, country_flag = 'Unknown', '🌍'
    if country_short:
        info = COUNTRY_MAP.get(country_short)
        if info:
            country_name, country_flag = info[0], info[1]
    if country_name == 'Unknown' and norm:
        _, country_name, country_flag = _detect_country_from_phone(norm)

    # ── Service resolution (4-step priority chain) ───────────────
    # 1. DB  →  2. service_hint (5sim product / panel tag)
    # 3. SMS sender field  →  4. keyword scan of message text
    if not service_key and service_hint:
        service_key = _detect_service_from_msg(service_hint) or service_hint.lower().strip()
    if not service_key and sender and sender.lower() not in ('unknown', ''):
        service_key = _detect_service_from_msg(sender)
    if not service_key:
        service_key = _detect_service_from_msg(raw_msg)

    # Clean up raw 5sim product names that already match our SVC keys
    if service_key and service_key not in SVC:
        sk = service_key.lower().replace(' ', '_').replace('-', '_')
        if sk in SVC:
            service_key = sk

    svc_name = SVC.get(service_key, service_key.replace('_', ' ').title() if service_key else '🔍 Unknown')
    masked   = mask_number(phone)

    # ── Extract OTP code — supports plain digits AND hyphenated (e.g. 377-302) ──
    # Priority 1: hyphenated format  "377-302"  (WhatsApp Business style)
    _hyph = re.findall(r'(?<!\d)(\d{3,4}-\d{3,4})(?!\d)', raw_msg)
    # Priority 2: plain digits not glued to letters  "43142"  "130737"
    _plain = re.findall(r'(?<!\d)(\d{4,8})(?!\d)', raw_msg)
    otp_code = _hyph[0] if _hyph else (_plain[0] if _plain else '')

    # ── Detect language of the incoming message ───────────────────
    lang_name, lang_emoji = _detect_language(raw_msg)

    # ── Build copy button row ─────────────────────────────────────
    def _copy_btn(code):
        if not code:
            return []
        return [[{"text": f"📋 Copy: {code}", "copy_text": {"text": code}}]]

    # ── Ensure number display always has + prefix ─────────────────
    def _fmt_num(num: str) -> str:
        digits = re.sub(r'[^\d]', '', num)
        return '+' + digits if digits else num

    # ── OTP message body ──────────────────────────────────────────
    def _otp_body(number_display, is_private=False):
        otp_display = otp_code if otp_code else raw_msg
        owner_line  = f"\n👤 *আপনার Number এর OTP*" if is_private else ""
        body = (
            f"🎉 *NEW OTP RECEIVED* 🎉\n\n"
            f"🌐 *Country:* {country_name} {country_flag}\n"
            f"📱 *Number:* `{number_display}`\n"
            f"🚨 *Service:* {svc_name}\n"
            f"📩 *OTP:* `{otp_display}`\n"
            f"{lang_emoji} {lang_name}"
            f"{owner_line}"
        )
        return body

    masked = mask_number(phone)                  # +992710××59  format
    full_fmt = _fmt_num(phone)                   # +992710009959 format (always + prefix)

    # ══════════════════════════════════════════════════════════════
    # 1. PRIVATE — number owner এর কাছে পাঠাও (FULL number দেখাবে)
    # ══════════════════════════════════════════════════════════════
    _priv_sent = False
    if uid_owner:
        priv_kbd = {"inline_keyboard":
            _copy_btn(otp_code) +
            [[{"text": "📢 OTP Group ↗", "url": support_link}]]
        }
        priv_body = _otp_body(full_fmt, is_private=True)
        for _retry in range(5):          # 5 attempts before queuing
            try:
                r = http_requests.post(tg_url, json={
                    "chat_id": uid_owner,
                    "text": priv_body,
                    "parse_mode": "Markdown",
                    "reply_markup": priv_kbd
                }, timeout=12)
                if r.status_code == 200:
                    _priv_sent = True
                    print(f"[OTP-PRIV] ✅ Sent to owner uid={uid_owner}")
                    break
                time.sleep(2)
            except Exception as _pe:
                time.sleep(2)
        if not _priv_sent:
            # Queue private delivery — retry task will re-attempt
            try:
                _qdb = sqlite3.connect(DB, check_same_thread=False)
                _qdb.execute(
                    "INSERT INTO failed_otp_queue(phone,raw_msg,sender,queued_at,retry_count,last_try) "
                    "VALUES(?,?,?,?,0,?)",
                    (phone, f"[PRIV:{uid_owner}] " + raw_msg,
                     sender, str(datetime.datetime.now()), str(datetime.datetime.now())))
                _qdb.commit(); _qdb.close()
                print(f"[OTP-QUEUE] Private delivery queued for uid={uid_owner}")
            except Exception as _qe:
                print(f"[OTP-QUEUE] Queue error: {_qe}")

    # ══════════════════════════════════════════════════════════════
    # 2. GROUP — masked number দেখাবে, সবার জন্য
    # ══════════════════════════════════════════════════════════════
    group_kbd = {"inline_keyboard":
        _copy_btn(otp_code) +
        [[{"text": "🚀 Community",  "url": support_link},
          {"text": "🤖 Get Number", "url": "https://t.me/SMS_MATRIXBOT"}]]
    }
    _group_sent = False
    for _retry in range(5):
        try:
            r = http_requests.post(tg_url, json={
                "chat_id": OTP_GROUP,
                "text": _otp_body(masked),
                "parse_mode": "Markdown",
                "reply_markup": group_kbd
            }, timeout=12)
            if r.status_code == 200:
                _group_sent = True; break
            time.sleep(2)
        except Exception:
            time.sleep(2)

    # ── Retry queue: group forward failed → persist ───────────────
    if not _group_sent:
        try:
            _qdb = sqlite3.connect(DB, check_same_thread=False)
            _qdb.execute(
                "INSERT INTO failed_otp_queue(phone,raw_msg,sender,queued_at,retry_count,last_try) "
                "VALUES(?,?,?,?,0,?)",
                (phone, raw_msg, sender,
                 str(datetime.datetime.now()), str(datetime.datetime.now())))
            _qdb.commit(); _qdb.close()
            print(f"[OTP-QUEUE] Group delivery queued for {phone[:8]}…")
        except Exception as _qe:
            print(f"[OTP-QUEUE] Queue error: {_qe}")

    # ══════════════════════════════════════════════════════════════
    # 3. ADMIN alert — full number + owner info
    # ══════════════════════════════════════════════════════════════
    admin_txt = (
        _otp_body(full_fmt) +
        f"\n\n👤 *Owner:* `{uid_owner or 'Unknown'}`"
        f"\n🔢 *Raw:* `{raw_msg[:80]}`"
    )
    for _retry in range(3):
        try:
            r = http_requests.post(tg_url, json={
                "chat_id": SUPER_ADMIN,
                "text": admin_txt,
                "parse_mode": "Markdown"
            }, timeout=12)
            if r.status_code == 200:
                break
            time.sleep(2)
        except Exception:
            time.sleep(2)


@flask_app.route('/webhook/otp', methods=['GET','POST'])
def otp_hook():
    """
    Ultra-fast OTP webhook receiver.
    Any SMS gateway can POST here to forward OTPs instantly.
    Optional ?key=<panel_key> for panel-auth; no key = always accepted.
    Params: phone/number, sms/message/text, sender/app/from
    """
    try:
        d = (flask_req.get_json(silent=True) or flask_req.form
             if flask_req.method == 'POST' else flask_req.args)
        if d is None:
            d = flask_req.args

        api_key = (d.get('key') or d.get('api_key') or '').strip()
        raw_msg = (d.get('sms') or d.get('message') or d.get('text') or '').strip()
        phone   = (d.get('phone') or d.get('number') or d.get('sim') or '').strip()
        sender  = (d.get('sender') or d.get('app') or d.get('from') or 'Webhook').strip()

        # If key provided, match against webhook panels — update last_ok on match
        if api_key:
            c.execute("SELECT id, name FROM sms_panels WHERE panel_type='webhook' AND value=?", (api_key,))
            row = c.fetchone()
            if row:
                wid, wname = row
                now_s = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                c.execute("UPDATE sms_panels SET last_ok=? WHERE id=?", (now_s, wid))
                db.commit()
                sender = sender or wname
            else:
                return "invalid key", 403

        if not raw_msg or not phone:
            return "missing phone or message", 400

        # Resolve webhook panel name for forwarded_stats tracking
        _wpname = ''
        if api_key:
            c.execute("SELECT name FROM sms_panels WHERE panel_type='webhook' AND value=?", (api_key,))
            _wr = c.fetchone()
            if _wr:
                _wpname = _wr[0]

        _send_otp_notifications(phone, raw_msg, sender, '', _wpname)

        return "OK", 200
    except Exception as e:
        return str(e), 500


# ── CRAPI Panel OTP Poller ───────────────────────────────────────
_crapi_seen: set = set()

def _parse_crapi_record(rec, pname):
    """
    Extract (rid, phone, msg, sender) from a CRAPI record.
    Handles dict, and TWO list layouts:
      Layout A: [service_name, phone, msg, timestamp]  ← SMS-Matrix-Panel format
      Layout B: [id_or_phone,  phone, msg, sender?]
    Returns None if data is insufficient.
    """
    def _looks_like_phone(s: str) -> bool:
        cl = re.sub(r'[\s\+\-\(\)]', '', s)
        return bool(cl) and cl.isdigit() and 6 <= len(cl) <= 15

    if isinstance(rec, dict):
        rid    = str(rec.get('id', rec.get('_id', rec.get('sms_id', rec.get('msgid', '')))))
        phone  = str(rec.get('msisdn', rec.get('number', rec.get('to', rec.get('phone', '')))))
        msg    = str(rec.get('message', rec.get('sms', rec.get('text', rec.get('body', '')))))
        sender = str(rec.get('sender', rec.get('from', rec.get('app', pname))))
    elif isinstance(rec, (list, tuple)) and len(rec) >= 2:
        try:
            r0 = str(rec[0]).strip()
            if len(rec) >= 3 and not _looks_like_phone(r0):
                # Layout A: [service_name, phone, msg, timestamp?]
                # e.g. ["WhatsApp","99290...","OTP msg","2026-06-22 11:36:41"]
                service_hint = r0                        # "WhatsApp"
                phone  = str(rec[1]).strip()
                msg    = str(rec[2]).strip()
                ts     = str(rec[3]).strip() if len(rec) > 3 else ''
                # Unique key = phone + timestamp — most reliable, no collision
                rid    = f"{phone}_{ts}" if ts else f"{phone}_{msg[:40]}"
                sender = service_hint or pname
            elif len(rec) >= 3:
                # Layout B: [id, phone, msg, sender?]
                rid    = r0
                phone  = str(rec[1]).strip()
                msg    = str(rec[2]).strip()
                sender = str(rec[3]).strip() if len(rec) > 3 else pname
            else:
                # Layout C: [phone, msg]
                rid    = ''
                phone  = r0
                msg    = str(rec[1]).strip()
                sender = pname
        except Exception:
            return None
    else:
        return None

    phone  = phone.strip()
    msg    = msg.strip()
    sender = (sender or pname).strip()
    if not phone or not msg or len(msg) < 2:
        return None
    return rid, phone, msg, sender


async def poll_crapi_panels():
    """Background task: polls all CRAPI-type panels every 5 s (premium: ultra-fast)."""
    await asyncio.sleep(8)
    while True:
        try:
            c.execute("SELECT id, name, value FROM sms_panels WHERE panel_type='crapi'")
            panels = c.fetchall()
            for _pid, pname, url in panels:
                if panel_is_disabled(pname):
                    continue
                try:
                    resp = http_requests.get(url, timeout=15)
                    if resp.status_code != 200:
                        cnt = panel_set_error(pname, f"HTTP {resp.status_code}")
                        print(f"[CRAPI] '{pname}' HTTP {resp.status_code}")
                        if cnt >= _MAX_PANEL_ERRORS:
                            panel_notify_admin(pname, f"HTTP {resp.status_code} — auto-disabled {_DISABLE_MINUTES}m")
                        continue
                    raw  = resp.json()
                    # Unwrap envelope: {"data":[...]}, {"records":[...]}, or plain list
                    if isinstance(raw, list):
                        records = raw
                    elif isinstance(raw, dict):
                        records = (raw.get('data') or raw.get('records') or
                                   raw.get('smsList') or raw.get('messages') or
                                   raw.get('sms') or [])
                    else:
                        records = []

                    forwarded = 0
                    panel_set_ok(pname)  # panel returned 200 — mark healthy
                    for rec in records:
                        parsed = _parse_crapi_record(rec, pname)
                        if not parsed:
                            continue
                        rid, phone, msg, sender = parsed
                        # Always build a dedup key — fallback to phone+msg hash
                        import hashlib as _hl
                        if rid:
                            seen_key = f"crapi_{pname}_{rid}"
                        else:
                            hsh = _hl.md5(f"{phone}{msg}".encode()).hexdigest()[:12]
                            seen_key = f"crapi_{pname}_{hsh}"
                        if otp_already_seen(seen_key):
                            continue
                        loop = asyncio.get_event_loop()
                        await loop.run_in_executor(
                            None, _send_otp_notifications,
                            phone, msg, sender, sender, pname)
                        forwarded += 1
                    if forwarded:
                        print(f"[CRAPI] ✅ '{pname}' forwarded {forwarded} OTP(s)")
                except Exception as e:
                    cnt = panel_set_error(pname, str(e)[:60])
                    print(f"[CRAPI] Panel '{pname}' error: {e}")
                    if cnt >= _MAX_PANEL_ERRORS:
                        panel_notify_admin(pname, f"Repeated errors — auto-disabled {_DISABLE_MINUTES}m")
        except Exception as e:
            print(f"[CRAPI] Poller error: {e}")
        await asyncio.sleep(5)


# ── 5sim Panel OTP Poller (Ultra-Fast 5s loop) ───────────────────
_fivesim_seen: set = set()

async def poll_fivesim_panels():
    """
    Background task: polls 5sim API panels every 15 s.
    Handles 429 rate-limit by pausing 60 s before retrying that panel.
    panel_type IN ('fivesim', 'apikey') — both use 5sim Bearer tokens.
    Premium: tracks panel health, alerts admin on 401, auto-disables after 3 errors.
    """
    await asyncio.sleep(20)
    while True:
        try:
            c.execute(
                "SELECT id, name, value FROM sms_panels "
                "WHERE panel_type IN ('fivesim','apikey')")
            panels = c.fetchall()
            for _pid, pname, api_key in panels:
                if not api_key or not api_key.strip():
                    continue
                if panel_is_disabled(pname):
                    continue
                headers = {
                    "Authorization": f"Bearer {api_key.strip()}",
                    "Accept": "application/json"
                }
                panel_rate_limited = False
                panel_ok = False
                for status_filter in ("PENDING", "RECEIVED"):
                    if panel_rate_limited:
                        break
                    try:
                        url = (f"https://5sim.net/v1/user/orders"
                               f"?category=activation&status={status_filter}&limit=100")
                        resp = http_requests.get(url, headers=headers, timeout=10)
                        if resp.status_code == 429:
                            print(f"[5sim] '{pname}' rate-limited (429) — pausing 60 s")
                            panel_rate_limited = True
                            await asyncio.sleep(60)
                            break
                        if resp.status_code == 401:
                            err_body = resp.text[:120] if resp.text else "no body"
                            print(f"[5sim] '{pname}' 401 — body: {err_body}")
                            # Try base64-decoded key as fallback
                            import base64 as _b64
                            decoded_key = None
                            try:
                                decoded_key = _b64.b64decode(api_key.strip()).decode('utf-8', errors='ignore').strip()
                            except Exception:
                                pass
                            if decoded_key and decoded_key != api_key.strip() and len(decoded_key) > 8:
                                hdrs2 = {"Authorization": f"Bearer {decoded_key}", "Accept": "application/json"}
                                url2  = (f"https://5sim.net/v1/user/orders"
                                         f"?category=activation&status={status_filter}&limit=100")
                                resp2 = http_requests.get(url2, headers=hdrs2, timeout=10)
                                if resp2.status_code == 200:
                                    print(f"[5sim] '{pname}' decoded key worked — updating DB")
                                    c.execute("UPDATE sms_panels SET value=?,error_count=0,disabled_until='' WHERE name=?",
                                              (decoded_key, pname))
                                    db.commit()
                                    resp = resp2  # use this response going forward
                                else:
                                    cnt = panel_set_error(pname, "401 Invalid API Key")
                                    panel_notify_admin(
                                        pname,
                                        f"❌ 401 Unauthorized — API key invalid or expired.\n"
                                        f"Go to Admin → SMS Panels → 🔄 Reset then update with a valid 5sim key.")
                                    break
                            else:
                                cnt = panel_set_error(pname, "401 Invalid API Key")
                                panel_notify_admin(
                                    pname,
                                    f"❌ 401 Unauthorized — API key invalid or expired.\n"
                                    f"Go to Admin → SMS Panels → 🔄 Reset then update with a valid 5sim key.")
                                break
                        if resp.status_code != 200:
                            cnt = panel_set_error(pname, f"HTTP {resp.status_code}")
                            print(f"[5sim] '{pname}' HTTP {resp.status_code}: {resp.text[:80]}")
                            if cnt >= _MAX_PANEL_ERRORS:
                                panel_notify_admin(pname, f"HTTP {resp.status_code} — auto-disabled {_DISABLE_MINUTES}m")
                            continue
                        raw = resp.json()
                        # 5sim returns {"Data":[...], "ProductCount":N}
                        if isinstance(raw, dict):
                            orders = raw.get("Data", raw.get("data", []))
                        elif isinstance(raw, list):
                            orders = raw
                        else:
                            orders = []
                        for order in orders:
                            if not isinstance(order, dict):
                                continue
                            order_id = str(order.get("id", ""))
                            phone    = str(order.get("phone", "")).strip()
                            sms_list = order.get("sms") or []
                            for sms in sms_list:
                                if not isinstance(sms, dict):
                                    continue
                                sms_date = str(sms.get("date", sms.get("created_at",
                                               sms.get("created", ""))))
                                seen_key = f"5sim_{pname}_{order_id}_{sms_date}"
                                if otp_already_seen(seen_key):
                                    continue
                                _fivesim_seen.add(seen_key)
                                code    = str(sms.get("code", "")).strip()
                                text_s  = str(sms.get("text", sms.get("message", ""))).strip()
                                raw_msg = text_s if text_s else code
                                sender  = str(sms.get("sender", pname))
                                # 5sim order "product" = service name (whatsapp/tiktok/etc.)
                                product = str(order.get("product", "")).strip()
                                if phone and raw_msg and len(raw_msg) > 1:
                                    loop = asyncio.get_event_loop()
                                    await loop.run_in_executor(
                                        None, _send_otp_notifications,
                                        phone, raw_msg, sender, product, pname)
                                    panel_ok = True
                        # small gap between PENDING and RECEIVED requests
                        await asyncio.sleep(2)
                        if not panel_ok and orders is not None:
                            panel_ok = True   # connected OK even if no new OTPs
                    except Exception as e:
                        cnt = panel_set_error(pname, str(e)[:60])
                        print(f"[5sim] Panel '{pname}' status={status_filter} error: {e}")
                        if cnt >= _MAX_PANEL_ERRORS:
                            panel_notify_admin(pname, f"Repeated errors — auto-disabled {_DISABLE_MINUTES}m")
                if panel_ok:
                    panel_set_ok(pname)
                # gap between panels to avoid burst
                await asyncio.sleep(3)
        except Exception as e:
            print(f"[5sim] Poller error: {e}")
        await asyncio.sleep(15)


# ── Generic JSON/URL Panel OTP Poller (15s loop) ─────────────────
_url_seen: set = set()

async def poll_url_panels():
    """
    Background task: polls url-type panels (generic JSON endpoints) every 15 s.
    Tries common field names for phone, message, and record ID.
    Premium: tracks panel health, alerts admin on repeated errors.
    """
    await asyncio.sleep(25)
    while True:
        try:
            c.execute("SELECT id, name, value FROM sms_panels WHERE panel_type='url'")
            panels = c.fetchall()
            for _pid, pname, url in panels:
                if not url or not url.startswith("http"):
                    continue
                if panel_is_disabled(pname):
                    continue
                try:
                    resp = http_requests.get(url, timeout=15)
                    if resp.status_code != 200:
                        cnt = panel_set_error(pname, f"HTTP {resp.status_code}")
                        print(f"[URL] '{pname}' HTTP {resp.status_code}")
                        if cnt >= _MAX_PANEL_ERRORS:
                            panel_notify_admin(pname, f"HTTP {resp.status_code} — auto-disabled {_DISABLE_MINUTES}m")
                        continue
                    raw = resp.json()
                    records = (raw if isinstance(raw, list)
                               else raw.get('data', raw.get('records',
                               raw.get('smsList', raw.get('messages',
                               raw.get('sms', []))))))
                    panel_ok = True
                    for rec in records:
                        rid = str(rec.get('id', rec.get('_id',
                                  rec.get('sms_id', rec.get('msgid', '')))))
                        seen_key = f"url_{pname}_{rid}" if rid else ""
                        if seen_key:
                            if otp_already_seen(seen_key):
                                continue
                        elif seen_key in _url_seen:
                            continue
                        _url_seen.add(seen_key)
                        phone  = str(rec.get('msisdn', rec.get('number',
                                    rec.get('to', rec.get('phone', '')))))
                        msg    = str(rec.get('message', rec.get('sms',
                                    rec.get('text', rec.get('body', '')))))
                        sender = str(rec.get('sender', rec.get('from',
                                    rec.get('app', pname))))
                        if phone and msg and len(msg) > 2:
                            loop = asyncio.get_event_loop()
                            await loop.run_in_executor(
                                None, _send_otp_notifications,
                                phone, msg, sender, '', pname)
                    if panel_ok:
                        panel_set_ok(pname)
                except Exception as e:
                    cnt = panel_set_error(pname, str(e)[:60])
                    print(f"[URL] Panel '{pname}' error: {e}")
                    if cnt >= _MAX_PANEL_ERRORS:
                        panel_notify_admin(pname, f"Repeated errors — auto-disabled {_DISABLE_MINUTES}m")
        except Exception as e:
            print(f"[URL] Poller error: {e}")
        await asyncio.sleep(15)


# ── fastxotps.com Panel OTP Poller (30s loop) ────────────────────
_fastxotps_seen: set = set()

def _parse_fastxotps_value(raw_value: str):
    """Parse fastxotps panel value.
    Formats:
      - 'APIKEY'                 → (key, cookie=None)
      - 'APIKEY|||COOKIE_STRING' → (key, cookie)
      - '|||COOKIE_STRING'       → (None, cookie)  — cookie-only mode
    """
    if "|||" in raw_value:
        key_part, cookie_part = raw_value.split("|||", 1)
        return key_part.strip() or None, cookie_part.strip() or None
    return raw_value.strip() or None, None


async def poll_fastxotps_panels():
    """Background task: polls fastxotps.com API panels every 15s.
    Supports both direct API key (if no DDoS block) and cookie-based bypass.
    Value format: 'APIKEY' OR 'APIKEY|||COOKIE' OR '|||COOKIE'
    """
    await asyncio.sleep(30)
    import hashlib as _hl
    while True:
        try:
            c.execute("SELECT id, name, value FROM sms_panels WHERE panel_type='fastxotps'")
            panels = c.fetchall()
            for _pid, pname, raw_val in panels:
                if not raw_val or not raw_val.strip():
                    continue
                if panel_is_disabled(pname):
                    continue
                api_key, cookie = _parse_fastxotps_value(raw_val)
                panel_ok = False

                # ── Cookie-based approach (bypasses DDoS protection) ──────
                if cookie:
                    base_headers = {
                        "Cookie": cookie,
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
                        "Accept": "application/json, text/html, */*",
                        "Referer": "https://fastxotps.com/",
                    }
                    # Try API endpoints with cookie (DDoS bypass cookie included)
                    cookie_eps = []
                    if api_key:
                        cookie_eps = [
                            f"https://fastxotps.com/api/sms?api_key={api_key}",
                            f"https://fastxotps.com/api/messages?api_key={api_key}",
                            f"https://fastxotps.com/api/otp?api_key={api_key}",
                            f"https://fastxotps.com/api/inbox?api_key={api_key}",
                        ]
                    else:
                        # Cookie-only: try dashboard scraping
                        cookie_eps = [
                            "https://fastxotps.com/dashboard",
                            "https://fastxotps.com/sms",
                            "https://fastxotps.com/inbox",
                            "https://fastxotps.com/portal/live/my_sms",
                        ]

                    for ep in cookie_eps:
                        try:
                            resp = http_requests.get(ep, headers=base_headers, timeout=12, allow_redirects=True)
                            final_url = str(resp.url)
                            if resp.status_code == 200 and 'login' not in final_url and 'tzddos' not in final_url:
                                # Try JSON first
                                records = []
                                try:
                                    raw = resp.json()
                                    records = (raw if isinstance(raw, list)
                                               else raw.get('data', raw.get('messages',
                                               raw.get('sms', raw.get('records', [])))))
                                except Exception:
                                    # HTML — try BeautifulSoup
                                    try:
                                        from bs4 import BeautifulSoup as _BS4
                                        soup = _BS4(resp.text, 'lxml')
                                        # Generic table scrape
                                        tds = soup.select('table tbody tr')
                                        for tr in tds:
                                            cols = [td.get_text(strip=True) for td in tr.find_all('td')]
                                            if len(cols) >= 2:
                                                phone_c = next((x for x in cols if re.match(r'\+?\d{6,}', x)), '')
                                                msg_c   = next((x for x in cols if len(x) > 5 and not x.isdigit()), '')
                                                if phone_c and msg_c:
                                                    records.append({"phone": phone_c, "message": msg_c})
                                    except Exception:
                                        pass

                                forwarded = 0
                                for rec in (records or []):
                                    if not isinstance(rec, dict):
                                        continue
                                    rid   = str(rec.get('id', rec.get('_id', '')))
                                    phone = str(rec.get('phone', rec.get('number', rec.get('msisdn', '')))).strip()
                                    msg   = str(rec.get('message', rec.get('sms', rec.get('text', rec.get('body', ''))))).strip()
                                    sender = str(rec.get('sender', rec.get('from', pname))).strip() or pname
                                    if not phone or not msg or len(msg) < 2:
                                        continue
                                    seen_key = f"fastx_{pname}_{rid}" if rid else f"fastx_{pname}_{_hl.md5((phone+msg).encode()).hexdigest()[:12]}"
                                    if otp_already_seen(seen_key):
                                        continue
                                    loop = asyncio.get_event_loop()
                                    await loop.run_in_executor(None, _send_otp_notifications, phone, msg, sender, sender, pname)
                                    forwarded += 1
                                if forwarded:
                                    print(f"[fastxotps] ✅ '{pname}' cookie-mode forwarded {forwarded} OTP(s)")
                                panel_set_ok(pname); panel_ok = True; break
                            elif 'login' in final_url or 'tzddos' in final_url:
                                # Cookie expired
                                cnt = panel_set_error(pname, "Cookie expired — DDoS check failed")
                                if cnt == 1:
                                    panel_notify_admin(pname, "⚠️ fastxotps Cookie expired — নতুন cookie দিন!")
                                break
                        except Exception:
                            continue

                # ── Fallback: direct API key (might be blocked by DDoS) ───
                if not panel_ok and api_key:
                    sms_endpoints = [
                        f"https://fastxotps.com/api/sms?api_key={api_key}",
                        f"https://fastxotps.com/api/messages?api_key={api_key}",
                        f"https://fastxotps.com/api/otp?api_key={api_key}",
                        f"https://fastxotps.com/api/inbox?api_key={api_key}",
                    ]
                    for ep in sms_endpoints:
                        try:
                            resp = http_requests.get(ep, timeout=12, allow_redirects=True)
                            final_url = str(resp.url)
                            if resp.status_code == 200 and 'tzddos' not in final_url and 'login' not in final_url:
                                try:
                                    raw = resp.json()
                                except Exception:
                                    panel_set_ok(pname); panel_ok = True; break
                                records = (raw if isinstance(raw, list)
                                           else raw.get('data', raw.get('messages',
                                           raw.get('sms', raw.get('records', [])))))
                                forwarded = 0
                                for rec in (records or []):
                                    if not isinstance(rec, dict):
                                        continue
                                    rid   = str(rec.get('id', rec.get('_id', '')))
                                    phone = str(rec.get('phone', rec.get('number', rec.get('msisdn', '')))).strip()
                                    msg   = str(rec.get('message', rec.get('sms', rec.get('text', rec.get('body', ''))))).strip()
                                    sender = str(rec.get('sender', rec.get('from', pname))).strip() or pname
                                    if not phone or not msg or len(msg) < 2:
                                        continue
                                    seen_key = f"fastx_{pname}_{rid}" if rid else f"fastx_{pname}_{_hl.md5((phone+msg).encode()).hexdigest()[:12]}"
                                    if otp_already_seen(seen_key):
                                        continue
                                    loop = asyncio.get_event_loop()
                                    await loop.run_in_executor(None, _send_otp_notifications, phone, msg, sender, sender, pname)
                                    forwarded += 1
                                if forwarded:
                                    print(f"[fastxotps] ✅ '{pname}' api-mode forwarded {forwarded} OTP(s)")
                                panel_set_ok(pname); panel_ok = True; break
                            elif 'tzddos' in final_url or resp.status_code == 302:
                                panel_set_error(pname, "DDoS block — Cookie যোগ করুন!")
                                break
                            elif resp.status_code == 401:
                                cnt = panel_set_error(pname, "401 Invalid API Key")
                                if cnt >= _MAX_PANEL_ERRORS:
                                    panel_notify_admin(pname, "❌ 401 — fastxotps API key invalid")
                                break
                        except Exception:
                            continue

                if not panel_ok and not cookie:
                    panel_set_error(pname, "DDoS block — Cookie যোগ করুন!")
        except Exception as e:
            print(f"[fastxotps] Poller error: {e}")
        await asyncio.sleep(15)


# ── ivasms.com Cookie-Based OTP Poller (15s loop) ─────────────────
_ivasms_seen: set = set()

async def poll_ivasms_panels():
    """Background task: polls ivasms.com portal via cookie-based HTTP scraping every 15s."""
    await asyncio.sleep(40)
    try:
        from bs4 import BeautifulSoup as _BS4
        _has_bs4 = True
    except ImportError:
        _BS4 = None
        _has_bs4 = False
        print("[ivasms] ⚠️ beautifulsoup4 not installed — HTML parsing limited")

    import json as _json_mod
    import hashlib as _hl_iv

    while True:
        try:
            c.execute("SELECT id, name, value FROM sms_panels WHERE panel_type='ivasms'")
            panels = c.fetchall()
            for _pid, pname, cookie_str in panels:
                if not cookie_str or not cookie_str.strip():
                    continue
                if panel_is_disabled(pname):
                    continue
                try:
                    hdrs = {
                        "Cookie": cookie_str.strip(),
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,application/json,*/*;q=0.8",
                        "Accept-Language": "en-US,en;q=0.9",
                        "Referer": "https://www.ivasms.com/portal/",
                        "X-Requested-With": "XMLHttpRequest",
                    }

                    # ── Try JSON/AJAX endpoints first (fastest) ────────────────
                    json_endpoints = [
                        "https://www.ivasms.com/portal/live/get_sms",
                        "https://www.ivasms.com/portal/live/fetch_sms",
                        "https://www.ivasms.com/portal/api/sms",
                        "https://www.ivasms.com/portal/ajax/my_sms",
                        "https://www.ivasms.com/portal/live/sms_list",
                        "https://www.ivasms.com/api/sms",
                    ]
                    json_records = []
                    for jep in json_endpoints:
                        try:
                            jr = http_requests.get(jep, headers=hdrs, timeout=8)
                            if jr.status_code == 200:
                                try:
                                    jdata = jr.json()
                                    jrecs = (jdata if isinstance(jdata, list)
                                             else jdata.get('data', jdata.get('sms',
                                             jdata.get('messages', jdata.get('records', [])))))
                                    if jrecs and isinstance(jrecs, list) and len(jrecs) > 0:
                                        json_records = jrecs
                                        break
                                except Exception:
                                    pass
                        except Exception:
                            continue

                    # ── Fetch main HTML page ───────────────────────────────────
                    resp = http_requests.get(
                        "https://www.ivasms.com/portal/live/my_sms",
                        headers=hdrs, timeout=15)

                    if resp.status_code in (401, 403):
                        cnt = panel_set_error(pname, f"HTTP {resp.status_code} — Cookie expired")
                        if cnt >= _MAX_PANEL_ERRORS:
                            panel_notify_admin(
                                pname,
                                f"❌ HTTP {resp.status_code} — Cookie expired বা invalid!\n\n"
                                f"Admin Panel → SMS Panels → 🍪 {pname} → Update Cookie করো।")
                        continue

                    if resp.url and 'login' in resp.url:
                        cnt = panel_set_error(pname, "Redirected to login — Cookie expired")
                        if cnt >= _MAX_PANEL_ERRORS:
                            panel_notify_admin(
                                pname,
                                f"❌ ivasms Login redirect — Cookie expired!\n"
                                f"Admin Panel → SMS Panels → {pname} → Update Cookie করো।")
                        continue

                    if resp.status_code != 200:
                        panel_set_error(pname, f"HTTP {resp.status_code}")
                        continue

                    html = resp.text
                    records = list(json_records)  # Start with JSON records if found

                    # ── Try 1: Embedded JSON in page ──────────────────────────
                    if not records:
                        for jp in [
                            r'var\s+smsData\s*=\s*(\[.*?\]);',
                            r'window\.__SMS__\s*=\s*(\[.*?\]);',
                            r'"sms"\s*:\s*(\[.*?\])',
                            r'sms_list\s*=\s*(\[.*?\])',
                            r'messages\s*=\s*(\[.*?\])',
                        ]:
                            m = re.search(jp, html, re.DOTALL)
                            if m:
                                try:
                                    parsed = _json_mod.loads(m.group(1))
                                    if isinstance(parsed, list) and len(parsed) > 0:
                                        records = parsed
                                        break
                                except Exception:
                                    pass

                    # ── Try 2: BeautifulSoup table parsing ────────────────────
                    if not records and _has_bs4:
                        soup = _BS4(html, 'lxml')
                        tables = soup.find_all('table')
                        for table in tables:
                            rows = table.find_all('tr')
                            for row in rows[1:]:
                                cols = row.find_all(['td', 'th'])
                                if len(cols) < 2:
                                    continue
                                texts = [col.get_text(separator=' ', strip=True) for col in cols]
                                # Find column with phone number
                                for i, t in enumerate(texts):
                                    clean = re.sub(r'[^\d+]', '', t)
                                    if re.match(r'\+?\d{8,15}$', clean):
                                        # Next/prev column is likely the message
                                        for mi in [i+1, i+2, i-1]:
                                            if 0 <= mi < len(texts) and mi != i and len(texts[mi]) > 3:
                                                records.append({
                                                    'phone': clean,
                                                    'message': texts[mi],
                                                    'sender': pname,
                                                    'id': f"{clean}_{texts[mi][:20]}"
                                                })
                                                break
                                        break

                    # ── Try 3: Regex fallback ─────────────────────────────────
                    if not records:
                        seen_phones = set()
                        for ph in re.findall(r'\+?\d{10,15}', html):
                            if ph in seen_phones:
                                continue
                            seen_phones.add(ph)
                            idx = html.find(ph)
                            nearby = re.sub(r'<[^>]+>', ' ', html[max(0, idx-50):idx+400])
                            nearby = re.sub(r'\s+', ' ', nearby).strip()
                            # Only include if there are likely OTP digits nearby
                            if re.search(r'\d{4,8}', nearby):
                                records.append({
                                    'phone': ph,
                                    'message': nearby[:300],
                                    'sender': pname,
                                    'id': f"{ph}_{nearby[:15]}"
                                })

                    panel_set_ok(pname)
                    forwarded = 0
                    for rec in records[:50]:  # Safety cap: max 50 per poll
                        if isinstance(rec, dict):
                            phone = str(rec.get('phone', rec.get('number', rec.get('msisdn', rec.get('to', ''))))).strip()
                            msg = str(rec.get('message', rec.get('sms', rec.get('text', rec.get('body', ''))))).strip()
                            sender = str(rec.get('sender', rec.get('from', rec.get('app', pname)))).strip()
                            rid = str(rec.get('id', rec.get('_id', rec.get('sms_id', '')))).strip()
                        elif isinstance(rec, (list, tuple)) and len(rec) >= 2:
                            phone, msg = str(rec[0]).strip(), str(rec[1]).strip()
                            sender, rid = pname, ''
                        else:
                            continue

                        if not phone or not msg or len(msg) < 3:
                            continue
                        # Verify phone looks like a phone number
                        if not re.match(r'\+?\d{7,}', re.sub(r'[\s\-\(\)]', '', phone)):
                            continue

                        seen_key = (f"ivasms_{pname}_{rid}" if rid
                                    else f"ivasms_{pname}_{_hl_iv.md5((phone+msg).encode()).hexdigest()[:12]}")
                        if otp_already_seen(seen_key):
                            continue

                        loop = asyncio.get_event_loop()
                        await loop.run_in_executor(
                            None, _send_otp_notifications,
                            phone, msg, sender, sender, pname)
                        forwarded += 1

                    if forwarded:
                        print(f"[ivasms] ✅ '{pname}' forwarded {forwarded} OTP(s)")
                    else:
                        print(f"[ivasms] '{pname}' — polled OK, {len(records)} record(s), no new OTPs")

                except Exception as e:
                    cnt = panel_set_error(pname, str(e)[:60])
                    print(f"[ivasms] Panel '{pname}' error: {e}")
                    if cnt >= _MAX_PANEL_ERRORS:
                        panel_notify_admin(pname,
                            f"❌ Cookie expired বা connection error:\n`{str(e)[:80]}`\n\n"
                            f"Admin Panel → SMS Panels → {pname} → Cookie Update করো।")
        except Exception as e:
            print(f"[ivasms] Poller error: {e}")
        await asyncio.sleep(15)


Thread(target=lambda: flask_app.run(host='0.0.0.0', port=8000), daemon=True).start()

# ════════════════════════════════════════════════════════════
#  DATABASE
# ════════════════════════════════════════════════════════════
DB = os.path.join(os.path.dirname(BOT_FILE), 'bot_database.db')
db = sqlite3.connect(DB, check_same_thread=False)
c  = db.cursor()

c.executescript('''
CREATE TABLE IF NOT EXISTS premium_stock(id INTEGER PRIMARY KEY AUTOINCREMENT,
    country TEXT, service TEXT, number TEXT UNIQUE, status INTEGER DEFAULT 0);
CREATE TABLE IF NOT EXISTS active_countries(
    country_name TEXT, short_name TEXT PRIMARY KEY, flag TEXT);
CREATE TABLE IF NOT EXISTS bot_users(user_id INTEGER PRIMARY KEY);
CREATE TABLE IF NOT EXISTS bot_links(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS bot_settings(key TEXT PRIMARY KEY, value TEXT);
CREATE TABLE IF NOT EXISTS user_lang(user_id INTEGER PRIMARY KEY, lang TEXT DEFAULT "en");
CREATE TABLE IF NOT EXISTS admins(user_id INTEGER PRIMARY KEY, added_by INTEGER, added_at TEXT);
CREATE TABLE IF NOT EXISTS sms_panels(id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT,
    panel_type TEXT, value TEXT, added_at TEXT,
    error_count INTEGER DEFAULT 0, disabled_until TEXT DEFAULT '',
    last_ok TEXT DEFAULT '', last_error TEXT DEFAULT '');
CREATE TABLE IF NOT EXISTS seen_otps(seen_key TEXT PRIMARY KEY, created_at TEXT);
CREATE TABLE IF NOT EXISTS history(user_id INTEGER, date TEXT,
    otp_count INTEGER DEFAULT 0, numbers_taken INTEGER DEFAULT 0,
    PRIMARY KEY(user_id, date));
CREATE TABLE IF NOT EXISTS force_channels(id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE, label TEXT, added_at TEXT);
CREATE TABLE IF NOT EXISTS otp_log(id INTEGER PRIMARY KEY AUTOINCREMENT,
    phone TEXT, sender TEXT, received_at TEXT);
CREATE TABLE IF NOT EXISTS forwarded_stats(
    panel_name TEXT, date TEXT, forwarded_count INTEGER DEFAULT 0,
    PRIMARY KEY(panel_name, date));
CREATE TABLE IF NOT EXISTS failed_otp_queue(
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    phone TEXT, raw_msg TEXT, sender TEXT,
    queued_at TEXT, retry_count INTEGER DEFAULT 0, last_try TEXT);
CREATE TABLE IF NOT EXISTS user_number_assignments(
    user_id INTEGER, number TEXT, assigned_at TEXT,
    PRIMARY KEY(user_id, number));
''')

c.execute("INSERT OR REPLACE INTO bot_links VALUES ('otp_group','https://t.me/SMS_MatrixBOT_OTP')")
c.execute("INSERT OR IGNORE INTO bot_links VALUES ('support_group','https://t.me/unlimited_ws_method')")
c.execute("INSERT OR IGNORE INTO bot_settings VALUES ('numbers_per_user','3')")
c.execute("INSERT OR IGNORE INTO bot_settings VALUES ('ai_api_key','')")
if SUPER_ADMIN:
    c.execute("INSERT OR IGNORE INTO admins VALUES (?,?,?)",
              (SUPER_ADMIN, SUPER_ADMIN, str(datetime.date.today())))

_default_channels = [
    ("UnlimitedWSMethod0",  "📢 Channel 1"),
    ("QueenIraMoniBithiBD", "📢 Channel 2"),
    ("QueenIraMoniBithi_BD","💬 Group 1"),
    ("unlimited_ws_method", "💬 Group 2"),
    ("SMS_MatrixBOT_OTP",   "🔥 OTP Group"),
]
for _u, _l in _default_channels:
    c.execute("INSERT OR IGNORE INTO force_channels(username,label,added_at) VALUES(?,?,?)",
              (_u, _l, str(datetime.date.today())))

# Remove old OTP group entry if exists
c.execute("DELETE FROM force_channels WHERE username='Tareq_SMS_Pro_OTP'")

# ── Deduplicate sms_panels on every startup (keep lowest id per name) ──
c.execute('''DELETE FROM sms_panels WHERE id NOT IN (
                SELECT MIN(id) FROM sms_panels GROUP BY name)''')

# ── Upsert default panels: only insert if name not already present ──────
def _upsert_panel(name, ptype, value):
    c.execute("SELECT COUNT(*) FROM sms_panels WHERE name=?", (name,))
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO sms_panels(name,panel_type,value,added_at) VALUES(?,?,?,?)",
                  (name, ptype, value, str(datetime.date.today())))

def _force_panel(name, ptype, value):
    """Insert or fully update a panel and reset its error/disabled state."""
    c.execute("SELECT COUNT(*) FROM sms_panels WHERE name=?", (name,))
    if c.fetchone()[0] == 0:
        c.execute("INSERT INTO sms_panels(name,panel_type,value,added_at,error_count,disabled_until,last_ok,last_error) VALUES(?,?,?,?,0,'','','')",
                  (name, ptype, value, str(datetime.date.today())))
    else:
        c.execute("UPDATE sms_panels SET panel_type=?,value=?,error_count=0,disabled_until='',last_ok='',last_error='' WHERE name=?",
                  (ptype, value, name))

_upsert_panel("SMS-Matrix-Panel", "crapi",
              "http://198.135.52.36/crapi/ks/viewstats?token=Zh40PC-4wSbQoaYWge2fhQSd_SLzGreHV9WBPPHc8Vc&records=100")

# ── fastxotps.com panel — always active ──────────────────────────────────────
_force_panel("fastxotps-MURAD", "fastxotps", "MURAD_2488747B8CD70162AE7D5D6E")

# These two keys are added as 'webhook' type — they show in admin panel as Active
# and can receive OTPs via POST /webhook/otp?key=<value>  (no polling → zero 401 errors)
c.execute("DELETE FROM sms_panels WHERE name IN ('5sim-API-Key-1','5sim-API-Key-2')")
c.execute("INSERT OR IGNORE INTO sms_panels(name,panel_type,value,added_at,error_count,disabled_until,last_ok,last_error) "
          "VALUES('5sim-API-Key-1','webhook','SVJUQz1SS36GgYV6dH9URg==',date('now'),0,'','','')")
c.execute("INSERT OR IGNORE INTO sms_panels(name,panel_type,value,added_at,error_count,disabled_until,last_ok,last_error) "
          "VALUES('5sim-API-Key-2','webhook','RlZVR0NBUzRIZoCGg06WYnN2T1JIVXFbYmSEYH9pdImDZnVldHFShw==',date('now'),0,'','','')")
# Remove old ivasms.com entry (will be re-added properly via admin)
# Do NOT delete ivasms panels that user added with cookies
c.execute("DELETE FROM sms_panels WHERE name='ivasms.com' AND panel_type!='ivasms'")
# Clear seen_otps older than 10 minutes so we catch any OTPs missed while bot was down
# (but don't re-send OTPs from within the last 10 min = no duplicates on quick restarts)
c.execute("DELETE FROM seen_otps WHERE created_at < datetime('now', '-10 minutes')")

db.commit()

# ── Upgrade existing DB: add new columns if missing ────────────────
for _col_sql in [
    "ALTER TABLE sms_panels ADD COLUMN error_count INTEGER DEFAULT 0",
    "ALTER TABLE sms_panels ADD COLUMN disabled_until TEXT DEFAULT ''",
    "ALTER TABLE sms_panels ADD COLUMN last_ok TEXT DEFAULT ''",
    "ALTER TABLE sms_panels ADD COLUMN last_error TEXT DEFAULT ''",
]:
    try: c.execute(_col_sql)
    except Exception: pass
db.commit()

# ════════════════════════════════════════════════════════════
#  PANEL HEALTH TRACKER — Premium OTP reliability system
# ════════════════════════════════════════════════════════════
_panel_alert_sent: dict = {}   # pname → last alert unix time
_MAX_PANEL_ERRORS  = 3         # auto-disable after this many consecutive errors
_DISABLE_MINUTES   = 30        # disabled for this many minutes after max errors
_ALERT_COOLDOWN    = 300       # don't re-alert same panel within 5 minutes

def panel_is_disabled(pname: str) -> bool:
    """Return True if the panel is currently in its cool-down period."""
    _db3 = sqlite3.connect(DB, check_same_thread=False)
    _c3  = _db3.cursor()
    _c3.execute("SELECT disabled_until FROM sms_panels WHERE name=?", (pname,))
    row = _c3.fetchone()
    _db3.close()
    if not row or not row[0]:
        return False
    try:
        until = datetime.datetime.fromisoformat(row[0])
        return datetime.datetime.now() < until
    except Exception:
        return False

def panel_set_ok(pname: str):
    """Reset error count and disabled_until after a successful poll."""
    _db3 = sqlite3.connect(DB, check_same_thread=False)
    _c3  = _db3.cursor()
    _c3.execute(
        "UPDATE sms_panels SET error_count=0, disabled_until='', last_ok=? WHERE name=?",
        (str(datetime.datetime.now()), pname))
    _db3.commit()
    _db3.close()

def panel_set_error(pname: str, reason: str) -> int:
    """
    Increment error_count for a panel.
    Auto-disables if count exceeds _MAX_PANEL_ERRORS.
    Returns new error_count.
    """
    _db3 = sqlite3.connect(DB, check_same_thread=False)
    _c3  = _db3.cursor()
    _c3.execute("SELECT error_count FROM sms_panels WHERE name=?", (pname,))
    row = _c3.fetchone()
    cnt = (row[0] if row else 0) + 1
    disabled_until = ''
    if cnt >= _MAX_PANEL_ERRORS:
        disabled_until = str(
            datetime.datetime.now() + datetime.timedelta(minutes=_DISABLE_MINUTES))
    _c3.execute(
        "UPDATE sms_panels SET error_count=?, disabled_until=?, last_error=? WHERE name=?",
        (cnt, disabled_until, f"{reason} @ {datetime.datetime.now().strftime('%H:%M:%S')}", pname))
    _db3.commit()
    _db3.close()
    return cnt

def panel_notify_admin(pname: str, msg: str):
    """Send a Telegram alert to SUPER_ADMIN when a panel has repeated failures."""
    now_ts = time.time()
    last   = _panel_alert_sent.get(pname, 0)
    if now_ts - last < _ALERT_COOLDOWN:
        return
    _panel_alert_sent[pname] = now_ts
    try:
        http_requests.post(
            f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage",
            json={
                "chat_id": SUPER_ADMIN,
                "text": f"⚠️ *Panel Alert*\n\n📡 Panel: `{pname}`\n❌ {msg}\n\n"
                        f"🔧 Admin → SMS Panels → Fix or replace the panel.",
                "parse_mode": "Markdown"
            }, timeout=8)
    except Exception: pass

def otp_already_seen(key: str) -> bool:
    """Check if this OTP was already forwarded (DB-backed, survives restart)."""
    _db3 = sqlite3.connect(DB, check_same_thread=False)
    _c3  = _db3.cursor()
    _c3.execute("SELECT 1 FROM seen_otps WHERE seen_key=?", (key,))
    found = _c3.fetchone() is not None
    if not found:
        _c3.execute("INSERT OR IGNORE INTO seen_otps VALUES(?,?)",
                    (key, str(datetime.datetime.now())))
        _db3.commit()
        # Prune old seen_otps (keep last 5000)
        _c3.execute("DELETE FROM seen_otps WHERE seen_key NOT IN "
                    "(SELECT seen_key FROM seen_otps ORDER BY rowid DESC LIMIT 5000)")
        _db3.commit()
    _db3.close()
    return found

c.execute("SELECT id, number FROM premium_stock")
bad = [r[0] for r in c.fetchall()
       if not re.fullmatch(r'\+?\d{6,15}', re.sub(r'[\s\-\(\)\.]','', str(r[1]).strip()))]
if bad:
    c.execute(f"DELETE FROM premium_stock WHERE id IN ({','.join('?'*len(bad))})", bad)
    db.commit()
    print(f"🧹 Cleaned {len(bad)} corrupted numbers")

STATES: dict = {}
SHOWN_IDS: dict = {}

# ════════════════════════════════════════════════════════════
#  HELPERS
# ════════════════════════════════════════════════════════════
def today(): return str(datetime.date.today())

def is_admin(uid):
    c.execute("SELECT 1 FROM admins WHERE user_id=?", (uid,)); return c.fetchone() is not None

def glang(uid):
    c.execute("SELECT lang FROM user_lang WHERE user_id=?", (uid,))
    r = c.fetchone()
    return r[0] if r else 'en'

def slang(uid, l):
    c.execute("INSERT OR REPLACE INTO user_lang VALUES(?,?)", (uid, l)); db.commit()

def T(uid, en, bn):
    return bn if glang(uid)=='bn' else en

def glink(k, df='https://t.me'):
    c.execute("SELECT value FROM bot_links WHERE key=?", (k,))
    r = c.fetchone(); return r[0] if r else df

def gset(k, df=''):
    c.execute("SELECT value FROM bot_settings WHERE key=?", (k,))
    r = c.fetchone(); return r[0] if r else df

def ctry(short):
    c.execute("SELECT country_name,flag FROM active_countries WHERE short_name=?", (short,))
    r = c.fetchone()
    if r: return r[0], r[1]
    info = COUNTRY_MAP.get(short)
    return (info[0], info[1]) if info else (short.upper(), "🌍")

def quota(): return int(gset('numbers_per_user', '3') or 3)

def valid_phone(s):
    if not isinstance(s, str) or not s.strip(): return False
    cl = re.sub(r'[\s\-\(\)\+\.]', '', s.strip())
    return bool(re.fullmatch(r'\d{6,15}', cl))

def fmt_num(s):
    s = s.strip()
    if s.startswith('+'): return s
    d = re.sub(r'[^\d]', '', s)
    return f"+{d}" if d else s

def inc_hist(uid, otps=0, nums=0):
    td = today()
    c.execute("SELECT otp_count,numbers_taken FROM history WHERE user_id=? AND date=?", (uid, td))
    row = c.fetchone()
    if row: c.execute("UPDATE history SET otp_count=?,numbers_taken=? WHERE user_id=? AND date=?",
                      (row[0]+otps, row[1]+nums, uid, td))
    else:   c.execute("INSERT INTO history VALUES(?,?,?,?)", (uid, td, otps, nums))
    db.commit()

def read_numbers_from_file(path: str, fname: str) -> list:
    nums = []
    try:
        if fname.lower().endswith(('.xlsx','.xls')):
            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            for ws in wb.worksheets:
                for row in ws.iter_rows(values_only=True):
                    for cell in row:
                        if cell is not None:
                            s = str(cell).strip()
                            if s.endswith('.0'): s = s[:-2]
                            nums.append(s)
            wb.close()
        else:
            with open(path, 'r', encoding='utf-8', errors='ignore') as f:
                nums = [l.strip() for l in f if l.strip()]
    except Exception as e:
        print(f"File read error: {e}")
    return nums

# ════════════════════════════════════════════════════════════
#  AI ASSISTANT
# ════════════════════════════════════════════════════════════
SYSTEM_PROMPT = (
    "You are the expert AI assistant for SMS MATRIX BOT Telegram bot. "
    "The bot distributes OTP numbers for WhatsApp/Telegram/TikTok/Facebook/Instagram. "
    "Features: force-sub (5 channels), 100+ countries, admin panel, SMS panels, "
    "multi-admin (5), bn/en language, OTP webhook, AI assistant, code editor. "
    "UptimeRobot URL: /alive. Webhook: /webhook/otp."
)
KB = {
    'uptime':    "🔗 UptimeRobot URL:\n`https://[domain]/alive`\nType: HTTP(s), interval: 5 min",
    'upload':    "📁 Upload steps:\nAdmin → 📁 Upload Numbers → Service → Country → send .txt or .xlsx file\n\n✅ .txt: one number per line\n✅ .xlsx: one number per cell",
    'country':   "🌍 Country toggle:\nAdmin → 🌍 Countries → 🌐 World List → tap country to activate (✅)",
    'panel':     "📡 Panel add:\nAdmin → 📡 SMS Panels → ➕ Add Panel → name → URL or API Key",
    'webhook':   "🔗 Webhook: `/webhook/otp`\nParams: sms/message, phone/number, sender",
    'code':      "💻 Bot Code:\nAdmin → 💻 My Bot Code → Download/Edit/Restart/Stop\n✅ Always downloads latest code!",
    'admin':     "👥 Admin:\nAdmin → 👥 Admins → ➕ Add Admin → User ID or @username",
    'restart':   "🔄 Restart:\nAdmin → 💻 My Bot Code → 🔄 Restart Bot",
    'number':    "📱 Get Number:\n/start → 📱 Get Number → Service → Country → tap number to copy it",
    'broadcast': "📣 Broadcast:\nAdmin → 📣 Broadcast → send your message",
}

def ai_reply(question: str, api_key: str = '') -> str:
    q = question.lower()
    if api_key and len(api_key) > 30:
        try:
            r = http_requests.post(
                "https://api.anthropic.com/v1/messages",
                headers={"x-api-key": api_key, "anthropic-version": "2023-06-01",
                         "content-type": "application/json"},
                json={"model": "claude-3-haiku-20240307", "max_tokens": 600,
                      "system": SYSTEM_PROMPT,
                      "messages": [{"role": "user", "content": question}]},
                timeout=20
            )
            if r.ok: return r.json()['content'][0]['text']
        except: pass
    for kw, ans in KB.items():
        if kw in q: return f"🤖 **AI Answer:**\n\n{ans}"
    if any(w in q for w in ['restart','stop','run']): return f"🤖\n\n{KB['restart']}"
    if any(w in q for w in ['number','copy','get']): return f"🤖\n\n{KB['number']}"
    if any(w in q for w in ['xlsx','excel','file']): return f"🤖\n\n{KB['upload']}"
    return ("🤖 Ask me about:\n"
            "• uptime • upload • country • panel • webhook • code • admin • restart • number • broadcast\n\n"
            "💡 Set Anthropic API key for detailed answers.")

# ════════════════════════════════════════════════════════════
#  BOT CLIENT
# ════════════════════════════════════════════════════════════
# Session path: use /tmp on Render/cloud, local otherwise
_SESSION_PATH = os.environ.get('SESSION_PATH', '/tmp/tareq_bot')
bot = TelegramClient(_SESSION_PATH, API_ID, API_HASH).start(bot_token=BOT_TOKEN)

async def set_commands():
    try:
        await bot(SetBotCommandsRequest(
            scope=BotCommandScopeDefault(), lang_code='',
            commands=[
                BotCommand('start',     'Start Bot'),
                BotCommand('myhistory', 'My History'),
                BotCommand('lang',      'Toggle Language / ভাষা পরিবর্তন'),
            ]
        ))
    except Exception as e: print(f"Commands error: {e}")

def test_custom_panel_url(panel_url: str, key: str) -> str:
    """
    Test an API key against a custom panel URL.
    Tries common SMS panel API patterns against the given base URL.
    Returns formatted result string.
    """
    panel_url = panel_url.rstrip('/')
    key = key.strip()
    results = []

    # Common SMS panel API patterns (CRAPI, handler_api, REST JSON)
    test_patterns = [
        # CRAPI style
        {"method":"GET",  "url": f"{panel_url}/crapi/ks/viewstats?token={key}&records=10"},
        {"method":"GET",  "url": f"{panel_url}/api/sms?api_key={key}"},
        {"method":"GET",  "url": f"{panel_url}/api/profile?api_key={key}"},
        {"method":"GET",  "url": f"{panel_url}/api/balance?api_key={key}"},
        {"method":"GET",  "url": f"{panel_url}/api/messages?api_key={key}"},
        {"method":"GET",  "url": f"{panel_url}/api/inbox?api_key={key}"},
        {"method":"GET",  "url": f"{panel_url}/api/user?api_key={key}"},
        {"method":"GET",  "url": f"{panel_url}/stubs/handler_api.php?api_key={key}&action=getBalance"},
        {"method":"GET",  "url": f"{panel_url}/api.php?api_key={key}&action=getBalance"},
        {"method":"GET",  "url": f"{panel_url}/api/getBalance?api_key={key}"},
        {"method":"GET",  "url": f"{panel_url}/api/v1/profile", "headers": {"Authorization": f"Bearer {key}"}},
    ]

    for pat in test_patterns:
        try:
            hdrs = pat.get("headers", {})
            if pat["method"] == "GET":
                r = http_requests.get(pat["url"], headers=hdrs, timeout=8)
            else:
                r = http_requests.post(pat["url"], headers=hdrs, timeout=8)

            if r.status_code == 200:
                ep_short = pat["url"].replace(panel_url, "").split("?")[0]
                try:
                    j = r.json()
                    # Detect success indicators
                    bal = j.get('balance', j.get('balance_bdt', j.get('Balance', None)))
                    user = j.get('email', j.get('username', j.get('user_id', j.get('id', None))))
                    status = j.get('status', j.get('response', j.get('success', None)))
                    if bal is not None or user or status in ('success','ok',1,True,200):
                        info = f"Balance: `{bal}`" if bal is not None else (f"User: `{user}`" if user else "Connected ✅")
                        return (f"✅ *Panel Connected!*\n\n"
                                f"🌐 URL: `{panel_url}`\n"
                                f"🔑 Key: `{key}`\n"
                                f"📡 Endpoint: `{ep_short}`\n"
                                f"💰 {info}\n\n"
                                f"→ Admin Panel → SMS Panels → ➕ Add Panel → CRAPI URL বা JSON/URL Panel")
                    # If no error keywords → likely connected
                    if not any(w in r.text.lower() for w in ('error','invalid','unauthorized','forbidden','denied','fail')):
                        rec_count = len(j) if isinstance(j, list) else len(j.get('data', j.get('records', j.get('messages', []))))
                        return (f"✅ *Panel Connected!*\n\n"
                                f"🌐 URL: `{panel_url}`\n"
                                f"🔑 Key: `{key}`\n"
                                f"📡 Endpoint: `{ep_short}`\n"
                                f"📦 Records: `{rec_count}`\n\n"
                                f"→ Admin Panel → SMS Panels → ➕ Add Panel → CRAPI URL বা JSON/URL Panel")
                except Exception:
                    if not any(w in r.text.lower() for w in ('error','invalid','unauthorized','forbidden')):
                        ep_short = pat["url"].replace(panel_url, "").split("?")[0]
                        return (f"✅ *Panel Connected (HTTP 200)!*\n\n"
                                f"🌐 URL: `{panel_url}`\n"
                                f"🔑 Key: `{key}`\n"
                                f"📡 Endpoint: `{ep_short}`\n\n"
                                f"→ Admin Panel → SMS Panels → ➕ Add Panel")
            results.append(f"• `{pat['url'].replace(panel_url,'')}` → HTTP {r.status_code}")
        except Exception as e:
            results.append(f"• Connection failed: {str(e)[:40]}")

    fail_lines = "\n".join(results[:6])
    return (f"❌ *Panel URL test failed*\n\n"
            f"🌐 URL: `{panel_url}`\n"
            f"🔑 Key: `{key}`\n\n"
            f"Tried endpoints:\n{fail_lines}\n\n"
            f"নিশ্চিত করুন URL এবং Key সঠিক আছে।")


def test_api_key(key: str) -> str:
    """
    Test an API key against multiple SMS panel services.
    Supports:
      - Plain key → auto-detect service
      - 'URL KEY' format → test against custom panel URL
    Returns a formatted result string.
    """
    text = key.strip()

    # ── Auto-detect URL+KEY format ────────────────────────────────
    # If input contains a URL (http/https) + space + key
    _url_match = re.match(r'(https?://\S+)\s+(\S+)', text)
    if _url_match:
        panel_url, api_key = _url_match.group(1), _url_match.group(2)
        return test_custom_panel_url(panel_url, api_key)
    # Or if input IS a URL (user sent just the full API URL)
    if text.startswith('http') and ('api_key=' in text or 'token=' in text or 'apikey=' in text):
        # extract key from URL param
        _kv = re.search(r'(?:api_key|token|apikey)=([^&\s]+)', text)
        if _kv:
            from urllib.parse import urlparse
            _parsed = urlparse(text)
            base = f"{_parsed.scheme}://{_parsed.netloc}"
            return test_custom_panel_url(base, _kv.group(1))

    key = text
    results = []

    # ── 1. 5sim.net (Bearer token) ───────────────────────────────
    for try_key in _unique([key, _try_decode64(key)]):
        if not try_key: continue
        try:
            r = http_requests.get(
                "https://5sim.net/v1/user/profile",
                headers={"Authorization": f"Bearer {try_key}", "Accept": "application/json"},
                timeout=8)
            if r.status_code == 200:
                info = r.json()
                bal  = info.get('balance', 'N/A')
                return (f"✅ *5sim.net — VALID KEY!*\n\n"
                        f"💰 Balance: `{bal}`\n"
                        f"🔑 Key: `{try_key}`\n\n"
                        f"→ Admin Panel → SMS Panels → ➕ Add Panel → 5sim API Key")
            elif r.status_code == 401:
                results.append("❌ 5sim.net → 401 Unauthorized")
            else:
                results.append(f"❌ 5sim.net → HTTP {r.status_code}")
        except Exception as e:
            results.append(f"❌ 5sim.net → {str(e)[:40]}")

    # ── 2. sms-activate.org ──────────────────────────────────────
    for try_key in _unique([key, _try_decode64(key)]):
        if not try_key: continue
        try:
            r = http_requests.get(
                f"https://api.sms-activate.org/stubs/handler_api.php?api_key={try_key}&action=getBalance",
                timeout=8)
            if r.status_code == 200 and r.text.startswith("ACCESS_BALANCE:"):
                bal = r.text.split(":")[1]
                return (f"✅ *sms-activate.org — VALID KEY!*\n\n"
                        f"💰 Balance: `{bal}`\n"
                        f"🔑 Key: `{try_key}`\n\n"
                        f"→ Admin Panel → SMS Panels → ➕ Add Panel → JSON/URL Panel\n"
                        f"→ URL: `https://api.sms-activate.org/stubs/handler_api.php?api_key={try_key}&action=getActiveActivations`")
            else:
                results.append(f"❌ sms-activate.org → {r.text[:50]}")
        except Exception as e:
            results.append(f"❌ sms-activate.org → {str(e)[:40]}")

    # ── 3. smshub.org ────────────────────────────────────────────
    for try_key in _unique([key, _try_decode64(key)]):
        if not try_key: continue
        try:
            r = http_requests.post(
                "https://smshub.org/stubs/handler_api.php",
                data={"api_key": try_key, "action": "getBalance"},
                timeout=8)
            if r.status_code == 200 and "ACCESS_BALANCE" in r.text:
                bal = r.text.split(":")[1] if ":" in r.text else "N/A"
                return (f"✅ *smshub.org — VALID KEY!*\n\n"
                        f"💰 Balance: `{bal}`\n"
                        f"🔑 Key: `{try_key}`")
            else:
                results.append(f"❌ smshub.org → {r.text[:50]}")
        except Exception as e:
            results.append(f"❌ smshub.org → {str(e)[:40]}")

    # ── 4. grizzlysms.com ────────────────────────────────────────
    for try_key in _unique([key, _try_decode64(key)]):
        if not try_key: continue
        try:
            r = http_requests.get(
                f"https://grizzlysms.com/stubs/handler_api.php?api_key={try_key}&action=getBalance",
                timeout=8)
            if r.status_code == 200 and "ACCESS_BALANCE" in r.text:
                bal = r.text.split(":")[1] if ":" in r.text else "N/A"
                return (f"✅ *grizzlysms.com — VALID KEY!*\n\n"
                        f"💰 Balance: `{bal}`\n"
                        f"🔑 Key: `{try_key}`")
            else:
                results.append(f"❌ grizzlysms.com → {r.text[:50]}")
        except Exception as e:
            results.append(f"❌ grizzlysms.com → {str(e)[:40]}")

    # ── 5. onlinesim.io ──────────────────────────────────────────
    for try_key in _unique([key, _try_decode64(key)]):
        if not try_key: continue
        try:
            r = http_requests.get(
                f"https://onlinesim.io/api/getBalance.php?apikey={try_key}",
                timeout=8)
            if r.status_code == 200:
                j = r.json()
                if j.get('response') == 1:
                    return (f"✅ *onlinesim.io — VALID KEY!*\n\n"
                            f"💰 Balance: `{j.get('balance','N/A')}`\n"
                            f"🔑 Key: `{try_key}`")
            results.append(f"❌ onlinesim.io → {r.text[:50]}")
        except Exception as e:
            results.append(f"❌ onlinesim.io → {str(e)[:40]}")

    # ── 6. fastxotps.com ─────────────────────────────────────────
    for try_key in _unique([key, _try_decode64(key)]):
        if not try_key: continue
        try:
            for ep in [
                f"https://fastxotps.com/api/profile?api_key={try_key}",
                f"https://fastxotps.com/api/balance?api_key={try_key}",
                f"https://fastxotps.com/api/user?api_key={try_key}",
                f"https://fastxotps.com/api/sms?api_key={try_key}",
            ]:
                try:
                    r = http_requests.get(ep, timeout=8)
                    if r.status_code == 200:
                        try:
                            j = r.json()
                            if (j.get('status') in ('success','ok',1,True) or
                                    j.get('success') or j.get('balance') is not None or
                                    j.get('user_id') or j.get('id') or j.get('email')):
                                bal = j.get('balance', j.get('balance_bdt','N/A'))
                                return (f"✅ *fastxotps.com — VALID KEY!*\n\n"
                                        f"💰 Balance: `{bal}`\n"
                                        f"🔑 Key: `{try_key}`\n\n"
                                        f"→ Admin Panel → SMS Panels → ➕ Add Panel → fastxotps.com")
                        except Exception:
                            pass
                        if 'error' not in r.text.lower() and 'invalid' not in r.text.lower():
                            return (f"✅ *fastxotps.com — KEY ACCEPTED!*\n\n"
                                    f"🔑 Key: `{try_key}`\n\n"
                                    f"→ Admin Panel → SMS Panels → ➕ Add Panel → fastxotps.com")
                except Exception:
                    continue
            results.append("❌ fastxotps.com → Key not accepted")
        except Exception as e:
            results.append(f"❌ fastxotps.com → {str(e)[:40]}")

    fail_lines = "\n".join(results[:8])
    return (f"❌ *Key NOT valid on any known service:*\n\n{fail_lines}\n\n"
            f"💡 *যেকোনো panel site test করতে:*\n"
            f"URL এবং Key একসাথে পাঠান:\n"
            f"`https://yourpanel.com YOURKEY`\n\n"
            f"অথবা পুরো API URL পাঠান:\n"
            f"`https://panel.com/api?api_key=YOURKEY`")

def _try_decode64(s: str) -> str:
    """Attempt base64 decode; return decoded string or empty."""
    import base64 as _b64
    try:
        d = _b64.b64decode(s).decode('utf-8', errors='ignore').strip()
        return d if len(d) > 8 else ''
    except Exception:
        return ''

def _unique(lst):
    """Deduplicate list preserving order."""
    seen = set(); out = []
    for x in lst:
        if x and x not in seen:
            seen.add(x); out.append(x)
    return out

def get_force_channels():
    c.execute("SELECT username, label FROM force_channels ORDER BY id")
    return c.fetchall()

async def check_sub(uid):
    bad = []
    for ch, lbl in get_force_channels():
        try: await bot(GetParticipantRequest(channel=ch, participant=uid))
        except: bad.append((ch, lbl))
    return bad

# ════════════════════════════════════════════════════════════
#  COMPACT KEYBOARD
# ════════════════════════════════════════════════════════════
_TGAPI = f"https://api.telegram.org/bot{BOT_TOKEN}"

def send_compact_kb(chat_id: int, text: str, lang: str = 'en', parse_mode: str = 'Markdown'):
    http_requests.post(f"{_TGAPI}/sendMessage", json={
        "chat_id": chat_id,
        "text": text,
        "parse_mode": parse_mode,
        "reply_markup": {
            "keyboard": [
                ["📱 Get Number",    "🌐 Available Country"],
                ["📊 Active Number", "🔍 Search Number"],
                ["💬 Support",       "👥 Refer"],
                ["👤 My Status"],
            ],
            "resize_keyboard": True,
            "persistent": True,
            "one_time_keyboard": False,
        }
    }, timeout=10)

def adm_kb():
    return [
        [Button.inline("📁 Upload Numbers","adm_upload"), Button.inline("📊 Stock Stats","adm_stats")],
        [Button.inline("📈 Daily Report","adm_daily"),    Button.inline("📣 Broadcast","adm_bc")],
        [Button.inline("🌍 Countries","adm_countries"),   Button.inline("🔢 Quota","adm_quota")],
        [Button.inline("🔗 Links","adm_links"),           Button.inline("👥 Admins","adm_admins")],
        [Button.inline("📡 SMS Panels","adm_panels"),     Button.inline("💻 My Bot Code","adm_code")],
        [Button.inline("🤖 AI Assistant","adm_ai"),       Button.inline("🔑 Test API Key","adm_keytest")],
        [Button.inline("👁 User View","view_user")],
    ]

# ════════════════════════════════════════════════════════════
#  SHOW NUMBERS  — copy_text buttons (tap = instant copy)
# ════════════════════════════════════════════════════════════
async def show_numbers(event, uid, svc, short, edit=True, reset_shown=False):
    c_name, c_flag = ctry(short)
    srv  = SVC.get(svc, svc.upper())
    lim  = quota()
    lang = glang(uid)

    if reset_shown:
        SHOWN_IDS.pop(uid, None)

    seen = SHOWN_IDS.get(uid, set())

    if seen:
        ph = ','.join('?' * len(seen))
        c.execute(
            f"SELECT id,number FROM premium_stock "
            f"WHERE country=? AND service=? AND status=0 AND id NOT IN ({ph}) LIMIT ?",
            (short, svc, *seen, lim))
    else:
        c.execute(
            "SELECT id,number FROM premium_stock "
            "WHERE country=? AND service=? AND status=0 LIMIT ?",
            (short, svc, lim))
    rows = c.fetchall()

    if not rows and seen:
        SHOWN_IDS.pop(uid, None)
        c.execute(
            "SELECT id,number FROM premium_stock "
            "WHERE country=? AND service=? AND status=0 LIMIT ?",
            (short, svc, lim))
        rows = c.fetchall()

    no_msg  = (f"{c_flag} <b>{c_name.upper()}</b> {srv}\n\n"
               + ("❌ No stock. Select another country."
                  if lang=='en' else "❌ স্টক নেই। অন্য দেশ বেছে নিন।"))
    if not rows:
        kbd = [[{"text":"🌍 Change Country","callback_data":f"svc_{svc}"}],
               [{"text":"◀️ Back to Services","callback_data":"select_svc"}]]
        _send_or_edit(event, uid, edit, no_msg, kbd); return

    SHOWN_IDS[uid] = seen | {r[0] for r in rows}
    inc_hist(uid, nums=len(rows))

    now_ts = str(datetime.datetime.now())
    for _did, num in rows:
        nf2 = fmt_num(num)
        c.execute("INSERT OR IGNORE INTO user_number_assignments VALUES(?,?,?)",
                  (uid, nf2, now_ts))
        # Mark number as taken — removes it from live stock permanently
        c.execute("UPDATE premium_stock SET status=1 WHERE number=? OR number=?",
                  (nf2, num))
    db.commit()

    caption = ("📋 <b>নাম্বারে ট্যাপ করলেই কপি হয়ে যাবে!</b>"
               if lang=='bn'
               else "📋 <b>Tap a number — it copies instantly!</b>")
    msg = f"{c_flag} <b>{c_name.upper()}</b> {srv}\n\n{caption}"

    kbd = []
    for _db_id, num in rows:
        nf = fmt_num(num)
        kbd.append([{"text": f"{c_flag} 📋  {nf}", "copy_text": {"text": nf}}])

    kbd += [
        [{"text":"🔄 Change Number","callback_data":f"chg_{svc}_{short}"},
         {"text":"🌍 Change Country","callback_data":f"svc_{svc}"}],
        [{"text":"📢 OTP Group ↗","url": glink("otp_group")}],
    ]
    if is_admin(uid):
        kbd.append([{"text":"🛠 Admin Panel","callback_data":"go_admin"}])

    _send_or_edit(event, uid, edit, msg, kbd)

def _send_or_edit(event, uid: int, edit: bool, text: str, inline_keyboard: list):
    """Send or edit a message using HTTP Bot API with HTML parse mode."""
    payload = {"chat_id": uid, "text": text,
                "parse_mode": "HTML",
                "reply_markup": {"inline_keyboard": inline_keyboard}}
    if edit:
        try:
            mid = event.query.msg_id
            http_requests.post(f"{_TGAPI}/editMessageText",
                               json={**payload, "message_id": mid}, timeout=10)
        except Exception:
            http_requests.post(f"{_TGAPI}/sendMessage", json=payload, timeout=10)
    else:
        http_requests.post(f"{_TGAPI}/sendMessage", json=payload, timeout=10)

# ════════════════════════════════════════════════════════════
#  /start
# ════════════════════════════════════════════════════════════
@bot.on(events.NewMessage(pattern=r'^/start$'))
async def on_start(event):
    if event.is_channel or event.is_group: return
    uid = event.sender_id
    c.execute("INSERT OR IGNORE INTO bot_users VALUES (?)", (uid,)); db.commit()

    if is_admin(uid):
        c.execute("SELECT COUNT(*) FROM premium_stock WHERE status=0"); stock = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM bot_users"); users = c.fetchone()[0]
        await event.respond(
            f"🛠 **SMS MATRIX BOT — Admin Panel**\n\n"
            f"📦 Live Stock: **{stock}**\n👥 Users: **{users}**\n\n👇 Select:",
            buttons=adm_kb(), parse_mode='md'); return

    bad = await check_sub(uid)
    if bad:
        btns = [[Button.url(lbl, f"https://t.me/{ch}")] for ch, lbl in bad]
        btns.append([Button.inline("🔄 I Have Joined — Verify ✅", "vsub")])
        await event.respond(
            "⚠️ **Join all channels & groups to use this bot:**",
            buttons=btns, parse_mode='md'); return

    lang = glang(uid)
    send_compact_kb(uid,
        "🔥 *SMS MATRIX BOT* ✅\n\n" +
        ("📲 Unlimited Method Active\n💰 Earn 500-1000 BDT daily"
         if lang=='en' else
         "📲 Unlimited Method চালু\n💰 প্রতিদিন ৫০০-১০০০ টাকা ইনকামের সুযোগ"),
        lang=lang)

# ════════════════════════════════════════════════════════════
#  Menu commands
# ════════════════════════════════════════════════════════════
@bot.on(events.NewMessage(pattern=r'^/myhistory$'))
async def on_history(event):
    if event.is_channel or event.is_group: return
    uid = event.sender_id; lang = glang(uid)
    c.execute("SELECT otp_count,numbers_taken FROM history WHERE user_id=? AND date=?", (uid, today()))
    row = c.fetchone(); o = row[0] if row else 0; n = row[1] if row else 0
    if lang == 'bn':
        txt = f"📝 **আজকের হিস্টোরি**\n\n📱 নাম্বার নিয়েছেন: **{n}** টি"
        if o > 0: txt += f"\n✅ OTP রিসিভ হয়েছে: **{o}** টি"
        else: txt += "\n⏳ এখনো কোনো OTP রিসিভ হয়নি"
    else:
        txt = f"📝 **Today's History**\n\n📱 Numbers taken: **{n}**"
        if o > 0: txt += f"\n✅ OTPs received: **{o}**"
        else: txt += "\n⏳ No OTP received yet today"
    await event.respond(txt, parse_mode='md')

@bot.on(events.NewMessage(pattern=r'^/lang$'))
async def on_lang(event):
    if event.is_channel or event.is_group: return
    uid = event.sender_id
    nl = 'bn' if glang(uid)=='en' else 'en'
    slang(uid, nl)
    send_compact_kb(uid,
        "✅ Language → English\\! Tap *Get Number* to start\\." if nl=='en'
        else "✅ ভাষা বাংলায় পরিবর্তন হয়েছে। *Get Number* চাপুন।",
        lang=nl)

# ════════════════════════════════════════════════════════════
#  /add  (manual)
# ════════════════════════════════════════════════════════════
@bot.on(events.NewMessage(pattern=r'^/add (.+) (.+) (.+)$'))
async def on_add(event):
    if not is_admin(event.sender_id): return
    short = event.pattern_match.group(1).strip().lower()
    svc   = event.pattern_match.group(2).strip().lower()
    num   = event.pattern_match.group(3).strip()
    if not valid_phone(num): await event.respond("❌ Invalid number."); return
    try:
        c.execute("INSERT INTO premium_stock(country,service,number,status)VALUES(?,?,?,0)",
                  (short, svc, fmt_num(num))); db.commit()
        cn, cf = ctry(short)
        await event.respond(f"✅ Added: {cf} {cn} — {fmt_num(num)}")
    except sqlite3.IntegrityError:
        await event.respond("⚠️ Already exists!")

# ════════════════════════════════════════════════════════════
#  MESSAGE HANDLER
# ════════════════════════════════════════════════════════════
@bot.on(events.NewMessage())
async def on_msg(event):
    if event.is_channel or event.is_group: return
    uid  = event.sender_id
    text = (event.text or '').strip()
    lang = glang(uid)

    # ── Reply Keyboard Buttons ─────────────────────────────────
    if text == "📱 Get Number":
        STATES.pop(uid, None)
        await event.respond(
            "⚙️ Select service:" if lang=='en' else "⚙️ সার্ভিস সিলেক্ট করুন:",
            buttons=[
                [Button.inline("💬 WhatsApp","svc_whatsapp"), Button.inline("🔹 Telegram","svc_telegram")],
                [Button.inline("🎵 TikTok",  "svc_tiktok"),  Button.inline("🌐 Facebook","svc_facebook")],
                [Button.inline("📸 Instagram","svc_instagram")]]); return

    if text == "🌐 Available Country":
        c.execute("""SELECT a.country_name, a.short_name, a.flag, COUNT(p.id) as cnt
                     FROM active_countries a
                     LEFT JOIN premium_stock p ON p.country=a.short_name AND p.status=0
                     GROUP BY a.short_name ORDER BY a.country_name""")
        rows = c.fetchall()
        if not rows:
            await event.respond(
                "❌ No active countries yet." if lang=='en' else "❌ কোনো দেশ এখনো সক্রিয় নেই।"); return
        txt = "🌍 **Available Countries:**\n\n" if lang=='en' else "🌍 **উপলব্ধ দেশসমূহ:**\n\n"
        for cname, cshort, cflag, cnt in rows:
            txt += f"{cflag} {cname} — **{cnt}** numbers\n"
        await event.respond(txt, parse_mode='md'); return

    if text == "📊 Active Number":
        c.execute("""SELECT una.number, ps.service, ps.country
                     FROM user_number_assignments una
                     LEFT JOIN premium_stock ps ON REPLACE(REPLACE(ps.number,'+',''),'-','')
                                                 = REPLACE(REPLACE(una.number,'+',''),'-','')
                     WHERE una.user_id=? ORDER BY una.assigned_at DESC LIMIT 10""", (uid,))
        rows = c.fetchall()
        if not rows:
            await event.respond(
                "📊 You haven't taken any numbers yet." if lang=='en'
                else "📊 আপনি এখনো কোনো নাম্বার নেননি।"); return
        txt = "📊 **Your Active Numbers:**\n\n" if lang=='en' else "📊 **আপনার নাম্বারসমূহ:**\n\n"
        for num, svc_k, cshort in rows:
            cn, cf = ctry(cshort or ''); srv = SVC.get(svc_k, svc_k or 'Unknown')
            txt += f"{cf} `{num}` — {srv}\n"
        await event.respond(txt, parse_mode='md'); return

    if text == "🔍 Search Number":
        STATES[uid] = "search"
        await event.respond(
            "🔍 Type a number or prefix:" if lang=='en' else "🔍 নাম্বার বা প্রিফিক্স টাইপ করুন:"); return

    if text == "💬 Support":
        sup = glink('support_group', 'https://t.me/unlimited_ws_method')
        chn = "https://t.me/UnlimitedWSMethod0"
        txt = (f"💬 **Support**\n\nFor help join our group or channel:" if lang=='en'
               else f"💬 **সাপোর্ট**\n\nযেকোনো সমস্যায় যোগাযোগ করুন:")
        await event.respond(txt, parse_mode='md', buttons=[
            [Button.url("💬 Support Group ↗", sup),
             Button.url("📢 Support Channel ↗", chn)]]); return

    if text == "👥 Refer":
        me = await bot.get_me()
        ref_link = f"https://t.me/{me.username}?start=ref_{uid}"
        txt = (f"👥 **Refer & Earn**\n\n🔗 Your referral link:\n`{ref_link}`\n\nShare with friends!"
               if lang=='en'
               else f"👥 **রেফার করুন**\n\n🔗 আপনার রেফারেল লিংক:\n`{ref_link}`\n\nবন্ধুদের সাথে শেয়ার করুন!")
        await event.respond(txt, parse_mode='md',
                            buttons=[[Button.url("🔗 Share Link ↗", ref_link)]]); return

    if text == "👤 My Status":
        c.execute("SELECT otp_count, numbers_taken FROM history WHERE user_id=? AND date=?",
                  (uid, today()))
        row = c.fetchone(); o = row[0] if row else 0; n = row[1] if row else 0
        c.execute("SELECT COUNT(*) FROM user_number_assignments WHERE user_id=?", (uid,))
        total_nums = c.fetchone()[0]
        lang_label = "বাংলা 🇧🇩" if lang=='bn' else "English 🇬🇧"
        if lang == 'bn':
            txt = f"👤 **ইউজার প্রোফাইল**\n\n🆔 আইডি: `{uid}`\n🌐 ভাষা: {lang_label}"
            if total_nums > 0: txt += f"\n\n📊 মোট নাম্বার: **{total_nums}** টি"
            if n > 0:          txt += f"\n📅 আজ নিয়েছেন: **{n}** টি"
            txt += f"\n📩 আজ OTP পেয়েছেন: **{o}** টি"
        else:
            txt = f"👤 **User Profile**\n\n🆔 ID: `{uid}`\n🌐 Language: {lang_label}"
            if total_nums > 0: txt += f"\n\n📊 Total Numbers: **{total_nums}**"
            if n > 0:          txt += f"\n📅 Taken today: **{n}**"
            txt += f"\n📩 OTPs received today: **{o}**"
        await event.respond(txt, parse_mode='md'); return

    if event.file and STATES.get(uid,'').startswith("up_"):
        rest  = STATES.pop(uid)[3:]
        parts = rest.split("_", 1)
        short = parts[0]; svc = parts[1] if len(parts)>1 else "whatsapp"
        fname = event.file.name or 'file.txt'
        ext   = os.path.splitext(fname.lower())[1]

        if ext not in ('.txt','.csv','.xlsx','.xls'):
            await event.respond(
                "❌ Send **.txt** or **.xlsx** file only!",
                buttons=[[Button.inline("🔙","adm_upload")]], parse_mode='md')
            STATES[uid] = f"up_{short}_{svc}"; return

        path   = await event.download_media()
        nums   = read_numbers_from_file(path, fname)
        added  = 0; skipped = 0; wrong_country = 0
        cn, cf = ctry(short)
        prefixes = DIALCODE.get(short, [])
        prefix_str = "/".join(prefixes[:3]) if prefixes else "any"
        for num in nums:
            if valid_phone(num):
                if not num_matches_country(num, short):
                    wrong_country += 1
                else:
                    try:
                        c.execute("INSERT INTO premium_stock(country,service,number,status)VALUES(?,?,?,0)",
                                  (short, svc, fmt_num(num))); added += 1
                    except sqlite3.IntegrityError: pass
            elif num: skipped += 1
        db.commit()
        try: os.remove(path)
        except: pass

        alert = (f"🎉 New Numbers Available!\n\n"
                 f"{cf} **{cn.upper()}** {SVC.get(svc,svc)}\n"
                 f"🆕 New stock: **{added}** numbers!\n\nUse /start to get your numbers!")
        c.execute("SELECT user_id FROM bot_users")
        for (u,) in c.fetchall():
            try: await bot.send_message(u, alert, parse_mode='md')
            except: pass

        result_txt = (f"✅ Upload complete!\n{cf} {cn} — {SVC.get(svc,svc)}\n"
                      f"➕ Added: **{added}** | ⏭ Skipped: **{skipped}**")
        if wrong_country:
            result_txt += f"\n❌ Wrong country: **{wrong_country}** (must start with `{prefix_str}`)"
        await event.respond(result_txt, buttons=adm_kb(), parse_mode='md'); return

    if event.file and STATES.get(uid) == "new_code":
        STATES.pop(uid)
        fname = event.file.name or ''
        if not fname.lower().endswith('.py'):
            await event.respond("❌ Send a .py file!", buttons=[[Button.inline("🔙","adm_code")]]); return
        path = await event.download_media()
        try:
            with open(path,'r',encoding='utf-8') as f: code = f.read()
            try: os.remove(path)
            except: pass
            with open(BOT_FILE,'w',encoding='utf-8') as f: f.write(code)
            await event.respond("✅ Code saved! Restarting...", buttons=[[Button.inline("🔙","adm_code")]])
            time.sleep(1); os.execv(sys.executable, [sys.executable]+sys.argv)
        except Exception as e:
            await event.respond(f"❌ {e}", buttons=[[Button.inline("🔙","adm_code")]]); return

    if not text or uid not in STATES: return
    state = STATES.pop(uid)

    if state == "search":
        c.execute("SELECT number,country,service FROM premium_stock "
                  "WHERE number LIKE ? AND status=0 LIMIT 10", (f"%{text}%",))
        rows = c.fetchall()
        if rows:
            out = f"🔍 Results ({len(rows)}):\n\n"
            for num, ct, sv in rows:
                cn, cf = ctry(ct)
                out += f"{cf} {cn} — {SVC.get(sv,sv)}\n📞 `{fmt_num(num)}`\n\n"
        else:
            out = f"❌ No result for '{text}'."
        await event.respond(out, parse_mode='md'); return

    if state == "bc":
        c.execute("SELECT user_id FROM bot_users"); cnt = 0
        for (u,) in c.fetchall():
            try: await bot.send_message(u, text); cnt += 1
            except: pass
        await event.respond(f"✅ Broadcast → {cnt} users.", buttons=adm_kb()); return

    if state == "otp_link":
        c.execute("INSERT OR REPLACE INTO bot_links VALUES('otp_group',?)", (text,)); db.commit()
        await event.respond("✅ OTP link updated.", buttons=adm_kb()); return

    if state == "sup_link":
        c.execute("INSERT OR REPLACE INTO bot_links VALUES('support_group',?)", (text,)); db.commit()
        await event.respond("✅ Support link updated.", buttons=adm_kb()); return

    if state == "quota":
        if text.isdigit() and 1 <= int(text) <= 10:
            c.execute("INSERT OR REPLACE INTO bot_settings VALUES('numbers_per_user',?)", (text,)); db.commit()
            await event.respond(f"✅ Quota set: {text} numbers/user.", buttons=adm_kb())
        else: await event.respond("❌ Enter a number 1-10.", buttons=adm_kb())
        return

    if state == "add_adm":
        uid_to_add = None
        if text.isdigit():
            uid_to_add = int(text)
        elif text.startswith('@') or re.match(r'^[a-zA-Z]', text):
            try:
                entity = await bot.get_entity(text.lstrip('@'))
                uid_to_add = entity.id
            except Exception as e:
                await event.respond(f"❌ Username not found: {e}", buttons=adm_kb()); return
        else:
            await event.respond("❌ Send User ID or @username.", buttons=adm_kb()); return
        c.execute("SELECT COUNT(*) FROM admins"); cnt = c.fetchone()[0]
        if cnt >= 5:
            await event.respond("❌ Max 5 admins reached.", buttons=adm_kb()); return
        c.execute("INSERT OR IGNORE INTO admins VALUES(?,?,?)", (uid_to_add, uid, today())); db.commit()
        await event.respond(f"✅ Admin added: `{uid_to_add}`", buttons=adm_kb(), parse_mode='md'); return

    if state == "pname":
        # legacy fallback — now redirects through type-aware flow
        STATES[uid] = f"pval_fivesim:{text}"
        await event.respond(f"📡 **{text}**\n\nAPI Key বা URL পাঠান:",
                            buttons=[[Button.inline("🔙","adm_panels")]], parse_mode='md'); return

    if state.startswith("pname_"):
        ptype = state[6:]   # fivesim / crapi / url / fastxotps / ivasms
        pname_val = text.strip()
        STATES[uid] = f"pval_{ptype}:{pname_val}"
        prompts = {
            "fivesim":   "5sim API Key পাঠান (Bearer token):",
            "crapi":     "CRAPI URL পাঠান (token সহ):",
            "url":       "JSON endpoint URL পাঠান:",
            "fastxotps": "fastxotps.com API Key পাঠান (e.g. MURAD_XXXX...):",
            "ivasms":    (
                "🍪 **ivasms.com Cookie পাঠান**\n\n"
                "কীভাবে Cookie পাবেন:\n"
                "1. Browser-এ ivasms.com login করুন\n"
                "2. F12 → Network tab → যেকোনো request click করুন\n"
                "3. Request Headers → Cookie value copy করুন\n"
                "4. নিচে paste করুন\n\n"
                "অথবা:\n"
                "F12 → Application → Cookies → www.ivasms.com\n"
                "সব cookie একসাথে copy করুন।\n\n"
                "Format:\n"
                "`sessionid=xxx; csrftoken=yyy; ...`"
            ),
        }
        prompt = prompts.get(ptype, "URL বা API Key পাঠান:")
        await event.respond(
            f"📡 **{pname_val}**\n\n{prompt}",
            buttons=[[Button.inline("🔙","adm_panels")]], parse_mode='md'); return

    if state.startswith("pval_"):
        rest = state[5:]
        if ":" in rest:
            ptype, pname = rest.split(":", 1)
        else:
            pname = rest
            tl = text.lower()
            if 'crapi' in tl or 'viewstats' in tl or 'crapi' in pname.lower():
                ptype = "crapi"
            elif text.startswith("http"):
                ptype = "url"
            else:
                ptype = "fivesim"
        val = text.strip()
        c.execute("SELECT COUNT(*) FROM sms_panels"); cnt = c.fetchone()[0]
        if cnt >= 20:
            await event.respond("❌ Max 20 panels.", buttons=adm_kb()); return

        STATES.pop(uid, None)
        # ── Live test before saving ──────────────────────────────────────
        test_msg = await event.respond("⏳ Testing API connection...", parse_mode='md')
        test_ok = False; test_detail = ""
        try:
            if ptype == "crapi":
                if not val.startswith("http"):
                    test_ok = False; test_detail = "❌ Invalid URL — must start with `http`."
                else:
                    r = http_requests.get(val, timeout=10)
                    if r.status_code == 200:
                        try:
                            d = r.json()
                            n = len(d) if isinstance(d, list) else len(d.get('data', d.get('records', [])))
                            test_detail = f"✅ Connected — {n} record(s) পাওয়া গেছে।"
                        except Exception:
                            test_detail = "✅ HTTP 200 — connected!"
                        test_ok = True
                    elif r.status_code == 403:
                        test_detail = f"❌ HTTP 403 — Token invalid বা inactive।"
                    elif r.status_code == 401:
                        test_detail = "❌ HTTP 401 — Unauthorized."
                    else:
                        test_detail = f"⚠️ HTTP {r.status_code} — {r.text[:80]}"
            elif ptype == "fivesim":
                headers = {"Authorization": f"Bearer {val}", "Accept": "application/json"}
                r = http_requests.get("https://5sim.net/v1/user/profile", headers=headers, timeout=10)
                if r.status_code == 200:
                    try: uname = r.json().get('username', '?')
                    except Exception: uname = '?'
                    test_detail = f"✅ 5sim key valid — Account: `{uname}`"
                    test_ok = True
                elif r.status_code == 401:
                    test_detail = "❌ HTTP 401 — Invalid 5sim API key।"
                else:
                    test_detail = f"⚠️ HTTP {r.status_code} — {r.text[:80]}"
            elif ptype == "url":
                if not val.startswith("http"):
                    test_detail = "❌ Invalid URL."
                else:
                    r = http_requests.get(val, timeout=10)
                    if r.status_code == 200:
                        test_detail = "✅ URL reachable — HTTP 200."; test_ok = True
                    else:
                        test_detail = f"⚠️ HTTP {r.status_code} — {r.text[:80]}"
            elif ptype == "fastxotps":
                fastx_endpoints = [
                    f"https://fastxotps.com/api/profile?api_key={val}",
                    f"https://fastxotps.com/api/balance?api_key={val}",
                    f"https://fastxotps.com/api/user?api_key={val}",
                ]
                for _ep in fastx_endpoints:
                    try:
                        r = http_requests.get(_ep, timeout=10)
                        if r.status_code == 200:
                            try:
                                j = r.json()
                                if (j.get('status') in ('success','ok',1,True) or
                                        j.get('success') or j.get('balance') is not None or
                                        j.get('user_id') or j.get('id') or j.get('email')):
                                    bal = j.get('balance', j.get('balance_bdt','N/A'))
                                    test_detail = f"✅ fastxotps.com valid — Balance: `{bal}`"
                                    test_ok = True; break
                            except Exception:
                                pass
                            if 'error' not in r.text.lower() and 'invalid' not in r.text.lower():
                                test_detail = "✅ fastxotps.com — Key accepted!"
                                test_ok = True; break
                        elif r.status_code == 401:
                            test_detail = "❌ HTTP 401 — Invalid fastxotps API key।"
                    except Exception:
                        continue
                if not test_ok and not test_detail:
                    test_detail = "⚠️ fastxotps.com — Could not verify key. Use Force Save if key is correct."
            elif ptype == "ivasms":
                # Test cookie by fetching ivasms portal page
                _iv_hdrs = {
                    "Cookie": val,
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "text/html,*/*;q=0.8",
                }
                try:
                    iv_r = http_requests.get(
                        "https://www.ivasms.com/portal/live/my_sms",
                        headers=_iv_hdrs, timeout=12, allow_redirects=True)
                    if iv_r.status_code == 200 and 'login' not in iv_r.url:
                        test_ok = True
                        test_detail = (f"✅ ivasms.com Cookie valid!\n"
                                       f"Page loaded — OTP forwarding শুরু হবে।")
                    elif 'login' in iv_r.url or iv_r.status_code in (301, 302):
                        test_detail = ("❌ Cookie invalid — Login page-এ redirect হচ্ছে।\n"
                                       "Browser থেকে fresh cookie copy করুন।")
                    elif iv_r.status_code == 403:
                        test_detail = ("⚠️ HTTP 403 — Cloudflare block।\n"
                                       "Cookie সঠিক হলে Force Save করুন।")
                        test_ok = True  # Allow save, poller will retry
                    else:
                        test_detail = f"⚠️ HTTP {iv_r.status_code} — Use Force Save যদি cookie সঠিক হয়।"
                        test_ok = True  # Allow force save
                except Exception as _ive:
                    test_detail = f"⚠️ Connection error: {str(_ive)[:60]}\nForce Save করুন।"
                    test_ok = True  # Allow save anyway
            else:
                test_ok = True; test_detail = "✅ Saved."
        except Exception as e:
            test_detail = f"❌ Connection failed: {str(e)[:80]}"

        await test_msg.delete()

        if test_ok:
            c.execute("INSERT INTO sms_panels(name,panel_type,value,added_at)VALUES(?,?,?,?)",
                      (pname, ptype, val, today())); db.commit()
            type_labels = {"fivesim":"🔑 5sim","crapi":"📡 CRAPI","url":"🌐 URL","fastxotps":"🔥 fastxotps","ivasms":"🍪 ivasms"}
            label = type_labels.get(ptype, ptype.upper())
            await event.respond(
                f"✅ **Panel Added & Tested!**\n\n"
                f"📛 Name: `{pname}`\n🔖 Type: {label}\n\n"
                f"{test_detail}\n\n"
                f"✔️ OTP auto-forward চালু হয়ে গেছে।",
                buttons=adm_kb(), parse_mode='md')
        else:
            # Save failed panel with error noted, but give user choice
            await event.respond(
                f"⚠️ **API Test Failed**\n\n"
                f"📛 Name: `{pname}`\n🔖 Type: `{ptype.upper()}`\n\n"
                f"{test_detail}\n\n"
                f"API সঠিক হলে **Force Save** করো, না হলে আবার চেষ্টা করো।",
                buttons=[
                    [Button.inline("💾 Force Save (Anyway)", f"fsave_{ptype}:{pname}:{val[:60]}")],
                    [Button.inline("🔄 আবার চেষ্টা করো", "add_panel")],
                    [Button.inline("🔙 Panels", "adm_panels")],
                ], parse_mode='md')
        return

    if state.startswith("crapi_url_"):
        pid = int(state[10:])
        server_url = text.strip().rstrip('/')
        c.execute("SELECT name, value FROM sms_panels WHERE id=?", (pid,))
        r = c.fetchone()
        if not r:
            await event.respond("❌ Panel not found."); STATES.pop(uid, None); return
        nm, key = r
        import urllib.parse as _urlp
        full_url = f"{server_url}/crapi/ks/viewstats?token={_urlp.quote(key, safe='')}&records=100"
        # Live test before converting
        await event.respond(f"⏳ Testing `{server_url}` ...", parse_mode='md')
        try:
            resp = http_requests.get(full_url, timeout=8)
            if resp.status_code == 200:
                try:
                    rdata = resp.json()
                    rec_count = len(rdata) if isinstance(rdata, list) else len(rdata.get('data', rdata.get('records', [])))
                    result_line = f"✅ **Connected!** {rec_count} records পাওয়া গেছে।"
                    convert_ok = True
                except Exception:
                    result_line = f"✅ HTTP 200 — connected!"
                    convert_ok = True
            elif resp.status_code == 403:
                result_line = f"❌ HTTP 403 — key invalid/inactive এই server-এ।"
                convert_ok = False
            elif resp.status_code == 401:
                result_line = f"❌ HTTP 401 — unauthorized।"
                convert_ok = False
            else:
                result_line = f"⚠️ HTTP {resp.status_code} — {resp.text[:60]}"
                convert_ok = False
        except Exception as e:
            result_line = f"❌ Connection failed: {str(e)[:60]}"
            convert_ok = False

        STATES.pop(uid, None)
        if convert_ok:
            c.execute("UPDATE sms_panels SET panel_type='crapi', value=?, error_count=0, disabled_until='', last_error='' WHERE id=?",
                      (full_url, pid))
            db.commit()
            await event.respond(
                f"📡 **{nm}** converted!\n\n"
                f"{result_line}\n\n"
                f"✅ Type: `webhook` → `CRAPI`\n"
                f"🔗 URL:\n`{full_url[:80]}...`\n\n"
                f"বট এখন থেকে প্রতি ৫ সেকেন্ডে এই server poll করবে।",
                buttons=[[Button.inline("📡 Panels দেখো", "adm_panels")]], parse_mode='md')
        else:
            await event.respond(
                f"❌ **Convert হয়নি**\n\n{result_line}\n\n"
                f"Server URL ভুল হতে পারে। যে জায়গা থেকে key নিয়েছ তাদের server address জিজ্ঞেস করো।",
                buttons=[
                    [Button.inline("🔄 আবার Try করো", f"conv2crapi_{pid}")],
                    [Button.inline("🔙 Panels", "adm_panels")]
                ], parse_mode='md')
        return

    if state.startswith("ivasms_newcookie:"):
        pid_str = state.split(":", 1)[1]
        try:
            pid_iv = int(pid_str)
        except ValueError:
            await event.respond("❌ Error.", buttons=adm_kb()); STATES.pop(uid, None); return
        new_cookie = text.strip()
        if len(new_cookie) < 10:
            await event.respond("❌ Cookie too short. আবার চেষ্টা করুন।",
                                buttons=[[Button.inline("🔙","adm_panels")]]); return
        STATES.pop(uid, None)
        # Update cookie in DB
        c.execute("UPDATE sms_panels SET value=?, error_count=0, disabled_until='', last_error='' WHERE id=?",
                  (new_cookie, pid_iv)); db.commit()
        c.execute("SELECT name FROM sms_panels WHERE id=?", (pid_iv,)); nr = c.fetchone()
        nm_iv = nr[0] if nr else "Panel"
        # Test immediately
        await event.respond("⏳ Cookie testing...", parse_mode='md')
        loop = asyncio.get_event_loop()
        def _test_new_iv():
            hdrs = {
                "Cookie": new_cookie,
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Accept": "text/html,*/*;q=0.8",
            }
            rv = http_requests.get(
                "https://www.ivasms.com/portal/live/my_sms",
                headers=hdrs, timeout=12, allow_redirects=True)
            return rv.status_code, rv.url
        try:
            sc, final_url = await loop.run_in_executor(None, _test_new_iv)
            if sc == 200 and 'login' not in str(final_url):
                res = "✅ Cookie valid! OTP forwarding চলছে।"
            elif 'login' in str(final_url):
                res = "⚠️ Cookie invalid — login redirect হচ্ছে। Poller retry করবে।"
            else:
                res = f"⚠️ HTTP {sc} — Poller retry করবে।"
        except Exception as e:
            res = f"⚠️ Test error: {str(e)[:60]}"
        await event.respond(
            f"🍪 **{nm_iv}** — Cookie Updated!\n\n{res}",
            buttons=[
                [Button.inline("🔙 Panel Info", f"opn_{pid_iv}")],
                [Button.inline("📡 Panels", "adm_panels")],
            ], parse_mode='md')
        return

    if state.startswith("fastxotps_newcookie:"):
        pid_str = state.split(":", 1)[1]
        try:
            pid_fx = int(pid_str)
        except ValueError:
            await event.respond("❌ Error.", buttons=adm_kb()); STATES.pop(uid, None); return
        new_cookie = text.strip()
        if len(new_cookie) < 10:
            await event.respond("❌ Cookie too short. আবার চেষ্টা করুন।",
                                buttons=[[Button.inline("🔙","adm_panels")]]); return
        STATES.pop(uid, None)
        # Read current value to keep API key
        c.execute("SELECT name, value FROM sms_panels WHERE id=?", (pid_fx,)); nr = c.fetchone()
        if not nr:
            await event.respond("❌ Panel not found.", buttons=adm_kb()); return
        nm_fx, cur_val = nr
        fkey_old, _ = _parse_fastxotps_value(cur_val)
        # New value: keep existing key, update cookie
        if fkey_old:
            new_val = f"{fkey_old}|||{new_cookie}"
        else:
            new_val = f"|||{new_cookie}"
        c.execute("UPDATE sms_panels SET value=?, error_count=0, disabled_until='', last_error='' WHERE id=?",
                  (new_val, pid_fx)); db.commit()
        # Quick test
        await event.respond("⏳ Cookie testing...", parse_mode='md')
        loop = asyncio.get_event_loop()
        def _test_fx_new():
            hdrs = {
                "Cookie": new_cookie,
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Referer": "https://fastxotps.com/",
            }
            ep = (f"https://fastxotps.com/api/sms?api_key={fkey_old}" if fkey_old
                  else "https://fastxotps.com/dashboard")
            rv = http_requests.get(ep, headers=hdrs, timeout=12, allow_redirects=True)
            return rv.status_code, str(rv.url)
        try:
            sc, final_url = await loop.run_in_executor(None, _test_fx_new)
            if sc == 200 and 'tzddos' not in final_url and 'login' not in final_url:
                res = "✅ Cookie VALID! DDoS bypass সফল। OTP polling চালু।"
            elif 'tzddos' in final_url:
                res = "⚠️ DDoS Cookie invalid — tzddos_verified cookie সহ দিন। Poller retry করবে।"
            elif 'login' in final_url:
                res = "⚠️ Session expired — login page redirect। নতুন session cookie দিন।"
            else:
                res = f"⚠️ HTTP {sc} — Poller retry করবে।"
        except Exception as e:
            res = f"⚠️ Test error: {str(e)[:60]}"
        await event.respond(
            f"🔥 **{nm_fx}** — Cookie Updated!\n\n{res}",
            buttons=[
                [Button.inline("🔙 Panel Info", f"opn_{pid_fx}")],
                [Button.inline("📡 Panels", "adm_panels")],
            ], parse_mode='md')
        return

    if state == "keytest":
        # Detect if user sent URL+KEY or just KEY
        _has_url = bool(re.match(r'https?://\S+', text.strip()))
        _wait_msg = (
            "⏳ Custom panel URL testing... please wait."
            if _has_url else
            "⏳ Testing your key on 6 services... please wait."
        )
        await event.respond(_wait_msg, parse_mode='md')
        loop = asyncio.get_event_loop()
        result = await loop.run_in_executor(None, test_api_key, text)
        await event.respond(
            f"🔑 **Key Test Result**\n\n{result}",
            buttons=[
                [Button.inline("🔑 Test Another Key","adm_keytest")],
                [Button.inline("🌐 Test Custom Panel URL","adm_keytest_custom")],
                [Button.inline("🔙 Admin Panel","adm_back")],
            ],
            parse_mode='md')
        return

    if state == "set_ai_key":
        c.execute("INSERT OR REPLACE INTO bot_settings VALUES('ai_api_key',?)", (text,)); db.commit()
        await event.respond("✅ AI API Key saved!", buttons=[[Button.inline("🔙","adm_ai")]]); return

    if state == "ai":
        ans = ai_reply(text, gset('ai_api_key',''))
        await event.respond(ans + "\n\n_Ask another question or /start_",
                            buttons=[[Button.inline("❌ Exit","adm_ai")]], parse_mode='md')
        STATES[uid] = "ai"; return

    if state == "fch_username":
        raw = text.strip()
        # accept full link: https://t.me/username  OR  @username  OR  username
        if 't.me/' in raw:
            uname = raw.split('t.me/')[-1].split('/')[0].split('?')[0].strip()
        else:
            uname = raw.lstrip('@').strip()
        if not re.match(r'^[a-zA-Z0-9_]{3,64}$', uname):
            await event.respond(
                "❌ Invalid username.\nSend @username, username, or https://t.me/username link.",
                buttons=[[Button.inline("🔙","adm_links")]]); return
        chs = get_force_channels()
        if len(chs) >= 10:
            await event.respond("❌ Max 10 force channels reached!", buttons=[[Button.inline("🔙","adm_links")]]); return
        STATES[uid] = f"fch_label_{uname}"
        await event.respond(
            f"✅ Username: @{uname}\n\nNow send a label (e.g. `📢 Channel 3`):",
            buttons=[[Button.inline("🔙","adm_links")]], parse_mode='md'); return

    if state.startswith("fch_label_"):
        uname = state[10:]
        label = text.strip()[:30]
        try:
            c.execute("INSERT OR IGNORE INTO force_channels(username,label,added_at) VALUES(?,?,?)",
                      (uname, label, today())); db.commit()
            await event.respond(
                f"✅ Force channel added!\n@{uname} — {label}",
                buttons=[[Button.inline("🔗 Links Menu","adm_links")]], parse_mode='md')
        except Exception as e:
            await event.respond(f"❌ Error: {e}", buttons=[[Button.inline("🔙","adm_links")]])
        return

# ════════════════════════════════════════════════════════════
#  CALLBACK HANDLER
# ════════════════════════════════════════════════════════════
@bot.on(events.CallbackQuery)
async def on_cb(event):
    data = event.data.decode(); uid = event.sender_id; lang = glang(uid)

    # ── Main Menu Inline Buttons ─────────────────────────────────
    if data == "btn_get_number":
        await event.answer()
        STATES.pop(uid, None)
        await bot.send_message(uid,
            "⚙️ Select service:" if lang=='en' else "⚙️ সার্ভিস সিলেক্ট করুন:",
            buttons=[
                [Button.inline("💬 WhatsApp","svc_whatsapp"), Button.inline("🔹 Telegram","svc_telegram")],
                [Button.inline("🎵 TikTok",  "svc_tiktok"),  Button.inline("🌐 Facebook","svc_facebook")],
                [Button.inline("📸 Instagram","svc_instagram")]]); return

    if data == "btn_countries":
        await event.answer()
        c.execute("""SELECT a.country_name, a.short_name, a.flag,
                            COUNT(p.id) as cnt
                     FROM active_countries a
                     LEFT JOIN premium_stock p ON p.country=a.short_name AND p.status=0
                     GROUP BY a.short_name ORDER BY a.country_name""")
        rows = c.fetchall()
        if not rows:
            await bot.send_message(uid,
                "❌ No active countries yet." if lang=='en' else "❌ কোনো দেশ এখনো সক্রিয় নেই।"); return
        txt = ("🌍 **Available Countries:**\n\n" if lang=='en'
               else "🌍 **উপলব্ধ দেশসমূহ:**\n\n")
        for cname, cshort, cflag, cnt in rows:
            txt += f"{cflag} {cname} — **{cnt}** numbers\n"
        await bot.send_message(uid, txt, parse_mode='md'); return

    if data == "btn_active":
        await event.answer()
        c.execute("""SELECT una.number, ps.service, ps.country
                     FROM user_number_assignments una
                     LEFT JOIN premium_stock ps ON REPLACE(REPLACE(ps.number,'+',''),'-','')
                                                 = REPLACE(REPLACE(una.number,'+',''),'-','')
                     WHERE una.user_id=?
                     ORDER BY una.assigned_at DESC LIMIT 10""", (uid,))
        rows = c.fetchall()
        if not rows:
            await bot.send_message(uid,
                "📊 You haven't taken any numbers yet." if lang=='en'
                else "📊 আপনি এখনো কোনো নাম্বার নেননি।"); return
        txt = ("📊 **Your Active Numbers:**\n\n" if lang=='en'
               else "📊 **আপনার নাম্বারসমূহ:**\n\n")
        for num, svc_k, cshort in rows:
            cn, cf = ctry(cshort or ''); srv = SVC.get(svc_k, svc_k or 'Unknown')
            txt += f"{cf} `{num}` — {srv}\n"
        await bot.send_message(uid, txt, parse_mode='md'); return

    if data == "btn_search":
        await event.answer()
        STATES[uid] = "search"
        await bot.send_message(uid,
            "🔍 Type a number or prefix:" if lang=='en' else "🔍 নাম্বার বা প্রিফিক্স টাইপ করুন:"); return

    if data == "btn_support":
        await event.answer()
        sup = glink('support_group', 'https://t.me/unlimited_ws_method')
        chn = "https://t.me/UnlimitedWSMethod0"
        txt = (f"💬 **Support**\n\nFor help join our group or channel:"
               if lang=='en'
               else f"💬 **সাপোর্ট**\n\nযেকোনো সমস্যায় যোগাযোগ করুন:")
        await bot.send_message(uid, txt, parse_mode='md', buttons=[
            [Button.url("💬 Support Group ↗", sup),
             Button.url("📢 Support Channel ↗", chn)]]); return

    if data == "btn_refer":
        await event.answer()
        me = await bot.get_me()
        ref_link = f"https://t.me/{me.username}?start=ref_{uid}"
        txt = (f"👥 **Refer & Earn**\n\n🔗 Your referral link:\n`{ref_link}`\n\nShare with friends!"
               if lang=='en'
               else f"👥 **রেফার করুন**\n\n🔗 আপনার রেফারেল লিংক:\n`{ref_link}`\n\nবন্ধুদের সাথে শেয়ার করুন!")
        await bot.send_message(uid, txt, parse_mode='md',
                               buttons=[[Button.url("🔗 Share Link ↗", ref_link)]]); return

    if data == "btn_status":
        await event.answer()
        c.execute("SELECT otp_count, numbers_taken FROM history WHERE user_id=? AND date=?",
                  (uid, today()))
        row = c.fetchone(); o = row[0] if row else 0; n = row[1] if row else 0
        c.execute("SELECT COUNT(*) FROM user_number_assignments WHERE user_id=?", (uid,))
        total_nums = c.fetchone()[0]
        lang_label = "বাংলা 🇧🇩" if lang=='bn' else "English 🇬🇧"
        if lang == 'bn':
            txt = f"👤 **ইউজার প্রোফাইল**\n\n🆔 আইডি: `{uid}`\n🌐 ভাষা: {lang_label}"
            if total_nums > 0: txt += f"\n\n📊 মোট নাম্বার: **{total_nums}** টি"
            if n > 0:          txt += f"\n📅 আজ নিয়েছেন: **{n}** টি"
            txt += f"\n📩 আজ OTP পেয়েছেন: **{o}** টি"
        else:
            txt = f"👤 **User Profile**\n\n🆔 ID: `{uid}`\n🌐 Language: {lang_label}"
            if total_nums > 0: txt += f"\n\n📊 Total Numbers: **{total_nums}**"
            if n > 0:          txt += f"\n📅 Taken today: **{n}**"
            txt += f"\n📩 OTPs received today: **{o}**"
        await bot.send_message(uid, txt, parse_mode='md'); return

    if data == "vsub":
        bad = await check_sub(uid)
        if bad:
            await event.answer("❌ You haven't joined yet!", alert=True); return
        await event.answer("✅ Verified! Welcome.", alert=False)
        await event.delete()
        send_compact_kb(uid,
            "🔥 *SMS MATRIX BOT* ✅\n\n"
            + ("📲 Unlimited Method Active\n💰 Earn 500-1000 BDT daily"
               if lang=='en' else
               "📲 Unlimited Method চালু\n💰 প্রতিদিন ৫০০-১০০০ টাকা"),
            lang=lang); return


    if data.startswith("chg_"):
        parts = data.split("_", 2)
        svc   = parts[1]; short = parts[2]
        await show_numbers(event, uid, svc, short, edit=True); return

    if data == "adm_back":
        STATES.pop(uid, None)
        c.execute("SELECT COUNT(*) FROM premium_stock WHERE status=0"); stock = c.fetchone()[0]
        await event.edit(f"🛠 **Admin Panel**\n\n📦 Live Stock: **{stock}**",
                         buttons=adm_kb(), parse_mode='md'); return

    if data == "view_user":
        await event.delete()
        send_compact_kb(uid,
            "🔥 *SMS MATRIX BOT* ✅\n\n"
            + ("📲 Unlimited Method Active\n💰 Earn 500-1000 BDT daily"
               if lang=='en' else "📲 Unlimited Method চালু\n💰 প্রতিদিন ৫০০-১০০০ টাকা"),
            lang=lang); return

    if data == "go_admin":
        if not is_admin(uid): await event.answer("❌ Access denied.", alert=True); return
        c.execute("SELECT COUNT(*) FROM premium_stock WHERE status=0"); stock = c.fetchone()[0]
        await event.edit(f"🛠 **Admin Panel**\n\n📦 Live Stock: **{stock}**",
                         buttons=adm_kb(), parse_mode='md'); return

    if data == "select_svc":
        await event.edit("⚙️ Select service:", buttons=[
            [Button.inline("💬 WhatsApp","svc_whatsapp"), Button.inline("🔹 Telegram","svc_telegram")],
            [Button.inline("🎵 TikTok",  "svc_tiktok"),  Button.inline("🌐 Facebook","svc_facebook")],
            [Button.inline("📸 Instagram","svc_instagram")]]); return

    if data.startswith("svc_"):
        svc = data[4:]
        c.execute("SELECT DISTINCT p.country FROM premium_stock p "
                  "JOIN active_countries a ON a.short_name=p.country "
                  "WHERE p.service=? AND p.status=0", (svc,))
        ws = {r[0] for r in c.fetchall()}
        if not ws:
            await event.edit(f"❌ No stock for {SVC.get(svc,svc)}.",
                             buttons=[[Button.inline("◀️ Back","select_svc")]]); return
        btns = []; row = []
        for short in sorted(ws):
            cn, cf = ctry(short)
            nm = cn[:8] if len(cn)>8 else cn
            row.append(Button.inline(f"{cf} {nm} [{short.upper()}]", f"ctry_{svc}_{short}"))
            if len(row)==2: btns.append(row); row=[]
        if row: btns.append(row)
        btns.append([Button.inline("◀️ Back to Services","select_svc")])
        await event.edit(f"🌍 Select Country for {SVC.get(svc,svc)}:", buttons=btns); return

    if data.startswith("ctry_"):
        _, svc, short = data.split("_", 2)
        SHOWN_IDS.pop(uid, None)
        await show_numbers(event, uid, svc, short, edit=True); return

    if not is_admin(uid):
        await event.answer("❌ Admin only.", alert=True); return

    if data == "adm_stats":
        c.execute("SELECT country,service,COUNT(*) FROM premium_stock WHERE status=0 GROUP BY country,service")
        rows = c.fetchall()
        c.execute("SELECT COUNT(*) FROM bot_users"); users = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM premium_stock WHERE status=0"); total = c.fetchone()[0]
        txt = f"📊 **Stock Report**\n👥 Users: {users} | 📦 Total: {total}\n\n"
        for ct, sv, cnt_val in rows:
            cn, cf = ctry(ct); txt += f"{cf} {cn} — {SVC.get(sv,sv)}: **{cnt_val}**\n"
        if not rows: txt += "❌ No stock."
        await event.edit(txt, buttons=[[Button.inline("🔙","adm_back")]], parse_mode='md'); return

    if data == "adm_daily":
        td = today()
        c.execute("SELECT COUNT(*) FROM otp_log WHERE received_at LIKE ?", (f"{td}%",))
        total_otp = c.fetchone()[0]
        c.execute("SELECT SUM(numbers_taken) FROM history WHERE date=?", (td,))
        total_nums = c.fetchone()[0] or 0
        c.execute("SELECT SUM(otp_count) FROM history WHERE date=?", (td,))
        total_otp_users = c.fetchone()[0] or 0
        # Premium: per-panel forwarded count today
        c.execute("SELECT panel_name, forwarded_count FROM forwarded_stats WHERE date=? ORDER BY forwarded_count DESC", (td,))
        panel_rows = c.fetchall()
        total_forwarded_today = sum(r[1] for r in panel_rows)
        # All-time total
        c.execute("SELECT SUM(forwarded_count) FROM forwarded_stats")
        all_time = c.fetchone()[0] or 0
        c.execute("""SELECT user_id, numbers_taken, otp_count FROM history
                     WHERE date=? ORDER BY numbers_taken DESC LIMIT 10""", (td,))
        user_rows = c.fetchall()
        txt = (f"📈 **Daily Report — {td}**\n\n"
               f"🚀 **OTP Forwarded Today: {total_forwarded_today}**\n"
               f"🏆 **All-Time Forwarded: {all_time}**\n\n"
               f"📡 **Per-Panel (Today):**\n")
        if panel_rows:
            for pn, pc in panel_rows:
                txt += f"  • {pn}: **{pc}** OTPs\n"
        else:
            txt += "  • No forwards yet today\n"
        txt += (f"\n📥 Total OTPs received: **{total_otp}**\n"
                f"📱 Total numbers taken: **{total_nums}**\n"
                f"✅ OTPs credited to users: **{total_otp_users}**\n\n"
                f"👥 **Top Users Today:**\n")
        if user_rows:
            for i, (u_id, u_nums, u_otp) in enumerate(user_rows, 1):
                txt += f"{i}. `{u_id}` — 📱{u_nums} nums | ✅{u_otp} OTPs\n"
        else:
            txt += "❌ No activity today."
        await event.edit(txt, buttons=[[Button.inline("🔄 Refresh","adm_daily"),
                                        Button.inline("🔙","adm_back")]], parse_mode='md'); return

    if data == "adm_quota":
        STATES[uid] = "quota"
        await event.edit(f"🔢 Numbers per user (current: {quota()})\n\nSend 1-10:",
                         buttons=[[Button.inline("🔙","adm_back")]]); return

    if data == "adm_bc":
        STATES[uid] = "bc"
        c.execute("SELECT COUNT(*) FROM bot_users"); n = c.fetchone()[0]
        await event.edit(f"📣 Broadcast to {n} users\n\nSend your message:",
                         buttons=[[Button.inline("🔙","adm_back")]]); return

    if data == "adm_links":
        chs = get_force_channels()
        txt = (f"🔗 **Links & Force Channels**\n\n"
               f"📢 OTP Link: {glink('otp_group')}\n"
               f"💬 Support: {glink('support_group')}\n\n"
               f"🔒 **Force Channels** ({len(chs)}/10):\n")
        for u, l in chs:
            txt += f"• {l} — @{u}\n"
        btns = [
            [Button.inline("✏️ OTP Link","edit_otp"), Button.inline("✏️ Support Link","edit_sup")],
            [Button.inline("➕ Add Force Channel","add_fch")],
        ]
        for ch_u, ch_l in chs:
            btns.append([Button.inline(f"❌ Remove {ch_l} (@{ch_u})", f"rmfch_{ch_u}")])
        btns.append([Button.inline("🔙","adm_back")])
        await event.edit(txt, buttons=btns, parse_mode='md'); return

    if data == "edit_otp":
        STATES[uid] = "otp_link"
        await event.edit("✏️ Send new OTP group link:", buttons=[[Button.inline("🔙","adm_links")]]); return

    if data == "edit_sup":
        STATES[uid] = "sup_link"
        await event.edit("✏️ Send new support link:", buttons=[[Button.inline("🔙","adm_links")]]); return

    if data == "add_fch":
        chs = get_force_channels()
        if len(chs) >= 10:
            await event.answer("❌ Max 10 channels reached!", alert=True); return
        STATES[uid] = "fch_username"
        await event.edit(
            "➕ **Add Force Channel/Group**\n\nSend the @username (without @):\nExample: `MyChannel`",
            buttons=[[Button.inline("🔙","adm_links")]], parse_mode='md'); return

    if data.startswith("rmfch_"):
        ch_u = data[6:]
        c.execute("DELETE FROM force_channels WHERE username=?", (ch_u,)); db.commit()
        await event.answer(f"✅ @{ch_u} removed!", alert=False)
        chs = get_force_channels()
        txt = (f"🔗 **Links & Force Channels**\n\n"
               f"📢 OTP Link: {glink('otp_group')}\n"
               f"💬 Support: {glink('support_group')}\n\n"
               f"🔒 **Force Channels** ({len(chs)}/10):\n")
        for u, l in chs:
            txt += f"• {l} — @{u}\n"
        btns = [
            [Button.inline("✏️ OTP Link","edit_otp"), Button.inline("✏️ Support Link","edit_sup")],
            [Button.inline("➕ Add Force Channel","add_fch")],
        ]
        for ch_u2, ch_l2 in chs:
            btns.append([Button.inline(f"❌ Remove {ch_l2} (@{ch_u2})", f"rmfch_{ch_u2}")])
        btns.append([Button.inline("🔙","adm_back")])
        await event.edit(txt, buttons=btns, parse_mode='md'); return

    if data == "adm_admins":
        c.execute("SELECT user_id FROM admins"); adms = c.fetchall()
        txt = f"👥 **Admins** ({len(adms)}/5)\n\n"; btns = []
        for (aid,) in adms:
            sup = aid == SUPER_ADMIN
            txt += f"{'👑' if sup else '🔹'} `{aid}`{'  (Super)' if sup else ''}\n"
            if not sup: btns.append([Button.inline(f"❌ Remove {aid}", f"rmadm_{aid}")])
        if len(adms) < 5:
            btns.append([Button.inline("➕ Add Admin (ID or @username)","addadm")])
        btns.append([Button.inline("🔙","adm_back")])
        await event.edit(txt, buttons=btns, parse_mode='md'); return

    if data == "addadm":
        STATES[uid] = "add_adm"
        await event.edit("➕ Send User ID or @username:",
                         buttons=[[Button.inline("🔙","adm_admins")]]); return

    if data.startswith("rmadm_"):
        if uid != SUPER_ADMIN: await event.answer("❌ Super Admin only.", alert=True); return
        rid = int(data[6:])
        c.execute("DELETE FROM admins WHERE user_id=?", (rid,)); db.commit()
        await event.answer(f"✅ Removed {rid}", alert=False)
        c.execute("SELECT user_id FROM admins"); adms = c.fetchall()
        txt = f"👥 **Admins** ({len(adms)}/5)\n\n"; btns = []
        for (aid,) in adms:
            sup = aid == SUPER_ADMIN
            txt += f"{'👑' if sup else '🔹'} `{aid}`{'  (Super)' if sup else ''}\n"
            if not sup: btns.append([Button.inline(f"❌ Remove {aid}", f"rmadm_{aid}")])
        if len(adms) < 5: btns.append([Button.inline("➕ Add Admin","addadm")])
        btns.append([Button.inline("🔙","adm_back")])
        await event.edit(txt, buttons=btns, parse_mode='md'); return

    if data == "adm_panels":
        c.execute("SELECT id,name,panel_type,value,error_count,disabled_until,last_ok,last_error FROM sms_panels ORDER BY id")
        pnls = c.fetchall()
        now_dt   = datetime.datetime.now()
        today_s  = datetime.date.today().isoformat()

        # ── Today's forwarded OTP counts per panel ─────────────────
        c.execute("SELECT panel_name, forwarded_count FROM forwarded_stats WHERE date=?", (today_s,))
        today_fwd = {r[0]: r[1] for r in c.fetchall()}

        # ── Pending OTP queue count ──────────────────────────────────
        c.execute("SELECT COUNT(*) FROM failed_otp_queue WHERE retry_count < 10")
        pending_q = c.fetchone()[0]

        # ── Per-panel status logic ───────────────────────────────────
        n_active = 0; n_pending = 0; n_error = 0; n_disabled = 0

        def _panel_status(nm, pt, ecnt, dis_until, fwd_today, last_ok_ts):
            """Returns (icon, status_label, counter_bucket)"""
            nonlocal n_active, n_pending, n_error, n_disabled
            if dis_until:
                try:
                    if now_dt < datetime.datetime.fromisoformat(dis_until):
                        until_dt = datetime.datetime.fromisoformat(dis_until)
                        mins_left = int((until_dt - now_dt).total_seconds() // 60)
                        n_disabled += 1
                        return "🔴", f"Disabled — {mins_left}m left ({ecnt} errors)"
                except Exception:
                    pass
            if ecnt and int(ecnt) > 0:
                n_error += 1
                return "🔴", f"Error ({ecnt}x)"
            if fwd_today and fwd_today > 0:
                n_active += 1
                return "🟢", f"Active — {fwd_today} OTPs today"
            # No OTPs today yet — check if panel pinged successfully today
            _seen_today = (last_ok_ts or '')[:10] == today_s
            if pt == "webhook":
                n_pending += 1
                if _seen_today:
                    return "🟡", "Connected — OTP অপেক্ষায় (webhook)"
                return "🟡", "Pending — অপেক্ষায় (webhook)"
            n_pending += 1
            if _seen_today:
                return "🟡", "Connected — আজ কোনো OTP নেই"
            return "🟡", "Pending — আজ কোনো OTP আসেনি"

        rows_info = []
        for pid, nm, pt, v, ecnt, dis_until, last_ok, last_err in pnls:
            fwd_today = today_fwd.get(nm, 0)
            icon, status_lbl = _panel_status(nm, pt, ecnt, dis_until, fwd_today, last_ok)
            rows_info.append((pid, nm, pt, v, ecnt, dis_until, last_ok, last_err, icon, status_lbl, fwd_today))

        txt  = f"📡 **SMS Panels** ({len(pnls)}/20)\n"
        txt += f"🟢 Active: {n_active}  🟡 Pending: {n_pending}  🔴 Error/Disabled: {n_error+n_disabled}\n"
        if pending_q:
            txt += f"⏳ Retry Queue: {pending_q} OTP(s) অপেক্ষায়\n"
        txt += "\n"

        btns = []
        for pid, nm, pt, v, ecnt, dis_until, last_ok, last_err, icon, status_lbl, fwd_today in rows_info:
            pt_label = {"crapi":"CRAPI","fivesim":"5SIM","webhook":"WEBHOOK","url":"URL","apikey":"APIKEY","fastxotps":"FASTXOTPS","ivasms":"IVASMS"}.get(pt, pt.upper())
            last_ok_s = last_ok[:16] if last_ok else "কখনো না"
            txt += f"{icon} **{nm}** `[{pt_label}]`\n"
            txt += f"   ╠ **{status_lbl}**\n"
            txt += f"   ╠ শেষ সফল: `{last_ok_s}`\n"
            if last_err and "Error" in status_lbl or "Disabled" in status_lbl:
                txt += f"   ╚ ❌ `{last_err[:45]}`\n"
            else:
                d = v[:30]+"…" if len(v)>30 else v
                txt += f"   ╚ `{d}`\n"
            txt += "\n"
            btns.append([Button.inline(f"{icon} {nm}", f"opn_{pid}"),
                         Button.inline("🔄", f"rstpnl_{pid}"),
                         Button.inline("❌", f"rmpnl_{pid}")])

        if len(pnls) < 20: btns.append([Button.inline("➕ Add Panel","add_panel")])
        btns.append([Button.inline("🔴🟢 Live Ping Test","pnl_live_status"),
                     Button.inline("📊 OTP Statistics","pnl_otp_stats")])
        btns.append([Button.inline("🔙","adm_back")])
        await event.edit(txt or "📡 No panels.", buttons=btns, parse_mode='md'); return

    if data == "pnl_live_status":
        await event.edit("🔄 **Live Ping Test চলছে...**\n\nসব প্যানেল চেক করা হচ্ছে, একটু অপেক্ষা করুন ⏳",
                         buttons=[], parse_mode='md')
        c.execute("SELECT id,name,panel_type,value FROM sms_panels ORDER BY id")
        pnls = c.fetchall()
        results = []
        for pid, nm, pt, v in pnls:
            if pt == "webhook":
                results.append((nm, pt, "🟢", "Webhook — সবসময় Active", 0))
                continue
            try:
                if pt in ("fivesim", "apikey"):
                    headers = {"Authorization": f"Bearer {v.strip()}", "Accept": "application/json"}
                    r = http_requests.get(
                        "https://5sim.net/v1/user/profile",
                        headers=headers, timeout=8)
                    if r.status_code == 200:
                        balance = r.json().get("balance", "?")
                        results.append((nm, pt, "🟢", f"Connected ✅ | Balance: {balance}", r.status_code))
                    elif r.status_code == 401:
                        results.append((nm, pt, "🔴", "❌ 401 — Invalid API Key", r.status_code))
                    elif r.status_code == 429:
                        results.append((nm, pt, "🟡", "⚠️ 429 — Rate Limited", r.status_code))
                    else:
                        results.append((nm, pt, "🔴", f"❌ HTTP {r.status_code}", r.status_code))
                elif pt == "ivasms":
                    # Test cookie-based access
                    hdrs_iv = {
                        "Cookie": v.strip(),
                        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    }
                    r = http_requests.get(
                        "https://www.ivasms.com/portal/live/my_sms",
                        headers=hdrs_iv, timeout=10, allow_redirects=True)
                    if r.status_code == 200 and 'login' not in r.url:
                        results.append((nm, pt, "🟢", "Cookie valid ✅ | Page loaded", r.status_code))
                    elif 'login' in r.url:
                        results.append((nm, pt, "🔴", "❌ Cookie expired — login redirect", 0))
                    else:
                        results.append((nm, pt, "🟡", f"⚠️ HTTP {r.status_code}", r.status_code))
                elif pt == "fastxotps":
                    # Test fastxotps key
                    r = http_requests.get(
                        f"https://fastxotps.com/api/profile?api_key={v.strip()}",
                        timeout=8)
                    if r.status_code == 200:
                        results.append((nm, pt, "🟢", "Key valid ✅ | HTTP 200", r.status_code))
                    elif r.status_code == 401:
                        results.append((nm, pt, "🔴", "❌ 401 Invalid API Key", r.status_code))
                    else:
                        results.append((nm, pt, "🟡", f"⚠️ HTTP {r.status_code}", r.status_code))
                else:
                    r = http_requests.get(v, timeout=8)
                    if r.status_code == 200:
                        results.append((nm, pt, "🟢", f"Connected ✅ | HTTP 200", r.status_code))
                    else:
                        results.append((nm, pt, "🔴", f"❌ HTTP {r.status_code}", r.status_code))
            except http_requests.exceptions.ConnectionError:
                results.append((nm, pt, "🔴", "❌ Connection Refused", 0))
            except http_requests.exceptions.Timeout:
                results.append((nm, pt, "🟡", "⚠️ Timeout (8s)", 0))
            except Exception as e:
                results.append((nm, pt, "🔴", f"❌ Error: {str(e)[:40]}", 0))

        active_c  = sum(1 for r in results if r[2] == "🟢")
        error_c   = sum(1 for r in results if r[2] == "🔴")
        warning_c = sum(1 for r in results if r[2] == "🟡")
        now_s = datetime.datetime.now().strftime("%H:%M:%S")

        txt  = f"🔴🟢 **Live Ping Results** `[{now_s}]`\n"
        txt += f"🟢 Active: {active_c}  🔴 Error: {error_c}  🟡 Warning: {warning_c}\n"
        txt += "━━━━━━━━━━━━━━━━━━━━━━\n\n"
        for nm, pt, icon, status_msg, code in results:
            pt_label = {"crapi":"CRAPI","fivesim":"5SIM","webhook":"WH","url":"URL","apikey":"5SIM","fastxotps":"FASTX","ivasms":"IVASMS"}.get(pt, pt.upper())
            txt += f"{icon} **{nm}** `[{pt_label}]`\n"
            txt += f"   └ {status_msg}\n\n"

        btns = [
            [Button.inline("🔄 আবার চেক করো", "pnl_live_status")],
            [Button.inline("📊 OTP Statistics", "pnl_otp_stats")],
            [Button.inline("🔙 Panels এ ফিরে যাও", "adm_panels")],
        ]
        await event.edit(txt, buttons=btns, parse_mode='md'); return

    if data == "pnl_otp_stats":
        today_s = datetime.date.today().isoformat()
        yesterday_s = (datetime.date.today() - datetime.timedelta(days=1)).isoformat()

        # per-panel forwarded stats
        c.execute("SELECT panel_name, SUM(forwarded_count) FROM forwarded_stats GROUP BY panel_name")
        totals = {r[0]: r[1] for r in c.fetchall()}
        c.execute("SELECT panel_name, forwarded_count FROM forwarded_stats WHERE date=?", (today_s,))
        today_counts = {r[0]: r[1] for r in c.fetchall()}
        c.execute("SELECT panel_name, forwarded_count FROM forwarded_stats WHERE date=?", (yesterday_s,))
        yest_counts = {r[0]: r[1] for r in c.fetchall()}

        # all panels
        c.execute("SELECT id, name, panel_type FROM sms_panels ORDER BY id")
        pnls = c.fetchall()

        # overall otp_log total
        c.execute("SELECT COUNT(*) FROM otp_log")
        total_log = c.fetchone()[0]
        c.execute("SELECT COUNT(*) FROM otp_log WHERE received_at LIKE ?", (today_s+'%',))
        today_log = c.fetchone()[0]

        txt  = "📊 **Panel OTP Statistics**\n"
        txt += f"━━━━━━━━━━━━━━━━━━━━━━\n"
        txt += f"📅 Today: `{today_s}`\n\n"

        grand_today = 0; grand_total = 0
        for _, nm, pt in pnls:
            td = today_counts.get(nm, 0)
            yd = yest_counts.get(nm, 0)
            tot = totals.get(nm, 0)
            grand_today += td; grand_total += tot
            pt_label = {"crapi":"CRAPI","fivesim":"5SIM","webhook":"WEBHOOK","url":"URL","apikey":"APIKEY","fastxotps":"FASTXOTPS","ivasms":"IVASMS"}.get(pt, pt.upper())
            if pt == "webhook" and tot == 0:
                status_line = "⏳ অপেক্ষায় (Webhook URL configure হয়নি)"
            else:
                status_line = f"🟢 আজ: **{td}** | গতকাল: {yd} | মোট: {tot}"
            txt += f"📡 **{nm}** `[{pt_label}]`\n"
            txt += f"   └ {status_line}\n\n"

        txt += f"━━━━━━━━━━━━━━━━━━━━━━\n"
        txt += f"🔢 **সব মিলিয়ে আজ:** `{grand_today}` OTPs\n"
        txt += f"🔢 **সব মিলিয়ে মোট:** `{grand_total}` OTPs\n"
        txt += f"📋 **OTP Log Total:** `{total_log}` entries\n"

        btns = [
            [Button.inline("🔄 Refresh", "pnl_otp_stats")],
            [Button.inline("🔙 Panels", "adm_panels")],
        ]
        await event.edit(txt, buttons=btns, parse_mode='md'); return

    if data.startswith("fsave_"):
        # Force Save a panel that failed live test
        rest = data[6:]   # format: {ptype}:{name}:{val_truncated}
        # val may be truncated (max 60 chars) — we need the real val from state cache
        # fallback: save what we have
        parts = rest.split(":", 2)
        if len(parts) == 3:
            ptype, pname, val_partial = parts
        else:
            await event.answer("Invalid state — please re-add the panel.", alert=True); return
        # Find full value from recently-entered message in STATES (already cleared).
        # We stored up to 60 chars in the button data — use that as the value.
        val = val_partial.strip()
        c.execute("SELECT COUNT(*) FROM sms_panels"); cnt = c.fetchone()[0]
        if cnt >= 20:
            await event.edit("❌ Max 20 panels reached.", buttons=[[Button.inline("🔙","adm_panels")]]); return
        c.execute("INSERT INTO sms_panels(name,panel_type,value,added_at,error_count,disabled_until,last_ok,last_error) VALUES(?,?,?,?,0,'','','Added via Force Save')",
                  (pname, ptype, val, today())); db.commit()
        type_labels = {"fivesim":"🔑 5sim","crapi":"📡 CRAPI","url":"🌐 URL","fastxotps":"🔥 fastxotps","ivasms":"🍪 ivasms"}
        label = type_labels.get(ptype, ptype.upper())
        await event.edit(
            f"💾 **Force Saved!**\n\n📛 Name: `{pname}`\n🔖 Type: {label}\n\n"
            f"⚠️ Panel saved — API test failed earlier.\n"
            f"Error থাকলে Admin → Panels থেকে দেখা যাবে।",
            buttons=[[Button.inline("📡 Panels","adm_panels")]], parse_mode='md'); return

    if data.startswith("testnow_"):
        pid = int(data[8:])
        c.execute("SELECT name, panel_type, value FROM sms_panels WHERE id=?", (pid,))
        r = c.fetchone()
        if not r: await event.answer("Not found.", alert=True); return
        nm, pt, val = r
        await event.answer("⏳ Testing...", alert=False)
        test_ok = False; test_detail = ""
        try:
            if pt in ("crapi",):
                resp = http_requests.get(val, timeout=10)
                if resp.status_code == 200:
                    try:
                        d = resp.json()
                        n = len(d) if isinstance(d, list) else len(d.get('data', d.get('records', [])))
                        test_detail = f"✅ Connected — {n} record(s)"
                    except Exception:
                        test_detail = "✅ HTTP 200 OK"
                    test_ok = True
                else:
                    test_detail = f"❌ HTTP {resp.status_code} — {resp.text[:60]}"
            elif pt in ("fivesim", "apikey"):
                headers = {"Authorization": f"Bearer {val}", "Accept": "application/json"}
                resp = http_requests.get("https://5sim.net/v1/user/profile", headers=headers, timeout=10)
                if resp.status_code == 200:
                    try: uname = resp.json().get('username', '?')
                    except Exception: uname = '?'
                    test_detail = f"✅ 5sim key valid — Account: `{uname}`"
                    test_ok = True
                else:
                    test_detail = f"❌ HTTP {resp.status_code} — {resp.text[:60]}"
            else:
                resp = http_requests.get(val, timeout=10)
                test_detail = f"HTTP {resp.status_code}"; test_ok = resp.status_code == 200
        except Exception as e:
            test_detail = f"❌ Error: {str(e)[:80]}"

        if test_ok:
            c.execute("UPDATE sms_panels SET error_count=0, disabled_until='', last_error='', last_ok=? WHERE id=?",
                      (datetime.datetime.now().strftime('%H:%M:%S'), pid)); db.commit()
        else:
            c.execute("UPDATE sms_panels SET last_error=? WHERE id=?", (test_detail[:80], pid)); db.commit()

        icon = "✅" if test_ok else "❌"
        await event.edit(
            f"🧪 **Test Result: {nm}**\n\n{icon} {test_detail}\n\n"
            f"{'Panel সুস্থ — errors reset হয়েছে।' if test_ok else 'Panel-এ সমস্যা আছে।'}",
            buttons=[
                [Button.inline("🔙 Panel Info", f"opn_{pid}")],
                [Button.inline("📡 Panels", "adm_panels")]
            ], parse_mode='md'); return

    if data.startswith("rstpnl_"):
        pid = int(data[7:])
        c.execute("SELECT name FROM sms_panels WHERE id=?", (pid,)); r = c.fetchone()
        if r:
            c.execute("UPDATE sms_panels SET error_count=0, disabled_until='', last_error='' WHERE id=?", (pid,))
            db.commit()
            await event.answer(f"✅ {r[0]} reset — polling resumed!", alert=True)
        await event.edit("✅ Panel reset. Go back to view.",
                         buttons=[[Button.inline("🔙 Panels","adm_panels")]]); return

    if data == "add_panel":
        await event.edit(
            "📡 **Add SMS Panel**\n\nPanel type select করুন:",
            buttons=[
                [Button.inline("🔑 5sim API Key",      "aptype_fivesim")],
                [Button.inline("📡 CRAPI URL",          "aptype_crapi")],
                [Button.inline("🌐 JSON/URL Panel",     "aptype_url")],
                [Button.inline("🔥 fastxotps.com",      "aptype_fastxotps")],
                [Button.inline("🍪 ivasms Cookie",      "aptype_ivasms")],
                [Button.inline("🔙","adm_panels")],
            ], parse_mode='md'); return

    if data.startswith("aptype_"):
        ptype = data[7:]
        labels = {
            "fivesim":    "5sim API Key",
            "crapi":      "CRAPI URL",
            "url":        "JSON/URL Panel",
            "fastxotps":  "fastxotps.com API Key",
            "ivasms":     "ivasms.com Cookie Panel",
        }
        STATES[uid] = f"pname_{ptype}"
        if ptype == "ivasms":
            await event.edit(
                "🍪 **ivasms.com Cookie Panel**\n\n"
                "Panel name পাঠান (e.g. ivasms-main):",
                buttons=[[Button.inline("🔙","adm_panels")]], parse_mode='md'); return
        await event.edit(
            f"📡 **{labels.get(ptype, ptype)}**\n\nPanel name পাঠান (e.g. fastxotps-main):",
            buttons=[[Button.inline("🔙","adm_panels")]], parse_mode='md'); return

    if data.startswith("opn_"):
        pid = int(data[4:])
        c.execute("SELECT name,panel_type,value,error_count,disabled_until,last_ok,last_error FROM sms_panels WHERE id=?", (pid,))
        r = c.fetchone()
        if not r: await event.answer("Not found.", alert=True); return
        nm, pt, v, errcnt, dis_until, last_ok, last_err = r

        # ── Status line ─────────────────────────────────────────────────
        if dis_until:
            status_icon = "🔴"; status_line = f"🔴 **Disabled** until `{dis_until}` ({errcnt} errors)"
        elif errcnt and errcnt > 0:
            status_icon = "🟡"; status_line = f"🟡 **Warning** — {errcnt} error(s)"
        else:
            status_icon = "🟢"; status_line = f"🟢 **Active** — last OK: `{last_ok or 'N/A'}`"

        err_section = ""
        if last_err:
            err_section = f"\n⚠️ **Last Error:** `{last_err}`\n"

        type_label = {"crapi":"📡 CRAPI","fivesim":"🔑 5sim","url":"🌐 URL","webhook":"🪝 Webhook","apikey":"🔑 APIKey","fastxotps":"🔥 fastxotps","ivasms":"🍪 ivasms"}.get(pt, pt.upper())

        if pt == "ivasms":
            ck_short = v[:40]+"…" if len(v)>40 else v
            await event.edit(
                f"🍪 **{nm}** — ivasms.com Cookie Panel\n\n"
                f"{status_line}{err_section}\n"
                f"🍪 Cookie (truncated):\n`{ck_short}`\n\n"
                f"📡 Polling: প্রতি ১৫ সেকেন্ডে OTP check হচ্ছে\n"
                f"🔗 URL: `https://www.ivasms.com/portal/live/my_sms`",
                buttons=[
                    [Button.inline("🔄 Cookie Update", f"ivasms_upd_{pid}")],
                    [Button.inline("🔄 Reset Errors", f"rstpnl_{pid}"), Button.inline("🧪 Test Cookie", f"ivasms_test_{pid}")],
                    [Button.inline("❌ Remove", f"rmpnl_{pid}")],
                    [Button.inline("🔙 Panels","adm_panels")]
                ], parse_mode='md')

        elif pt == "url":
            val_short = v[:60]+"…" if len(v)>60 else v
            await event.edit(
                f"📡 **{nm}** — {type_label}\n\n"
                f"{status_line}{err_section}\n"
                f"🔗 URL:\n`{val_short}`",
                buttons=[
                    [Button.url(f"🔗 Open {nm}", v)],
                    [Button.inline("🔄 Reset Errors", f"rstpnl_{pid}"), Button.inline("❌ Remove", f"rmpnl_{pid}")],
                    [Button.inline("🔙 Panels","adm_panels")]
                ], parse_mode='md')

        elif pt == "webhook":
            domain = os.environ.get('REPLIT_DOMAINS', os.environ.get('REPLIT_DEV_DOMAIN', 'your-domain'))
            wh_url = f"https://{domain}/api/webhook/otp"
            await event.edit(
                f"📡 **{nm}** — {type_label}\n\n"
                f"🔑 **API Key:**\n`{v}`\n\n"
                f"━━━━━━━━━━━━━━━━━━\n"
                f"**Option 1 — Webhook:** Private panel-এ নিচের URL set করো:\n"
                f"`{wh_url}`\n\n"
                f"━━━━━━━━━━━━━━━━━━\n"
                f"**Option 2 — Polling:** Server URL দিলে বট নিজে poll করবে:",
                buttons=[
                    [Button.inline("🔄 Polling-এ Convert করো","conv2crapi_"+str(pid))],
                    [Button.inline("❌ Remove", f"rmpnl_{pid}")],
                    [Button.inline("🔙 Panels","adm_panels")]
                ], parse_mode='md')

        elif pt in ("crapi", "url"):
            val_short = v[:70]+"…" if len(v)>70 else v
            await event.edit(
                f"📡 **{nm}** — {type_label}\n\n"
                f"{status_line}{err_section}\n"
                f"🔗 URL:\n`{val_short}`",
                buttons=[
                    [Button.inline("🔄 Reset Errors", f"rstpnl_{pid}"), Button.inline("🧪 Test Now", f"testnow_{pid}")],
                    [Button.inline("❌ Remove", f"rmpnl_{pid}")],
                    [Button.inline("🔙 Panels","adm_panels")]
                ], parse_mode='md')

        elif pt == "fastxotps":
            fkey, fcookie = _parse_fastxotps_value(v)
            key_short = (fkey[:25]+"…" if fkey and len(fkey)>25 else fkey) or "—"
            ck_short  = (fcookie[:35]+"…" if fcookie and len(fcookie)>35 else fcookie) or "❌ Cookie নেই (DDoS block হবে!)"
            cookie_warn = "" if fcookie else "\n\n⚠️ **Cookie নেই!** fastxotps DDoS protection-এর কারণে API কাজ করছে না। নিচের বাটন দিয়ে Cookie যোগ করুন।"
            await event.edit(
                f"🔥 **{nm}** — fastxotps Panel\n\n"
                f"{status_line}{err_section}\n"
                f"🔑 API Key: `{key_short}`\n"
                f"🍪 DDoS Cookie: `{ck_short}`{cookie_warn}",
                buttons=[
                    [Button.inline("🍪 Cookie যোগ/Update", f"fastxotps_upd_{pid}")],
                    [Button.inline("🔄 Reset Errors", f"rstpnl_{pid}"), Button.inline("🧪 Test", f"fastxotps_test_{pid}")],
                    [Button.inline("❌ Remove", f"rmpnl_{pid}")],
                    [Button.inline("🔙 Panels","adm_panels")]
                ], parse_mode='md')

        elif pt in ("fivesim", "apikey"):
            val_short = v[:30]+"…" if len(v)>30 else v
            await event.edit(
                f"📡 **{nm}** — {type_label}\n\n"
                f"{status_line}{err_section}\n"
                f"🔑 API Key:\n`{val_short}`",
                buttons=[
                    [Button.inline("🔄 Reset Errors", f"rstpnl_{pid}"), Button.inline("🧪 Test Now", f"testnow_{pid}")],
                    [Button.inline("❌ Remove", f"rmpnl_{pid}")],
                    [Button.inline("🔙 Panels","adm_panels")]
                ], parse_mode='md')

        else:
            await event.edit(
                f"📡 **{nm}** — {type_label}\n\n"
                f"{status_line}{err_section}\n"
                f"🔑 Value:\n`{v[:60]}`",
                buttons=[
                    [Button.inline("🔄 Reset Errors", f"rstpnl_{pid}"), Button.inline("❌ Remove", f"rmpnl_{pid}")],
                    [Button.inline("🔙 Panels","adm_panels")]
                ], parse_mode='md')
        return

    if data.startswith("ivasms_upd_"):
        pid = int(data[11:])
        c.execute("SELECT name FROM sms_panels WHERE id=?", (pid,)); r = c.fetchone()
        if not r: await event.answer("Not found.", alert=True); return
        nm = r[0]
        STATES[uid] = f"ivasms_newcookie:{pid}"
        await event.edit(
            f"🍪 **Cookie Update: {nm}**\n\n"
            f"নতুন cookie পাঠান:\n\n"
            f"**কীভাবে Cookie নিবেন:**\n"
            f"1. Chrome → ivasms.com → Login করুন\n"
            f"2. F12 → Network tab → যেকোনো request click\n"
            f"3. **Request Headers** → **Cookie** value copy করুন\n"
            f"4. এখানে paste করুন\n\n"
            f"অথবা:\n"
            f"F12 → **Application** → **Cookies** → `www.ivasms.com`\n"
            f"সব row-এর name=value; এই format-এ copy করুন।\n\n"
            f"Format: `sessionid=xxx; csrftoken=yyy; ...`",
            buttons=[[Button.inline("❌ Cancel", f"opn_{pid}")]], parse_mode='md')
        return

    if data.startswith("ivasms_test_"):
        pid = int(data[12:])
        c.execute("SELECT name, value FROM sms_panels WHERE id=?", (pid,)); r = c.fetchone()
        if not r: await event.answer("Not found.", alert=True); return
        nm, cookie_str = r
        await event.answer("⏳ Testing cookie...", alert=False)
        try:
            loop = asyncio.get_event_loop()
            def _test_iv():
                hdrs = {
                    "Cookie": cookie_str.strip(),
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Accept": "text/html,*/*;q=0.8",
                }
                rv = http_requests.get(
                    "https://www.ivasms.com/portal/live/my_sms",
                    headers=hdrs, timeout=12, allow_redirects=True)
                return rv.status_code, rv.url
            sc, final_url = await loop.run_in_executor(None, _test_iv)
            if sc == 200 and 'login' not in str(final_url):
                res = f"✅ Cookie VALID — HTTP {sc}\nPage loaded successfully!"
                c.execute("UPDATE sms_panels SET error_count=0, disabled_until='', last_error='' WHERE id=?", (pid,)); db.commit()
            elif 'login' in str(final_url):
                res = "❌ Cookie EXPIRED — Login page-এ redirect হচ্ছে।\nনতুন cookie দিন।"
            elif sc == 403:
                res = f"⚠️ HTTP 403 — Cloudflare block। Cookie সঠিক কিন্তু IP block।"
            else:
                res = f"⚠️ HTTP {sc} — {str(final_url)[:60]}"
        except Exception as e:
            res = f"❌ Error: {str(e)[:80]}"
        await event.edit(
            f"🍪 **{nm}** — Cookie Test\n\n{res}",
            buttons=[
                [Button.inline("🔄 Cookie Update", f"ivasms_upd_{pid}")],
                [Button.inline("🔙 Panel", f"opn_{pid}")],
            ], parse_mode='md')
        return

    if data.startswith("fastxotps_upd_"):
        pid = int(data[14:])
        c.execute("SELECT name, value FROM sms_panels WHERE id=?", (pid,)); r = c.fetchone()
        if not r: await event.answer("Not found.", alert=True); return
        nm, cur_val = r
        fkey, _ = _parse_fastxotps_value(cur_val)
        key_info = f"`{fkey}`" if fkey else "—"
        STATES[uid] = f"fastxotps_newcookie:{pid}"
        await event.edit(
            f"🍪 **fastxotps Cookie Update: {nm}**\n\n"
            f"🔑 API Key (unchanged): {key_info}\n\n"
            f"**কীভাবে Cookie নিবেন:**\n"
            f"1. Chrome-এ **fastxotps.com** যান → Login করুন\n"
            f"2. Site load হলে **F12** চাপুন → **Network** tab\n"
            f"3. যেকোনো request-এ click → **Request Headers** → **Cookie** value copy করুন\n"
            f"4. এখানে paste করুন\n\n"
            f"অথবা: F12 → **Application** → **Cookies** → `fastxotps.com`\n"
            f"সব cookie name=value; format-এ copy করুন\n\n"
            f"⚠️ **Important:** fastxotps.com-এর cookie + tzddos verification cookie\n"
            f"দুটোই copy করতে হবে। Format:\n"
            f"`PHPSESSID=xxx; tzddos_verified=yyy; ...`",
            buttons=[[Button.inline("❌ Cancel", f"opn_{pid}")]], parse_mode='md')
        return

    if data.startswith("fastxotps_test_"):
        pid = int(data[15:])
        c.execute("SELECT name, value FROM sms_panels WHERE id=?", (pid,)); r = c.fetchone()
        if not r: await event.answer("Not found.", alert=True); return
        nm, raw_val = r
        fkey, fcookie = _parse_fastxotps_value(raw_val)
        await event.answer("⏳ Testing...", alert=False)
        if not fcookie:
            await event.edit(
                f"🔥 **{nm}** — Test\n\n"
                f"❌ Cookie নেই! DDoS protection bypass করতে cookie লাগবে।\n\n"
                f"👇 Cookie Update বাটন দিয়ে cookie দিন।",
                buttons=[
                    [Button.inline("🍪 Cookie যোগ করুন", f"fastxotps_upd_{pid}")],
                    [Button.inline("🔙 Panel", f"opn_{pid}")],
                ], parse_mode='md')
            return
        try:
            loop = asyncio.get_event_loop()
            def _test_fx():
                hdrs = {
                    "Cookie": fcookie,
                    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                    "Referer": "https://fastxotps.com/",
                }
                ep = (f"https://fastxotps.com/api/sms?api_key={fkey}" if fkey
                      else "https://fastxotps.com/dashboard")
                rv = http_requests.get(ep, headers=hdrs, timeout=12, allow_redirects=True)
                return rv.status_code, str(rv.url), rv.text[:200]
            sc, final_url, body_preview = await loop.run_in_executor(None, _test_fx)
            if sc == 200 and 'tzddos' not in final_url and 'login' not in final_url:
                res = f"✅ Cookie VALID — HTTP {sc}\nDDoS bypass সফল! OTP polling চলছে।"
                c.execute("UPDATE sms_panels SET error_count=0, disabled_until='', last_error='' WHERE id=?", (pid,)); db.commit()
            elif 'tzddos' in final_url:
                res = "❌ DDoS Cookie INVALID — Verification page-এ redirect হচ্ছে।\nনতুন cookie দিন (tzddos_verified cookie সহ)।"
            elif 'login' in final_url:
                res = "❌ Session Cookie EXPIRED — Login redirect হচ্ছে।\nনতুন cookie দিন।"
            else:
                res = f"⚠️ HTTP {sc}\nURL: {final_url[:80]}\nBody: {body_preview[:80]}"
        except Exception as e:
            res = f"❌ Error: {str(e)[:80]}"
        await event.edit(
            f"🔥 **{nm}** — fastxotps Cookie Test\n\n{res}",
            buttons=[
                [Button.inline("🍪 Cookie Update", f"fastxotps_upd_{pid}")],
                [Button.inline("🔙 Panel", f"opn_{pid}")],
            ], parse_mode='md')
        return

    if data.startswith("rmpnl_"):
        pid = int(data[6:])
        c.execute("SELECT name FROM sms_panels WHERE id=?", (pid,)); r = c.fetchone()
        if r:
            c.execute("DELETE FROM sms_panels WHERE id=?", (pid,)); db.commit()
            await event.answer(f"✅ {r[0]} removed!", alert=False)
        c.execute("SELECT id,name,panel_type,value FROM sms_panels ORDER BY id"); pnls = c.fetchall()
        txt = f"📡 **SMS Panels** ({len(pnls)}/20)\n\n"; btns = []
        for pid2, nm2, pt2, v2 in pnls:
            d = v2[:38]+"…" if len(v2)>38 else v2
            txt += f"🔹 **{nm2}** `{pt2.upper()}`\n`{d}`\n\n"
            btns.append([Button.inline(f"🔗 {nm2}", f"opn_{pid2}"),
                         Button.inline("❌", f"rmpnl_{pid2}")])
        if len(pnls) < 20: btns.append([Button.inline("➕ Add Panel","add_panel")])
        btns.append([Button.inline("🔙","adm_back")])
        await event.edit(txt or "📡 No panels.", buttons=btns, parse_mode='md'); return

    if data.startswith("conv2crapi_"):
        pid = int(data[11:])
        c.execute("SELECT name, value FROM sms_panels WHERE id=?", (pid,)); r = c.fetchone()
        if not r: await event.answer("Not found.", alert=True); return
        nm, key = r
        STATES[uid] = f"crapi_url_{pid}"
        await event.edit(
            f"🔄 **Polling-এ Convert: {nm}**\n\n"
            f"🔑 Key: `{key[:30]}...`\n\n"
            f"এই key-টা যে server-এ কাজ করে তার **base URL** পাঠাও:\n"
            f"উদাহরণ:\n"
            f"• `http://198.135.52.36`\n"
            f"• `http://45.12.34.56`\n"
            f"• `https://sms.example.com`\n\n"
            f"বট নিজে নিচের format-এ try করবে:\n"
            f"`{{server}}/crapi/ks/viewstats?token={{key}}&records=100`\n\n"
            f"⚠️ যদি সঠিক server না জানো, যে জায়গা থেকে key নিয়েছ তাদের জিজ্ঞেস করো।",
            buttons=[[Button.inline("❌ Cancel", f"opn_{pid}")]], parse_mode='md')
        return

    if data == "adm_countries":
        await event.edit("🌍 **Country Management**", buttons=[
            [Button.inline("🌐 World List (Toggle)","world_0")],
            [Button.inline("📋 Active List","list_c")],
            [Button.inline("🔙","adm_back")]], parse_mode='md'); return

    if data.startswith("world_"):
        pg = int(data[6:]); per = 12; st = pg*per; chunk = COUNTRIES[st:st+per]
        c.execute("SELECT short_name FROM active_countries"); active = {r[0] for r in c.fetchall()}
        btns = []; row = []
        for cn, sh, fl in chunk:
            on = sh in active
            nm = cn[:9] if len(cn)>9 else cn
            row.append(Button.inline(f"{'✅' if on else ''}{fl} {nm} [{sh.upper()}]", f"tgl_{sh}"))
            if len(row)==2: btns.append(row); row=[]
        if row: btns.append(row)
        nav = []
        if pg > 0:          nav.append(Button.inline("◀️ Prev", f"world_{pg-1}"))
        if st+per<len(COUNTRIES): nav.append(Button.inline("Next ▶️", f"world_{pg+1}"))
        if nav: btns.append(nav)
        btns.append([Button.inline("🔙","adm_countries")])
        tp = (len(COUNTRIES)+per-1)//per
        await event.edit(f"🌍 Countries — Page {pg+1}/{tp}  (✅ = Active, tap to toggle)",
                         buttons=btns); return

    if data.startswith("tgl_"):
        sh = data[4:]
        c.execute("SELECT 1 FROM active_countries WHERE short_name=?", (sh,))
        if c.fetchone():
            c.execute("DELETE FROM active_countries WHERE short_name=?", (sh,))
            c.execute("DELETE FROM premium_stock WHERE country=?", (sh,))
            db.commit(); await event.answer(f"✅ {sh.upper()} deactivated", alert=False)
        else:
            info = COUNTRY_MAP.get(sh)
            if info:
                c.execute("INSERT OR IGNORE INTO active_countries VALUES(?,?,?)",
                          (info[0], sh, info[1])); db.commit()
                await event.answer(f"✅ {info[1]} {info[0]} activated!", alert=False)
        idx = next((i for i,(_, s, _) in enumerate(COUNTRIES) if s==sh), 0)
        pg = idx//12; per = 12; st = pg*per; chunk = COUNTRIES[st:st+per]
        c.execute("SELECT short_name FROM active_countries"); active = {r[0] for r in c.fetchall()}
        btns = []; row = []
        for cn, s2, fl in chunk:
            on = s2 in active
            nm = cn[:9] if len(cn)>9 else cn
            row.append(Button.inline(f"{'✅' if on else ''}{fl} {nm} [{s2.upper()}]", f"tgl_{s2}"))
            if len(row)==2: btns.append(row); row=[]
        if row: btns.append(row)
        nav = []
        if pg > 0:            nav.append(Button.inline("◀️ Prev", f"world_{pg-1}"))
        if st+per<len(COUNTRIES): nav.append(Button.inline("Next ▶️", f"world_{pg+1}"))
        if nav: btns.append(nav)
        btns.append([Button.inline("🔙","adm_countries")])
        tp = (len(COUNTRIES)+per-1)//per
        await event.edit(f"🌍 Countries — Page {pg+1}/{tp}  (✅ = Active):", buttons=btns); return

    if data == "list_c":
        c.execute("SELECT country_name,short_name,flag FROM active_countries ORDER BY country_name")
        rows = c.fetchall(); txt = "📋 **Active Countries:**\n\n"; btns = []
        for cn, sh, fl in rows:
            c.execute("SELECT COUNT(*) FROM premium_stock WHERE country=? AND status=0", (sh,))
            cnt_val = c.fetchone()[0]
            txt += f"{fl} {cn} [{sh.upper()}] — **{cnt_val}**\n"
            btns.append([Button.inline(f"❌ {fl} {cn} [{sh.upper()}]", f"delc_{sh}")])
        if not rows: txt += "❌ None active."
        btns.append([Button.inline("🔙","adm_countries")])
        await event.edit(txt, buttons=btns, parse_mode='md'); return

    if data.startswith("delc_"):
        sh = data[5:]
        c.execute("DELETE FROM active_countries WHERE short_name=?", (sh,))
        c.execute("DELETE FROM premium_stock WHERE country=?", (sh,)); db.commit()
        await event.answer(f"✅ {sh.upper()} removed!", alert=False)
        c.execute("SELECT country_name,short_name,flag FROM active_countries ORDER BY country_name")
        rows = c.fetchall(); txt = "📋 **Active Countries:**\n\n"; btns = []
        for cn, sh2, fl in rows:
            c.execute("SELECT COUNT(*) FROM premium_stock WHERE country=? AND status=0", (sh2,))
            cnt_val = c.fetchone()[0]
            txt += f"{fl} {cn} [{sh2.upper()}] — **{cnt_val}**\n"
            btns.append([Button.inline(f"❌ {fl} {cn} [{sh2.upper()}]", f"delc_{sh2}")])
        if not rows: txt += "❌ None active."
        btns.append([Button.inline("🔙","adm_countries")])
        await event.edit(txt, buttons=btns, parse_mode='md'); return

    if data == "adm_upload":
        await event.edit("📁 Select service:", buttons=[
            [Button.inline("💬 WhatsApp","up_whatsapp"), Button.inline("🔹 Telegram","up_telegram")],
            [Button.inline("🎵 TikTok",  "up_tiktok"),  Button.inline("🌐 Facebook","up_facebook")],
            [Button.inline("📸 Instagram","up_instagram")],
            [Button.inline("🔙","adm_back")]]); return

    if data.startswith("up_") and not data.startswith("upc_"):
        svc = data[3:]
        c.execute("SELECT country_name,short_name,flag FROM active_countries ORDER BY country_name")
        rows = c.fetchall()
        if not rows:
            await event.answer("❌ No active countries! Go to Countries first.", alert=True); return
        btns = []; row = []
        for cn, sh, fl in rows:
            nm = cn[:8] if len(cn)>8 else cn
            row.append(Button.inline(f"{fl} {nm} [{sh.upper()}]", f"upc_{svc}_{sh}"))
            if len(row)==2: btns.append(row); row=[]
        if row: btns.append(row)
        btns.append([Button.inline("🔙","adm_upload")])
        await event.edit(f"📁 {SVC.get(svc,svc)} — Select Country:", buttons=btns); return

    if data.startswith("upc_"):
        rest = data[4:]; pts = rest.split("_",1); svc = pts[0]; sh = pts[1] if len(pts)>1 else "eg"
        STATES[uid] = f"up_{sh}_{svc}"
        cn, cf = ctry(sh)
        await event.edit(
            f"📥 **Upload Numbers**\n{cf} **{cn}** [{sh.upper()}] — {SVC.get(svc,svc)}\n\n"
            f"✅ Send **.txt** file (one number per line)\n"
            f"✅ Send **.xlsx** file (one number per cell in Excel)",
            buttons=[[Button.inline("🔙","adm_upload")]], parse_mode='md'); return

    if data == "adm_code":
        try:
            sz    = os.path.getsize(BOT_FILE)
            with open(BOT_FILE,'r',encoding='utf-8') as f: lines = f.read().count('\n')
            mt    = datetime.datetime.fromtimestamp(os.path.getmtime(BOT_FILE)).strftime("%Y-%m-%d %H:%M")
        except: sz=0; lines=0; mt="N/A"
        await event.edit(
            f"💻 **My Bot Code**\n\n"
            f"📄 main.py | 📏 {lines} lines | 💾 {sz//1024} KB\n"
            f"🕐 Modified: {mt}\n\n"
            f"✅ **Auto-updated** — always downloads latest code!\n\n"
            f"Features: Force Sub • 5 Services • 100+ Countries\n"
            f"SMS Panels (20) • Multi-Admin (5) • TXT+XLSX Upload\n"
            f"@username Admin • OTP Webhook • AI Assistant",
            buttons=[
                [Button.inline("📄 Download Code","code_dl")],
                [Button.inline("✏️ Upload New Code (.py)","code_edit")],
                [Button.inline("🔄 Restart Bot","code_restart")],
                [Button.inline("⏹ Stop Bot","code_stop")],
                [Button.inline("🔙","adm_back")],
            ], parse_mode='md'); return

    if data == "code_dl":
        await event.answer("📄 Sending file...", alert=False)
        try:
            await bot.send_file(uid, BOT_FILE,
                caption=f"💻 **SMS MATRIX BOT Code**\n📅 {today()}\n📏 Complete — always up to date!",
                parse_mode='md')
            await event.edit("✅ Code file sent! Check above ☝️",
                             buttons=[[Button.inline("🔙","adm_code")]])
        except Exception as e:
            await event.edit(f"❌ Error: {e}", buttons=[[Button.inline("🔙","adm_code")]])
        return

    if data == "code_edit":
        STATES[uid] = "new_code"
        await event.edit(
            "✏️ **Edit Bot Code**\n\nSend updated .py file.\nBot will auto-restart after save.\n\n"
            "⚠️ Caution: wrong code may stop the bot!",
            buttons=[[Button.inline("❌ Cancel","adm_code")]], parse_mode='md'); return

    if data == "code_restart":
        await event.edit("🔄 Restarting bot...", buttons=[])
        time.sleep(1); os.execv(sys.executable, [sys.executable]+sys.argv)

    if data == "code_stop":
        await event.edit("⏹ Bot stopping. Workflow will restart automatically.", buttons=[])
        time.sleep(2); sys.exit(0)

    if data == "adm_keytest":
        STATES.pop(uid, None)
        STATES[uid] = "keytest"
        await event.edit(
            "🔑 **Universal API Key Tester**\n\n"
            "**Option 1 — শুধু Key পাঠান:**\n"
            "বট স্বয়ংক্রিয়ভাবে এই সাইটগুলোতে test করবে:\n"
            "• 5sim.net • sms-activate.org\n"
            "• smshub.org • grizzlysms.com\n"
            "• onlinesim.io • fastxotps.com\n\n"
            "**Option 2 — যেকোনো Panel Site test করতে:**\n"
            "URL এবং Key একসাথে পাঠান:\n"
            "`https://yourpanel.com YOURKEY`\n\n"
            "**Option 3 — পুরো API URL পাঠান:**\n"
            "`https://panel.com/api?api_key=YOURKEY`\n\n"
            "⏳ Testing সাধারণত ৫-২০ সেকেন্ড লাগে।",
            buttons=[
                [Button.inline("🌐 Custom Panel URL Test","adm_keytest_custom")],
                [Button.inline("❌ Cancel","adm_back")],
            ], parse_mode='md'); return

    if data == "adm_keytest_custom":
        STATES.pop(uid, None)
        STATES[uid] = "keytest"
        await event.edit(
            "🌐 **Custom Panel URL Tester**\n\n"
            "আপনার panel URL এবং API Key একসাথে পাঠান:\n\n"
            "**Format:**\n"
            "`https://yourpanel.com YOURKEY`\n\n"
            "**উদাহরণ:**\n"
            "`https://fastxotps.com MURAD_2488...`\n"
            "`https://mypanel.net abc123key`\n\n"
            "বট সব common API endpoint test করবে।",
            buttons=[[Button.inline("❌ Cancel","adm_back")]], parse_mode='md'); return

    if data == "adm_ai":
        STATES.pop(uid, None)
        key = gset('ai_api_key','')
        ks = "✅ Anthropic API set" if key else "⚠️ No API key — built-in KB active"
        await event.edit(
            f"🤖 **AI Bot Assistant**\n\nStatus: {ks}\n\n"
            f"Ask about: uptime • upload • country\npanel • webhook • code • admin • restart",
            buttons=[
                [Button.inline("💬 Ask a Question","ai_chat")],
                [Button.inline("🔑 Set Anthropic API Key","ai_setkey")],
                [Button.inline("🔙","adm_back")],
            ], parse_mode='md'); return

    if data == "ai_setkey":
        STATES[uid] = "set_ai_key"
        await event.edit("🔑 Send Anthropic API Key (https://console.anthropic.com):",
                         buttons=[[Button.inline("🔙","adm_ai")]]); return

    if data == "ai_chat":
        STATES[uid] = "ai"
        await event.edit("💬 **AI Chat Active**\n\nType your question...\n_/start to exit_",
                         buttons=[[Button.inline("❌ Exit","adm_ai")]]); return

# ════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════
import asyncio

async def retry_failed_otps():
    """Background: every 45s retry OTPs that failed to forward. Max 10 retries per OTP."""
    await asyncio.sleep(30)
    while True:
        try:
            _rdb = sqlite3.connect(DB, check_same_thread=False)
            _rc  = _rdb.cursor()
            _rc.execute(
                "SELECT id, phone, raw_msg, sender FROM failed_otp_queue "
                "WHERE retry_count < 10 ORDER BY id LIMIT 20")
            rows = _rc.fetchall()
            _rdb.close()
            for qid, phone, raw_msg, sender in rows:
                tg_url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
                ok = False
                try:
                    r = http_requests.post(tg_url, json={
                        "chat_id": OTP_GROUP,
                        "text": f"🔁 *RETRY OTP*\n📱 `{phone}`\n📩 `{raw_msg}`",
                        "parse_mode": "Markdown"
                    }, timeout=12)
                    ok = r.status_code == 200
                except Exception:
                    pass
                _rdb2 = sqlite3.connect(DB, check_same_thread=False)
                if ok:
                    _rdb2.execute("DELETE FROM failed_otp_queue WHERE id=?", (qid,))
                    print(f"[OTP-RETRY] ✅ Re-sent queued OTP id={qid}")
                else:
                    _rdb2.execute(
                        "UPDATE failed_otp_queue SET retry_count=retry_count+1, last_try=? WHERE id=?",
                        (str(datetime.datetime.now()), qid))
                _rdb2.commit(); _rdb2.close()
        except Exception as e:
            print(f"[OTP-RETRY] Error: {e}")
        await asyncio.sleep(45)


async def _supervised(coro_fn, name, delay=5):
    """Restart a coroutine whenever it exits or crashes — keeps pollers alive forever."""
    while True:
        try:
            await coro_fn()
        except Exception as e:
            print(f"[SUPERVISOR] '{name}' crashed: {e} — restarting in {delay}s")
        except asyncio.CancelledError:
            print(f"[SUPERVISOR] '{name}' cancelled.")
            break
        await asyncio.sleep(delay)


async def main():
    await set_commands()
    # Supervised pollers — auto-restart on any crash
    asyncio.ensure_future(_supervised(poll_crapi_panels,    "CRAPI-Poller",    delay=3))
    asyncio.ensure_future(_supervised(poll_fivesim_panels, "5sim-Poller",     delay=5))
    asyncio.ensure_future(_supervised(poll_url_panels,     "URL-Poller",      delay=5))
    asyncio.ensure_future(_supervised(poll_fastxotps_panels,"fastxotps-Poller",delay=10))
    asyncio.ensure_future(_supervised(poll_ivasms_panels,  "ivasms-Poller",   delay=10))
    asyncio.ensure_future(_supervised(retry_failed_otps,   "OTP-Retry",       delay=10))
    print("✅ SMS MATRIX BOT is online!")

    # ── Auto-reconnect loop ────────────────────────────────────────────────
    # run_until_disconnected() returns (or raises) when Telegram drops the connection.
    # We catch it and reconnect immediately so OTP forwarding never stops.
    reconnect_delay = 5
    while True:
        try:
            await bot.run_until_disconnected()
            print("[BOT] Disconnected — reconnecting in 5s…")
        except Exception as e:
            print(f"[BOT] Connection error: {e} — reconnecting in {reconnect_delay}s…")
        await asyncio.sleep(reconnect_delay)
        try:
            await bot.connect()
            print("[BOT] Reconnected ✅")
        except Exception as ce:
            print(f"[BOT] Reconnect failed: {ce} — will retry…")


if __name__ == '__main__':
    bot.loop.run_until_complete(main())
