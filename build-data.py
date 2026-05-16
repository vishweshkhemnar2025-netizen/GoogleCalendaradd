#!/usr/bin/env python3
"""
Build calendars.json (cleartext) and students.enc.json (ENCRYPTED) for the
Term IV calendar page.

Privacy model
-------------
The page is a static file with no backend, so any data it can read, a visitor
can also read. Therefore each student's record (name + courses) is encrypted
with a key derived from their OWN roll number + 4-digit PIN. The shipped
students.enc.json contains NO readable names, courses, PINs or emails - only
ciphertext indexed by a salted hash of the roll number.

Crypto (all browser-native via WebCrypto on decrypt):
  key64 = PBKDF2-HMAC-SHA256( "roll:pin", saltPerStudent, ITERATIONS, 64 bytes )
  kEnc  = key64[:32]   kMac = key64[32:]
  ct    = AES-256-CTR( kEnc, counter=IV(16B), plaintext=JSON(record) )
  tag   = HMAC-SHA256( kMac, IV || ct )                # encrypt-then-MAC
  recordId = SHA-256( lookupSalt || roll )             # lookup handle only

Honest limitation: a 4-digit PIN is 10,000 possibilities and roll numbers are
enumerable, so a determined, technically skilled attacker could brute-force one
targeted student's record offline. This stops casual snooping and bulk
scraping (you cannot just open the file and read the roster); it is NOT
bulletproof. Real protection needs a server, which is out of scope.

Inputs
------
1. calendars.template.csv              course_code,display_name,calendar_id
2. ENROLL_DIR/*.xlsx                    one per course-section, named exactly as
                                        course_code; sheet has Roll No./Name/Email
3. PASSWORDS_JS                         the SR-elections password file
                                        ("email": "1234") - SENSITIVE, never commit

Outputs (next to this script)
-----------------------------
   calendars.json        38 calendars: key -> { name, calId }   (safe to commit)
   students.enc.json      encrypted student bundle               (safe to commit)

Never commit: the PASSWORDS file, or any cleartext enrollment file.

Usage:  python3 build-data.py ["/path/to/Enrollment details"] ["/path/to/PASSWORDS_new.js"]
"""

import base64
import csv
import hashlib
import hmac
import json
import os
import re
import sys
from collections import OrderedDict
from pathlib import Path

import openpyxl

# ---------------------------------------------------------------------------
# CONFIG
# ---------------------------------------------------------------------------
HERE = Path(__file__).parent
CALENDARS_CSV = HERE / "calendars.template.csv"
ENROLL_DIR = Path(sys.argv[1]) if len(sys.argv) > 1 else (HERE.parent / "Enrollment details")
PASSWORDS_JS = Path(sys.argv[2]) if len(sys.argv) > 2 else (ENROLL_DIR / "script" / "PASSWORDS_new.js")

OUT_CALENDARS = HERE / "calendars.json"
OUT_STUDENTS = HERE / "students.enc.json"

PBKDF2_ITERATIONS = 200_000      # also hard-coded in index.html - keep in sync

# ---------------------------------------------------------------------------
# The 38 Google Calendar IDs. ACAD edits these each term.
# Key = slug(course_code).  Calendar settings -> Integrate calendar -> Calendar ID.
# ---------------------------------------------------------------------------
CALENDAR_IDS = {
    'B2B_M_SecA': 'c_5628375a0e79765e81c15da372de8b9186cbfc8c73412d120078aae0b599c36b@group.calendar.google.com',
    'B2B_M_SecB': 'c_d96a6ac8539d1d484fac51a68bece9e2261a0296512ebbdaf3605dbd12417990@group.calendar.google.com',
    'B2B_M_SecC': 'c_e86aaccd94a101fdf0c4bd1c1e9a06697453fe211e12a9304d50718fd3c8982a@group.calendar.google.com',
    'BM_SecA':    'c_01f1ef38f3d8da76339cbcd6505c9758e2f8b340280d686f7bd7bb648dc0ffb0@group.calendar.google.com',
    'BM_SecB':    'c_6288cb3281a8fb44873971abc787f70d2f0e9e1e9561a02bcea908e13c3a1e68@group.calendar.google.com',
    'CB_SecA':    'c_5a8afd8f03f3017eed0f36d813cb2f90e604e50f970c6a87cd611f42224e9c63@group.calendar.google.com',
    'CB_SecB':    'c_3c67e353459118f6bff4a4dc3c5dbe13d1e5e93c479aaac30bc050ba95a8e199@group.calendar.google.com',
    'CV':         'c_37065ae030f7b74036b46184db3712d2dfcdcb389ecfde6c255e420d409a5e61@group.calendar.google.com',
    'DSDT_SecA':  'c_4e429deea3e49efeed3d8dc9d667b918555348e99564f7663e95a3a001409bb8@group.calendar.google.com',
    'DSDT_SecB':  'c_0ca3d3e0399cf33f2ce6bf48fcdeb452637a76b22c9925b01b822f6f6fb869bd@group.calendar.google.com',
    'DSDT_SecC':  'c_f51547fd0c78aee3aed11ee20d2cca6490743d95a69f5e7d6c081d93b67ac0eb@group.calendar.google.com',
    'FD_SecA':    'c_c9f228bb1a16ec0f06c64b6b5ec336a58fc58f4e3b2150e44841dd7fefd8cc36@group.calendar.google.com',
    'FD_SecB':    'c_5a8c3f42cd6add3691460179bfd9a9f5e003e6e485eb6906b67d9246ec690d09@group.calendar.google.com',
    'FSA_SecA':   'c_85e2de9cd5600c2226122b120824d2e53130da69ce05cbc9cc8167451355f4e9@group.calendar.google.com',
    'FSA_SecB':   'c_473ed0a5e107dd892122bbb75cc4bba4744438bbe30a7a27410d476f209fd92b@group.calendar.google.com',
    'HRM_IR':     'c_f85fc819f6e26abdbea3f0aac90e5796d6fdb8ef1ab77d1e0cec49a4d0abcf78@group.calendar.google.com',
    'INV_SecA':   'c_d4d30791a48d84133d90fb67c7f70c9d76208a12555e827921f609ad3a93434a@group.calendar.google.com',
    'INV_SecB':   'c_5f08704f2e46d13f013941395f10c04cd9f01317c119d453358073074887b107@group.calendar.google.com',
    'MG_SecA':    'c_8df3ab99be903dfda41e14aad4c192c1df647f7c9c86f07030346c6374b5ec73@group.calendar.google.com',
    'MG_SecB':    'c_15135148b62b579e1d79b20d1c158dfddcdcd310b8c9f1c0331e890ae2e30c3c@group.calendar.google.com',
    'MSAIC_SecA': 'c_e8392a03d7baf699e37377899aefafe60aa32d48427699877aa8791a05bab893@group.calendar.google.com',
    'MSAIC_SecB': 'c_c01ba4eab309387d9a944bf6113b1a0e8376af277c2c8dc3c9a104ca412338d0@group.calendar.google.com',
    'MoB':        'c_f8251bb6f8fffbfa77d3c6816ac86075e51cea95f0b2735fb450acc8b03e41d9@group.calendar.google.com',
    'PA_SecA':    'c_c2814b31f6e703dfdad71bf7869121690d217ff3b9d92502dc1825153efe0a9c@group.calendar.google.com',
    'PA_SecB':    'c_59053100bac60669e2e98218576138dccd36a2a3bfbcb83990c4fd92894ac5ec@group.calendar.google.com',
    'PCG_SecA':   'c_9f5fe19b100d107511bef20eff6ad0de5741a6e167c027ddc605034ef73529d7@group.calendar.google.com',
    'PCG_SecB':   'c_c8a8aedad08f3c0a2b412645d7b2cdf6f0a06dc502e34a6653e32bf00932662d@group.calendar.google.com',
    'PEVC_SecA':  'c_9abec2f6e5e96316ec96fe732cde597ead16cebe3f42452ccb44fe6ef1540082@group.calendar.google.com',
    'PEVC_SecB':  'c_afbfe0e6b20e80ab43b3f39f82edc541978cf94f719d5df3c3b715d27bd16b2a@group.calendar.google.com',
    'PM':         'c_f39ff0d8b5daacd5e1226b26f9ffeefb6f3e1c92e3eef5fa1c916d47becf39e4@group.calendar.google.com',
    'PSM_SecA':   'c_6517d9726bf27f11badb2f472a7fa9ec4418c2f125a01a3c8a320fd1f9372d50@group.calendar.google.com',
    'PSM_SecB':   'c_6b549d0f97df88fe7f6b99f0d355dfe5d2874c12e9bd6965db500054534e9cfd@group.calendar.google.com',
    'Rev_Mgmt':   'c_353b39ea152788830d4d55a2fa963913f24cd021fbd9486c1a573e94b76fd521@group.calendar.google.com',
    'SCM_SecA':   'c_f4b0e34879e5e8332901e835612d43534c08bd0b6fd502e02792173e46069b6e@group.calendar.google.com',
    'SCM_SecB':   'c_9463410989735c9bf7ce10420ac3afbadee693b72d1f15808941c4fdf6025e5a@group.calendar.google.com',
    'SDM':        'c_ba77c76f5248eb9a02ff84130a6322c8feec297f61ceb80a3621e0dd526e7f2e@group.calendar.google.com',
    'SRC':        'c_5df13e3685dcc307447ca7cd9c8069afde9038d9b3498bf7aafbb55e1d5e85eb@group.calendar.google.com',
    'TS_ADR':     'c_05ff81e3c091d91e777ddd3243433016add2be587a4b52ee91d7106e724b09e2@group.calendar.google.com',
}

SECTION_SUFFIX = {" (A)": "_SecA", " (B)": "_SecB", " (C)": "_SecC"}


def slug(course_code: str) -> str:
    base, suffix = course_code, ""
    for raw, key_suffix in SECTION_SUFFIX.items():
        if base.endswith(raw):
            base, suffix = base[: -len(raw)], key_suffix
            break
    return re.sub(r"[^A-Za-z0-9]+", "_", base).strip("_") + suffix


# ===========================================================================
# Minimal, self-tested AES-256 (forward direction only - enough for CTR mode).
# ===========================================================================
_SBOX = bytes.fromhex(
    "637c777bf26b6fc53001672bfed7ab76ca82c97dfa5947f0add4a2af9ca472c0"
    "b7fd9326363ff7cc34a5e5f171d8311504c723c31896059a071280e2eb27b275"
    "09832c1a1b6e5aa0523bd6b329e32f8453d100ed20fcb15b6acbbe394a4c58cf"
    "d0efaafb434d338545f9027f503c9fa851a3408f929d38f5bcb6da2110fff3d2"
    "cd0c13ec5f974417c4a77e3d645d197360814fdc222a908846eeb814de5e0bdb"
    "e0323a0a4906245cc2d3ac629195e479e7c8376d8dd54ea96c56f4ea657aae08"
    "ba78252e1ca6b4c6e8dd741f4bbd8b8a703eb5664803f60e613557b986c11d9e"
    "e1f8981169d98e949b1e87e9ce5528df8ca1890dbfe6426841992d0fb054bb16"
)


def _xtime(a):
    a <<= 1
    if a & 0x100:
        a ^= 0x11B
    return a & 0xFF


def _expand_key_256(key: bytes):
    assert len(key) == 32
    rcon = 1
    w = [list(key[i:i + 4]) for i in range(0, 32, 4)]  # 8 words
    for i in range(8, 60):
        t = list(w[i - 1])
        if i % 8 == 0:
            t = t[1:] + t[:1]
            t = [_SBOX[b] for b in t]
            t[0] ^= rcon
            rcon = _xtime(rcon)
        elif i % 8 == 4:
            t = [_SBOX[b] for b in t]
        w.append([w[i - 8][j] ^ t[j] for j in range(4)])
    # round keys as 15 x 16-byte blocks
    rk = []
    for r in range(15):
        blk = []
        for c in range(4):
            blk += w[r * 4 + c]
        rk.append(blk)
    return rk


def _aes256_encrypt_block(block16: bytes, round_keys):
    s = list(block16)

    def add(rk):
        for i in range(16):
            s[i] ^= rk[i]

    add(round_keys[0])
    for rnd in range(1, 15):
        # SubBytes
        for i in range(16):
            s[i] = _SBOX[s[i]]
        # ShiftRows (state is column-major: s[r + 4c])
        ns = s[:]
        for r in range(4):
            for c in range(4):
                ns[r + 4 * c] = s[r + 4 * ((c + r) % 4)]
        s = ns
        if rnd != 14:
            # MixColumns
            for c in range(4):
                col = s[4 * c:4 * c + 4]
                t = col[0] ^ col[1] ^ col[2] ^ col[3]
                u = col[0]
                s[4 * c + 0] ^= t ^ _xtime(col[0] ^ col[1])
                s[4 * c + 1] ^= t ^ _xtime(col[1] ^ col[2])
                s[4 * c + 2] ^= t ^ _xtime(col[2] ^ col[3])
                s[4 * c + 3] ^= t ^ _xtime(col[3] ^ u)
        add(round_keys[rnd])
    return bytes(s)


def _aes256_ctr(key: bytes, iv16: bytes, data: bytes) -> bytes:
    """CTR with a full 128-bit big-endian counter (matches WebCrypto length=128)."""
    rk = _expand_key_256(key)
    out = bytearray()
    counter = int.from_bytes(iv16, "big")
    for off in range(0, len(data), 16):
        ks = _aes256_encrypt_block(
            (counter % (1 << 128)).to_bytes(16, "big"), rk)
        chunk = data[off:off + 16]
        out += bytes(b ^ ks[i] for i, b in enumerate(chunk))
        counter += 1
    return bytes(out)


def _selftest_crypto():
    # FIPS-197 AES-256 known-answer
    key = bytes.fromhex("000102030405060708090a0b0c0d0e0f101112131415161718191a1b1c1d1e1f")
    pt = bytes.fromhex("00112233445566778899aabbccddeeff")
    ct = _aes256_encrypt_block(pt, _expand_key_256(key))
    assert ct.hex() == "8ea2b7ca516745bfeafc49904b496089", "AES-256 KAT failed: " + ct.hex()
    # CTR round-trip is just XOR with keystream -> encrypt twice returns plaintext
    iv = os.urandom(16)
    msg = b"the quick brown fox jumps over 13 lazy dogs!! 1234567890"
    enc = _aes256_ctr(key, iv, msg)
    assert _aes256_ctr(key, iv, enc) == msg, "AES-CTR round-trip failed"


def encrypt_record(roll: str, pin: str, plaintext: bytes, lookup_salt: bytes):
    salt = os.urandom(16)
    iv = os.urandom(16)
    dk = hashlib.pbkdf2_hmac("sha256", f"{roll}:{pin}".encode(),
                             salt, PBKDF2_ITERATIONS, 64)
    k_enc, k_mac = dk[:32], dk[32:]
    ct = _aes256_ctr(k_enc, iv, plaintext)
    tag = hmac.new(k_mac, iv + ct, hashlib.sha256).digest()
    rid = hashlib.sha256(lookup_salt + roll.encode()).hexdigest()
    b64 = lambda b: base64.b64encode(b).decode()
    return rid, {"s": b64(salt), "i": b64(iv), "c": b64(ct), "t": b64(tag)}


# ===========================================================================
# Inputs
# ===========================================================================
def read_calendars_csv():
    rows = []
    with open(CALENDARS_CSV, newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            code = (r["course_code"] or "").strip()
            if code:
                rows.append((code, (r["display_name"] or "").strip()))
    return rows


def parse_passwords(path: Path):
    txt = path.read_text(encoding="utf-8", errors="replace")
    pw = dict(re.findall(r'"([^"]+@[^"]+)"\s*:\s*"(\d+)"', txt))
    if not pw:
        raise SystemExit(f"No passwords parsed from {path}")
    return {k.strip().lower(): v for k, v in pw.items()}


def col_index(header, *names):
    for i, cell in enumerate(header):
        label = str(cell).strip().lower() if cell is not None else ""
        if any(label == n.lower() for n in names):
            return i
    return None


def read_enrollment_xlsx(path: Path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = ws.iter_rows(values_only=True)
    header = next(rows)
    ri = col_index(header, "Roll No.", "Roll No", "Roll Number")
    ni = col_index(header, "Name")
    ei = col_index(header, "Email id", "Email", "Email ID", "Email Id")
    if ri is None or ni is None:
        raise SystemExit(f"{path.name}: missing Roll No./Name header")
    for row in rows:
        if row is None or ri >= len(row) or row[ri] in (None, ""):
            continue
        roll = str(row[ri]).strip()
        if roll.endswith(".0"):
            roll = roll[:-2]
        name = str(row[ni]).strip() if ni < len(row) and row[ni] else ""
        email = (str(row[ei]).strip().lower()
                 if ei is not None and ei < len(row) and row[ei] else "")
        yield roll, name, email
    wb.close()


# ===========================================================================
def main():
    _selftest_crypto()

    cal_rows = read_calendars_csv()
    if len(cal_rows) != 38:
        print(f"WARNING: expected 38 calendars, CSV has {len(cal_rows)}")

    calendars, code_to_key = OrderedDict(), OrderedDict()
    for code, display in cal_rows:
        key = slug(code)
        code_to_key[code] = key
        if key not in CALENDAR_IDS:
            raise SystemExit(f"No calendar ID for key '{key}' (course '{code}').")
        calendars[key] = {"name": display, "calId": CALENDAR_IDS[key]}

    passwords = parse_passwords(PASSWORDS_JS)

    enroll = {}      # roll -> {"name","email","sections":set}
    cal_order = list(calendars.keys())
    for code, _ in cal_rows:
        xlsx = ENROLL_DIR / f"{code}.xlsx"
        if not xlsx.exists():
            raise SystemExit(f"Missing enrollment file: {xlsx.name}")
        for roll, name, email in read_enrollment_xlsx(xlsx):
            rec = enroll.setdefault(roll, {"name": name, "email": email,
                                           "sections": set()})
            if not rec["name"] and name:
                rec["name"] = name
            if not rec["email"] and email:
                rec["email"] = email
            rec["sections"].add(code_to_key[code])

    lookup_salt = os.urandom(16)
    records, missing_pin = OrderedDict(), []
    for roll in sorted(enroll):
        rec = enroll[roll]
        pin = passwords.get(rec["email"])
        if not pin:
            missing_pin.append(f"{roll} <{rec['email'] or 'no-email'}>")
            continue
        ordered = [k for k in cal_order if k in rec["sections"]]
        payload = json.dumps({"name": rec["name"], "roll": roll,
                              "sections": ordered}, ensure_ascii=False).encode()
        rid, blob = encrypt_record(roll, pin, payload, lookup_salt)
        records[rid] = blob

    if missing_pin:
        raise SystemExit("No PIN for these students (fix PASSWORDS file or "
                          "emails), refusing to ship a half-built bundle:\n  "
                          + "\n  ".join(missing_pin))

    bundle = {
        "v": 1,
        "kdf": {"name": "PBKDF2", "hash": "SHA-256", "iters": PBKDF2_ITERATIONS,
                "dkLen": 64},
        "cipher": "AES-256-CTR",
        "mac": "HMAC-SHA-256",
        "lookupSalt": base64.b64encode(lookup_salt).decode(),
        "records": records,
    }

    OUT_CALENDARS.write_text(json.dumps(calendars, indent=2, ensure_ascii=False)
                             + "\n", encoding="utf-8")
    OUT_STUDENTS.write_text(json.dumps(bundle, separators=(",", ":")) + "\n",
                            encoding="utf-8")

    # remove any stale cleartext enrollment file so it can never be published
    stale = HERE / "enrollments.json"
    if stale.exists():
        stale.unlink()
        print("removed stale cleartext enrollments.json")

    counts = [len(v["sections"]) for v in enroll.values()]
    print(f"calendars.json     : {len(calendars)} calendars")
    print(f"students.enc.json  : {len(records)} students (encrypted)")
    print(f"courses/student    : min {min(counts)}, max {max(counts)}, "
          f"avg {sum(counts)/len(counts):.1f}")
    blank = [k for k, v in calendars.items() if not v["calId"]]
    if blank:
        print("WARNING: calendars with blank calId: " + ", ".join(blank))
    # sanity: no plaintext name leaked into the bundle
    raw = OUT_STUDENTS.read_text()
    leak = [enroll[r]["name"] for r in list(enroll)[:50]
            if enroll[r]["name"] and enroll[r]["name"] in raw]
    print("plaintext-name leak check:", "FAIL " + str(leak) if leak else "clean")


if __name__ == "__main__":
    main()
