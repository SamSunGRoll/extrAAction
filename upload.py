import os
import time
import re
import json
import pandas as pd
import requests
from tqdm import tqdm

# ===== CONFIG =====
TEST_MODE = False
TEST_LIMIT = 20
MAX_RETRIES = 3
RETRY_DELAY_SECONDS = 2

REGISTER_CATTLE_API_URL = "http://107.210.222.39:9000/registerCattle/manual"
BENEFICIARY_API_URL = "http://107.210.222.39:9000/beneficiaries/"
MILCHANIMALS_API_URL = "http://107.210.222.39:9000/milchanimals/"
QR_GENERATE_API_URL = "http://107.210.222.39:9000/qr/generate"
DEFAULT_DEV_API_KEY = "7f2d9b8c-3a1e-4f6b-9c2d-8e7a6b5c4d3f"
CHECKPOINT_FILE = "upload_checkpoint.jsonl"

main_excel = "/home/codedreamer/Documents/GitHub/extrAAction/Tirupati_Rural.xlsx"
farmer_excel = "/home/codedreamer/Documents/GitHub/extrAAction/Tirupati_rural_farmers.xlsx"

faces_dir = "/home/codedreamer/Downloads/Tirupati_Rural_Faces"
muzzle_dir = "/home/codedreamer/Downloads/Tirupati_Rural_Muzzle_fACES/Tirupati_Rural_Muzzle"
sides_dir = "/home/codedreamer/Downloads/Tirupati_Rural_Sides"

# ===== LOAD DATA =====
df = pd.read_excel(main_excel, dtype=str)
farmers_df = pd.read_excel(farmer_excel, dtype=str)

df["godhaar"] = df["godhaar"].str.strip()
farmers_df["godhaar"] = farmers_df["godhaar"].str.strip()

farmers_map = {row["godhaar"]: row for _, row in farmers_df.iterrows()}

# ===== HELPERS =====
def parse_age(age_str):
    try:
        age_str = str(age_str).upper()
        years = 0
        months = 0

        if "Y" in age_str:
            years = int(age_str.split("Y")[0].strip() or 0)

        if "M" in age_str:
            if "Y" in age_str:
                months = int(age_str.split("Y")[1].split("M")[0].strip() or 0)
            else:
                months = int(age_str.split("M")[0].strip() or 0)

        return years * 12 + months
    except:
        return 0

def clean_str(value):
    value = str(value).strip()
    return "" if value in ("nan", "None") else value

def to_int_or_default(value, default=0):
    try:
        text = str(value).strip()
        if text in ("", "nan", "None"):
            return default
        return int(float(text))
    except Exception:
        return default

def to_float_or_default(value, default=0.0):
    try:
        text = str(value).strip()
        if text in ("", "nan", "None"):
            return default
        return float(text)
    except Exception:
        return default

def get_from_sources(default, *values):
    for value in values:
        cleaned = clean_str(value)
        if cleaned != "":
            return cleaned
    return default

def get_ci(record, *keys):
    if record is None:
        return ""
    lowered = {str(k).strip().lower(): v for k, v in record.items()}
    for key in keys:
        value = lowered.get(str(key).strip().lower(), "")
        if clean_str(value) != "":
            return value
    return ""

def build_location(row, farmer_info):
    lat = get_from_sources(
        "",
        get_ci(row, "latitude", "lat"),
        get_ci(farmer_info, "latitude", "lat"),
    )
    lon = get_from_sources(
        "",
        get_ci(row, "longitude", "lon", "lng"),
        get_ci(farmer_info, "longitude", "lon", "lng"),
    )
    if lat != "" and lon != "":
        return f"{lat},{lon}"

    raw_location = get_from_sources("", row.get("location", ""), farmer_info.get("location", ""))
    if raw_location == "":
        # Always return API-valid fallback format: "lat,lon"
        return "0,0"

    # Handle dict/json-like text from Excel such as:
    # "{'longitude': 79.47, 'latitude': 13.58, 'timestamp': '...'}"
    lat_match = re.search(r"latitude['\"]?\s*[:=]\s*([+-]?\d+(?:\.\d+)?)", raw_location)
    lon_match = re.search(r"longitude['\"]?\s*[:=]\s*([+-]?\d+(?:\.\d+)?)", raw_location)
    if lat_match and lon_match:
        return f"{lat_match.group(1)},{lon_match.group(1)}"

    # Final fallback to always satisfy location regex expected by API.
    return "0,0"

def map_register_state(value):
    text = clean_str(value).lower()
    if text in ("andhra pradesh", "ap"):
        return "AP"
    return value if clean_str(value) else "AP"

def map_register_district(value):
    text = clean_str(value).lower()
    if text in ("tirupathi", "tirupati", "tpt"):
        return "TPT"
    return value if clean_str(value) else "TPT"

def map_register_mandal(value):
    text = clean_str(value).lower()
    if text in ("tirupathi rural", "tirupati rural", "tpt02"):
        return "TPT02"
    return value if clean_str(value) else "TPT01"

def upsert_beneficiary(farmer_info, row):
    num_of_items = to_int_or_default(
        farmer_info.get("num_of_items", row.get("num_of_items", 0)),
        default=0,
    )

    beneficiary_id = clean_str(farmer_info.get("farmerId", row.get("farmerId", "")))
    if beneficiary_id == "":
        beneficiary_id = f"BEN_{clean_str(row.get('godhaar', 'UNKNOWN'))}"

    beneficiary_state = get_from_sources("Andhra Pradesh", row.get("state", ""), farmer_info.get("state", ""))
    beneficiary_district = get_from_sources("Tirupathi", row.get("district", ""), farmer_info.get("district", ""))
    beneficiary_mandal = get_from_sources("Tirupathi Rural", row.get("mandal", ""), farmer_info.get("mandal", ""))

    beneficiary_village = get_from_sources(
        "Unknown Village",
        row.get("village", ""),
        farmer_info.get("village", ""),
    )
    if len(clean_str(beneficiary_village)) < 2:
        beneficiary_village = "Unknown Village"

    payload = {
        "beneficiary_id": beneficiary_id,
        "name": get_from_sources("UNKNOWN", farmer_info.get("farmerName", ""), row.get("farmer_name", "")),
        "village": beneficiary_village,
        "mandal": beneficiary_mandal,
        "district": beneficiary_district,
        "state": beneficiary_state,
        "phone_number": clean_str(farmer_info.get("phonenumber", "")),
        "num_of_items": num_of_items,
    }
    response = request_with_retry(BENEFICIARY_API_URL, json=payload)

    # Treat duplicate/already-exists as success to keep script idempotent.
    if response.status_code in [200, 201, 409]:
        return True, payload, response

    return False, payload, response

def upsert_milchanimal(farmer_info, row, godhaar, animal_type):
    village = get_from_sources("", row.get("village", ""), farmer_info.get("village", ""))
    beneficiary_id = get_from_sources(
        "",
        farmer_info.get("beneficiary_id", ""),
        farmer_info.get("farmerId", ""),
        row.get("beneficiary_id", ""),
        row.get("farmerId", ""),
    )
    breed = get_from_sources("", row.get("breed", ""), farmer_info.get("breed", ""))
    seller_id = get_from_sources("", row.get("seller_id", ""), farmer_info.get("seller_id", ""))
    location = build_location(row, farmer_info)
    cost = to_float_or_default(row.get("cost", farmer_info.get("cost", 0)), default=0)
    insurance_premium = to_float_or_default(
        row.get("insurance_premium", farmer_info.get("insurance_premium", 0)),
        default=0,
    )
    pregnant_text = get_from_sources("", row.get("pregnant", ""), farmer_info.get("pregnant", ""))
    pregnant = str(pregnant_text).strip().lower() in ("1", "true", "yes", "y")
    pregnancy_months = to_int_or_default(
        row.get("pregnancy_months", farmer_info.get("pregnancy_months", 0)),
        default=0,
    )
    calf_type = get_from_sources("", row.get("calf_type", ""), farmer_info.get("calf_type", ""))
    milk_yield_per_day = to_float_or_default(
        row.get("milk_yield_per_day", farmer_info.get("milk_yield_per_day", 0)),
        default=0,
    )
    tag_no = get_from_sources(godhaar, row.get("tag_no", ""), farmer_info.get("tag_no", ""), godhaar)
    animal_id = godhaar
    animal_photo = get_from_sources("", row.get("animal_photo", ""), farmer_info.get("animal_photo", ""))
    health_cert = get_from_sources("", row.get("health_cert", ""), farmer_info.get("health_cert", ""))
    valuation_cert = get_from_sources("", row.get("valuation_cert", ""), farmer_info.get("valuation_cert", ""))

    payload = {
        "animal_id": animal_id,
        "beneficiary_id": beneficiary_id,
        "seller_id": seller_id,
        "purchase_place": village,
        "location": location,
        "cost": cost,
        "insurance_premium": insurance_premium,
        "type": animal_type,
        "breed": breed,
        "pregnant": pregnant,
        "pregnancy_months": pregnancy_months,
        "calf_type": calf_type,
        "milk_yield_per_day": milk_yield_per_day,
        "tag_no": tag_no,
        "animal_photo": animal_photo,
        "health_cert": health_cert,
        "valuation_cert": valuation_cert,
    }
    response = request_with_retry(MILCHANIMALS_API_URL, json=payload)
    if response.status_code in [200, 201, 409]:
        return True, payload, response
    return False, payload, response

def generate_qr(animal_id, state, district, mandal):
    payload = {
        "state": state,
        "district": district,
        "mandal": mandal,
    }
    url = f"{QR_GENERATE_API_URL}/{animal_id}"
    response = request_with_retry(url, json=payload, timeout=30)
    if response.status_code in [200, 201]:
        return True, payload, response

    # Backend may be eventually consistent after register/create; retry on 404.
    if response.status_code == 404:
        for _ in range(3):
            time.sleep(2)
            response = request_with_retry(url, json=payload, timeout=30)
            if response.status_code in [200, 201]:
                return True, payload, response

    return False, payload, response

def request_with_retry(url, data=None, files=None, json=None, timeout=30):
    api_key = os.getenv("API_X_KEY", DEFAULT_DEV_API_KEY).strip()
    bearer_token = os.getenv("API_BEARER_TOKEN", "").strip()
    headers = {}
    if api_key:
        headers["X-API-Key"] = api_key
    if bearer_token:
        headers["Authorization"] = f"Bearer {bearer_token}"

    last_response = None
    last_exception = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = requests.post(
                url,
                data=data,
                files=files,
                json=json,
                headers=headers,
                timeout=timeout,
            )
            # retry only transient server/network style statuses
            if response.status_code in (500, 502, 503, 504, 429):
                last_response = response
                if attempt < MAX_RETRIES:
                    time.sleep(RETRY_DELAY_SECONDS)
                    continue
            return response
        except Exception as exc:
            last_exception = exc
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_DELAY_SECONDS)
                continue
            raise last_exception
    return last_response

def load_checkpoint(path):
    processed = {}
    if not os.path.exists(path):
        return processed
    with open(path, "r") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                obj = json.loads(line)
                gid = clean_str(obj.get("godhaar", ""))
                if gid:
                    processed[gid] = obj
            except Exception:
                continue
    return processed

def save_checkpoint(path, godhaar, payload):
    entry = {"godhaar": godhaar, "ts": int(time.time())}
    if isinstance(payload, dict):
        entry.update(payload)
    else:
        entry["status"] = str(payload)
    with open(path, "a") as f:
        f.write(json.dumps(entry) + "\n")

def append_unique_lines(path, items):
    existing = set()
    if os.path.exists(path):
        with open(path, "r") as f:
            existing = {line.strip() for line in f if line.strip()}
    new_items = [item for item in items if item and item not in existing]
    if not new_items:
        return
    with open(path, "a") as f:
        for item in new_items:
            f.write(item + "\n")

def get_face_path(godhaar):
    path = os.path.join(faces_dir, godhaar, f"{godhaar}.jpg")
    return path if os.path.exists(path) else None

def get_muzzle_paths(godhaar):
    folder = os.path.join(muzzle_dir, godhaar)
    if not os.path.exists(folder):
        return []
    return [
        os.path.join(folder, f)
        for f in os.listdir(folder)
        if f.lower().endswith((".jpg", ".jpeg", ".png"))
    ]

def get_side_paths(godhaar, face_path):
    folder = os.path.join(sides_dir, godhaar)
    left, right = None, None

    if os.path.exists(folder):
        for f in os.listdir(folder):
            fname = f.lower()
            full = os.path.join(folder, f)
            if "left" in fname:
                left = full
            elif "right" in fname:
                right = full

    if not left:
        left = face_path
    if not right:
        right = face_path

    return left, right

# ===== INIT LOGS =====
success_ids = []
failed_ids = []
skipped_ids = []
beneficiary_failed_ids = []
milchanimal_failed_ids = []
qr_failed_ids = []

checkpoint_map = load_checkpoint(CHECKPOINT_FILE)
processed_ids = set(checkpoint_map.keys())

# Recover prior run counters from checkpoint for resumable summaries/logs.
for gid, entry in checkpoint_map.items():
    status = entry.get("status", "")
    if status == "success":
        success_ids.append(gid)
    elif status == "failed":
        failed_ids.append(gid)
    elif status == "skipped":
        skipped_ids.append(gid)
    if entry.get("beneficiary_failed"):
        beneficiary_failed_ids.append(gid)
    if entry.get("milchanimal_failed"):
        milchanimal_failed_ids.append(gid)
    if entry.get("qr_failed"):
        qr_failed_ids.append(gid)

# ===== TEST MODE =====
if TEST_MODE:
    df = df.head(TEST_LIMIT)
    print(f"\n🧪 TEST MODE ENABLED → Processing {len(df)} rows\n")

# ===== MAIN LOOP =====
for _, row in tqdm(df.iterrows(), total=len(df), desc="Uploading"):
    godhaar = str(row["godhaar"]).strip()
    if godhaar in processed_ids:
        continue
    if clean_str(godhaar) == "":
        skipped_ids.append("MISSING_GODHAAR")
        save_checkpoint(CHECKPOINT_FILE, "MISSING_GODHAAR", "skipped")
        continue

    # ---- Farmer Mapping ----
    farmer_info = farmers_map.get(godhaar)
    if farmer_info is None:
        skipped_ids.append(godhaar)
        continue

    farmer_name = str(farmer_info["farmerName"]).strip()

    # ---- Beneficiary Upsert ----
    try:
        ok, ben_payload, ben_response = upsert_beneficiary(farmer_info, row)
        if not ok:
            beneficiary_failed_ids.append(godhaar)
            skipped_ids.append(godhaar)
            print(f"\n❌ Beneficiary create failed for {godhaar} -> {ben_response.status_code} | {ben_response.text}")
            print(f"   sent beneficiary payload: {ben_payload}")
            save_checkpoint(CHECKPOINT_FILE, godhaar, {
                "status": "skipped",
                "beneficiary_failed": True,
                "milchanimal_failed": False,
                "qr_failed": False,
            })
            continue
    except Exception as e:
        beneficiary_failed_ids.append(godhaar)
        skipped_ids.append(godhaar)
        print(f"\n❌ Beneficiary exception for {godhaar}: {e}")
        save_checkpoint(CHECKPOINT_FILE, godhaar, {
            "status": "skipped",
            "beneficiary_failed": True,
            "milchanimal_failed": False,
            "qr_failed": False,
        })
        continue

    # Normalize animal type with fallback default.
    animal_type = clean_str(farmer_info.get("animalType", "")).lower()
    if animal_type not in ("cow", "buffalo"):
        animal_type = "cow"

    # ---- Milchanimal Upsert ----
    milch_payload = {}
    milchanimal_ok = False
    try:
        ok, milch_payload, milch_response = upsert_milchanimal(farmer_info, row, godhaar, animal_type)
        if not ok:
            milchanimal_failed_ids.append(godhaar)
            if milch_response is None:
                print(f"\n❌ Milchanimal create failed for {godhaar} -> validation | {milch_payload}")
                save_checkpoint(CHECKPOINT_FILE, godhaar, {
                    "status": "failed",
                    "beneficiary_failed": False,
                    "milchanimal_failed": True,
                    "qr_failed": False,
                })
                continue
            print(f"\n❌ Milchanimal create failed for {godhaar} -> {milch_response.status_code} | {milch_response.text}")
            print(f"   sent milchanimal payload: {milch_payload}")
        else:
            milchanimal_ok = True
    except Exception as e:
        milchanimal_failed_ids.append(godhaar)
        print(f"\n❌ Milchanimal exception for {godhaar}: {e}")

    # ---- Images ----
    face_path = get_face_path(godhaar)
    muzzle_paths = sorted(get_muzzle_paths(godhaar))
    if not face_path:
        skipped_ids.append(godhaar)
        save_checkpoint(CHECKPOINT_FILE, godhaar, {
            "status": "skipped",
            "beneficiary_failed": False,
            "milchanimal_failed": not milchanimal_ok,
            "qr_failed": False,
        })
        continue
    if len(muzzle_paths) == 0:
        muzzle_paths = [face_path]
    muzzle_paths = muzzle_paths[:5]
    left_path, right_path = get_side_paths(godhaar, face_path)

    files = []
    try:
        files.append(("front_image", open(face_path, "rb")))
        files.append(("left_image", open(left_path, "rb")))
        files.append(("right_image", open(right_path, "rb")))
        for i, m_path in enumerate(muzzle_paths, start=1):
            files.append((f"muzzle_image_{i}", open(m_path, "rb")))

        village = get_from_sources("", row.get("village", ""), farmer_info.get("village", ""))
        phone = clean_str(farmer_info.get("phonenumber", ""))
        beneficiary_id = get_from_sources(
            "",
            farmer_info.get("beneficiary_id", ""),
            farmer_info.get("farmerId", ""),
            row.get("beneficiary_id", ""),
            row.get("farmerId", ""),
        )

        register_state = map_register_state(
            get_from_sources("AP", row.get("state", ""), farmer_info.get("state", ""))
        )
        register_district = map_register_district(
            get_from_sources("TPT", row.get("district", ""), farmer_info.get("district", ""))
        )
        register_mandal = map_register_mandal(
            get_from_sources("TPT01", row.get("mandal", ""), farmer_info.get("mandal", ""))
        )

        data = {
            "beneficiary_id": beneficiary_id,
            "godhaar_id": godhaar,
            "farmer_name": farmer_name,
            "animal_type": animal_type,
            "breed": get_from_sources("", row.get("breed", ""), farmer_info.get("breed", "")),
            "age": parse_age(row.get("age", "")),
            "state": register_state,
            "district": register_district,
            "mandal": register_mandal,
            "village": village,
            "phone": phone,
        }

        # ---- API CALL ----
        response = request_with_retry(REGISTER_CATTLE_API_URL, data=data, files=files, timeout=30)

        if response.status_code in [200, 201]:
            success_ids.append(godhaar)
            if milchanimal_ok:
                qr_animal_id = clean_str(milch_payload.get("animal_id", "")) or godhaar
                qr_ok, qr_payload, qr_response = generate_qr(
                    qr_animal_id,
                    register_state,
                    register_district,
                    register_mandal,
                )
                if not qr_ok:
                    qr_failed_ids.append(godhaar)
                    print(f"\n❌ QR generation failed for {godhaar} -> {qr_response.status_code} | {qr_response.text}")
                    print(f"   qr animal_id: {qr_animal_id}")
                    print(f"   sent qr payload: {qr_payload}")
            else:
                qr_failed_ids.append(godhaar)
                print(f"\n⚠️ QR skipped for {godhaar} because milchanimal creation failed.")
            save_checkpoint(CHECKPOINT_FILE, godhaar, {
                "status": "success",
                "beneficiary_failed": False,
                "milchanimal_failed": not milchanimal_ok,
                "qr_failed": godhaar in qr_failed_ids,
            })
        else:
            failed_ids.append(godhaar)
            print(f"\n❌ {godhaar} → {response.status_code} | {response.text}")
            print(f"   sent data: {data}")
            save_checkpoint(CHECKPOINT_FILE, godhaar, {
                "status": "failed",
                "beneficiary_failed": False,
                "milchanimal_failed": not milchanimal_ok,
                "qr_failed": False,
            })

    except Exception as e:
        failed_ids.append(godhaar)
        print(f"\n❌ Exception for {godhaar}: {e}")
        save_checkpoint(CHECKPOINT_FILE, godhaar, {
            "status": "failed",
            "beneficiary_failed": False,
            "milchanimal_failed": not milchanimal_ok,
            "qr_failed": False,
        })
    finally:
        for _, f in files:
            try:
                f.close()
            except Exception:
                pass

# ===== SAVE LOGS =====
append_unique_lines("success_ids.txt", success_ids)
append_unique_lines("failed_ids.txt", failed_ids)
append_unique_lines("skipped_ids.txt", skipped_ids)
append_unique_lines("beneficiary_failed_ids.txt", beneficiary_failed_ids)
append_unique_lines("milchanimal_failed_ids.txt", milchanimal_failed_ids)
append_unique_lines("qr_failed_ids.txt", qr_failed_ids)

# ===== SUMMARY =====
print("\n" + "=" * 50)
print("🔥 FINAL UPLOAD SUMMARY 🔥")
print("=" * 50)
print(f"📊 Total Processed : {len(df)}")
print(f"✅ Success         : {len(success_ids)}")
print(f"❌ Failed          : {len(failed_ids)}")
print(f"⚠️ Skipped         : {len(skipped_ids)}")
print(f"👤 Beneficiary Fail: {len(beneficiary_failed_ids)}")
print(f"🐄 Milchanimal Fail: {len(milchanimal_failed_ids)}")
print(f"🔳 QR Fail         : {len(qr_failed_ids)}")

if len(df) > 0:
    print(f"📈 Success Rate    : {round((len(success_ids)/len(df))*100, 2)}%")

print("=" * 50)
print("📄 Logs saved:")
print(" - success_ids.txt")
print(" - failed_ids.txt")
print(" - skipped_ids.txt")
print(" - beneficiary_failed_ids.txt")
print(" - milchanimal_failed_ids.txt")
print(" - qr_failed_ids.txt")
print("🚀 Done bro!")
