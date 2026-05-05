import os
import pandas as pd
import requests
from tqdm import tqdm

# ===== CONFIG =====
TEST_MODE = True
TEST_LIMIT = 10

API_URL = "http://107.210.222.39:9000/registerCattle/manual"

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

# ===== TEST MODE =====
if TEST_MODE:
    df = df.head(TEST_LIMIT)
    print(f"\n🧪 TEST MODE ENABLED → Processing {len(df)} rows\n")

# ===== MAIN LOOP =====
for _, row in tqdm(df.iterrows(), total=len(df), desc="Uploading"):
    godhaar = str(row["godhaar"]).strip()

    # ---- Farmer Mapping ----
    farmer_info = farmers_map.get(godhaar)
    if farmer_info is None:
        skipped_ids.append(godhaar)
        continue

    farmer_name = str(farmer_info["farmerName"]).strip()

    # Normalize animal_type for API: must be "cow" or "buffalo"
    animal_type = str(farmer_info["animalType"]).strip().lower()
    if animal_type not in ("cow", "buffalo"):
        skipped_ids.append(godhaar)
        continue

    # ---- Images ----
    face_path = get_face_path(godhaar)
    muzzle_paths = sorted(get_muzzle_paths(godhaar))  # stable order

    # require at least 1 muzzle + face
    if not face_path or len(muzzle_paths) == 0:
        skipped_ids.append(godhaar)
        continue

    # API supports max 5 muzzle images (muzzle_image_1 required, 2..5 optional)
    muzzle_paths = muzzle_paths[:5]

    left_path, right_path = get_side_paths(godhaar, face_path)

    # ---- FILES ----
    files = []
    try:
        files.append(("front_image", open(face_path, "rb")))
        files.append(("left_image", open(left_path, "rb")))
        files.append(("right_image", open(right_path, "rb")))

        for i, m_path in enumerate(muzzle_paths, start=1):
            files.append((f"muzzle_image_{i}", open(m_path, "rb")))

        # ---- DATA ----
        # village: from main excel row
        village = str(row.get("village", "")).strip()

        # phone: from farmers excel, keyed by godhaar
        phone = str(farmer_info.get("phonenumber", "")).strip()
        if phone in ("nan", "None"):
            phone = ""

        data = {
            "beneficiary_id": str(farmer_info.get("farmerId", row.get("farmerId", ""))),
            "godhaar_id": godhaar,
            "farmer_name": farmer_name,
            "animal_type": animal_type,
            "breed": str(row["breed"]),
            "age": parse_age(row["age"]),
            "state": "AP",
            "district": "TPT",
            "mandal": "TPT01",
            "village": village,
            "phone": phone,
        }

        # ---- API CALL ----
        response = requests.post(API_URL, data=data, files=files, timeout=30)

        if response.status_code in [200, 201]:
            success_ids.append(godhaar)
        else:
            failed_ids.append(godhaar)
            file_keys = [k for k, _ in files]
            print(f"\n❌ {godhaar} → {response.status_code} | {response.text}")
            print(f"   sent files: {file_keys}")
            print(f"   sent data: {data}")

    except Exception as e:
        failed_ids.append(godhaar)
        print(f"\n❌ Exception for {godhaar}: {e}")

    finally:
        for _, f in files:
            try:
                f.close()
            except:
                pass

# ===== SAVE LOGS =====
with open("success_ids.txt", "w") as f:
    f.write("\n".join(success_ids))

with open("failed_ids.txt", "w") as f:
    f.write("\n".join(failed_ids))

with open("skipped_ids.txt", "w") as f:
    f.write("\n".join(skipped_ids))

# ===== SUMMARY =====
print("\n" + "=" * 50)
print("🔥 FINAL UPLOAD SUMMARY 🔥")
print("=" * 50)
print(f"📊 Total Processed : {len(df)}")
print(f"✅ Success         : {len(success_ids)}")
print(f"❌ Failed          : {len(failed_ids)}")
print(f"⚠️ Skipped         : {len(skipped_ids)}")

if len(df) > 0:
    print(f"📈 Success Rate    : {round((len(success_ids)/len(df))*100, 2)}%")

print("=" * 50)
print("📄 Logs saved:")
print(" - success_ids.txt")
print(" - failed_ids.txt")
print(" - skipped_ids.txt")
print("🚀 Done bro!")
