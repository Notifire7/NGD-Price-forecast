#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build ไฟล์สำหรับ GitHub Pages ลง docs/ โดย **ลิงก์ของลูกค้าไม่เปลี่ยนทุกเดือน**

    python3 publish.py

ได้:
    docs/fo/index.html       -> https://<user>.github.io/<repo>/fo/
    docs/cogen-8/index.html  -> .../cogen-8/
    docs/lpg/index.html      -> .../lpg/
    docs/cogen3/index.html   -> .../cogen3/

แจกลิงก์ครั้งเดียวจบ เดือนหน้าแค่ push ทับ ลูกค้ากดลิงก์เดิมได้เลย
ไฟล์ internal ไม่ถูกเผยแพร่ (ตั้งใจ — มันเห็นครบทุกกลุ่ม)
"""
import json, os, re, subprocess, sys

SLUG = {"ngd_fo": "fo", "cogen_m8": "cogen-8", "ngd_lpg": "lpg", "cogen_p3": "cogen3"}
DOCS = "docs"

# หน้าแรกกลาง ๆ ไม่ลิสต์ว่ามีกลุ่มไหนบ้าง
INDEX = """<!DOCTYPE html><html lang="th"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<meta name="robots" content="noindex,nofollow"><title>PTTNGD</title>
<style>body{margin:0;min-height:100vh;display:flex;align-items:center;justify-content:center;
background:#0A2429;color:#8FB3B6;font-family:system-ui,sans-serif;font-size:14px;
text-align:center;padding:24px;line-height:1.7}</style></head><body><div>
กรุณาใช้ลิงก์เฉพาะที่ได้รับจากเจ้าหน้าที่การตลาด<br>
หากเข้าถึงไม่ได้ กรุณาติดต่อผู้ดูแลบัญชีของท่าน
</div></body></html>
"""


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else "profiles.json"
    profiles = json.load(open(src, encoding="utf-8"))
    profiles = profiles["profiles"] if isinstance(profiles, dict) else profiles

    os.makedirs(DOCS, exist_ok=True)
    open(f"{DOCS}/.nojekyll", "w").close()
    open(f"{DOCS}/index.html", "w", encoding="utf-8").write(INDEX)

    links = []
    for pr in profiles:
        slug = SLUG.get(pr["view"][0]) if len(pr["view"]) == 1 else None
        if not slug:
            slug = re.sub(r"[^a-z0-9]+", "-", pr["label"].lower()).strip("-") or "x"
        d = f"{DOCS}/{slug}"
        os.makedirs(d, exist_ok=True)
        json.dump({"profiles": [pr]}, open("_tmp_profile.json", "w"), ensure_ascii=False)
        subprocess.run(["python3", "build.py", "--profiles", "_tmp_profile.json",
                        "--out", f"{d}/index.html"], check=True,
                       stdout=subprocess.DEVNULL)
        links.append((pr["label"], slug, pr["password"]))
    os.remove("_tmp_profile.json")

    print(f"เขียนลง {DOCS}/ แล้ว {len(links)} หน้า\n")
    print("ลิงก์ + รหัส สำหรับแจกลูกค้า (แทน <user>/<repo> ด้วยของจริง)\n")
    w = max(len(l) for l, _, _ in links)
    for label, slug, pw in links:
        print(f"  {label:<{w}}  https://<user>.github.io/<repo>/{slug}/")
        print(f"  {'':<{w}}  รหัส: {pw}\n")
    print("ขั้นต่อไป:  git add docs && git commit -m 'update' && git push")


if __name__ == "__main__":
    main()
