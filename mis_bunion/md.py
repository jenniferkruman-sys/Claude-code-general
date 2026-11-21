"""Post-process CSV (manual overrides) and write the markdown report. usage: python3 -I md.py <out_dir>"""
import sys, csv
OUT = sys.argv[1]
P = f"{OUT}/mis_bunion_maccabi.csv"
rows = list(csv.DictReader(open(P, encoding="utf-8-sig")))

# Manual override: file row for Feldman is listed as hand surgery in Ashkelon; the paper author is at Meir (Kfar Saba).
for r in rows:
    if r["name"].startswith("פלדמן ויקטור"):
        r["in_maccabi_file"] = "unclear"
        r["notes"] = (r["notes"] + " Same name in the file, but listed as hand surgery in Ashkelon. Identity not confirmed.").strip()

RANK = {"strong": 3, "moderate": 2, "weak": 1, "none": 0}; MR = {"Y": 0, "unclear": 1, "N": 2}
rows.sort(key=lambda d: (-RANK[d["mis_evidence"]], MR[d["in_maccabi_file"]], 0 if d["haifa_area"] == "Y" else 1))
with open(P, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)

def esc(s): return s.replace("|", "/")

L = []
A = L.append
A("# MIS bunion surgeons × Maccabi surgeon-choice file\n")
A("_Compiled 2026-11-21. Maccabi file: `doctors-list.xls`, sheet `10.2026`, last saved 2026-10-20._\n")
A("## Read this first\n")
A("1. **The Maccabi file has no hospital columns, no co-pay columns and no subspecialty for foot & ankle.** Its columns are: title, first name, last name, license, medical field, specialization (the only orthopedic subspecialty listed is hand surgery), active, location count, and towns 1–4. **None of the co-pays below come from the file.** Every Sheli/Zahav amount has to be looked up per procedure in the online directory (serguide.maccabi4u.co.il) or by calling the clinic. The doctor page I opened in serguide only shows the *visit* fee (quarterly co-pay), not the surgery co-pay.")
A("2. **Rule from the Aug-2026 regulations:** Zahav §6: you can choose a private surgeon in a private hospital only if *both* the surgeon and the hospital are in agreement with Maccabi *for that procedure*. The co-pay is derived from the cost of the operation, minimum ₪300. Sheli §34: with an in-agreement surgeon you get an extra discount of up to 50% off the Zahav co-pay. In some cases the Sheli and Zahav co-pays are the same, and never below ₪300. Both have a 12-month waiting period. Sources: maccabi_sheli_regulations.pdf, maccabi_zahav_regulations.pdf on maccabi4u.co.il.")
A("3. **The file looks incomplete, or it is a narrower list than 'all surgeon-choice surgeons'.** It has only 149 orthopedists. Only 1 of the 9 surgeons on Raphael's MICA list (Feldbrin) is in it. Bermant's own site says he works in agreement with Maccabi, yet he is not in it. Before relying on an 'N', confirm in serguide or with Maccabi (*3555).")
A("4. Evidence scale: **strong** = the surgeon is first or senior author of a MIS-HV outcome or technique paper, or has their own MIS-HV page *plus* an independent MIS listing. **moderate** = own MIS-HV page, or listed on a hospital MIS/MICA page, or co-author of MIS-HV papers. **weak** = a generic 'minimally invasive foot & ankle' claim with nothing bunion-specific. **none** = foot & ankle surgeon, no MIS found. Being a foot & ankle surgeon was never counted as MIS evidence.\n")

A("## Main table\n")
A("Sorted by MIS evidence, then Maccabi file status, then Haifa area.\n")
A("| # | Name | Public hospital role | Private hospitals | MIS evidence | Evidence + URL | In Maccabi file | Sheli co-pay | Zahav co-pay | Haifa area |")
A("|---|---|---|---|---|---|---|---|---|---|")
for i, r in enumerate(rows, 1):
    fm = f' — {esc(r["maccabi_file_match"])}' if r["in_maccabi_file"] != "N" else ""
    A(f'| {i} | {esc(r["name"])} | {esc(r["public_hospital_role"])} | {esc(r["private_hospitals"])} | **{r["mis_evidence"]}** | {esc(r["mis_evidence_detail_and_url"])}{(" _Note: " + esc(r["notes"]) + "_") if r["notes"] else ""} | {r["in_maccabi_file"]}{fm} | not in file | not in file | {r["haifa_area"]} |')
A("")

A("## Takeaways\n")
A("- **Strongest MIS evidence (all outside the Maccabi file):** Palmanovich (Meir; built the Israeli MICA literature), Ganot/Genuth (Wolfson, plus Raphael/Assuta privately; recent MICA+DMMO series), Tenenbaum, Burg, Dreiangel, Michlin, Green Halimi. Brook and Koifman are co-authors of the 2020 MIS HV series (with Tenenbaum and Dreiangel) and are on Raphael's MICA list.")
A("- **In the Maccabi file and some MIS evidence:** Feldbrin (Wolfson F&A head; on Raphael's MICA list, hospital-level evidence only). Yaron Brin (Hadera/Kfar Saba; co-author of 3 MICA papers, but his own MIS bunion practice isn't confirmed). Snir Heller (2010 co-author only; weak). Salman (Haifa/Afula; weak, his MIS claim is about toes). Eidelman (Haifa; pediatric percutaneous only).")
A("- **Haifa reality:** no Haifa surgeon has strong MIS-bunion evidence. The best is Ben Zvi (Carmel): explicitly MIS for HV, but with temporary K-wires and ~6 weeks of restrictions. That is an older-generation technique, not 3rd/4th-generation screw-fixed MICA, and he is not in the file. Margulis (Carmel) trained at Meir in 2008-09, Palmanovich's unit. That's worth asking him about directly, but it is not evidence.")
A("- **Quote check, not just name check:** surgeon choice needs surgeon + hospital + procedure in agreement. Several MIS surgeons operate at Raphael (Tel Aviv). Ask Maccabi which hospital the agreement covers for *hallux valgus correction* with that surgeon.\n")

# (a) Haifa-area candidates
A("## (a) Haifa-area surgeons in the Maccabi file with no MIS evidence (candidates to call)\n")
A("Foot & ankle confirmed from their profiles:\n")
A("- **מלמד איל (Eyal Melamed)** — Haifa, Kiryat Tivon. Founded the Rambam F&A service. Site: https://eyalmelamed.co.il/")
A("- **רוביצקי אלכסיי (Alexei Rovitsky)** — Haifa. Rambam deputy director of orthopedics, F&A & trauma. Operates at Assuta and Elisha.")
A("- (Weak MIS claim, already in the main table: **סלמן סרגיי** — Haifa/Afula; **אידלמן מרק** — pediatric.)\n")
A("Other orthopedists in the file with a Haifa / Krayot / Tivon / Afula / Ramat Yishai location. Subspecialty **not** verified, and the file can't tell you. Ask whether they do forefoot/MIS:\n")
A("ארנרייך אדם · ארסן שפיגלמן (first and last names flipped in the file) · בור נועם · בן לולו אורן · ברקוביץ ירון · גנסין אייל · גרינטל ארנן · דבי דן · חדאד מרואן · חדש אורי · יונאי יניב · לוין דניאל · מרום ליאור · נירנברג גבריאל · סגל ברק · פינקלשטיין מאוריציו (Tiberias/Afula) · פסקין בצלאל · קשת דורון · שזר אבי. The full orthopedist extract is in `maccabi_orthopedists.csv`.\n")

A("## (b) Could not verify\n")
A("| Item | Reason |")
A("|---|---|")
for a, b in [
 ("All Sheli/Zahav surgery co-pays", "Not in the file. serguide doctor pages show only the visit fee. Needs a per-procedure lookup or a call."),
 ("Whether the file is the complete surgeon-choice list", "No header or metadata says what it is. It contains 149 orthopedists and misses surgeons who say they're in agreement (Bermant)."),
 ("Herzliya Medical Center MIS surgeons", "The HV page describes the open and MIS approaches but names no surgeons (the doctor list loads by JavaScript)."),
 ("Assuta, Clalit (Carmel/Emek) hospital pages", "403 bot protection for scripted access."),
 ("Elisha (Haifa)", "www.elisha.co.il now redirects to alumpaz.co.il. No HV/MIS surgeon page found."),
 ("Bnei Zion", "Site redirects to gov.il landing page. No HV/MIS page found."),
 ("Ramat Aviv MC (NORMED)", "Not checked: domain not found or reached in this session."),
 ("Duns100 foot & ankle list (roles for Brook, Elias, Salman, Rovitsky)", "403 when fetched directly. Roles quoted from search snippets only, marked as such."),
 ("Tenenbaum and Koifman public-hospital roles", "The 2020 abstract carries no affiliations, and their pages I read don't state one."),
 ("Ganot = Genuth identity", "Strong circumstantial match (Swiss fellowship, Wolfson/Assuta, Iselin and Moses co-authors), not explicitly stated. The 2021 Rambam pediatric paper's 'Gil Gannot' may be the same person during residency. Unconfirmed."),
 ("Margulis, Portal-Banker, Liberzon, Salman MIS-bunion", "Only generic 'MIS F&A' claims found. Ask each directly: technique (MICA/PECA/4th-gen?), fixation (screws vs K-wires), volume per year."),
 ("Surgeon volume (cases/year)", "No surgeon publishes it. Palmanovich's 'MICA: A Learning Curve' and the Genuth series give cohort sizes only."),
 ("Search coverage", "Bing/DuckDuckGo returned bot challenges, so web discovery relied on one search API. Smaller clinic sites may be missed."),
]:
    A(f"| {a} | {b} |")
A("")
A("## Questions to ask each surgeon\n")
A("1. Which MIS generation and fixation (MICA with 2 screws / 4th-gen / PECA / K-wires)? 2. How many MIS HV cases per year, and your recurrence rate? 3. Is DMMO added routinely? (The Genuth 2026 papers report *high radiographic recurrence* with MICA+DMMO, so ask how they select patients.) 4. At which hospital is *my* procedure in agreement with Maccabi, and what is the exact Sheli co-pay?")
open(f"{OUT}/mis_bunion_maccabi.md", "w", encoding="utf-8").write("\n".join(L))
print("ok", len(rows))
