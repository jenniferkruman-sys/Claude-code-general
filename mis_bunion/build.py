"""Match MIS-bunion surgeon research against the Maccabi agreement file and write outputs.
usage: python3 -I build.py <doctors-list.xls> <out_dir>
"""
import sys, re, csv, difflib
import pandas as pd

XLS, OUT = sys.argv[1], sys.argv[2]

# ---------------------------------------------------------------- Maccabi file
df = pd.read_excel(XLS, engine="xlrd")
df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
TOWN_COLS = ["ישוב 1", "ישוב 2", "ישוב 3", "ישוב 4"]

def norm(s):
    s = str(s)
    s = re.sub(r"[\"'׳״`’()\-–.,]", " ", s)
    s = s.replace("פרופ", " ").replace("דר ", " ")
    return " ".join(s.split())

def tokens(s):
    return [t for t in norm(s).split() if t not in {"ד", "ר", "פרופ", "דר"}]

rows = []
for _, r in df.iterrows():
    rows.append({
        "first": norm(r["שם פרטי"]), "last": norm(r["שם משפחה"]),
        "field": r["תחום רפואי"], "sub": r["תחום התמחות"] if pd.notna(r["תחום התמחות"]) else "",
        "license": r["רישיון ראשי"],
        "towns": ", ".join(str(r[c]) for c in TOWN_COLS if pd.notna(r[c])),
    })

ORTHO_ROWS = [r for r in rows if "אורתופד" in str(r["field"]) or "אורתופד" in str(r["sub"])]

def sim(a, b):
    return difflib.SequenceMatcher(None, a, b).ratio()

def match(heb_last, heb_first):
    """Return (status, matched_row, score). Tries both name orders, fuzzy on each token."""
    L, F = norm(heb_last), norm(heb_first)
    best, brow = 0.0, None
    for r in ORTHO_ROWS:
        for (a, b) in [(r["last"], r["first"]), (r["first"], r["last"])]:  # order may be flipped
            s_last = max(sim(L, a), max((sim(L, t) for t in a.split()), default=0))
            s_first = max(sim(F, b), max((sim(F, t) for t in b.split()), default=0)) if F else 1.0
            s = 0.6 * s_last + 0.4 * s_first
            if s > best:
                best, brow = s, r
    if best >= 0.92:
        return "Y", brow, best
    if best >= 0.80:
        return "unclear", brow, best
    return "N", brow, best

# ---------------------------------------------------------------- research data
# evidence: strong = own MIS-HV publication (first/senior author) and/or own site + hospital MIS list;
#           moderate = own site describing MIS HV, or hospital MIS-HV surgeon list, or co-authorship;
#           weak = generic "minimally invasive foot & ankle" claim without bunion-specific MIS.
R = "https://www.raphaelhospitals.co.il/surgeries/hallux-valgus/"
PM = "https://pubmed.ncbi.nlm.nih.gov/"
S = [
 dict(he_last="פלמנוביץ'", he_first="אסקיאל", en="Ezequiel Palmanovich", public="Director, Foot & Ankle Service, Meir MC (Kfar Saba)",
      private="Private clinic, Kfar Saba", ev="strong",
      evd=f"Senior/first author of 8+ MIS hallux valgus papers incl. 'MICA: A Learning Curve' (2020) and 'MICA for all HV severity grades' (2025); MedReviews profile: HV by MIS. {PM}32340840/ ; {PM}40639433/ ; https://www.medreviews.co.il/provider/dr-palmanovich-ezequiel",
      haifa="N"),
 dict(he_last="גנוט", he_first="גיל", en="Gil Ganot / Genuth", public="Senior, Foot & Ankle Unit, Wolfson MC (per own site)",
      private="Raphael, Assuta, Ramat Aviv MC (per own site)", ev="strong",
      evd=f"First author of 4 MICA / MICA+DMMO outcome papers 2025-26 (Assuta TA affiliation); own site: MIS HV, F&A fellowships Australia+Switzerland. {PM}41254892/ ; {PM}41768428/ ; https://drganot.co.il/",
      haifa="N", note="Ganot (drganot.co.il) and Genuth (PubMed) treated as one person: same Swiss fellowship, Wolfson/Assuta links, co-author Moses (Wolfson). Not 100% confirmed."),
 dict(he_last="טננבאום", he_first="שי", en="Shay Tenenbaum", public="Not verified (abstract has no affiliation)",
      private="Raphael (MICA list); private clinic Herzliya", ev="strong",
      evd=f"First author 'Minimal Invasive Hallux Valgus Surgery' series, 40 feet (Foot Ankle Orthop 2020); own site page on MIS HV with screws; Raphael MICA list. https://pmc.ncbi.nlm.nih.gov/articles/PMC8702946/ ; https://tenenbaum.co.il/hv-mis/ ; {R}",
      haifa="N"),
 dict(he_last="בורג", he_first="אלון", en="Alon Burg", public="Senior, F&A unit, Beilinson (Rabin MC)",
      private="Raphael (MICA list); private clinics", ev="strong",
      evd=f"First author MIS HV papers (HV+metatarsus adductus 2020, 2022; MIS HV early results 2010 co-author); own site: percutaneous burr + screws, immediate weight-bearing; Raphael MICA list. {PM}35260325/ ; {PM}32381319/ ; https://dr.alon-burg.co.il/ ; {R}",
      haifa="N"),
 dict(he_last="דריאנגל", he_first="ניב", en="Niv Dreiangel", public="Senior orthopedist, Sheba (per own site)",
      private="Raphael (MICA list)", ev="strong",
      evd=f"Co-author 2020 MIS HV series; own page 'MICA / PERC' HV; Raphael MICA list. https://pmc.ncbi.nlm.nih.gov/articles/PMC8702946/ ; https://doctorniv.co.il/ ; {R}",
      haifa="N"),
 dict(he_last="ברוק", he_first="נתן (נוני)", en="Natan (Noni) Brook / Bruck", public="Head, F&A Unit, Sheba (Duns100 via search snippet; not opened)",
      private="Raphael (MICA list); Assuta, Herzliya MC (Duns100 snippet)", ev="moderate",
      evd=f"Co-author 2020 MIS HV series (as 'Bruck N'); Raphael MICA list. https://pmc.ncbi.nlm.nih.gov/articles/PMC8702946/ ; {R}",
      haifa="N"),
 dict(he_last="קויפמן", he_first="עודד", en="Oded Koifman / Coifman", public="Not verified",
      private="Raphael (MICA list)", ev="moderate",
      evd=f"Senior (last) author 2020 MIS HV series ('Coifman OR'); Raphael MICA list. https://pmc.ncbi.nlm.nih.gov/articles/PMC8702946/ ; {R}",
      haifa="N"),
 dict(he_last="מיכלין", he_first="מיכאל", en="Michael Michlin", public="Orthopedics, Meir MC (per papers)",
      private="Not verified", ev="strong",
      evd=f"First author 'MICA for all HV severity grades' (2025) and MICA >50% translation (2025). {PM}40639433/ ; {PM}40131848/",
      haifa="N", note="Found via PubMed; not in seed list."),
 dict(he_last="גרין חלימי", he_first="יאיר", en="Yair Green Halimi", public="Orthopedics, Tel Aviv Sourasky (per paper)",
      private="Givatayim clinic (MedReviews)", ev="strong",
      evd=f"First author 'percutaneous scarf-like screwless osteotomy' technique + mid-term outcomes (2024); MedReviews MIS F&A list. {PM}39059549/",
      haifa="N", note="Percutaneous, but screwless (different from MICA screw fixation)."),
 dict(he_last="פרט", he_first="דן", en="Dan Prat", public="Senior Deputy Director, Orthopedics, Sheba; F&A unit",
      private="Raphael (MICA list)", ev="moderate",
      evd=f"Own site: HV & hammer toe by MIS; Raphael MICA list. https://danprat.co.il/ ; {R}",
      haifa="N", note="Seed 'פרט דן' = Dan Prat. Not the same as Dan Perl (Meir MC, co-author PMID 40131848)."),
 dict(he_last="פרידמן", he_first="גיא", en="Guy Friedman", public="Not verified (residency Barzilai)",
      private="Raphael (MICA list)", ev="moderate",
      evd=f"Own page 'bunion & hammer toe by MIS'; Raphael MICA list. https://ankle.co.il/ ; {R}",
      haifa="N"),
 dict(he_last="ברמנט", he_first="אלכסנדר", en="Alexander (Alex) Bermant", public="Senior F&A surgeon, Barzilai (own site; previously Wolfson)",
      private="Raphael (MICA list), Assuta (own site)", ev="moderate",
      evd=f"Raphael MICA list; own site: bunion 'classic and minimal (no bone exposure)' methods. {R} ; https://dr-bermant.co.il/",
      haifa="N", note="Own site claims an agreement with Maccabi, but he is not in the Maccabi file."),
 dict(he_last="פלדברין", he_first="זאב", en="Zeev Feldbrin", public="Director, F&A Service, Wolfson MC",
      private="Raphael (MICA list)", ev="moderate",
      evd=f"Raphael MICA surgeon list (hospital-level evidence only); Raphael profile lists HV correction. {R} ; https://www.raphaelhospitals.co.il/doctors/feldbrin-zeev/",
      haifa="N"),
 dict(he_last="בן צבי", he_first="ליאור", en="Lior Ben Zvi", public="Senior F&A surgeon, Carmel MC (Haifa)",
      private="Private clinic Haifa area; Clalit", ev="moderate",
      evd="Own site: HV corrected 'minimally invasive', osteotomy fixed with K-wires removed after weeks, ~6 wk recovery. That is an older-generation percutaneous technique, not screw-fixed MICA. https://www.drliorbenzvi.co.il/",
      haifa="Y"),
 dict(he_last="עצמון", he_first="רן", en="Ran Atzmon", public="Orthopedics, Assuta Ashdod (per papers)",
      private="Not verified", ev="moderate", evd=f"Co-author of 4 MICA/MIS HV papers. {PM}32340840/ ; {PM}37202550/", haifa="N"),
 dict(he_last="סגל", he_first="דוד", en="David Segal", public="Orthopedics, Meir MC (per papers)", private="Not verified",
      ev="moderate", evd=f"Co-author of 5 MIS HV papers. {PM}32340840/ ; {PM}35260325/", haifa="N"),
 dict(he_last="מוזס", he_first="נדב", en="Nadav Moses", public="Orthopedics, Wolfson MC (per papers)", private="Not verified",
      ev="moderate", evd=f"Co-author of 3 MICA+DMMO papers 2025-26. {PM}41254892/", haifa="N"),
 dict(he_last="ברין", he_first="ירון", en="Yaron Brin", public="Orthopedics, Meir MC (per papers)", private="Not verified",
      ev="moderate", evd=f"Co-author of 3 MICA papers (Meir). Own MIS-bunion practice not confirmed. {PM}32340840/ ; {PM}40639433/", haifa="N"),
 dict(he_last="יאסין", he_first="מוסטפא", en="Mustafa Yassin", public="Orthopedics, Hasharon (Rabin MC)", private="Not verified",
      ev="moderate", evd=f"First author 'Percutaneous surgery of the forefoot vs open' (2020). {PM}30737141/", haifa="N"),
 dict(he_last="אידלמן", he_first="מרק", en="Mark Eidelman", public="Pediatric Orthopedics, Rambam (per paper)", private="Not verified",
      ev="moderate", evd=f"Senior author: percutaneous distal metatarsal osteotomy for ADOLESCENT HV (2021). Pediatric, not adult MIS. {PM}32228233/",
      haifa="Y"),
 dict(he_last="פלדמן", he_first="ויקטור", en="Viktor Feldman", public="Orthopedics, Meir MC (per papers)", private="Not verified",
      ev="weak", evd=f"Co-author of 2 MICA papers only. {PM}32340840/ ; {PM}37202550/", haifa="N",
      note="The Maccabi file has a 'פלדמן ויקטור' listed as hand surgery, Ashkelon. May not be the same person."),
 dict(he_last="הלר", he_first="שניר", en="Snir Heller", public="Not verified", private="Not verified",
      ev="weak", evd=f"Co-author of 2010 MIS HV early-results paper only. {PM}20951570/", haifa="N"),
 dict(he_last="סלמן", he_first="סרגיי", en="Sergey Salman", public="Head, F&A Service, Emek MC (Afula)",
      private="Medica RMC / Elisha Haifa (rmc-med.co.il); Sharap Plus Haifa", ev="weak",
      evd="Own site lists HV treatment and, separately, 'MIS for toe deformities'. No bunion-specific MIS statement. https://dr-salman.zapweb.co.il/ ; https://rmc-med.co.il/doctors/",
      haifa="Y"),
 dict(he_last="מרגוליס", he_first="מיכאל", en="Michael Margulis", public="Senior, in charge of F&A surgery, Carmel MC (Haifa)",
      private="Sharap Plus Haifa; Nesher clinic", ev="weak",
      evd="Profiles list bunion surgery and 'minimally invasive F&A surgery' as separate items. No bunion-specific MIS. F&A fellowship at Meir MC 2008-09. https://www.medico.co.il/doctors/ ; https://sharapplus.co.il/doctors/michail-margolis/",
      haifa="Y"),
 dict(he_last="פורטל בנקר", he_first="טלי", en="Tali Portal-Banker", public="Senior F&A & trauma, Rambam",
      private="Private clinics Haifa & Lohamei HaGetaot", ev="weak",
      evd="MedReviews: 'arthroscopic and minimally invasive surgery' (generic) and tagged MIS F&A. No bunion-specific MIS. https://www.medreviews.co.il/provider/dr-portal-banker-tali",
      haifa="Y"),
 dict(he_last="ליברזון", he_first="אהרון", en="Aharon Liberzon", public="F&A unit, Laniado (Netanya)",
      private="Clinics Netanya, Herzliya, Haifa", ev="weak",
      evd="Only a MedReviews 'MIS F&A' tag. No bunion-specific MIS. https://www.medreviews.co.il/search/treatment/minimally-invasive-surgery-foot-and-ankle",
      haifa="Y", note="Has a Haifa clinic per directories; main post in Netanya."),
 dict(he_last="אליאס", he_first="שלמה", en="Shlomo Elias", public="Director, F&A Surgery Unit, Assuta Ashdod",
      private="Ramat Aviv MC (NORMED), Herzliya MC; Sportopedia TA", ev="none",
      evd="Own site and MedReviews list 'HV correction' only. No MIS statement found. https://footdr.co.il/", haifa="N"),
 dict(he_last="רוביצקי", he_first="אלכסיי", en="Alexei Rovitsky", public="Deputy Director, Orthopedics, Rambam; F&A & trauma",
      private="Assuta, Elisha (medico)", ev="none",
      evd="Medico profile: deformity correction & complex fractures. No MIS-bunion evidence. https://www.medico.co.il/doctors/", haifa="Y"),
 dict(he_last="מלמד", he_first="איל", en="Eyal Melamed", public="Founder, F&A service, Rambam; past president, Israeli Foot Society",
      private="Sharap Plus Haifa; Kiryat Tivon", ev="none",
      evd="Own site: F&A surgeon, forefoot deformities. No MIS statement. https://eyalmelamed.co.il/", haifa="Y"),
]

RANK = {"strong": 3, "moderate": 2, "weak": 1, "none": 0}
MRANK = {"Y": 0, "unclear": 1, "N": 2}

out = []
for s in S:
    st, r, sc = match(s["he_last"], s["he_first"])
    out.append({
        "name": f'{s["he_last"]} {s["he_first"]} ({s["en"]})',
        "public_hospital_role": s["public"],
        "private_hospitals": s["private"],
        "mis_evidence": s["ev"],
        "mis_evidence_detail_and_url": s["evd"],
        "in_maccabi_file": st,
        "maccabi_file_match": (f'{r["last"]} {r["first"]} | lic {r["license"]} | {r["field"]} {r["sub"]} | {r["towns"]} | score {sc:.2f}' if st != "N" else f"best non-match: {r['last']} {r['first']} ({sc:.2f})"),
        "sheli_copay": "not in file - check serguide.maccabi4u.co.il / clinic",
        "zahav_copay": "not in file - check serguide.maccabi4u.co.il / clinic",
        "haifa_area": s["haifa"],
        "non_ortho_namesake_in_file": "; ".join(f'{x["last"]} {x["first"]} ({x["field"]}, {x["towns"]})' for x in rows if x not in ORTHO_ROWS and norm(s["he_last"]) == x["last"] and norm(s["he_first"]).split()[0] == x["first"].split()[0]),
        "notes": s.get("note", ""),
    })
out.sort(key=lambda d: (-RANK[d["mis_evidence"]], MRANK[d["in_maccabi_file"]], 0 if d["haifa_area"] == "Y" else 1))

cols = list(out[0].keys())
with open(f"{OUT}/mis_bunion_maccabi.csv", "w", newline="", encoding="utf-8-sig") as f:
    w = csv.DictWriter(f, fieldnames=cols); w.writeheader(); w.writerows(out)

# Haifa-area orthopedists in the Maccabi file
HAIFA_TOWNS = {"חיפה", "קריית מוצקין", "קריית טבעון", "קריית אתא", "קריית ביאליק", "קריית ים", "נשר", "טירת כרמל", "רמת ישי", "עפולה"}
ortho = [r for r in rows if "אורתופד" in str(r["field"]) or "אורתופד" in str(r["sub"])]
haifa_ortho = [r for r in ortho if any(t.strip() in HAIFA_TOWNS for t in r["towns"].split(","))]
pd.DataFrame(ortho).to_csv(f"{OUT}/maccabi_orthopedists.csv", index=False, encoding="utf-8-sig")

for d in out:
    print(d["in_maccabi_file"], "|", d["name"], "|", d["maccabi_file_match"])
print("\nHaifa-area ortho in file:", len(haifa_ortho))
for r in haifa_ortho:
    print(" ", r["last"], r["first"], "|", r["towns"], "|", r["sub"])
