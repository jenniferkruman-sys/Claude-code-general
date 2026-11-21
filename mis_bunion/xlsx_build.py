"""Build the MIS bunion surgeon call tracker. usage: python3 -I xlsx_build.py <out.xlsx>"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment

OUT = sys.argv[1]
F = "Arial"
HDR = Font(name=F, bold=True, color="FFFFFF", size=10)
BODY = Font(name=F, size=10)
BOLD = Font(name=F, size=10, bold=True)
FILL_INFO = PatternFill("solid", fgColor="1F4E78")
FILL_CALL = PatternFill("solid", fgColor="7F6000")
FILL_INPUT = PatternFill("solid", fgColor="FFFF00")
FILL_T1 = PatternFill("solid", fgColor="E2EFDA")
FILL_T2 = PatternFill("solid", fgColor="EDF3F9")
THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(wrap_text=True, vertical="top")

RAPH = "Raphael switchboard 03-775-2000 (all Raphael surgeons)"
# rank, tier, he, en, evidence, track record, public role, private hospitals, clinic address, phones, email, web/booking, maccabi file, haifa drive, contact source/notes
D = [
 (1, "1", "פרופ' אסקיאל פלמנוביץ'", "Prof. Ezequiel Palmanovich", "Strong",
  "8+ MIS hallux valgus papers: 'MICA: A Learning Curve' (2020), RCT SERI vs chevron (2021), modified MIS osteotomy for sesamoids (2023), 'MICA for all severity grades' (2025). US F&A fellowship (Myerson, Baltimore). Meir's F&A service is the main Israeli MICA team.",
  "Director, Foot & Ankle Service, Meir MC (Kfar Saba)", "Private clinic; hospital for private surgery not confirmed (ask)",
  "Beit Mansour, HaNufar 2, floor 7, room 709, Ra'anana (MedReviews)", "055-4548291 (MedReviews forwarding line)", "", "medreviews.co.il/provider/dr-palmanovich-ezequiel",
  "N", "~1h", "Ask whether he operates himself or with his team."),
 (2, "1", "ד\"ר גיל גנוט (Genuth)", "Dr. Gil Ganot / Genuth", "Strong",
  "First author of 4 MICA / MICA+DMMO papers 2025-26 incl. Foot & Ankle International (12-month outcomes, recurrence analysis). F&A fellowships in Australia and Switzerland. Publishes his own recurrence data. MedReviews: 423 reviews.",
  "Senior, Foot & Ankle Unit, Wolfson MC (Holon)", "Raphael, Assuta, Ramat Aviv MC",
  "Raphael Specialist Clinics, Park Atidim, Building 4, floor 1, Tel Aviv", "077-804-2884 (own site); 055-4548275 (MedReviews)", "", "drganot.co.il",
  "N", "~1h15", "MedReviews URL 'dr-genuth-gil' confirms Ganot = Genuth. His 2026 paper reports high X-ray recurrence with MICA+DMMO: ask about DMMO."),
 (3, "2", "ד\"ר אלון בורג", "Dr. Alon Burg", "Strong",
  "MIS hallux valgus since at least 2010 (early-results paper). Papers on severe HV and HV with metatarsus adductus (2020, 2022). Fellowship at the Royal National Orthopaedic Hospital, London (2009). Site: percutaneous burr + screws, walking immediately.",
  "Senior, F&A unit, Beilinson (Rabin MC, Petah Tikva)", "Raphael",
  "Sokolov 46, floor 5, Hod HaSharon (MedReviews)", "072-3939553 (private, own site); 055-4548480 (MedReviews)", "dr.alonburg@gmail.com",
  "dr.alon-burg.co.il ; online booking rofim.org.il/ZimunScriptPage.aspx?YomanID=3315",
  "N", "~1h", "The 073-3949510 line on his site is for Clalit Mushlam only."),
 (4, "2", "ד\"ר שי טננבאום", "Dr. Shay Tenenbaum", "Strong",
  "First author of a 40-foot MIS hallux valgus series (Foot Ankle Orthop 2020, conference abstract). MIS HV page on his site since 2016.",
  "Not verified", "Raphael",
  "HaNadiv 71, floor 1 (Beit HaNadiv), Herzliya", "052-7737769; 053-8411225", "office@tenenbaum.co.il", "tenenbaum.co.il",
  "N", "~1h", ""),
 (5, "2", "ד\"ר נתן (נוני) ברוק", "Dr. Natan (Noni) Brook", "Moderate",
  "Co-author of the 2020 MIS hallux valgus series. On Raphael's MICA surgeon list. Reported Head of F&A Unit at Sheba (search snippet, not verified).",
  "Head, F&A Unit, Sheba (unverified)", "Raphael, Herzliya MC, Assuta",
  "Medinat HaYehudim 89, Herzliya (footdoc.co.il); Hadarim clinic, Kfar Saba", "058-6293998; 09-9509697 (footdoc); Herzliya MC 09-9592999", "", "footdoc.co.il",
  "N", "~1h", "Contacts are from search results; I didn't open the pages. Verify."),
 (6, "2", "ד\"ר עודד קויפמן", "Dr. Oded Koifman (Coifman)", "Moderate",
  "Senior (last) author of the 2020 MIS hallux valgus series. On Raphael's MICA surgeon list. Formerly in Sheba orthopedics (Raphael profile, per search).",
  "Independent; formerly Sheba (per search)", "Raphael, Assuta",
  "Hagefen clinic, HaMaapil 2A, Ramat Gan; Athletic Clinic, Stadium 1, Netanya", "03-7573999 / 053-9956349 (Ramat Gan); 054-8279900 (Netanya)", "", "doctors.assuta.co.il/doctor/oded-coifman-115127",
  "N", "~1h", "Listed in the Maccabi *directory* (general, not the surgeon-choice file). Contacts are from search results; verify."),
 (7, "2", "ד\"ר ניב דריאנגל", "Dr. Niv Dreiangel", "Strong",
  "Co-author of the 2020 MIS hallux valgus series. Own page on MICA / percutaneous HV. On Raphael's MICA list. Israel national basketball team doctor (2018).",
  "Senior orthopedist, Sheba", "Raphael",
  "HaBarzel 15, floor 3, Ramat HaHayal, Tel Aviv", "054-9905433", "niv.dreiangelmd@gmail.com", "doctorniv.co.il",
  "N", "~1h15", ""),
 (8, "3", "ד\"ר מיכאל מיכלין", "Dr. Michael Michlin", "Strong (papers)",
  "First author of 2 MICA papers (2025): MICA for all severity grades; MICA with >50% head translation. Part of Palmanovich's Meir team.",
  "Orthopedics, Meir MC (Kfar Saba)", "Not found",
  "Not found", "Not found: call the Meir MC orthopedic clinic", "", "",
  "N", "~1h", "No independent practice found."),
 (9, "3", "ד\"ר יאיר גרין חלימי", "Dr. Yair Green Halimi", "Strong (papers)",
  "First author: percutaneous scarf-like screwless osteotomy, technique + mid-term outcomes (2024). Different technique from MICA: percutaneous, no screws.",
  "Orthopedics, Tel Aviv Sourasky (Ichilov)", "Not verified",
  "HaTaasia 4, Givatayim (MedReviews)", "055-4548374 (MedReviews forwarding line)", "", "medreviews.co.il/provider/dr-green-halimi-yair",
  "N", "~1h15", ""),
 (10, "3", "ד\"ר דן פרט", "Dr. Dan Prat", "Moderate",
  "Own site: hallux valgus & hammer toe by MIS. On Raphael's MICA list. Fellowship at University of Pennsylvania. No MIS papers found.",
  "Senior Deputy Director, Orthopedics, Sheba; F&A unit", "Raphael",
  "HaDuvdevan 7, floor 3 (AB Physio), Kiryat Ono", "050-7740177", "drdanprat@gmail.com", "danprat.co.il",
  "N", "~1h15", ""),
 (11, "3", "ד\"ר גיא פרידמן", "Dr. Guy Friedman", "Moderate",
  "Own page: bunion & hammer toe by MIS. On Raphael's MICA list. No MIS papers found.",
  "Not verified", "Raphael; Ramat Aviv MC",
  "Ramat Aviv MC, Tel Aviv; HaShiryonaim 11, Ashkelon", "050-7341691", "friguy10@gmail.com", "ankle.co.il",
  "N", "~1h15", ""),
 (12, "3", "ד\"ר אלכסנדר ברמנט", "Dr. Alexander (Alex) Bermant", "Moderate",
  "On Raphael's MICA list. Own site: bunion by 'classic and minimal (no bone exposure)' methods. No MIS papers found.",
  "Senior F&A surgeon, Barzilai (previously Wolfson)", "Raphael, Assuta",
  "HaLohamim 53, floor 4, Holon; HaMelachot 21, Building B, floor 5, Modi'in (Ayala 360)", "050-303-9237; 073-321-4181", "bermantmd@gmail.com", "dr-bermant.co.il",
  "N (claims Maccabi agreement)", "~1h15", "His site says he works in agreement with Maccabi: ask exactly which agreement."),
 (13, "3", "ד\"ר זאב פלדברין", "Dr. Zeev Feldbrin", "Moderate",
  "On Raphael's MICA surgeon list (hospital-level evidence only). Raphael profile lists HV correction.",
  "Director, F&A Service, Wolfson MC", "Raphael",
  "Raphael, Park Atidim 3, Tel Aviv", "03-775-2000 (Raphael)", "", "raphaelhospitals.co.il/doctors/feldbrin-zeev",
  "Y (Tel Aviv)", "~1h15", "The only one in the Maccabi surgeon-choice file."),
 (14, "3", "ד\"ר ליאור בן צבי", "Dr. Lior Ben Zvi", "Moderate",
  "Own site: MIS hallux valgus fixed with K-wires removed after a few weeks, ~6 weeks recovery. Older-generation technique, not screw-fixed MICA.",
  "Senior F&A surgeon, Carmel MC (Haifa)", "Not verified; Clalit",
  "Vardiya 12, Haifa", "052-3786722", "Liorbenzvimd@gmail.com", "drliorbenzvi.co.il",
  "N", "Haifa", "Most convenient, but last on technique."),
]

INFO = ["Rank", "Tier", "Name (HE)", "Name (EN)", "MIS evidence", "Track record", "Public hospital role",
        "Private hospitals (operates)", "Clinic address(es)", "Phone(s)", "Email", "Website / booking",
        "In Maccabi surgeon-choice file", "Approx. drive from Haifa", "Notes"]
CALL = ["Call date", "Spoke with", "Visit price - private (₪)", "Visit covered by Maccabi Sheli? (Y/N)",
        "In Sheli surgeon-choice for HV? (Y/N)", "Hospital for surgery", "Surgery cost - private (₪)",
        "Sheli co-pay (₪)", "Zahav co-pay (₪)", "Private insurance coverage / reimbursement",
        "Est. out-of-pocket: visit + Sheli co-pay (₪)", "Earliest consult date", "Surgery wait time",
        "Technique & fixation (MICA? screws/wires? DMMO?)", "MIS HV cases / year", "Recurrence rate quoted",
        "Operates himself? (Y/N)", "Status / next step"]
WIDTH_INFO = [6, 6, 22, 24, 12, 60, 32, 26, 36, 30, 26, 34, 16, 12, 40]
WIDTH_CALL = [12, 16, 14, 14, 14, 20, 16, 14, 14, 28, 18, 14, 14, 28, 12, 14, 12, 28]

wb = Workbook()
ws = wb.active
ws.title = "Surgeons"
ws.sheet_view.rightToLeft = False

# group header row
n_info, n_call = len(INFO), len(CALL)
ws.cell(row=1, column=1, value="SURGEON INFO (researched 2026-11-21)").font = HDR
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_info)
ws.cell(row=1, column=n_info + 1, value="FILL IN AS YOU CALL (yellow cells)").font = HDR
ws.merge_cells(start_row=1, start_column=n_info + 1, end_row=1, end_column=n_info + n_call)
for c in range(1, n_info + n_call + 1):
    ws.cell(row=1, column=c).fill = FILL_INFO if c <= n_info else FILL_CALL
    ws.cell(row=1, column=c).alignment = Alignment(horizontal="center")

for i, h in enumerate(INFO + CALL, 1):
    c = ws.cell(row=2, column=i, value=h)
    c.font = HDR; c.fill = FILL_INFO if i <= n_info else FILL_CALL
    c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center"); c.border = BOX
ws.row_dimensions[2].height = 45

col = {h: n_info + i + 1 for i, h in enumerate(CALL)}
L = get_column_letter
for r_i, row in enumerate(D, start=3):
    for c_i, v in enumerate(row, 1):
        c = ws.cell(row=r_i, column=c_i, value=v)
        c.font = BODY; c.alignment = WRAP; c.border = BOX
        c.fill = FILL_T1 if row[1] == "1" else (FILL_T2 if row[1] == "2" else PatternFill())
    ws.cell(row=r_i, column=3).font = BOLD
    for h in CALL:
        c = ws.cell(row=r_i, column=col[h])
        c.fill = FILL_INPUT; c.border = BOX; c.font = Font(name=F, size=10, color="0000FF"); c.alignment = WRAP
    v, s = L(col["Visit price - private (₪)"]), L(col["Sheli co-pay (₪)"])
    f = ws.cell(row=r_i, column=col["Est. out-of-pocket: visit + Sheli co-pay (₪)"],
                value=f'=IF(AND({v}{r_i}="",{s}{r_i}=""),"",N({v}{r_i})+N({s}{r_i}))')
    f.fill = PatternFill(); f.font = BODY; f.number_format = '₪#,##0'
    for h in ["Visit price - private (₪)", "Surgery cost - private (₪)", "Sheli co-pay (₪)", "Zahav co-pay (₪)"]:
        ws.cell(row=r_i, column=col[h]).number_format = '₪#,##0'
    for h in ["Call date", "Earliest consult date"]:
        ws.cell(row=r_i, column=col[h]).number_format = 'dd/mm/yyyy'
    ws.cell(row=r_i, column=col["Recurrence rate quoted"]).number_format = '0%'
    ws.row_dimensions[r_i].height = 110

last = 2 + len(D)
yn = DataValidation(type="list", formula1='"Y,N,Unclear"', allow_blank=True)
ws.add_data_validation(yn)
for h in ["Visit covered by Maccabi Sheli? (Y/N)", "In Sheli surgeon-choice for HV? (Y/N)", "Operates himself? (Y/N)"]:
    yn.add(f"{L(col[h])}3:{L(col[h])}{last}")
st = DataValidation(type="list", formula1='"To call,Left message,Called - waiting,Consult booked,Not pursuing"', allow_blank=True)
ws.add_data_validation(st); st.add(f"{L(col['Status / next step'])}3:{L(col['Status / next step'])}{last}")

ws.cell(row=2, column=col["Est. out-of-pocket: visit + Sheli co-pay (₪)"]).comment = Comment(
    "Formula: private visit price + Sheli co-pay. Blank until you enter one of them. Doesn't include private-insurance reimbursement.", "Claude")
ws.cell(row=2, column=col["Sheli co-pay (₪)"]).comment = Comment(
    "Per Maccabi Sheli regulations (Aug 2026) §34: up to 50% off the Zahav co-pay with an in-agreement surgeon AND hospital; never below ₪300.", "Claude")

for i, w in enumerate(WIDTH_INFO + WIDTH_CALL, 1):
    ws.column_dimensions[L(i)].width = w
ws.freeze_panes = "E3"
ws.auto_filter.ref = f"A2:{L(n_info + n_call)}{last}"

# ---------------- How to use sheet
h = wb.create_sheet("How to use")
lines = [
 ("How to use this tracker", True),
 ("• Columns A–O (blue header): researched info. Columns P onward (brown header): yellow cells are for you to fill in as you call.", False),
 ("• 'Est. out-of-pocket' fills itself in from visit price + Sheli co-pay. Don't type in it.", False),
 ("• Rows are ranked by MIS track record. Green = tier 1, light blue = tier 2, white = tier 3.", False),
 ("• 'MedReviews forwarding line' numbers (055-454xxxx) go through MedReviews' call routing. The clinic's own number is better when listed.", False),
 ("• Contacts marked 'from search results' (Brook, Koifman) are from pages I didn't open. Verify them.", False),
 ("", False),
 ("Example of a filled row (format only — invented values, not real data):", True),
]
for i, (t, b) in enumerate(lines, 1):
    c = h.cell(row=i, column=1, value=t); c.font = Font(name=F, size=11 if b else 10, bold=b)
ex_hdr = ["Call date", "Spoke with", "Visit price - private (₪)", "In Sheli surgeon-choice for HV?", "Hospital for surgery", "Sheli co-pay (₪)", "Technique & fixation", "Status"]
ex_val = ["24/11/2026", "Secretary (Michal)", 750, "Y", "Raphael", 4200, "MICA, 2 screws + Akin, no DMMO", "Consult booked"]
for j, (a, b) in enumerate(zip(ex_hdr, ex_val), 1):
    c = h.cell(row=9, column=j, value=a); c.font = HDR; c.fill = FILL_CALL; c.alignment = Alignment(wrap_text=True)
    c = h.cell(row=10, column=j, value=b); c.font = Font(name=F, size=10, color="0000FF", italic=True); c.fill = FILL_INPUT
    if isinstance(b, int): c.number_format = '₪#,##0'
qs = [
 ("Questions to ask the clinic", True),
 ("1. Price of a private consultation? Covered by Maccabi Sheli / Zahav?", False),
 ("2. Is the doctor in the Maccabi Sheli surgeon-choice agreement for hallux valgus correction (ניתוח לתיקון הלוקס ולגוס)? At which hospital?", False),
 ("3. Private surgery cost if not covered? What does it include (hospital, anesthesia, implants, follow-up)?", False),
 ("4. Earliest consultation, and typical wait for surgery?", False),
 ("Questions to ask Maccabi (*3555)", True),
 ("5. Confirm the surgeon AND hospital are in agreement for this procedure, and the exact Sheli co-pay. (Zahav §6: both must be in agreement; Sheli §34: up to 50% off the Zahav co-pay, minimum ₪300.)", False),
 ("Questions for the surgeon at the consultation", True),
 ("6. MIS generation & fixation (MICA with screws? K-wires?). Will you add DMMO, and why?", False),
 ("7. How many MIS bunion cases per year? Your recurrence rate? Do you do the whole operation yourself?", False),
 ("8. Weight-bearing, return to normal shoes, return to driving?", False),
 ("", False),
 ("Sources: own surgeon sites, MedReviews, Raphael MICA page (raphaelhospitals.co.il/surgeries/hallux-valgus), PubMed, Maccabi Sheli/Zahav regulations (Aug 2026). Full URLs in mis_bunion_maccabi.md.", False),
]
for i, (t, b) in enumerate(qs, 12):
    c = h.cell(row=i, column=1, value=t); c.font = Font(name=F, size=11 if b else 10, bold=b)
h.column_dimensions["A"].width = 18
for j in range(2, 9): h.column_dimensions[L(j)].width = 20

wb.move_sheet("How to use", offset=-1)
wb.active = 1
wb.save(OUT)
print("saved", OUT)
