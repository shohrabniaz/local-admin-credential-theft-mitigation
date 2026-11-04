"""Build local-admin / credential-theft mitigation submission pack (defensive)."""
from __future__ import annotations

from datetime import date
from pathlib import Path
import json
import shutil

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from pptx import Presentation
from pptx.dml.color import RGBColor as PRgb
from pptx.util import Inches, Pt as PPt

ROOT = Path(__file__).resolve().parent
OUT = ROOT / "Final_Submission"
ASSESS = ROOT / "01_Assessment"
MIT = ROOT / "02_Mitigation"
EVAL = ROOT / "03_Evaluation"
EVID = ROOT / "evidence"
NAVY = RGBColor(0x1F, 0x4E, 0x79)
TODAY = date.today().isoformat()


def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs:
        r.font.color.rgb = NAVY


def load_latest(pattern: str):
    files = sorted(EVID.glob(pattern), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        return None
    return json.loads(files[0].read_text(encoding="utf-8-sig"))


def build_docs(score, inv, laps):
    ASSESS.mkdir(parents=True, exist_ok=True)
    MIT.mkdir(parents=True, exist_ok=True)

    d = Document()
    d.add_heading("Endpoint privilege and credential exposure assessment", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | Local admin hardening | {TODAY}")
    heading(d, "Ethics", 1)
    d.add_paragraph(
        "Read-only Windows inventory only. No credential dumping, no attack-path "
        "graph tooling, no lateral-movement procedures."
    )
    heading(d, "Lab host signals", 1)
    if inv:
        d.add_paragraph(
            f"Host {inv.get('Host')}; current user elevated admin: "
            f"{inv.get('CurrentUserIsElevatedAdmin')}; Administrators members: "
            + ", ".join(
                (m.get("Name") if isinstance(m, dict) else str(m))
                for m in inv.get("LocalAdministrators", [])
            )
        )
    if laps:
        cg = laps.get("CredentialGuard") or {}
        lp = laps.get("LAPS") or {}
        d.add_paragraph(
            f"LAPS policy present (legacy/windows): "
            f"{lp.get('LegacyAdmPwdPolicyPresent')}/{lp.get('WindowsLapsPolicyPresent')}; "
            f"Credential Guard running: {cg.get('CredentialGuardRunning')}"
        )
    if score:
        heading(d, "Scored exposure", 1)
        d.add_paragraph(
            f"Pre-mitigation exposure {score.get('PreExposurePercent')}%; "
            f"post-mitigation design residual {score.get('PostMitigationDesignPercent')}% "
            f"(reduction {score.get('ReductionPoints')} points)."
        )
        for f in score.get("Factors", []):
            d.add_paragraph(
                f"{f.get('Id')} {f.get('Title')}: pre {f.get('Pre')} → design {f.get('PostDesign')} — {f.get('Note')}",
                style="List Bullet",
            )
    d.save(ASSESS / "Endpoint_Privilege_Credential_Exposure.docx")

    d = Document()
    d.add_heading("Mitigation design — least privilege, LAPS, credential protections", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | {TODAY}")
    heading(d, "Controls", 1)
    for x in [
        "Remove daily interactive users from local Administrators; keep built-in Administrator for break-glass only.",
        "Deploy Windows LAPS (Entra-joined or AD) so every device has a unique, rotating local admin password.",
        "IT elevation via temporary membership or an approved elevation tool — not standing rights.",
        "Enable Credential Guard where hardware allows; confirm VBS running.",
        "Defender ASR rules that reduce credential theft and Office child-process abuse (pilot then enforce).",
        "Deny network logon for local accounts where operationally feasible to blunt lateral reuse.",
        "Measure helpdesk tickets for 'need admin' in the two weeks after rollout.",
    ]:
        d.add_paragraph(x, style="List Number")
    heading(d, "User and IT impact", 1)
    d.add_paragraph(
        "Standard users lose install rights; package common apps via winget/Intune. "
        "IT needs a clear elevation request path. Expect a short ticket spike, then decline "
        "if software packaging keeps up."
    )
    d.save(MIT / "Mitigation_Design.docx")

    d = Document()
    d.add_heading("SME endpoint privilege hardening guide", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | {TODAY}")
    for x in [
        "[ ] Inventory local Administrators on a sample of PCs",
        "[ ] Identify shared local admin passwords — plan LAPS",
        "[ ] Pilot: remove daily users from Administrators on 5–10 PCs",
        "[ ] Deploy Windows LAPS to pilot OU / Entra group",
        "[ ] Confirm Credential Guard on supported hardware",
        "[ ] Publish elevation request process for staff",
        "[ ] ASR pilot (audit → warn → block)",
        "[ ] Re-score exposure after 30 days",
    ]:
        d.add_paragraph(x, style="List Bullet")
    d.save(MIT / "SME_Hardening_Guide.docx")

    d = Document()
    d.add_heading("Lab validation results", 0)
    d.add_paragraph(f"Md Shohrab Hossen Niaz | {TODAY}")
    d.add_paragraph(
        "Validation uses configuration inventory on the laboratory Windows host plus "
        "design-coverage scoring. It does not execute privilege-escalation or lateral-movement tooling."
    )
    if score:
        sig = score.get("LabSignals") or {}
        d.add_paragraph(
            f"Observed: admin member count={sig.get('AdminMemberCount')}, "
            f"daily-admin likely={sig.get('DailyAdminLikely')}, "
            f"LAPS present={sig.get('LapsPolicyPresent')}, "
            f"Credential Guard running={sig.get('CredentialGuardRunning')}."
        )
        d.add_paragraph(
            f"Score: {score.get('PreExposurePercent')}% → {score.get('PostMitigationDesignPercent')}%."
        )
    d.save(ASSESS / "Lab_Validation_Results.docx")


def build_workbook(score):
    EVAL.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    fill = PatternFill("solid", fgColor="1F4E79")
    white = Font(color="FFFFFF", bold=True)
    ws = wb.active
    ws.title = "Exposure_Score"
    ws["A1"] = "Local privilege / credential exposure"
    ws["A1"].font = Font(bold=True, color="1F4E79", size=13)
    ws["A3"] = "Pre %"
    ws["B3"] = score.get("PreExposurePercent") if score else ""
    ws["A4"] = "Post design %"
    ws["B4"] = score.get("PostMitigationDesignPercent") if score else ""
    ws["A5"] = "Reduction"
    ws["B5"] = score.get("ReductionPoints") if score else ""
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 40

    fac = wb.create_sheet("Factors")
    fac.append(["Id", "Title", "Pre", "PostDesign", "Note"])
    for c in fac[1]:
        c.fill = fill
        c.font = white
    if score:
        for f in score.get("Factors", []):
            fac.append([f.get("Id"), f.get("Title"), f.get("Pre"), f.get("PostDesign"), f.get("Note")])
    for col in fac.columns:
        fac.column_dimensions[col[0].column_letter].width = 28
    fac.column_dimensions["B"].width = 42
    fac.column_dimensions["E"].width = 55

    e8 = wb.create_sheet("EssentialEight_Map")
    e8.append(["E8 theme", "This pack control", "Lab claim"])
    for c in e8[1]:
        c.fill = fill
        c.font = white
    for row in [
        ("Restrict administrative privileges", "Remove daily local admins; JIT elevation", "Design + inventory on one host"),
        ("User application hardening", "ASR pilot; least privilege installs", "Guidance — not full fleet ASR attestation"),
        ("Multi-factor authentication", "Out of scope here — see MFA hardening pack", "Linked theme only"),
        ("Regular backups", "Out of scope", "Not claimed"),
    ]:
        e8.append(list(row))
    for col in e8.columns:
        e8.column_dimensions[col[0].column_letter].width = 36

    path = EVAL / "Local_Admin_Exposure_Workbook.xlsx"
    wb.save(path)
    return path


def build_report(score, inv, laps):
    pre = score.get("PreExposurePercent") if score else "n/a"
    post = score.get("PostMitigationDesignPercent") if score else "n/a"
    red = score.get("ReductionPoints") if score else "n/a"
    sig = (score or {}).get("LabSignals") or {}
    cg = ((laps or {}).get("CredentialGuard") or {})

    sections = [
        (
            "Abstract",
            "Widespread local administrator rights and shared local admin passwords remain "
            "a common path to malware execution and lateral movement in SMEs. I assessed "
            "privilege and credential-theft exposure on a laboratory Windows host, designed "
            "mitigations around least privilege, Windows LAPS, and Credential Guard / ASR, "
            f"and scored residual risk by control coverage. Live inventory found "
            f"{sig.get('AdminMemberCount', 'n/a')} local Administrators members "
            f"(daily-admin likely={sig.get('DailyAdminLikely')}), LAPS policy "
            f"{sig.get('LapsPolicyPresent')}, and Credential Guard running="
            f"{sig.get('CredentialGuardRunning', cg.get('CredentialGuardRunning'))}. "
            f"Exposure scored {pre}% → design residual {post}% (reduction {red} points). "
            "No offensive lateral-movement tooling was used.",
        ),
        (
            "1. Introduction",
            "Catalogue Project 39 asks what local-privilege and credential-theft exposure "
            "exists in a typical SME fleet, which mitigations reduce lateral-movement risk, "
            "and what user/IT impact follows. ASD Essential Eight emphasises restricting "
            "administrative privileges. This pack answers those questions with defensive "
            "inventory and a mitigation design that an MSP can pilot without standing "
            "local admin for every staff PC.",
        ),
        (
            "2. Method",
            "Design science with lab validation limited to configuration evidence: enumerate "
            "the local Administrators group; check Windows LAPS / legacy AdmPwd policy "
            "signals; query Device Guard / Credential Guard status; score five exposure "
            "factors before and after the designed control set. Attack-path tooling is "
            "explicitly out of scope under lab ethics.",
        ),
        (
            "3. Exposure findings",
            f"On host {(inv or {}).get('Host', 'lab')}, the current interactive session was "
            f"not elevated, yet the Administrators group still included a daily-style "
            f"account pattern (member count {sig.get('AdminMemberCount')}). LAPS policy "
            f"was not detected, leaving local Administrator password uniqueness unproven. "
            f"Credential Guard appeared running on this hardware, which lowers memory-"
            "scraping residual relative to an unprotected host but does not excuse standing "
            "local admin or shared passwords.",
        ),
        (
            "4. Mitigation design",
            "Remove interactive users from Administrators; deploy Windows LAPS; provide "
            "JIT elevation for IT; keep Credential Guard on supported devices; pilot "
            "Defender ASR; consider denying network logon for local accounts. Package "
            "software so users do not need admin to work. Expect a short helpdesk spike, "
            "then lower risk if packaging and elevation SLAs are clear.",
        ),
        (
            "5. Validation and limits",
            f"Coverage scoring moves exposure from {pre}% to {post}%. This validates that "
            "the control set addresses the measured gaps on one host; it does not prove "
            "fleet-wide Essential Eight maturity or withstand a live adversary simulation. "
            "One PC is not an estate attestation.",
        ),
        (
            "6. Conclusion",
            "SMEs should treat local admin as an exception with an expiry, not a default. "
            "LAPS plus least privilege cuts the lateral-movement value of a single "
            "compromised PC. Credential Guard helps where available, but identity hygiene "
            "on the endpoint remains the first win.",
        ),
        (
            "References",
            "ASD Essential Eight — Restrict administrative privileges; User application "
            "hardening. Microsoft Learn: Windows LAPS; Credential Guard; Attack surface "
            "reduction rules. MITRE ATT&CK privilege escalation / lateral movement "
            "(conceptual). NIST SP 800-53 AC-6 least privilege. ACSC endpoint hardening guidance.",
        ),
    ]
    bw = sum(len(b.split()) for t, b in sections if t != "References")
    m = Document()
    tp = m.add_paragraph()
    tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    tr = tp.add_run(
        "Securing local administrator rights\n"
        "and credential-theft mitigation for SMEs"
    )
    tr.bold = True
    tr.font.size = Pt(16)
    tr.font.color.rgb = NAVY
    mp = m.add_paragraph()
    mp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    mr = mp.add_run(
        f"Md Shohrab Hossen Niaz\nEndpoint privilege hardening (defensive lab)\n"
        f"Live host inventory | No attack-path tooling\n"
        f"Word count (excluding title block and reference list): {bw}"
    )
    mr.font.size = Pt(11)
    for title, body in sections:
        heading(m, title, 1)
        m.add_paragraph(body)
    OUT.mkdir(parents=True, exist_ok=True)
    path = OUT / "01_Local_Admin_Hardening_Report.docx"
    m.save(path)
    print("BODY_WORDS", bw)
    return path


def build_slides(score):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    navy = PRgb(0x1F, 0x4E, 0x79)

    def add(title, bullets, notes):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        b = s.shapes.add_textbox(Inches(0.6), Inches(0.4), Inches(12), Inches(1))
        r = b.text_frame.paragraphs[0].add_run()
        r.text = title
        r.font.size = PPt(28)
        r.font.bold = True
        r.font.color.rgb = navy
        body = s.shapes.add_textbox(Inches(0.6), Inches(1.5), Inches(12), Inches(5))
        tf = body.text_frame
        tf.word_wrap = True
        first = True
        for line in bullets:
            p = tf.paragraphs[0] if first else tf.add_paragraph()
            first = False
            run = p.add_run()
            run.text = line
            run.font.size = PPt(18)
            p.space_after = PPt(10)
        s.notes_slide.notes_text_frame.text = notes

    pre = score.get("PreExposurePercent") if score else "n/a"
    post = score.get("PostMitigationDesignPercent") if score else "n/a"
    red = score.get("ReductionPoints") if score else "n/a"
    sig = (score or {}).get("LabSignals") or {}

    add(
        "Local admin is a lateral-movement gift",
        ["SME PCs often grant standing admin", "Shared local passwords spread fast", "Defensive lab — inventory, not attack tools"],
        "Project 39.",
    )
    add(
        "What the lab host showed",
        [
            f"Administrators members: {sig.get('AdminMemberCount', 'n/a')}",
            f"LAPS policy: {sig.get('LapsPolicyPresent')}",
            f"Credential Guard running: {sig.get('CredentialGuardRunning')}",
        ],
        "Live signals.",
    )
    add(
        "Mitigations that matter",
        ["Least privilege + JIT elevation", "Windows LAPS unique passwords", "Credential Guard / ASR where supported"],
        "Design.",
    )
    add(
        "Score and ask",
        [f"Exposure {pre}% → {post}% (≈{red} pts)", "Pilot 5–10 PCs; watch helpdesk friction", "Questions?"],
        "Close.",
    )
    path = OUT / "02_Local_Admin_Hardening_Presentation.pptx"
    prs.save(path)
    return path


def assemble():
    score = load_latest("*_Privilege_Exposure_Score.json")
    inv = load_latest("*_Local_Admin_Inventory.json")
    laps = load_latest("*_LAPS_CredentialGuard.json")
    build_docs(score, inv, laps)
    build_workbook(score)
    build_report(score, inv, laps)
    build_slides(score)

    pairs = [
        (EVAL / "Local_Admin_Exposure_Workbook.xlsx", OUT / "Appendix" / "A_Evaluation" / "Local_Admin_Exposure_Workbook.xlsx"),
        (ASSESS / "Endpoint_Privilege_Credential_Exposure.docx", OUT / "Appendix" / "B_Assessment" / "Endpoint_Privilege_Credential_Exposure.docx"),
        (ASSESS / "Lab_Validation_Results.docx", OUT / "Appendix" / "B_Assessment" / "Lab_Validation_Results.docx"),
        (MIT / "Mitigation_Design.docx", OUT / "Appendix" / "C_Mitigation" / "Mitigation_Design.docx"),
        (MIT / "SME_Hardening_Guide.docx", OUT / "Appendix" / "C_Mitigation" / "SME_Hardening_Guide.docx"),
    ]
    for src, dst in pairs:
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dst)
    app_d = OUT / "Appendix" / "D_Scripts"
    app_e = OUT / "Appendix" / "E_Evidence"
    app_d.mkdir(parents=True, exist_ok=True)
    app_e.mkdir(parents=True, exist_ok=True)
    for f in ROOT.glob("*.ps1"):
        if not f.name.startswith("_"):
            shutil.copy2(f, app_d / f.name)
    for f in EVID.glob("*.json"):
        shutil.copy2(f, app_e / f.name)

    readme = Document()
    readme.add_heading("Project submission", 0)
    readme.add_paragraph(
        "Md Shohrab Hossen Niaz\nLocal administrator and credential-theft mitigation\n"
        "Defensive lab inventory + LAPS/least-privilege design."
    )
    readme.save(OUT / "00_Read_Me_First.docx")
    print("assembled", OUT)


if __name__ == "__main__":
    assemble()
