# ESG Rating Skill for Gold Mining Companies

A disclosure-based ESG rating methodology for gold mining companies, built as a Claude skill. Scores companies on **33 indicators** across Environmental, Social, and Governance pillars on a 0–999 scale with AAA–C grades.

---

## Rating Scale

| Score | Grade |
|---|---|
| 849–999 | AAA (ESG-1) |
| 749–848 | AA (ESG-2) |
| 649–748 | A (ESG-3) |
| 549–648 | BBB (ESG-4) |
| 449–548 | BB (ESG-5) |
| 299–448 | B (ESG-6) |
| < 299 | C (ESG-7) |

---

## Indicator Structure

Each pillar contributes equally (333 pts max):

**E — Environment (11 indicators)**
GHG emissions (Scope 1+2+3), energy intensity, renewable energy share, water consumption, tailings management (GISTM), hazardous substances (cyanide/mercury), biodiversity, land reclamation, environmental penalties, ISO 14001.

**S — Social (10 indicators)**
LTIFR, fatalities, ISO 45001, occupational diseases, staff turnover, training hours, indigenous peoples & communities rights, supplier due diligence, grievance mechanism, social investments.

**G — Governance (12 indicators)**
Board independence, CEO/Chair separation, anti-corruption system, confirmed corruption cases, ESG board oversight & remuneration, risk management, whistleblower channel, responsible gold standards (LBMA/WGC/RJC/ICMM), reserves disclosure, tax transparency & EITI, ESG disclosure quality, financial transparency (ACRA methodology).

---

## Core Principle

**Disclosure-based only.** The skill scores what is explicitly disclosed in the report. No assumptions, no external data. Undisclosed indicator → x = 0, score = 0, reason recorded in the calculation passport.

---

## Supported Companies

| Russian peer-set | Global peer-set |
|---|---|
| Polyus, Polymetal, UGC, Seligdar, Nordgold | Newmont, Barrick, Agnico Eagle, AngloGold Ashanti, Kinross, Gold Fields |

Peer benchmarks are used for quantitative indicators (E1, E3, E5, E10, S1, S2, S4, S5, S10). Russian companies use the Russian peer-set; international companies use the global peer-set.

---

## Key Features

- **Peer benchmarks** — percentile-based scoring against a calibrated peer group for 2024; falls back to ICMM absolute thresholds when fewer than 5 peers are available
- **Hard ceilings** — fatalities > 0 caps S2 at x ≤ 0.5; confirmed corruption caps G4 at x ≤ 0.5; absence of IFRS sets G12 to x = 0
- **Sanctions adjustment** — companies barred from international standards (OFAC/EU/UK sanctions, ICMM/WGC/LBMA/RJC suspension) receive a +0.20 bonus on G8, G9, G10 with an anti-washing cap
- **ACRA financial transparency (G12)** — integrates ACRA's financial disclosure sub-factor: IFRS reporting, audit openness, IR materials, group ownership structure
- **Dynamic analysis** — multi-year ΔScore, CAGR, E/S/G decomposition, top-5 drivers of change, separating real improvement from disclosure improvement
- **Automated Word report** — full calculation passport with per-indicator breakdown, source page references, peer benchmarks, and risk matrix

---

## Repository Structure

```
├── SKILL.md                          — skill entry point & workflow instructions
├── references/
│   ├── methodology.md                — 33 indicators: formulas, weights, ceilings, sanctions bonus, rating scale
│   ├── extraction_guide.md           — where to find each indicator in reports, key phrases, PDF check
│   ├── peer_benchmarks.md            — Russian & global peer-set values (2024)
│   └── report_template.md            — Word report structure
└── scripts/
    ├── calculate_rating.py           — takes extracted JSON, outputs full score decomposition
    └── build_docx_report.py          — generates the Word report from results
```

---

## How to Install

1. Download the `.skill` file from [Releases](../../releases)
2. In Claude.ai: go to **Settings → Skills → Add skill** and upload the file

Or clone this repository and pack it manually:
```bash
git clone https://github.com/Reus3/Skill-For-Claude.-ESG-rating-for-gold-mining-companies.git
cd Skill-For-Claude.-ESG-rating-for-gold-mining-companies
zip -r esg-gold-rating.skill SKILL.md references/ scripts/
```

---

## How to Use

1. Upload an annual report or sustainability report (PDF or DOCX) to Claude
2. Ask: *"Calculate the ESG rating using your methodology"* or *"Apply my criteria to this report"*
3. Claude will extract data, score all 33 indicators, compare against peer benchmarks, and generate a Word report with the full calculation passport

---

## Output

The skill produces:
- An ESG score (0–999) with a letter grade (AAA–C)
- Per-indicator breakdown with source page references
- Peer benchmark comparison
- 4×10 risk matrix
- Dynamic analysis (if multiple years provided)
- A downloadable Word report

---

## Notes

- Methodology is calibrated for gold mining only — do not apply to other industries
- Peer benchmark values are updated annually; 2024 values are included in `references/peer_benchmarks.md`
- The sanctions adjustment applies only to G8, G9, G10 and does not affect G12 (ACRA financial transparency can be achieved on Russian platforms without Western standards)
