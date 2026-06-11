# ADPC / ADNCA Output Data QC Skill

## Purpose
Comprehensive QC of the final ADPC or ADNCA dataset after `adnca.sas` has run.
Verifies flag derivation logic, population integrity, treatment mappings, erroneous
records, MRRLT duplicate timepoints, and correct inclusion/exclusion across the
NCA and summary statistics populations.

Produces `qc_output_results.json`, `qc_output_report.md`, `qc_output_report.html`,
and `adpc_output_qc.xlsx`.

## When to invoke
- "Run QC on the output ADNCA / ADPC dataset"
- "Check the NCA01FL and PKSUMXFL populations"
- "Verify the CRIT and NCAxxXRS flags in the final dataset"
- Any request to QC the ADPC/ADNCA dataset before submitting to Phoenix WinNonlin / Certara

## Quick start

```bash
# Fully interactive
python adpc_output_qc_run.py

# Pre-fill paths
python adpc_output_qc_run.py \
    --data  path/to/adnca.xpt \
    --sas   path/to/adnca.sas \
    --phase "PHASE 1"

# Non-interactive (accept all auto-detected mappings)
python adpc_output_qc_run.py \
    --data  path/to/adnca.csv \
    --sas   path/to/adnca.sas \
    --yes
```

## Required packages
```
pip install pandas pyreadstat openpyxl markdown
```

## Input files

| File | Required | Description |
|------|----------|-------------|
| ADPC / ADNCA dataset | ✅ Yes | CSV, XPT, or SAS7BDAT |
| Production SAS program | Recommended | Used to parse `whichc()` / `cmiss()` derivation logic |

## Key design decisions

### MRRLT vs NRRLT for duplicate detection
Duplicate timepoints are checked at the **MRRLT** level (not NRRLT) because MRRLT
is the actual relative time passed to the NCA software. After `adnca.sas`
de-duplication the NCA population must have zero MRRLT duplicates.

### Population definitions
| Population | Filter | Includes troughs? |
|---|---|---|
| NCA | `NCA01FL = 'Y'` | No — profile records only |
| Summary stats | `PKSUMXFL` missing/empty | Yes — profile + trough records |
| Excluded from everything | `PKSUMXFL = 'Y'` | — |

### Flag derivation (parsed from SAS program)
The agent automatically extracts which CRITs drive `PKSUMXFL` and which NCAxxXRS
variables drive `NCAXFL` from `whichc()` and `cmiss()` calls in the production
SAS program. This is shown verbatim in **Section 0** of the report.

**Important:** Not all CRITs and not all NCAxxXRS drive NCA01FL or PKSUMXFL.
Some are informational only or reserved for future analyses. The report
explicitly lists which variables drive each derived flag and which are
informational.

## Sections and checks

| Section | What it checks |
|---------|----------------|
| 0 Flag Derivation Logic | Parsed `whichc()` / `cmiss()` lists; raw SAS statements; informational-only flags |
| A Dataset Overview | Record counts, column list, unmapped variables |
| B Treatment Mapping | TREAT / SATRT / ATRT / AVISIT / VISITCD / PHASE / ACTARM cross-reference |
| C Route Table | Unique EXROUTE / ECROUTE values |
| D CRIT / NCAxxXRS Crossref | All flags with Y/N/missing counts; DRIVES_PKSUMXFL / DRIVES_NCAXFL |
| E Population Checks | NCA population, summary population, trough records, contradiction check |
| F DTYPE Flags | DELETED / MODIFIED / COPY distribution; DELETED in NCA population |
| G MRRLT Duplicates | Duplicate MRRLT within USUBJID × PCTEST × VISITCD, by population |
| H Erroneous Records | Bad-year PCDTC, PCDTCEFL='Y', missing ARFTDTM, ATM=00:00 |
| I Dose Flags | CRIT3FL (missed dose), CRIT7FL (incomplete dose) |
| J Pre/Post Flags | CRIT2FL, CRIT10FL, RECORD_PRE_POST, RECORD_POST_PRE |
| K Vomiting | CRIT12FL, VOMFL by visit |
| L Profile Types | PROFILE_TYPE distribution; RICHRFL |
| M RICHRFL / LISTRFL | Consistency check; LISTRFL and NCA01FL must not co-occur |
| Summary | Prioritised action items: 🔴 Critical / ⚠️ Review / ℹ️ Info |

## XLSX sheets

| Sheet | Contents |
|-------|----------|
| CRIT_NCA_Crossref | All CRIT/NCAxxXRS with counts and driver flags |
| Flag_Derivation_Logic | Parsed whichc/cmiss variable lists |
| Treatment_Mapping | TREAT/SATRT/ATRT/AVISIT/PHASE/ACTARM unique combinations |
| Route_Table | EXROUTE / ECROUTE unique values |
| NCA01FL_Y | All NCA records; yellow highlight if any CRITxFL='Y' |
| NCA01FL_with_CRIT | Records in NCA with at least one CRITxFL='Y' (intentional inclusions) |
| Summary_Pop_with_CRIT | Records in summary pop with CRITxFL='Y' — red highlight (unexpected) |
| Trough_Records | PKSUMXFL missing AND NCA01FL != 'Y' |
| MRRLT_Dups_All | All MRRLT duplicates |
| MRRLT_Dups_NCA | MRRLT duplicates within NCA01FL='Y' (must be zero) |
| MRRLT_Dups_Summary | MRRLT duplicates within summary population |
| Erroneous_Records | Bad dates, PCDTCEFL='Y' |
| DELETED_Records | All DTYPE='DELETED' for audit trail |
| Missing_Dose | CRIT3FL / MISSED_DOSE_FL records |
| Vomiting | CRIT12FL / VOMFL records |
| SubjectCounts_Phase | N subjects/records by PHASE |
| SubjectCounts_ARM | N subjects/records by ACTARM |
| SubjectCounts_PCTEST | N subjects/records by PCTEST |
| Profile_Types_Overall | PROFILE_TYPE distribution |
| Profile_Types_ByGroup | PROFILE_TYPE by PHASE × ACTARM × PCTEST |
| Pop_Contradiction | PKSUMXFL='Y' AND NCA01FL='Y' (must be empty) |

## Variable mapping
Variables are auto-detected from column names (case-insensitive) with common
aliases. If a required variable is missing, the user is prompted. If the user
provides a name not in the dataset, it is logged as a QC finding in the report.

## Companion tools
- `pk_input_qc` — QC of SDTM/ADaM *input* datasets before building ADNCA
- `adpc_review_v30.r` — Interactive R Shiny ARRLT review (use with NCA01FL or PKSUMXFL filter)
