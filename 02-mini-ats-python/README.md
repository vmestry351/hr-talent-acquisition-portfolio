# Mini ATS (Python)

A rule-based Applicant Tracking System prototype. Give it resumes (PDF, DOCX or TXT) and a job profile.
It runs a set of test cases, scores each candidate out of 100, decides **SELECTED / HOLD / REJECTED**
(or **MANUAL REVIEW** if the resume cannot be read) and saves everything to an Excel file.

## Setup
```
pip install -r requirements.txt
```
Needs Python 3.9 or newer.

## Run
```
python ats.py --list-jobs
python ats.py --job junior_data_analyst --resumes sample_resumes
python ats.py --job sales_executive --resumes "C:\path\to\resume.pdf" --output my_results.xlsx
python ats.py --job hr_intern --resumes sample_resumes --dry-run      # preview, no Excel written
```
Running again with the same `--output` file **adds** new candidates. A candidate with the same email or phone
for the same job is skipped as a duplicate.

## The test cases
| # | Test | Points | Knockout? (fail = reject) |
|---|------|--------|---------------------------|
| 1 | Contact details (email, phone) | 5 | Yes, if neither is found |
| 2 | Minimum education | 10 | Yes |
| 3 | Experience range | 15 | Yes, only if the role has a minimum above 0 |
| 4 | Must-have skills | 35 | Yes, if fewer than the required share (default 60%) are found |
| 5 | Nice-to-have skills | 10 | No |
| 6 | Role relevance (title keywords) | 10 | No |
| 7 | Resume quality (length, sections) | 10 | No |
| 8 | Location | 5 | Only if `location_required` is true |

**Decision rules**
1. Any failed knockout test: REJECTED.
2. Experience above the role's maximum: HOLD (possibly over-qualified), if the score is high enough.
3. Score 70 or more: SELECTED. Score 55 to 69: HOLD. Below 55: REJECTED.
4. Unreadable resume: MANUAL REVIEW.

Thresholds are in `jobs.json` under `settings`.

## The Excel file
- **Summary:** counts, average score and selection rate per job (formulas, update automatically).
- **Candidates:** one row per candidate with score, decision and the reasons. The Decision column is a dropdown,
  so a recruiter can override the system's decision.
- **Test_Results:** every test, pass or fail, with the score and detail for each candidate.

## Add or change a job
Edit `jobs.json`. Each job has `must_have_skills`, `nice_to_have_skills`, `title_keywords`, `min_experience`,
`max_experience`, `min_education` (diploma, bachelor, master or phd) and optional `locations`.
Skill names map to the words searched for in the resume, under the `skills` section. Add aliases there
(for example, "power bi" also matches "powerbi").

## Limitations (please read)
- It matches **keywords**, not meaning. For example, "sales dashboard" can count as a sales skill for a
  sales role. Always review borderline cases.
- Scanned or image-only PDFs have no readable text and go to MANUAL REVIEW (OCR is not included).
- Name, education and experience are found with simple patterns, so unusual resume formats may be misread.
- Experience from dates is an estimate. It only counts date ranges under an Experience or Internship heading.

## Use it responsibly
- It never uses name, gender, age, photo or college name to score a candidate. Keep it that way if you extend it.
- A person should review every rejection before the candidate is told, and should be able to explain the reasons
  (see the Reasons column).
- Resumes contain personal data. Use mock data for demos and follow your company's privacy policy for real data.

## Ideas to extend
Resume ranking across many jobs, an email draft for each decision, a simple web page (Streamlit or Flask),
or OCR for scanned resumes.
