# HireFlow ATS: web app with hiring dashboard

A browser-based Applicant Tracking System demo. Paste a resume, pick a job, and 8 automatic test cases give a score out of 100 and a decision (Selected, Hold or Rejected). Add candidates to a pipeline board and watch the dashboard update.

**Live demo:** https://vmestry351.github.io/hr-talent-acquisition-portfolio/04-ats-web-app/

## Features
- **Screen a resume:** paste text or upload a .txt file. Same logic as the Python version in `02-mini-ats-python` (checked against it on 7 sample resumes, with identical scores).
- **Pipeline board:** Applied, Screening, Shortlisted, Interview, Offer, Joined, Rejected.
- **Dashboard:** KPIs, hiring funnel, screening decisions, candidates by source, average score by job.
- **Export** the pipeline to CSV.
- **Light and dark theme.** Data is saved in your own browser (localStorage) and is never sent anywhere.

## Limits
Keyword-based screening misses context, so a recruiter must review every decision. PDF and DOCX files are not read in the browser (use the Python tool for those). Use mock or anonymised resumes only.
