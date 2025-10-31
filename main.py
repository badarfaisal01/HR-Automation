import pandas as pd
import os
from cv_parser import parse_cv

required_skills = {"Python", "Machine Learning", "SQL"}

data = []

for file in os.listdir("attachments"):
    if file.endswith(".pdf"):
        cv_data = parse_cv(f"attachments/{file}")
        applicant_skills = set(cv_data["Skills"])
        missing = required_skills - applicant_skills
        extra = applicant_skills - required_skills

        cv_data["Missing Skills"] = ", ".join(missing)
        cv_data["Extra Skills"] = ", ".join(extra)
        data.append(cv_data)

df = pd.DataFrame(data)
df.to_excel("structured_applicant_data.xlsx", index=False)
print("✅ Data saved to structured_applicant_data.xlsx")
