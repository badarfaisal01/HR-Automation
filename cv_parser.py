import fitz
import re

def extract_text_from_pdf(file_path):
    text = ""
    with fitz.open(file_path) as pdf:
        for page in pdf:
            text += page.get_text()
    return text

def parse_cv(file_path):
    text = extract_text_from_pdf(file_path)
    name = re.search(r'Name[:\s]*([A-Za-z ]+)', text)
    email = re.search(r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+', text)
    skills = re.findall(r'Python|Java|AI|Machine Learning|React|SQL', text, re.IGNORECASE)

    return {
        "Name": name.group(1) if name else "Unknown",
        "Email": email.group(0) if email else "Not Found",
        "Skills": list(set(skills))
    }
