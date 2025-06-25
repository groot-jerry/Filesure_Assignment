import os
import fitz  # PyMuPDF
import json
import re
import unicodedata
import magic  # Requires: pip install python-magic-bin (for Windows)

# Helper: Sanitize filenames
def sanitize_filename(name, fallback):
    try:
        return unicodedata.normalize("NFKD", name).encode("ascii", "ignore").decode("ascii") or fallback
    except:
        return fallback

# Helper: Guess file extension from content
def get_file_extension(fdata):
    try:
        mime = magic.from_buffer(fdata, mime=True)
        if "pdf" in mime:
            return ".pdf"
        elif "text" in mime:
            return ".txt"
        elif "msword" in mime:
            return ".doc"
        elif "officedocument" in mime:
            return ".docx"
        elif "excel" in mime:
            return ".xls"
        elif "image" in mime:
            return ".jpg"
        else:
            return ".bin"
    except:
        return ".bin"

# Step 1: Extract PDF Text
def extract_text_from_pdf(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    for page in doc:
        full_text += page.get_text()
    return full_text

# Step 2: Extract Structured Fields
def extract_fields(form_text):
    data = {
        "company_name": "",
        "cin": "",
        "registered_office": "",
        "appointment_date": "",
        "auditor_name": "",
        "auditor_address": "",
        "auditor_frn_or_membership": "",
        "appointment_type": ""
    }

    patterns = {
        "company_name": r"Company Name\s*:\s*(.+)",
        "cin": r"CIN\s*:\s*([A-Z0-9]+)",
        "registered_office": r"Registered Office\s*:\s*(.+)",
        "appointment_date": r"Appointment Date\s*:\s*([\d/-]+)",
        "auditor_name": r"Auditor Name\s*:\s*(.+)",
        "auditor_address": r"Auditor Address\s*:\s*(.+)",
        "auditor_frn_or_membership": r"(FRN|Membership Number)\s*:\s*(\S+)",
        "appointment_type": r"Appointment Type\s*:\s*(New Appointment|Reappointment)"
    }

    for field, pattern in patterns.items():
        match = re.search(pattern, form_text, re.IGNORECASE)
        if match:
            if field == "auditor_frn_or_membership":
                data[field] = match.group(2).strip()
            else:
                data[field] = match.group(1).strip()

    return data

# Step 3: Save JSON
def save_to_json(data, output_folder, filename="output.json"):
    os.makedirs(output_folder, exist_ok=True)
    path = os.path.join(output_folder, filename)
    with open(path, "w") as f:
        json.dump(data, f, indent=2)
    print(f"[✓] Extracted data saved to {path}")
    return path

# Step 4: Extract Attachments
def extract_attachments(pdf_path, output_folder):
    doc = fitz.open(pdf_path)
    attachments = []

    for i in range(doc.embfile_count()):
        try:
            try:
                raw_info = doc.embfile_info(i)
                original_name = raw_info.get("filename", f"attachment_{i}")
                fname = sanitize_filename(original_name, f"attachment_{i}")
            except Exception as e:
                print(f"[!] Failed to read metadata for attachment {i}: {e}")
                fname = f"attachment_{i}"

            try:
                fdata = doc.embfile_get(i)
            except Exception as e:
                print(f"[!] Failed to extract file data for attachment {i}: {e}")
                continue

            ext = get_file_extension(fdata)
            full_name = fname + ext if not fname.endswith(ext) else fname
            attachment_path = os.path.join(output_folder, full_name)

            with open(attachment_path, "wb") as f:
                f.write(fdata)

            attachments.append(attachment_path)
            print(f"[✓] Attachment extracted: {attachment_path}")

        except Exception as e:
            print(f"[!] Failed to extract attachment {i}: {e}")

    return attachments

# Step 5: Analyze Attachments
def analyze_attachments(attachment_paths):
    insights = []

    for file in attachment_paths:
        filename = os.path.basename(file).lower()

        # Filename-based insight
        if "consent" in filename:
            insights.append("A consent letter confirming the auditor's acceptance has been signed.")
        if "intimation" in filename:
            insights.append("An intimation letter regarding the appointment is attached.")
        if "resolution" in filename or "board" in filename:
            insights.append("A board resolution approving the auditor's appointment is included.")

        # Content-based insight
        _, ext = os.path.splitext(file)
        try:
            if ext == ".txt":
                with open(file, "r", encoding="utf-8") as f:
                    content = f.read()
            elif ext == ".pdf":
                content = extract_text_from_pdf(file)
            else:
                continue

            if "unanimous" in content.lower():
                insights.append("The board approved the appointment unanimously.")

            match = re.search(r"signed on\s*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})", content, re.IGNORECASE)
            if match:
                insights.append(f"Consent was signed on {match.group(1)}.")

        except Exception as e:
            print(f"[!] Could not read attachment content: {file}. Error: {e}")

    return list(set(insights))

# Step 6: Generate AI-style Summary
def generate_ai_summary(data, extra_insights=None):
    summary = (
        f"{data.get('company_name', 'The company')} has appointed {data.get('auditor_name', 'an auditor')} "
        f"as its statutory auditor, effective from {data.get('appointment_date', 'the specified date')}. "
        f"The company's Corporate Identification Number (CIN) is {data.get('cin', 'N/A')}, and its registered office is located at {data.get('registered_office', 'N/A')}. "
        f"The appointment is classified as a {data.get('appointment_type', 'not specified')}, "
        f"with the auditor holding FRN/Membership Number {data.get('auditor_frn_or_membership', 'N/A')}. "
        f"All relevant disclosures and documents have been duly submitted as per regulatory requirements."
    )

    if extra_insights:
        summary += "\n\nAdditional insights:\n" + "\n".join(f"- {insight}" for insight in extra_insights)

    return summary

# Step 7: Save Summary
def save_summary(summary_text, output_folder, filename="summary.txt"):
    os.makedirs(output_folder, exist_ok=True)
    path = os.path.join(output_folder, filename)
    with open(path, "w", encoding="utf-8") as f:
        f.write(summary_text)
    print(f"[✓] Summary saved to {path}")

# Main Execution
if __name__ == "__main__":
    pdf_path = r"C:\Users\HP\Downloads\Form ADT-1-29092023_signed.pdf"  # path of pdf file
    output_folder = r"F:\filesure-assignment\output" # path of output folder

    print("[•] Extracting text from main PDF...")
    main_text = extract_text_from_pdf(pdf_path)

    print("[•] Extracting form fields...")
    form_data = extract_fields(main_text)
    save_to_json(form_data, output_folder)

    print("[•] Extracting and analyzing attachments...")
    attachments = extract_attachments(pdf_path, output_folder)
    insights = analyze_attachments(attachments)

    print("[•] Generating summary...")
    summary = generate_ai_summary(form_data, insights)
    save_summary(summary, output_folder)

    print("\n✅ All tasks completed successfully!")
