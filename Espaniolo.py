def create_vendor_email(note, lang_code, subject_text):
    tone = "in English (US)" if lang_code == "en" else "in Spanish"
    subject_clean = subject_text or "Request for Invoice Submission"

    # Updated instruction for accuracy + structure
    prompt = (
        f"You are an Accounts Payable specialist writing to a vendor. "
        f"The input may be in English, Spanish, or Greek — detect and translate automatically. "
        f"Identify the vendor name (e.g., 'Iberostar') and include it naturally in the greeting (e.g., 'Dear Iberostar,'). "
        f"If no vendor name is found, use 'Dear Vendor,'. "
        f"Preserve all invoice numbers, amounts, and codes *exactly as written* by the user — do not reformat, simplify, or expand them. "
        f"Write a concise, polite vendor email {tone} following this exact layout:\n\n"
        f"Dear [Vendor],\n\n"
        f"[Short body — 2 or 3 clear paragraphs.]\n\n"
        f"Thank you for your attention to this matter.\n\n"
        f"Best regards,\n\n"
        f"[Signature block]\n\n"
        f"Do not include markdown syntax (no ```html, no ``` blocks). "
        f"Use <p> and <br> for spacing and append this signature once:\n{signature_block}\n"
        f"User note:\n{note}"
    )

    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a bilingual AP email expert writing polished HTML vendor emails."},
            {"role": "user", "content": prompt}
        ]
    )

    email_body = completion.choices[0].message.content.strip()

    # 🧹 Clean out any stray code blocks or fences
    email_body = (
        email_body.replace("```html", "")
        .replace("```", "")
        .replace("\n\n", "</p><p>")
        .replace("\n", " ")
        .replace("Dear ", "<p>Dear ")
        .replace("Thank you for your attention to this matter.", "</p><p>Thank you for your attention to this matter.</p>")
        .replace("Best regards,", "<br><br><strong>Best regards,</strong><br>")
    )

    if not email_body.startswith("<p>"):
        email_body = f"<p>{email_body}</p>"

    # ✨ Polished HTML template with good readability
    email_html = f"""
<html>
<head>
<meta charset='utf-8'>
<title>{subject_clean}</title>
<style>
body {{
    font-family: 'Segoe UI', Calibri, Arial, sans-serif;
    font-size: 15px;
    color: #222;
    line-height: 1.6;
    margin: 0;
    background-color: #f2f4f7;
}}
.container {{
    max-width: 720px;
    margin: 40px auto;
    background: #ffffff;
    border-radius: 10px;
    padding: 35px 45px;
    box-shadow: 0 3px 12px rgba(0,0,0,0.08);
}}
h2 {{
    font-size: 18px;
    color: #003366;
    border-bottom: 1px solid #ddd;
    padding-bottom: 8px;
    margin-bottom: 25px;
}}
p {{
    margin: 12px 0;
}}
br {{
    line-height: 1.8;
}}
.signature {{
    margin-top: 35px;
}}
</style>
</head>
<body>
<div class="container">
    <h2>{subject_clean}</h2>
    {email_body}
</div>
</body>
</html>
"""
    return email_html
