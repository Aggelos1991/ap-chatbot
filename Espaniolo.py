import streamlit as st
from openai import OpenAI
import re

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(page_title="📧 Vendor Email Creator – Sani Ikos Group", layout="wide")
st.title("📧 Vendor Email Creator – Sani Ikos Group")

# =========================================================
# API KEY
# =========================================================
api_key = st.secrets.get("OPENAI_API_KEY")
if not api_key:
    st.error("❌ Please add your API key in Streamlit → Secrets → OPENAI_API_KEY")
    st.stop()

client = OpenAI(api_key=api_key)

# =========================================================
# LOGO + SIGNATURE
# =========================================================
logo_url = "https://career.unipi.gr/career_cv/logo_comp/81996-new-logo.png"
signature_block = f"""
<br><br>
<table style='margin-top:10px;'>
<tr>
<td style='vertical-align:top; padding-right:10px;'>
    <img src='{logo_url}' width='180'>
</td>
<td style='vertical-align:top;'>
    <b>Angelos Keramaris</b><br>
    AP Process Officer – Sani Ikos Group
</td>
</tr>
</table>
"""

# =========================================================
# HELPER FUNCTIONS
# =========================================================
def transcribe_audio(uploaded_file):
    with uploaded_file as f:
        result = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return result.text.strip()

def create_vendor_email(note, lang_code, subject_text):
    tone = "in English (US)" if lang_code == "en" else "in Spanish"
    subject_clean = subject_text or "Request for Invoice Submission"

    prompt = (
        f"You are an Accounts Payable specialist writing directly to a vendor. "
        f"The input may be in English, Spanish, or Greek — detect and translate automatically. "
        f"Detect the vendor name mentioned by the user and use it in the greeting (e.g., 'Dear Iberostar,'). "
        f"If no vendor name is found, use 'Dear Vendor,'. "
        f"Write a clear, polite, and professional vendor email {tone}. "
        f"If invoices or credit notes are mentioned, request them to be sent to ap.iberia@ikosresorts.com. "
        f"Include a greeting, concise body, and one closing only. "
        f"Append this HTML signature (no duplicates):\n{signature_block}\n"
        f"User note:\n{note}"
    )

    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are a bilingual AP email expert."},
            {"role": "user", "content": prompt}
        ]
    )

    email_body = completion.choices[0].message.content.strip()

    # ✨ Modern formatted HTML wrapper
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
    margin: 40px;
    background-color: #f6f7f9;
}}
.container {{
    max-width: 720px;
    margin: auto;
    background: #fff;
    border-radius: 10px;
    padding: 35px 45px;
    box-shadow: 0 0 10px rgba(0,0,0,0.08);
}}
h2 {{
    font-size: 18px;
    color: #003366;
    border-bottom: 1px solid #ddd;
    padding-bottom: 8px;
    margin-bottom: 25px;
}}
.signature {{
    margin-top: 30px;
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

# =========================================================
# UI
# =========================================================
st.subheader("🎙️ Upload a voice memo or type your message for the vendor")

col1, col2 = st.columns([2, 1])
with col1:
    audio_file = st.file_uploader(
        "Upload voice memo (.wav, .mp3, .mp4, .m4a)",
        type=["wav", "mp3", "mp4", "m4a"]
    )
    user_input = st.text_area("Or type your note (in English / Español / Ελληνικά):", height=150)

with col2:
    target_lang = st.radio("Email language:", ["🇺🇸 English (US)", "🇪🇸 Español (ES)"])
    lang_code = "en" if "English" in target_lang else "es"
    subject_text = st.text_input("✏️ Subject line:", "")

# =========================================================
# AUDIO TRANSCRIPTION
# =========================================================
if audio_file:
    st.audio(audio_file)
    with st.spinner("🎧 Transcribing..."):
        try:
            spoken_text = transcribe_audio(audio_file)
            st.success("✅ Transcribed successfully.")
            st.write(f"🗣 **You said:** {spoken_text}")
            user_input = spoken_text
        except Exception as e:
            st.error(f"Transcription failed: {e}")
            st.stop()

# =========================================================
# GENERATE EMAIL
# =========================================================
if st.button("✉️ Generate Vendor Email") and user_input.strip():
    with st.spinner("🤖 Creating email..."):
        email_html = create_vendor_email(user_input, lang_code, subject_text)

    st.markdown("### 📩 Preview (HTML email)")
    st.markdown(email_html, unsafe_allow_html=True)

    # ---- Download HTML version for Outlook
    st.download_button(
        label="⬇️ Download HTML Email (for Outlook)",
        data=email_html.encode("utf-8"),
        file_name=f"{subject_text or 'vendor_email'}.html",
        mime="text/html"
    )

    with st.expander("ℹ️ How to use this file in Outlook (Mac or Windows)"):
        st.markdown("""
**💻 If you're using macOS Outlook (Legacy or New):**
1. Click **⬇️ Download HTML Email (for Outlook)** above.  
2. Open the `.html` file in **Safari**.  
3. Press **Cmd +A → Cmd +C** to copy the rendered content.  
4. Paste directly into your Outlook message body. ✅  

**💼 If using Windows Outlook:**
1. In a new email → **Insert → Attach File →** browse to your `.html` file.  
2. Click the small arrow beside **Insert** → select **Insert as Text**.  

Your email will display perfectly formatted with logo, spacing, and signature.  
        """)
