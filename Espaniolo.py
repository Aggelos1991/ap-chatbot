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
# HELPERS
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
    subject_html = f"<b>Subject:</b> {subject_text or 'Vendor communication'}<br><br>"

    prompt = (
        f"You are an Accounts Payable specialist writing directly to a vendor. "
        f"The input may be in English, Spanish, or Greek — detect and translate automatically. "
        f"Detect the vendor name mentioned by the user and use it in the greeting. "
        f"If invoices or credit notes are mentioned, request them to be sent to ap.iberia@ikosresorts.com. "
        f"Write a clear, polite, and professional vendor email {tone}. "
        f"Include a greeting, concise body, and a single 'Best regards' closing. "
        f"Do NOT include another signature; append this block at the end:\n{signature_block}\n"
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
    return subject_html + email_body

# =========================================================
# UI
# =========================================================
st.subheader("🎙️ Upload a voice memo or type your note")

col1, col2 = st.columns([2, 1])
with col1:
    audio_file = st.file_uploader(
        "Upload voice (.wav, .mp3, .mp4, .m4a)",
        type=["wav", "mp3", "mp4", "m4a"]
    )
    user_input = st.text_area("Or type your note (English / Español / Ελληνικά):", height=150)

with col2:
    target_lang = st.radio("Email language:", ["🇺🇸 English (US)", "🇪🇸 Español (ES)"])
    lang_code = "en" if "English" in target_lang else "es"
    subject_text = st.text_input("✏️ Subject line (optional):", "")

# =========================================================
# TRANSCRIBE
# =========================================================
if audio_file:
    st.audio(audio_file)
    with st.spinner("🎧 Transcribing..."):
        try:
            spoken = transcribe_audio(audio_file)
            st.success("✅ Transcribed successfully.")
            st.write(f"🗣 **You said:** {spoken}")
            user_input = spoken
        except Exception as e:
            st.error(f"Transcription failed: {e}")
            st.stop()

# =========================================================
# GENERATE
# =========================================================
if st.button("✉️ Generate Vendor Email") and user_input.strip():
    with st.spinner("🤖 Creating email..."):
        email_html = create_vendor_email(user_input, lang_code, subject_text)

    st.markdown("### 📩 Generated Vendor Email")
    st.markdown(email_html, unsafe_allow_html=True)

    # ---- Plain text version (for preview)
    plain = re.sub("<[^>]*>", "", email_html)

    # ---- Download HTML version
    st.download_button(
        label="⬇️ Download HTML Email (for Outlook)",
        data=email_html.encode("utf-8"),
        file_name="vendor_email.html",
        mime="text/html"
    )

    st.info("💡 To use in Outlook: New Email → Insert → Attach File → Insert as Text. "
            "Your email will render perfectly with logo and formatting.")
