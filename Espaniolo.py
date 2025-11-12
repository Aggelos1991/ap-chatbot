import streamlit as st
from openai import OpenAI
from st_audio_recorder import st_audio_recorder
import tempfile

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
def transcribe_audio_from_file(file_path):
    """Transcribe audio file (Greek, English, or Spanish)."""
    with open(file_path, "rb") as f:
        result = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return result.text.strip()


def create_vendor_email(note, lang_code, subject_text):
    tone = "in English (US)" if lang_code == "en" else "in Spanish"
    subject_clean = subject_text or "Request for Invoice Submission"

    prompt = (
        f"You are an Accounts Payable specialist writing to a vendor. "
        f"The input may be in English, Spanish, or Greek — detect and translate automatically. "
        f"Identify the vendor name (e.g., 'Iberostar') and include it naturally in the greeting (e.g., 'Dear Iberostar,'). "
        f"If no vendor name is found, use 'Dear Vendor,'. "
        f"Preserve all invoice numbers, amounts, and codes *exactly as written* by the user — "
        f"do not reformat, simplify, or expand them. "
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

    # 🧹 Clean any leftover formatting or code fences
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

    # ✨ Clean professional HTML wrapper
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


# =========================================================
# MAIN UI
# =========================================================
st.subheader("🎙 Record your voice or type your message")

col1, col2 = st.columns([2, 1])
with col1:
    st.markdown("**🎤 Record your message below:**")
    audio_bytes = audio_recorder(text="Click to record or stop", icon_size="2x")

    user_input = ""
    if audio_bytes:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".wav") as tmp_file:
            tmp_file.write(audio_bytes)
            tmp_path = tmp_file.name

        st.audio(tmp_path)
        with st.spinner("🎧 Transcribing your recording..."):
            text = transcribe_audio_from_file(tmp_path)
            st.success("✅ Transcription complete.")
            st.write(f"🗣 **You said:** {text}")
            user_input = text

    st.markdown("---")
    audio_file = st.file_uploader("Or upload an existing file", type=["wav", "mp3", "mp4", "m4a"])
    if audio_file:
        st.audio(audio_file)
        with st.spinner("🎧 Transcribing uploaded audio..."):
            text = transcribe_audio_from_file(audio_file.name)
            st.success("✅ Transcription complete.")
            st.write(f"🗣 **You said:** {text}")
            user_input = text

    st.markdown("---")
    manual_text = st.text_area("Or type manually:", height=150)
    if manual_text.strip():
        user_input = manual_text.strip()

with col2:
    target_lang = st.radio("Email language:", ["🇺🇸 English (US)", "🇪🇸 Español (ES)"])
    lang_code = "en" if "English" in target_lang else "es"
    subject_text = st.text_input("✏️ Subject line:", "")

# =========================================================
# GENERATE EMAIL
# =========================================================
if st.button("✉️ Generate Vendor Email") and user_input.strip():
    with st.spinner("🤖 Creating vendor email..."):
        email_html = create_vendor_email(user_input, lang_code, subject_text)

    st.markdown("### 📩 Preview (HTML email)")
    st.markdown(email_html, unsafe_allow_html=True)

    st.download_button(
        label="⬇️ Download HTML Email (for Outlook)",
        data=email_html.encode("utf-8"),
        file_name=f"{subject_text or 'vendor_email'}.html",
        mime="text/html"
    )

    with st.expander("ℹ️ How to use this file in Outlook (Mac or Windows)"):
        st.markdown("""
**💻 macOS Outlook (Legacy or New):**
1. Click **⬇️ Download HTML Email (for Outlook)** above.  
2. Open the `.html` file in **Safari**.  
3. Press **Cmd + A → Cmd + C** to copy the rendered content.  
4. Paste directly into Outlook’s email body. ✅  

**💼 Windows Outlook:**
1. In a new email, go to **Insert → Attach File →** choose the `.html` file.  
2. Click the small arrow next to **Insert** → select **Insert as Text**.  
3. The formatted email will appear perfectly with logo and spacing.
        """)
