import streamlit as st
from openai import OpenAI

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(page_title="📧 Vendor Email Creator – Sani Ikos Group", layout="wide")
st.title("📧 Vendor Email Creator – Sani Ikos Group")

# =========================================================
# API KEY CHECK
# =========================================================
api_key = st.secrets.get("OPENAI_API_KEY", None)
if not api_key:
    st.error("❌ Add your OpenAI key in Settings → Secrets → `OPENAI_API_KEY=\"sk-...\"`")
    st.stop()

client = OpenAI(api_key=api_key)

# =========================================================
# BRANDING
# =========================================================
logo_url = "https://upload.wikimedia.org/wikipedia/commons/thumb/1/13/Sani_Resort_logo.png/320px-Sani_Resort_logo.png"

signature_block = f"""
<br><br>
<table style='margin-top:10px;'>
<tr>
<td style='vertical-align:top; padding-right:10px;'>
    <img src='{logo_url}' width='140'>
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
    """Transcribe audio in any language (Greek, English, Spanish)."""
    with uploaded_file as f:
        result = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return result.text.strip()

def create_vendor_email(note, lang_code):
    """Generate bilingual vendor email with clean format and correct tone."""
    tone = "in English" if lang_code == "en" else "in Spanish"
    prompt = (
        f"You are an Accounts Payable specialist writing directly to a vendor. "
        f"The input may be in English, Spanish, or Greek — detect and translate automatically. "
        f"Compose a clear, professional, and polite vendor email {tone}. "
        f"If invoices or credit notes are mentioned, request them to be sent to ap.iberia@ikosresorts.com. "
        f"Include a subject line, greeting, concise body, and one single 'Best regards,' closing. "
        f"End with the following signature block exactly as shown, without repeating 'Best regards':\n\n"
        f"{signature_block}\n\n"
        f"User note:\n{note}"
    )

    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You are an expert bilingual AP vendor communication specialist."},
            {"role": "user", "content": prompt},
        ]
    )
    return completion.choices[0].message.content.strip()

# =========================================================
# UI
# =========================================================
st.subheader("🎙️ Upload your voice memo or type your message for the vendor")

col1, col2 = st.columns([2, 1])
with col1:
    audio_file = st.file_uploader(
        "Upload audio (.wav, .mp3, .mp4, .m4a)",
        type=["wav", "mp3", "mp4", "m4a"]
    )
    user_input = st.text_area("Or type your note (in English, Español, or Ελληνικά):", height=150)

with col2:
    target_lang = st.radio("Output email language:", ["English 🇬🇧", "Español 🇪🇸"])
    lang_code = "en" if "English" in target_lang else "es"

# =========================================================
# PROCESS
# =========================================================
if audio_file:
    st.audio(audio_file)
    with st.spinner("🧠 Transcribing your recording..."):
        try:
            text = transcribe_audio(audio_file)
            st.success("✅ Transcription complete.")
            st.write(f"**You said:** {text}")
            user_input = text
        except Exception as e:
            st.error(f"Transcription failed: {e}")
            st.stop()

if st.button("✉️ Generate Vendor Email") and user_input.strip():
    with st.spinner("🤖 Creating vendor email..."):
        email_text = create_vendor_email(user_input, lang_code)

    st.markdown("### 📩 Generated Vendor Email")
    st.markdown(email_text, unsafe_allow_html=True)
