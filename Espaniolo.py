import streamlit as st
from openai import OpenAI
from io import BytesIO
from gtts import gTTS
import base64

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
# BRANDING – INLINE LOGO
# =========================================================
logo_url = "https://upload.wikimedia.org/wikipedia/commons/1/13/Sani_Resort_logo.png"  # example placeholder
logo_html = f"<img src='{logo_url}' width='160' style='margin-top:15px;'>"

signature_block = f"""
<br><br>
Best regards,<br>
<b>Angelos Keramaris</b><br>
AP Process Officer – Sani Ikos Group<br>
{logo_html}
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

def create_vendor_email(note, lang_code):
    tone = "in English" if lang_code == "en" else "in Spanish"

    prompt = (
        f"You are an Accounts Payable specialist writing directly to a vendor. "
        f"The user may speak in English, Spanish, or Greek — detect it automatically. "
        f"Translate the content if needed and write a professional, polite, and concise vendor email {tone}. "
        f"If invoices or credit notes are mentioned, request them to be sent to ap.iberia@ikosresorts.com. "
        f"Include a proper subject line and greeting suitable for external vendors. "
        f"Finish the email with the signature block for Angelos Keramaris, AP Process Officer – Sani Ikos Group, "
        f"and leave two line breaks before it. Do not use placeholders. "
        f"User note:\n\n{note}"
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

    # Inject the signature visually for Streamlit (HTML)
    styled_email = email_text + signature_block
    st.markdown("### 📩 Generated Vendor Email")
    st.markdown(styled_email, unsafe_allow_html=True)

    # Optional voice playback of the email
    try:
        tts = gTTS(email_text, lang=lang_code)
        out = BytesIO()
        tts.write_to_fp(out)
        st.audio(out.getvalue(), format="audio/mp3")
    except Exception:
        st.warning("Voice playback unavailable.")
