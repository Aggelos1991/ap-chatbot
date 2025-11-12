import streamlit as st
from openai import OpenAI
from io import BytesIO
from gtts import gTTS

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(page_title="📧 AI Email Creator (EN ↔ ES)", layout="wide")
st.title("📧 AI Email Creator (Voice + Text + Bilingual)")

# =========================================================
# API KEY CHECK
# =========================================================
api_key = st.secrets.get("OPENAI_API_KEY", None)
if not api_key:
    st.error("❌ Add your OpenAI key in Settings → Secrets → `OPENAI_API_KEY=\"sk-...\"`")
    st.stop()

client = OpenAI(api_key=api_key)

# =========================================================
# HELPERS
# =========================================================
def transcribe_audio(uploaded_file):
    """Transcribe voice using Whisper (GPT-4o-mini)."""
    with uploaded_file as f:
        result = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return result.text.strip()

def create_email(content, target_lang):
    """Generate professional email from content in chosen language."""
    lang_instruction = "in English" if target_lang == "en" else "in Spanish"
    prompt = (
        f"You are a highly skilled professional email writer. "
        f"Rewrite the following note as a well-structured, polite, and natural business email {lang_instruction}. "
        "Keep it concise and accurate, include a subject line and closing signature. "
        f"Text:\n\n{content}"
    )
    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": "You write professional emails quickly and accurately."},
            {"role": "user", "content": prompt},
        ]
    )
    return completion.choices[0].message.content.strip()

# =========================================================
# UI
# =========================================================
st.subheader("🎙️ Upload your voice memo or type your draft")

col1, col2 = st.columns([2, 1])
with col1:
    audio_file = st.file_uploader(
        "Upload audio (.wav, .mp3, .mp4, .m4a)",
        type=["wav", "mp3", "mp4", "m4a"]
    )
    user_input = st.text_area("Or type your rough message here:", height=150)

with col2:
    target_lang = st.radio("Output language:", ["English 🇬🇧", "Español 🇪🇸"])
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

if st.button("✉️ Generate Email") and user_input.strip():
    with st.spinner("🤖 Creating email..."):
        email_text = create_email(user_input, lang_code)
    st.markdown("### 📩 Generated Email")
    st.markdown(email_text)

    # Voice playback of the email (optional)
    try:
        tts = gTTS(email_text, lang=lang_code)
        out = BytesIO()
        tts.write_to_fp(out)
        st.audio(out.getvalue(), format="audio/mp3")
    except Exception:
        st.warning("Voice playback unavailable.")
