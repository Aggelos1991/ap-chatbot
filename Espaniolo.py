import streamlit as st
from openai import OpenAI
from io import BytesIO
from gtts import gTTS

st.set_page_config(page_title="🎙️ Bilingual Voice Chat", layout="wide")
st.title("🎙️ English ↔ Español Voice Chat (Cloud Version)")

# --- API key check ---
api_key = st.secrets.get("OPENAI_API_KEY", None)
if not api_key:
    st.error("❌ Add your OpenAI key in Settings → Secrets → `OPENAI_API_KEY=\"sk-...\"`")
    st.stop()

client = OpenAI(api_key=api_key)

# --- Helpers ---
def transcribe_audio(uploaded_file):
    with uploaded_file as f:
        result = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return result.text.strip()

def bilingual_chat(message):
    prompt = (
        "You are a friendly bilingual AI that chats naturally in English or Spanish. "
        "Detect the language of the user's message and respond in the same language."
    )
    completion = client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[
            {"role": "system", "content": prompt},
            {"role": "user", "content": message}
        ]
    )
    return completion.choices[0].message.content.strip()

# --- UI ---
st.subheader("🎧 Speak or type to chat!")

# ✅ now supports m4a, mp3, mp4, wav
audio_file = st.file_uploader(
    "Upload audio (.wav, .mp3, .mp4, .m4a)",
    type=["wav", "mp3", "mp4", "m4a"]
)
user_input = st.text_input("Or type your message (English or Español):")

if audio_file:
    st.audio(audio_file)
    with st.spinner("🧠 Transcribing..."):
        try:
            text = transcribe_audio(audio_file)
            st.write(f"🗣 You said: **{text}**")
            user_input = text
        except Exception as e:
            st.error(f"Transcription failed: {e}")
            st.stop()

if user_input:
    with st.spinner("🤖 Thinking..."):
        reply = bilingual_chat(user_input)
    st.markdown(f"**🤖 Bot:** {reply}")

    # Voice output
    try:
        lang = "es" if any(
            w in user_input.lower() for w in ["el", "la", "de", "que", "y", "un"]
        ) else "en"
        tts = gTTS(reply, lang=lang)
        out = BytesIO()
        tts.write_to_fp(out)
        st.audio(out.getvalue(), format="audio/mp3")
    except Exception:
        st.warning("Voice playback unavailable.")
