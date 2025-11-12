import streamlit as st
import torch
import sounddevice as sd
import wavio
from io import BytesIO
from gtts import gTTS
from openai import OpenAI
from transformers import (
    M2M100ForConditionalGeneration, 
    M2M100Tokenizer, 
    AutoModelForCausalLM, 
    AutoTokenizer
)

# =========================
# CONFIGURATION
# =========================
st.set_page_config(page_title="🎙️ Bilingual Voice Chat", layout="wide")
st.title("🎙️ English ↔ Español Voice Chat")

client = OpenAI(api_key=st.secrets.get("OPENAI_API_KEY"))

# =========================
# LOAD MODELS
# =========================
@st.cache_resource
def load_models():
    trans_model = M2M100ForConditionalGeneration.from_pretrained("facebook/m2m100_418M")
    trans_tokenizer = M2M100Tokenizer.from_pretrained("facebook/m2m100_418M")
    conv_model = AutoModelForCausalLM.from_pretrained("microsoft/DialoGPT-medium")
    conv_tokenizer = AutoTokenizer.from_pretrained("microsoft/DialoGPT-medium")
    return trans_model, trans_tokenizer, conv_model, conv_tokenizer

trans_model, trans_tokenizer, conv_model, conv_tokenizer = load_models()

# =========================
# TRANSLATION HELPERS
# =========================
def translate(text, src_lang, tgt_lang):
    trans_tokenizer.src_lang = src_lang
    encoded = trans_tokenizer(text, return_tensors="pt")
    generated = trans_model.generate(
        **encoded,
        forced_bos_token_id=trans_tokenizer.get_lang_id(tgt_lang),
        max_length=256
    )
    return trans_tokenizer.decode(generated[0], skip_special_tokens=True)

def detect_language(text):
    spanish = {"el", "la", "los", "las", "un", "una", "de", "en", "y", "que", "por", "para"}
    words = set(text.lower().split())
    return "es" if words & spanish else "en"

# =========================
# RECORD VOICE
# =========================
def record_voice(seconds=6, samplerate=44100):
    st.info("🎤 Recording... speak now!")
    audio = sd.rec(int(seconds * samplerate), samplerate=samplerate, channels=1, dtype='int16')
    sd.wait()
    st.success("✅ Recording finished!")
    filename = "voice_input.wav"
    wavio.write(filename, audio, samplerate, sampwidth=2)
    return filename

def transcribe_audio(file_path):
    with open(file_path, "rb") as f:
        transcript = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return transcript.text.strip()

# =========================
# CHAT MEMORY + LOGIC
# =========================
if "chat_history" not in st.session_state:
    st.session_state.chat_history = []

def bilingual_chat(user_input):
    lang = detect_language(user_input)
    src = "es" if lang == "es" else "en"
    tgt = "en" if lang == "es" else "es"

    input_en = translate(user_input, src, "en") if src == "es" else user_input
    st.session_state.chat_history.append(input_en + conv_tokenizer.eos_token)

    input_ids = conv_tokenizer.encode("".join(st.session_state.chat_history), return_tensors="pt")
    response_ids = conv_model.generate(input_ids, max_length=256, pad_token_id=conv_tokenizer.eos_token_id)
    response_en = conv_tokenizer.decode(response_ids[:, input_ids.shape[-1]:][0], skip_special_tokens=True)

    st.session_state.chat_history.append(response_en + conv_tokenizer.eos_token)
    return translate(response_en, "en", tgt) if src == "es" else response_en

# =========================
# STREAMLIT UI
# =========================
col1, col2 = st.columns([2, 1])

with col1:
    user_input = st.text_input("💬 Type your message (English or Español):")

with col2:
    if st.button("🎙️ Record Voice"):
        audio_path = record_voice()
        st.audio(audio_path)
        text_from_audio = transcribe_audio(audio_path)
        st.write(f"🗣 You said: **{text_from_audio}**")
        user_input = text_from_audio

if user_input:
    response = bilingual_chat(user_input)
    st.markdown(f"**🤖 Bot:** {response}")

    # Speak the reply
    lang_reply = "es" if detect_language(user_input) == "es" else "en"
    tts = gTTS(response, lang=lang_reply)
    audio_out = BytesIO()
    tts.write_to_fp(audio_out)
    st.audio(audio_out.getvalue(), format="audio/mp3")
