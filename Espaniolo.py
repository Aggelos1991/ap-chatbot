import streamlit as st
import torch
import speech_recognition as sr
from transformers import M2M100ForConditionalGeneration, M2M100Tokenizer, AutoModelForCausalLM, AutoTokenizer
from gtts import gTTS
from io import BytesIO

st.set_page_config(page_title="Bilingual Chat (EN ↔ ES)", layout="wide")
st.title("💬 Bilingual Chat (English ↔ Español)")

# ================================
# Load models (cached)
# ================================
@st.cache_resource
def load_models():
    trans_model = M2M100ForConditionalGeneration.from_pretrained("facebook/m2m100_418M")
    trans_tokenizer = M2M100Tokenizer.from_pretrained("facebook/m2m100_418M")
    conv_model = AutoModelForCausalLM.from_pretrained("microsoft/DialoGPT-medium")
    conv_tokenizer = AutoTokenizer.from_pretrained("microsoft/DialoGPT-medium")
    return trans_model, trans_tokenizer, conv_model, conv_tokenizer

trans_model, trans_tokenizer, conv_model, conv_tokenizer = load_models()

# ================================
# Translation helper
# ================================
def translate(text, src_lang, tgt_lang):
    trans_tokenizer.src_lang = src_lang
    encoded = trans_tokenizer(text, return_tensors="pt")
    generated = trans_model.generate(**encoded, forced_bos_token_id=trans_tokenizer.get_lang_id(tgt_lang), max_length=256)
    return trans_tokenizer.decode(generated[0], skip_special_tokens=True)

# ================================
# Language detection
# ================================
def detect_language(text):
    spanish_keywords = {"el", "la", "los", "las", "un", "una", "de", "en", "y", "que", "por", "para", "con"}
    words = set(text.lower().split())
    return "es" if words & spanish_keywords else "en"

# ================================
# Voice input
# ================================
def record_audio():
    recognizer = sr.Recognizer()
    with sr.Microphone() as source:
        st.info("🎙️ Speak now...")
        audio_data = recognizer.listen(source)
        st.success("✅ Audio captured!")
    try:
        text = recognizer.recognize_google(audio_data, language="es-ES")
        return text
    except sr.UnknownValueError:
        st.error("Could not understand audio.")
    except sr.RequestError:
        st.error("Speech recognition service unavailable.")
    return ""

# ================================
# Conversation with memory
# ================================
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

# ================================
# Streamlit UI
# ================================
col1, col2 = st.columns([2, 1])
with col1:
    user_input = st.text_input("💬 Type your message (English or Español):")
with col2:
    if st.button("🎤 Voice Input (Spanish)"):
        voice_text = record_audio()
        if voice_text:
            st.text(f"🗣 You said: {voice_text}")
            user_input = voice_text

if user_input:
    response = bilingual_chat(user_input)
    st.markdown(f"**🤖 Bot:** {response}")

    # Voice output
    tts = gTTS(response, lang="es" if detect_language(user_input) == "es" else "en")
    audio_bytes = BytesIO()
    tts.write_to_fp(audio_bytes)
    st.audio(audio_bytes.getvalue(), format="audio/mp3")
