import streamlit as st
from openai import OpenAI
from transformers import (
    M2M100ForConditionalGeneration,
    M2M100Tokenizer,
    AutoModelForCausalLM,
    AutoTokenizer
)
from io import BytesIO
from gtts import gTTS

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(page_title="🎙️ Bilingual Voice Chat", layout="wide")
st.title("🎙️ English ↔ Español Voice Chat")

# =========================================================
# API KEY CHECK
# =========================================================
api_key = None
try:
    api_key = st.secrets.get("OPENAI_API_KEY")
except Exception:
    pass

if not api_key:
    st.error(
        "❌ OpenAI API key not found.\n\n"
        "Please add it in **Settings → Secrets** like this:\n\n"
        '```\nOPENAI_API_KEY="sk-your-key-here"\n```'
    )
    st.stop()

client = OpenAI(api_key=api_key)

# =========================================================
# LOAD MODELS
# =========================================================
@st.cache_resource
def load_models():
    trans_model = M2M100ForConditionalGeneration.from_pretrained("facebook/m2m100_418M")
    trans_tokenizer = M2M100Tokenizer.from_pretrained("facebook/m2m100_418M")
    conv_model = AutoModelForCausalLM.from_pretrained("microsoft/DialoGPT-medium")
    conv_tokenizer = AutoTokenizer.from_pretrained("microsoft/DialoGPT-medium")
    return trans_model, trans_tokenizer, conv_model, conv_tokenizer

trans_model, trans_tokenizer, conv_model, conv_tokenizer = load_models()

# =========================================================
# HELPERS
# =========================================================
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

def transcribe_audio(uploaded_file):
    with uploaded_file as f:
        transcript = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return transcript.text.strip()

# =========================================================
# MEMORY + CHAT LOGIC
# =========================================================
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

# =========================================================
# STREAMLIT UI
# =========================================================
st.subheader("🎧 Speak or type to chat!")

audio_file = st.file_uploader("Upload an audio file (.wav, .mp3)", type=["wav", "mp3"])
user_input = st.text_input("Or type your message (English or Español):")

if audio_file is not None:
    st.audio(audio_file)
    with st.spinner("🧠 Transcribing your voice..."):
        try:
            spoken_text = transcribe_audio(audio_file)
            st.write(f"🗣 You said: **{spoken_text}**")
            user_input = spoken_text
        except Exception as e:
            st.error(f"Transcription failed: {e}")
            st.stop()

if user_input:
    with st.spinner("🤖 Thinking..."):
        response = bilingual_chat(user_input)
    st.markdown(f"**🤖 Bot:** {response}")

    lang_reply = "es" if detect_language(user_input) == "es" else "en"
    try:
        tts = gTTS(response, lang=lang_reply)
        audio_out = BytesIO()
        tts.write_to_fp(audio_out)
        st.audio(audio_out.getvalue(), format="audio/mp3")
    except Exception:
        st.warning("Unable to generate voice output.")
