import streamlit as st
from transformers import pipeline, Conversation

@st.cache_resource
def load_pipelines():
    translator_en_es = pipeline("translation", model="alirezamsh/small100", src_lang="en", tgt_lang="es")
    translator_es_en = pipeline("translation", model="alirezamsh/small100", src_lang="es", tgt_lang="en")
    conversational = pipeline("conversational", model="microsoft/DialoGPT-small")
    return translator_en_es, translator_es_en, conversational

translator_en_es, translator_es_en, conversational = load_pipelines()

def detect_language(text):
    spanish_words = {"el", "la", "los", "las", "un", "una", "de", "en", "y"}
    words = set(text.lower().split())
    return "es" if words & spanish_words else "en"

def bilingual_chat(user_input):
    lang = detect_language(user_input)
    if lang == "es":
        input_en = translator_es_en(user_input, max_length=128)[0]['translation_text']
    else:
        input_en = user_input
    
    conv = Conversation(input_en)
    response_en = conversational(conv, max_length=128)[-1].generated_responses[-1]
    
    if lang == "es":
        return translator_en_es(response_en, max_length=128)[0]['translation_text']
    return response_en

st.title("Bilingual Chatbot")

if "messages" not in st.session_state:
    st.session_state.messages = []

for message in st.session_state.messages:
    with st.chat_message(message["role"]):
        st.markdown(message["content"])

if prompt := st.chat_input("Message"):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)
    
    response = bilingual_chat(prompt)
    st.session_state.messages.append({"role": "assistant", "content": response})
    with st.chat_message("assistant"):
        st.markdown(response)
