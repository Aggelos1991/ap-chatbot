import streamlit as st
from transformers import M2M100ForConditionalGeneration, Conversation, AutoModelForCausalLM, AutoTokenizer
from tokenization_small100 import SMALL100Tokenizer

@st.cache_resource
def load_models():
    trans_model = M2M100ForConditionalGeneration.from_pretrained("alirezamsh/small100")
    tokenizer = SMALL100Tokenizer.from_pretrained("alirezamsh/small100")
    conv_model = AutoModelForCausalLM.from_pretrained("microsoft/DialoGPT-small")
    conv_tokenizer = AutoTokenizer.from_pretrained("microsoft/DialoGPT-small")
    return trans_model, tokenizer, conv_model, conv_tokenizer

trans_model, tokenizer, conv_model, conv_tokenizer = load_models()

def translate(text, src_lang, tgt_lang):
    tokenizer.src_lang = src_lang
    tokenizer.tgt_lang = tgt_lang
    encoded = tokenizer(text, return_tensors="pt")
    generated = trans_model.generate(**encoded, max_length=128)
    return tokenizer.decode(generated[0], skip_special_tokens=True)

def detect_language(text):
    spanish_words = {"el", "la", "los", "las", "un", "una", "de", "en", "y"}
    words = set(text.lower().split())
    return "es" if words & spanish_words else "en"

def bilingual_chat(user_input):
    lang = detect_language(user_input)
    src = "es" if lang == "es" else "en"
    tgt = "en" if lang == "es" else "es"
    input_en = translate(user_input, src, "en") if lang == "es" else user_input
    
    inputs = conv_tokenizer.encode(input_en + conv_tokenizer.eos_token, return_tensors="pt")
    response = conv_model.generate(inputs, max_length=128)
    response_en = conv_tokenizer.decode(response[:, inputs.shape[-1]:][0], skip_special_tokens=True)
    
    return translate(response_en, "en", tgt) if lang == "es" else response_en

# Rest of Streamlit code remains the same
