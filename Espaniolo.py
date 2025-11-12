# Install if needed: pip install gtts speechrecognition pyaudio

import speech_recognition as sr
from gtts import gTTS
import os
import playsound

# Existing: translators, conversational, detect_language

def speak(text, lang='en'):
    tts = gTTS(text=text, lang=lang)
    tts.save("response.mp3")
    playsound.playsound("response.mp3")
    os.remove("response.mp3")

def listen():
    r = sr.Recognizer()
    with sr.Microphone() as source:
        audio = r.listen(source)
    try:
        return r.recognize_google(audio, language="en-US,es-ES")
    except:
        return ""

def bilingual_voice_chat():
    while True:
        user_input = listen()
        if user_input.lower() in ["exit", "salir"]:
            break
        lang = detect_language(user_input)
        response = bilingual_chat(user_input)  # From previous
        speak(response, lang)

bilingual_voice_chat()
