import streamlit as st
from openai import OpenAI
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
# INLINE LOGO (your uploaded file)
# =========================================================
with open("3116FEF9-419C-47D2-A832-15A27AD9D1F8.jpeg", "rb") as img:
    logo_base64 = base64.b64encode(img.read()).decode("utf-8")

signature_block = f"""
<br><br>
<table style='margin-top:10px;'>
<tr>
<td style='vertical-align:top; padding-right:10px;'>
    <img src='data:image/jpeg;base64,{logo_base64}' width='180'>
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
    """Transcribe voice in Greek, English, or Spanish."""
    with uploaded_file as f:
        result = client.audio.transcriptions.create(
            model="gpt-4o-mini-transcribe",
            file=f
        )
    return result.text.strip()

def create_vendor_email(note, lang_code):
    """Generate polished vendor email with vendor name detection."""
    tone = "in English" if lang_code == "en" else "in Spanish"

    prompt = (
        f"You are an Accounts Payable specialist writing directly to a vendor. "
        f"The input may be in English, Spanish, or Greek — detect it automatically. "
        f"Detect the vendor name mentioned by the user and use it in the greeting (e.g., 'Dear Iberostar,'). "
        f"If no vendor name is found, use 'Dear Vendor,'. "
        f"Translate and rewrite the content as a clear, polite vendor email {tone}. "
        f"If invoices or credit notes are mentioned, request them to be sent to ap.iberia@ikosresorts.com. "
        f"Include a subject line, greeting, concise body, and one single 'Best regards' closing. "
        f"End with the provided signature block (no duplicates):\n\n{signature_block}\n\n"
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

    # --- Copy to Clipboard button ---
    copy_script = f"""
    <script>
    function copyToClipboard() {{
        navigator.clipboard.writeText({email_text!r});
        alert("📋 Email copied to clipboard!");
    }}
    </script>
    <button onclick="copyToClipboard()" style="background-color:#0066cc;color:white;border:none;
        padding:10px 18px;border-radius:6px;cursor:pointer;font-weight:bold;">
        Copy Email to Clipboard
    </button>
    """
    st.markdown(copy_script, unsafe_allow_html=True)
