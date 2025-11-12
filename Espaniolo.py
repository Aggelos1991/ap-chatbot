import streamlit as st
from openai import OpenAI

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
# LOGO + SIGNATURE (inline HTML)
# =========================================================
logo_url = "https://career.unipi.gr/career_cv/logo_comp/81996-new-logo.png"

signature_block = f"""
<br><br>
<table style='margin-top:10px;'>
<tr>
<td style='vertical-align:top; padding-right:10px;'>
    <img src='{logo_url}' width='180'>
</td>
<td style='vertical-align:top;'>
    <b>Angelos Keramaris</b><br>
    AP Process Officer – Sani Ikos Group
</td>
</tr>
</table>
"""

# =========================================================
# HELPER
# =========================================================
def create_vendor_email(note, lang_code):
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
user_input = st.text_area("Write or speak your note (Greek, English, Spanish):", height=150)
target_lang = st.radio("Output language:", ["English 🇬🇧", "Español 🇪🇸"])
lang_code = "en" if "English" in target_lang else "es"

if st.button("✉️ Generate Vendor Email") and user_input.strip():
    with st.spinner("🤖 Creating vendor email..."):
        email_html = create_vendor_email(user_input, lang_code)

    st.markdown("### 📩 Formatted Vendor Email")
    st.markdown(email_html, unsafe_allow_html=True)

    # ✅ Clean HTML for Outlook copy (no <html> / <body> tags)
    html_clean = email_html.replace("\n", " ").replace("'", "\\'")

    copy_script = f"""
    <button id="copyHTML" style="background-color:#0066cc;color:white;border:none;
        padding:10px 18px;border-radius:6px;cursor:pointer;font-weight:bold;">
        📋 Copy Formatted Email
    </button>
    <script>
    const btn = document.getElementById("copyHTML");
    btn.addEventListener("click", async () => {{
        try {{
            const html = '{html_clean}';
            const blob = new Blob([html], {{ type: 'text/html' }});
            const item = new ClipboardItem({{'text/html': blob}});
            await navigator.clipboard.write([item]);
            btn.innerText = "✅ Copied!";
            setTimeout(() => btn.innerText = "📋 Copy Formatted Email", 2000);
        }} catch (err) {{
            alert("Clipboard blocked by browser security settings.");
        }}
    }});
    </script>
    """
    st.markdown(copy_script, unsafe_allow_html=True)
