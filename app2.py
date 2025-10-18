import streamlit as st
import pandas as pd
import re
from streamlit.components.v1 import html

# ======================================
# CONFIGURATION
# ======================================
st.set_page_config(page_title="🦖 ReconRaptor — Vendor Reconciliation", layout="wide")
st.title("🦖 ReconRaptor — Vendor Invoice Reconciliation")

# ======================================
# INTERACTIVE 3D SECTION
# ======================================
three_js = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <style>
    html, body {
      margin: 0;
      height: 100%;
      overflow: hidden;
      background-color: #d9e2dd;
    }
    #container {
      width: 100%;
      height: 100%;
    }
    canvas {
      width: 100% !important;
      height: 100% !important;
      display: block;
    }
  </style>
</head>
<body>
  <div id="container"></div>

  <script src="https://unpkg.com/three@0.158.0/build/three.min.js"></script>
  <script src="https://unpkg.com/three@0.158.0/examples/js/controls/OrbitControls.js"></script>

  <script>
    const container = document.getElementById('container');

    const renderer = new THREE.WebGLRenderer({ antialias: true });
    renderer.setSize(container.clientWidth, container.clientHeight);
    renderer.setPixelRatio(window.devicePixelRatio);
    container.appendChild(renderer.domElement);

    const scene = new THREE.Scene();
    scene.background = new THREE.Color('#e9ede8');

    const camera = new THREE.PerspectiveCamera(50, container.clientWidth / container.clientHeight, 0.1, 100);
    camera.position.set(3, 2, 5);

    const controls = new THREE.OrbitControls(camera, renderer.domElement);
    controls.enableDamping = true;

    const hemiLight = new THREE.HemisphereLight(0xffffff, 0x444444, 0.9);
    hemiLight.position.set(0, 20, 0);
    scene.add(hemiLight);

    const dirLight = new THREE.DirectionalLight(0xffffff, 0.8);
    dirLight.position.set(5, 10, 7);
    scene.add(dirLight);

    const planeGeometry = new THREE.PlaneGeometry(20, 20);
    const planeMaterial = new THREE.MeshStandardMaterial({ color: '#cfd8dc', roughness: 1 });
    const plane = new THREE.Mesh(planeGeometry, planeMaterial);
    plane.rotation.x = -Math.PI / 2;
    plane.position.y = -0.5;
    scene.add(plane);

    const geometry = new THREE.BoxGeometry(1, 1, 1);
    const material = new THREE.MeshStandardMaterial({
      color: '#009688',
      metalness: 0.3,
      roughness: 0.4,
      emissive: '#00bfa5',
      emissiveIntensity: 0.2
    });
    const cube = new THREE.Mesh(geometry, material);
    cube.position.y = 0.5;
    scene.add(cube);

    let t = 0;
    function animate() {
      requestAnimationFrame(animate);
      t += 0.02;
      cube.rotation.y += 0.01;
      cube.rotation.x += 0.005;
      cube.position.y = 0.5 + Math.sin(t) * 0.15;
      controls.update();
      renderer.render(scene, camera);
    }

    window.addEventListener('resize', () => {
      const width = container.clientWidth;
      const height = container.clientHeight;
      renderer.setSize(width, height);
      camera.aspect = width / height;
      camera.updateProjectionMatrix();
    });

    animate();
  </script>
</body>
</html>
"""

with st.container():
    st.subheader("🌴 Resort Interactive 3D Visualization")
    html(three_js, height=600)

# ======================================
# HELPERS
# ======================================
def normalize_number(v):
    """Convert numeric strings like '1.234,56' or '1,234.56' safely to float."""
    if v is None or str(v).strip() == "":
        return 0.0
    s = str(v).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") == 1:
        if s.find(",") > s.find("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif s.count(",") == 1:
        s = s.replace(",", ".")
    elif s.count(".") > 1:
        s = s.replace(".", "", s.count(".") - 1)
    try:
        return float(s)
    except:
        return 0.0


def normalize_columns(df, tag):
    """Map multilingual headers to unified names."""
    mapping = {
        "invoice": [
            "invoice", "factura", "fact", "nº", "num", "numero", "número",
            "document", "doc", "ref", "referencia", "nº factura", "num factura", "alternative document",
            "αρ.", "αριθμός", "νουμερο", "νούμερο", "no", "παραστατικό", "αρ. τιμολογίου", "αρ. εγγράφου"
        ],
        "credit": [
            "credit", "haber", "credito", "crédito", "nota de crédito", "nota crédito",
            "abono", "abonos", "importe haber", "valor haber",
            "πίστωση", "πιστωτικό", "πιστωτικό τιμολόγιο", "πίστωση ποσού"
        ],
        "debit": [
            "debit", "debe", "cargo", "importe", "importe total", "valor", "monto",
            "amount", "document value", "charge", "total", "totale", "totales", "totals",
            "base imponible", "importe factura", "importe neto",
            "χρέωση", "αξία", "αξία τιμολογίου"
        ],
        "reason": [
            "reason", "motivo", "concepto", "descripcion", "descripción",
            "detalle", "detalles", "razon", "razón",
            "observaciones", "comentario", "comentarios", "explicacion",
            "αιτιολογία", "περιγραφή", "παρατηρήσεις", "σχόλια", "αναφορά", "αναλυτική περιγραφή"
        ],
        "cif": [
            "cif", "nif", "vat", "iva", "tax", "id fiscal", "número fiscal", "num fiscal", "code",
            "αφμ", "φορολογικός αριθμός", "αριθμός φορολογικού μητρώου"
        ],
        "date": [
            "date", "fecha", "fech", "data", "fecha factura", "fecha doc", "fecha documento",
            "ημερομηνία", "ημ/νία", "ημερομηνία έκδοσης", "ημερομηνία παραστατικού"
        ],
    }

    rename_map = {}
    cols_lower = {c: str(c).strip().lower() for c in df.columns}

    for key, aliases in mapping.items():
        for col, low in cols_lower.items():
            if any(a in low for a in aliases):
                rename_map[col] = f"{key}_{tag}"

    out = df.rename(columns=rename_map)

    # Ensure debit/credit exist
    for required in ["debit", "credit"]:
        cname = f"{required}_{tag}"
        if cname not in out.columns:
            out[cname] = 0.0

    return out


# ======================================
# CORE MATCHING + PAYMENTS FUNCTIONS
# ======================================
# (Your entire invoice-matching and extract_payments logic stays unchanged here)
# [Paste all from your previous version — nothing to modify in logic below]

# --- Paste your match_invoices() function here ---
# --- Paste your extract_payments() function here ---

# ======================================
# STREAMLIT UI SECTION
# ======================================
uploaded_erp = st.file_uploader("📂 Upload ERP Export (Excel)", type=["xlsx"])
uploaded_vendor = st.file_uploader("📂 Upload Vendor Statement (Excel)", type=["xlsx"])

if uploaded_erp and uploaded_vendor:
    erp_raw = pd.read_excel(uploaded_erp, dtype=str)
    ven_raw = pd.read_excel(uploaded_vendor, dtype=str)

    erp_df = normalize_columns(erp_raw, "erp")
    ven_df = normalize_columns(ven_raw, "ven")

    with st.spinner("Reconciling invoices..."):
        matched, erp_missing, ven_missing = match_invoices(erp_df, ven_df)
        erp_pay, ven_pay, matched_pay = extract_payments(erp_df, ven_df)

    st.success("✅ Reconciliation complete")

    def highlight_row(row):
        if row["Status"] == "Match":
            return ['background-color: #2e7d32; color: white'] * len(row)
        elif row["Status"] == "Difference":
            return ['background-color: #f9a825; color: black'] * len(row)
        return [''] * len(row)

    st.subheader("📊 Matched / Differences")
    if not matched.empty:
        st.dataframe(matched.style.apply(highlight_row, axis=1), use_container_width=True)
    else:
        st.info("No matches found.")

    st.subheader("❌ Missing in ERP (found in vendor but not in ERP)")
    if not erp_missing.empty:
        st.dataframe(erp_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in ERP.")

    st.subheader("❌ Missing in Vendor (found in ERP but not in vendor)")
    if not ven_missing.empty:
        st.dataframe(ven_missing.style.applymap(lambda _: "background-color: #c62828; color: white"), use_container_width=True)
    else:
        st.success("✅ No missing invoices in Vendor.")

    st.subheader("🏦 Payment Transactions (Identified in both sides)")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**💼 ERP Payments**")
        if not erp_pay.empty:
            st.dataframe(erp_pay.style.applymap(lambda _: "background-color: #004d40; color: white"), use_container_width=True)
            st.markdown(f"**Total ERP Payments:** {erp_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No ERP payments found.")

    with col2:
        st.markdown("**🧾 Vendor Payments**")
        if not ven_pay.empty:
            st.dataframe(ven_pay.style.applymap(lambda _: "background-color: #1565c0; color: white"), use_container_width=True)
            st.markdown(f"**Total Vendor Payments:** {ven_pay['Amount'].sum():,.2f} EUR")
        else:
            st.info("No Vendor payments found.")

    st.markdown("### ✅ Matched Payments")
    if not matched_pay.empty:
        st.dataframe(matched_pay.style.applymap(lambda _: "background-color: #2e7d32; color: white"), use_container_width=True)
        total_erp = matched_pay["ERP Amount"].sum()
        total_vendor = matched_pay["Vendor Amount"].sum()
        diff_total = abs(total_erp - total_vendor)
        st.markdown(f"**Total Matched ERP Payments:** {total_erp:,.2f} EUR")
        st.markdown(f"**Total Matched Vendor Payments:** {total_vendor:,.2f} EUR")
        st.markdown(f"**Difference Between ERP and Vendor Payments:** {diff_total:,.2f} EUR")
    else:
        st.info("No matching payments found.")
