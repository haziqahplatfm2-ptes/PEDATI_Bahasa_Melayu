import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from io import BytesIO

# --- 1. CONFIGURATION ---
genai.configure(api_key=st.secrets["GOOGLE_API_KEY"])

@st.cache_resource
def find_working_model():
    """Mencari model yang tersedia untuk mengelakkan ralat versi."""
    try:
        for m in genai.list_models():
            if 'generateContent' in m.supported_generation_methods:
                return m.name
    except:
        return "models/gemini-1.5-flash"
    return "models/gemini-1.5-flash"

selected_model_name = find_working_model()
model = genai.GenerativeModel(selected_model_name)

# --- 2. AI LOGIC (BAHASA MELAYU) ---
def generate_pedati_plan_bm(topic, syllabus, extra_context):
    prompt = f"""
    Topik: {topic}. Kod Sukatan Pelajaran: {syllabus}. Konteks: {extra_context}.
    Hasilkan rancangan pengajaran profesional dalam BAHASA MELAYU sepenuhnya. 
    Pastikan tiada istilah Bahasa Inggeris digunakan kecuali jika perlu untuk istilah teknikal.

    Jika ada disebutkan di dalam Ruang Konteks: penyata seperti berikut **tukar ke tulisan JAWI Melayu**
    Pastikan tiada istilah Bahasa Inggeris dan Tulisan Bahasa Melayu digunakan kecuali jika perlu untuk istilah teknikal.
    
    Gunakan nama peringkat PEDATI berikut:
    P [Pengetahuan Sedia Ada], E [Empati/Engage], D [Daya Usaha/Develop], A [Aplikasi], T [Taksiran], I [Imbas Kembali/Improve].

    Gunakan penanda (SECTION:) untuk penstrukturan dokumen:
    SECTION: OBJEKTIF PEMBELAJARAN
    [4 poin utama]
    SECTION: HASIL PEMBELAJARAN
    [4 poin utama]
    SECTION: KRITERIA KEJAYAAN
    [4 poin utama]
    SECTION: PRASYARAT
    [1 poin utama]
    SECTION: KATA KUNCI
    [6 item kata kunci]
    SECTION: KBAT (KEMAHIRAN BERFIKIR ARAS TINGGI)
    [4 domain mengikut Taksonomi Bloom]
    SECTION: KEWARGANEGARAAN DIGITAL
    [4 poin mengenai etika digital, penggunaan Chromebook, Canva, YouTube, atau peranti digital secara bertanggungjawab]

    SECTION: PERINGKAT PEDATI
    STAGE: P [Pengetahuan Sedia Ada] | SB: [Aktiviti Guru] | CB: [Aktiviti Murid]
    STAGE: E [Engage] | SB: [Aktiviti Guru] | CB: [Aktiviti Murid]
    STAGE: D [Develop] | SB: [Aktiviti Guru] | CB: [Aktiviti Murid]
    STAGE: A [Apply] | SB: [Aktiviti Guru] | CB: [Aktiviti Murid]
    STAGE: T [Test] | SB: [Aktiviti Guru] | CB: [Aktiviti Murid]
    STAGE: I [Improve] | SB: [Aktiviti Guru] | CB: [Aktiviti Murid]
    """
    try:
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Ralat Sistem: {str(e)}"

# --- 3. WORD EXPORT LOGIC (BAHASA MELAYU) ---
def create_word_export_bm(topic, syllabus, text):
    doc = Document()
    doc.add_heading(f'Rancangan Pengajaran: {topic} ({syllabus})', 0)

    # 1. Jadual Pentadbiran (6 ruangan)
    admin_table = doc.add_table(rows=3, cols=4)
    admin_table.style = 'Table Grid'
    labels = [["Minggu No:", "Tarikh:"], ["Bil. Pelajar:", "Hari:"], ["Tempat/Makmal:", "Durasi (minit):"]]
    for r in range(3):
        admin_table.cell(r, 0).text = labels[r][0]
        admin_table.cell(r, 2).text = labels[r][1]
    doc.add_paragraph()

    # 2. Jadual Sumber
    doc.add_heading("Sumber & Bahan Bantu Mengajar", level=1)
    res_table = doc.add_table(rows=1, cols=1)
    res_table.style = 'Table Grid'
    res_table.cell(0, 0).text = "Papan Pintar (Smart Board), Chromebook, Meja Menulis, Projektor, Perkongsian Skrin"

    # 3. Analisis Kandungan & Kotak Jadual
    sections = text.split('SECTION:')
    for section in sections:
        if not section.strip(): continue
        lines = section.strip().split('\n')
        title = lines[0].strip()
        content_lines = lines[1:]
        doc.add_heading(title.title(), level=1)

        if "|" in section and "PEDATI" in title.upper():
            table = doc.add_table(rows=1, cols=3)
            table.style = 'Table Grid'
            hdr = table.rows[0].cells
            hdr[0].text, hdr[1].text, hdr[2].text = 'Peringkat (PEDATI)', 'Fasilitator (Guru)', 'Pelajar (Murid)'

            for line in content_lines:
                if "|" in line:
                    p = line.split("|")
                    row = table.add_row().cells
                    row[0].text = p[0].split(":")[-1].strip()
                    row[1].text = p[1].split(":")[-1].strip()
                    row[2].text = p[2].split(":")[-1].strip()
        else:
            table = doc.add_table(rows=1, cols=1)
            table.style = 'Table Grid'
            table.cell(0, 0).text = "\n".join([l.strip() for l in content_lines if l.strip()])

    # 4. Halaman Pengesahan HOD
    doc.add_page_break()
    doc.add_heading("Pengesahan & Ulasan HOD", level=1)
    hod_table = doc.add_table(rows=3, cols=2)
    hod_table.style = 'Table Grid'
    hod_table.cell(0, 0).text = "Ulasan";
    hod_table.cell(0, 1).text = "Tandatangan / Cop Rasmi"
    hod_table.rows[1].height = Pt(60)
    hod_table.cell(2, 0).text = "Tarikh:";
    hod_table.cell(2, 1).text = "Nama:"

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 4. GUI SECTION ---
st.set_page_config(page_title="Penjana PEDATI Master", layout="wide")

with st.sidebar:
    st.title("📖 Panduan Pengguna")
    st.info("Cara menggunakan portal ini:")
    
    st.markdown("""
    ### 1. Isi Maklumat
    Masukkan **Topik Pelajaran** dan **Kod Sukatan**. Gunakan kotak **Konteks** untuk keperluan khusus seperti "Kerja Berkumpulan", "Kajian Kes", "Pautan YouTube" atau "LMS Dalam Talian".
    
    ### 2. Jana
    Klik **🚀 JANA** dan tunggu AI merangka pelan anda.
    
    ### 3. Semak & Simpan
    Semak **Pratonton AI**. Jika anda berpuas hati, klik **📥 Muat Turun Word** untuk mendapatkan dokumen profesional anda.
    
    ---
    ### 💡 Tip Profesional
    Jika jawapan AI terhenti, cuba tambah kata kunci yang lebih spesifik dalam kotak Konteks untuk membantu AI!
    """)
    st.markdown("---")
    st.caption("Versi Aplikasi 2.0 (BM) | Inovasi PTES")

# --- DASHBOARD UTAMA ---
st.title("🎓 Penjana Rancangan Pengajaran PEDATI")
st.info(f"Sistem dihubungkan melalui: {selected_model_name}")

c1, c2 = st.columns(2)
with c1: u_topic = st.text_input("Topik Pelajaran:")
with c2: u_syllabus = st.text_input("Kod Sukatan Pelajaran:")
u_extra = st.text_area("Konteks/Kata Kunci Spesifik (Pilihan):")

if st.button("🚀 JANA RANCANGAN PENGAJARAN PEDATI"):
    if u_topic and u_syllabus:
        with st.spinner("AI sedang membina rancangan PEDATI anda..."):
            result = generate_pedati_plan_bm(u_topic, u_syllabus, u_extra)
            st.session_state['pedati_out_bm'] = result
    else:
        st.warning("Sila masukkan Topik dan Kod Sukatan.")

if 'pedati_out_bm' in st.session_state:
    st.divider()
    st.subheader("📝 Pratonton AI")
    st.text_area("Kandungan Rangka", st.session_state['pedati_out_bm'], height=300)
    doc_file = create_word_export_bm(u_topic, u_syllabus, st.session_state['pedati_out_bm'])
    st.download_button("📥 Muat Turun Word (.docx)", doc_file, f"PEDATI_BM_{u_topic}.docx")

# --- FOOTER SECTION ---
st.markdown("---") 
st.markdown(
    """
    <div style='text-align: center; color: grey; font-size: 0.8em;'>
        <p><b>Smart PEDATI Lesson Plan AI-Generator v2.0 (Bahasa Melayu)</b></p>
        <p>Dibangunkan & Dikonsepkan oleh: <b>Hajah Nurul Haziqah @ Hjh Hartini Hj Nordin</b></p>
        <p>© 2026 PTES Academic Innovation Computer Science</p>
    </div>
    """,
    unsafe_allow_html=True
)
