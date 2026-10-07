import streamlit as st
from google import genai
from docx import Document
from docx.shared import Mm, Pt
import io

# Mengatur konfigurasi halaman web
st.set_page_config(page_title="Generator RPP Deep Learning", layout="wide")

st.title("⛪ Generator Modul Ajar Pendidikan Agama Katolik SD")
st.subheader("Berbasis Pembelajaran Mendalam (Deep Learning) - Multi-Pertemuan")
st.write("Isi formulir di bawah ini untuk merancang modul ajar secara otomatis menggunakan AI.")

# --- SIDEBAR INPUT ---
st.sidebar.header("🔑 Pengaturan API Key")
api_key = st.sidebar.text_input("Masukkan Google GenAI API Key Anda:", type="password")

st.sidebar.header("✍️ Identitas Sekolah & Penulis")
sekolah = st.sidebar.text_input("Nama Sekolah:", "SMP Negeri 1 Metro")
kepala_sekolah = st.sidebar.text_input("Nama Kepala Sekolah:", "Fatimah, S.Pd. M.M.")
nip_kepala_sekolah = st.sidebar.text_input("NIP Kepala Sekolah:", "19670705 199202 2 002")
penulis = st.sidebar.text_input("Nama Penulis Modul:", "Antonius Tamtama, S.S")
nip_penulis = st.sidebar.text_input("NIP Penulis:", "198211212024211003")

# --- FORM UTAMA INPUT ---
col1, col2 = st.columns(2)

with col1:
    mapel = st.text_input("Mata Pelajaran:", "Pendidikan Agama Katolik dan Budi Pekerti")
    kelas_fase = st.text_input("Kelas / Fase:", "Kelas 4 / Fase B")
    elemen = st.selectbox("Elemen Pembelajaran:", ["Yesus Kristus", "Pribadi Murid", "Gereja", "Masyarakat"])
    topik_bahasan = st.text_input("Topik / Pokok Bahasan:", "Aku Bangga Sebagai Bangsa Indonesia")

with col2:
    tujuan_pembelajaran = st.text_area("Tujuan Pembelajaran:", "Murid mampu memahami nilai-nilai kebangsaan...")
    kktp = st.text_area("Kriteria Ketercapaian Pembelajaran (KKTP):", "1. Menyebutkan keanekaragaman suku dan budaya\n2. Menjelaskan ajaran Gereja tentang cinta tanah air")
    waktu = st.number_input("Jumlah Pertemuan (1 Pertemuan = 105 Menit):", min_value=1, max_value=10, value=2)

st.header("⚙️ Parameter Pembelajaran Mendalam")
col3, col4 = st.columns(2)

with col3:
    opsi_dimensi = [
        "Keimanan dan Ketakwaan terhadap Tuhan YME",
        "Kewargaan",
        "Penalaran Kritis",
        "Kreativitas",
        "Kolaborasi",
        "Kemandirian",
        "Kesehatan",
        "Komunikasi"
    ]
    
    dimensi_terpilih = st.multiselect(
        "Dimensi Profil Lulusan (Bisa Pilih Lebih dari 1):",
        options=opsi_dimensi,
        default=["Keimanan dan Ketakwaan terhadap Tuhan YME", "Penalaran Kritis"]
    )
    
    dimensi_profil_lulusan = ", ".join(dimensi_terpilih) if dimensi_terpilih else "Tidak ada dimensi yang dipilih"

    praktik_pedagogis = st.text_input("Praktik Pedagogis:", "Diskusi, Kateketis")
    lingkungan_pembelajaran = st.text_input("Lingkungan Pembelajaran:", "Ruang Kelas")

with col4:
    kemitraan_pembelajaran = st.text_input("Kemitraan Pembelajaran:", "Orang Tua, Lingkungan / Kring")
    pemanfaatan_digital = st.text_input("Pemanfaatan Digital:", "Canva for Education, Video, LCD Projector")
    persiapan_pembelajaran = st.text_area("Persiapan Guru:", "Guru menyiapkan presentasi materi pembelajaran, LKPD")

# --- FUNGSI MERUBAH TEKS MENJADI DOCX DI MEMORI ---
def buat_file_docx(teks_rpp):
    doc = Document()
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)
    
    bersih_teks = (
        teks_rpp.replace("<p>", "")
        .replace("</p>", "\n")
        .replace("<h1>", "\n\n")
        .replace("</h1>", "\n")
        .replace("<h2>", "\n\n")
        .replace("</h2>", "\n")
        .replace("<h3>", "\n\n")
        .replace("</h3>", "\n")
        .replace("<br>", "\n")
        .replace("<br/>", "\n")
    )
    
    for baris in bersih_teks.split('\n'):
        doc.add_paragraph(baris)
        
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- PROSES GENERATE ---
if st.button("🚀 Generate RPP / Modul Ajar", type="primary"):
    if not api_key:
        st.error("Silakan masukkan API Key Anda di sidebar terlebih dahulu!")
    else:
        with st.spinner(f"Sedang merancang RPP untuk {waktu} pertemuan... Mohon tunggu beberapa detik."):
            try:
                client = genai.Client(api_key=api_key.strip())
                
                prompt_text = f"""
                Saya adalah guru Pendidikan Agama Katolik jenjang Sekolah Dasar. Buatkan Rencana Pelaksanaan Pembelajaran (RPP) / Modul Ajar berbasis Pembelajaran Mendalam (Deep Learning) secara LENGKAP untuk {waktu} pertemuan. Rancangan pembelajaran harus sesuai dengan tahapan belajar perkembangan anak usia sekolah dasar, gunakan bahasa yang sederhana dan mudah dipahami oleh anak usia sekolah dasar.
                Gunakan Alkitab, Ajaran Sosial Gereja Katolik, Katekismus Gereja Katolik, dan Kitab Hukum Kanonik sebagai referensi utama.
                Setiap 1 pertemuan terdiri dari 105 menit. Bagi waktu di setiap pertemuan agar sesuai dengan kegiatan awal, kegiatan inti, dan penutup. Gunakan rumusan Dimensi Profil Lulusan bukan Dimensi Profil Pelajar Pancasila.
                
                PENTING: Anda harus menyusun output ini menggunakan format HTML murni yang rapi dan elegan agar langsung siap dicetak di kertas A4.
                Jangan gunakan markdown biasa (seperti ## atau **). Gunakan tag HTML seperti <h1>, <h2>, <p>, <ul>, <li>, dan <table>.
                
                Detail Kelas & Desain:
                - Sekolah: {sekolah} | Penulis: {penulis}
                - Mata Pelajaran: {mapel} | Kelas/Fase: {kelas_fase} | Elemen: {elemen}
                - Topik/Pokok Bahasan: {topik_bahasan} | Tujuan Pembelajaran: {tujuan_pembelajaran}
                - Total Waktu Rencana: {waktu} Pertemuan | KKTP: {kktp}
                - Dimensi Profil Lulusan: {dimensi_profil_lulusan} | Praktik Pedagogis: {praktik_pedagogis}
                - Lingkungan: {lingkungan_pembelajaran} | Kemitraan: {kemitraan_pembelajaran}
                - Digital: {pemanfaatan_digital} | Persiapan: {persiapan_pembelajaran}
                
                Struktur RPP harus mengikuti susunan berikut:
                1. Judul Modul yang menarik dan inspiratif di bagian atas.
                2. Identitas modul (mapel, Kelas/Fase, elemen, Sekolah, Penulis, topik/Pokok Bahasan, Tujuan Pembelajaran, Kriteria Ketercapaian Pembelajaran, Dimensi Profil Lulusan, Alokasi Waktu) di dalam tabel HTML rapi.
                3. Desain Pembelajaran (Tujuan pembelajaran, Kriteria ketercapaian pembelajaran, Praktik pedagogis, Lingkungan Pembelajaran, Kemitraan Pembelajaran, Pemanfaatan Digital, Persiapan Pembelajaran dan Tahapan Pembelajaran: Memahami/Understanding, Mengaplikasi/Applying, dan Merefleksi/Reflecting) di dalam tabel HTML.
                4. Langkah Pembelajaran detail untuk Pertemuan 1 sampai ke-{waktu}. Setiap pertemuan berisi:
                   - Kegiatan Awal (Apersepsi, Motivasi, Asesmen Diagnostik, Tujuan & Manfaat Pembelajaran)
                   - Kegiatan Inti (Meaningful, Eksplorasi Mendalam, Diskusi/Kolaborasi)
                   - Kegiatan Akhir (Rangkuman/Kesimpulan, Joyful, Refleksi, Apresiasi)
                5. Asesmen Formatif & Lembar Kerja Murid (LKM) untuk tiap pertemuan.
                6. Asesmen Sumatif (5 soal menjodohkan, 5 soal benar/salah, 10 soal pilihan ganda HOTS dengan 4 opsi jawaban dan Kunci Jawaban sesuai KKTP).
                7. Referensi / Daftar Pustaka sesuai kaidah penulisan daftar pustaka.
                
                Di akhir halaman dokumen, buatlah layout tanda tangan kiri-kanan menggunakan tabel HTML transparan (tanpa border):
                Sebelah kiri: Mengetahui, Kepala Sekolah {kepala_sekolah} (NIP: {nip_kepala_sekolah})
                Sebelah kanan: Metro, Penulis {penulis} (NIP: {nip_penulis})
                """
                
                # Menggunakan model gemini-2.5-flash
                response = client.models.generate_content(
                    model='gemini-2.5-flash',
                    contents=prompt_text,
                )
                
                st.session_state['rpp_html'] = response.text
                st.success("🎉 RPP Berhasil Dibuat!")
                
            except Exception as e:
                st.error(f"Terjadi kesalahan saat menghubungi Gemini API: {e}")

# --- TAMPILAN HASIL & TOMBOL AKSI ---
if 'rpp_html' in st.session_state:
    st.markdown("---")
    st.header("📄 Menu Aksi & Pratinjau Dokumen")
    
    btn_col1, btn_col2, btn_col3 = st.columns(3)
    
    with btn_col1:
        st.info("💡 **Cetak ke A4 / Simpan PDF:** Tekan **Ctrl + P** (Windows) atau **Cmd + P** (Mac).")
        
    with btn_col2:
        file_docx = buat_file_docx(st.session_state['rpp_html'])
        st.download_button(
            label="📥 Unduh File Word (.DOCX)",
            data=file_docx,
            file_name=f"RPP_{mapel.replace(' ', '_')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        
    with btn_col3:
        st.download_button(
            label="🌐 Unduh File Kode HTML",
            data=st.session_state['rpp_html'],
            file_name=f"RPP_{mapel.replace(' ', '_')}.html",
            mime="text/html"
        )
    
    html_content = f"""
    <div style="padding: 30px; border: 1px solid #ccc; background-color: white; color: black; font-family: Arial, sans-serif; line-height: 1.6; max-width: 800px; margin: 0 auto;">
        {st.session_state['rpp_html']}
    </div>
    """
    st.markdown(html_content, unsafe_allow_html=True)
