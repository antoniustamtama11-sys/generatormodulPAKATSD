import streamlit as st
from google import genai
from docx import Document
from docx.shared import Mm, Pt
import io

# Mengatur konfigurasi halaman web
st.set_page_config(page_title="Generator RPP Deep Learning", layout="wide")

st.title("⛪ Generator Modul Ajar Pendidikan Agama Katolik SD")
st.subheader("Berbasis Pembelajaran Mendalam (Deep Learning) - Multi-Pertemuan")
st.write("Isi formulir di bawah ini untuk merancang modulajar secara otomatis menggunakan AI.")

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
    kelas_fase = st.text_input("Kelas / Fase:", "Kelas ... / Fase ...")
    elemen = st.selectbox("Elemen Pembelajaran:", ["Yesus Kristus", "Pribadi Murid", "Gereja", "Masyarakat"])
    topik_bahasan = st.text_input("Topik / Pokok Bahasan:", "...")

with col2:
    tujuan_pembelajaran = st.text_area("Tujuan Pembelajaran:", "Murid Memahami ...")
    kktp = st.text_area("Kriteria Ketercapaian Pembelajaran (KKTP):", "1. ...")
    waktu = st.number_input("Jumlah Pertemuan (1 Pertemuan = 105 Menit):", min_value=1, max_value=10, value=2)

st.header("⚙️ Parameter Pembelajaran Mendalam")
col3, col4 = st.columns(2)

with col3:
    # --- PERUBAHAN DI SINI: Menggunakan st.multiselect untuk 8 Profil Dimensi Lulusan ---
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
    
    # Menggabungkan hasil list multiselect menjadi satu string yang dipisahkan koma
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
    # Atur ukuran halaman A4
    section = doc.sections[0]
    section.page_width = Mm(210)
    section.page_height = Mm(297)
    
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(11)
    
    # Hapus tag HTML dasar agar tidak mengotori Word dokumen
    bersih_teks = teks_rpp.replace("<p>", "").replace("</p>", "\n").replace("<h1>", "\n\n").replace("</h1>", "\n").replace("<h2>", "\n\n").replace("</h2>", "\n").replace("<h3>", "\n\n").replace("</h3>", "\n")
    
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
        with st.spinner(f"Sedang merancang RPP untuk {waktu} pertemuan... Mohon tunggu sekitar 15-20 detik."):
            try:
                client = genai.Client(api_key=api_key)
                
                prompt_text = f"""
                saya adalah guru pendidikan agama katolik jenjang Sekolah Dasar, Buatkan Rencana Pelaksanaan Pembelajaran (RPP) / Modul Ajar berbasis Pembelajaran Mendalam (Deep Learning) secara LENGKAP untuk {waktu} pertemuan. rancangan pembelajaran harus sesuai dengan tahapan belajar perkembangan anak usia sekolah dasar, buat bahasa yang sederhana dan mudah dipahami oleh anak  usia sekolah dasar.
                Gunakan Alkitab, ajaran sosial gereja Katolik, Katekismus gereja katolik dan Kitab hukum kanonik sebagai referensi utama. gunakan juga file buku pendidikan agama kalolik di link berikut sebagai referesi sesuai dengan jenjang kelas: https://drive.google.com/drive/folders/1OVx6e_Dek1oDs_jBgHPbMEEFAVtpk4EW?usp=sharing. 
                Setiap 1 pertemuan terdiri dari 105 menit. Bagi waktu di setiap pertemuan agar sesuai dengan kegiatan awal, kegiatan inti, dan penutup. gunakan rumusan Dimensi Profil Lulusan bukan Dimensi profil Pelajar pancasila
                
                PENTING: Anda harus menyusun output ini menggunakan format HTML murni yang rapi dan elegan agar langsung siap dicetak di kertas A4.
                Jangan gunakan markdown biasa (seperti ## atau **). Gunakan tag HTML seperti <h1>, <h2>, <p>, <ul>, <li>, dan <table>.
                
                Detail Kelas & Desain:
                - Sekolah: {sekolah} | Penulis: {penulis}
                - Mata Pelajaran: {mapel} | Kelas/Fase: {kelas_fase} | Elemen: {elemen}
                - Topik/Pokok Bahasan: {topik_bahasan} | Tujuan Pembelajaran: {tujuan_pembelajaran}
                - Total Waktu Rencana: {waktu} Pertemuan | KKTP: {kktp}
                - Dimensi Profil Lulusan: {dimensi_profil_lulusan} | Praktik Pedagogis: sesuaikan dengan rancangan pembelajaran {praktik_pedagogis}
                - Lingkungan: {lingkungan_pembelajaran} | Kemitraan: {kemitraan_pembelajaran}
                - Digital: {pemanfaatan_digital} | Persiapan: {persiapan_pembelajaran}
                
                Struktur RPP harus mengikuti susunan berikut:
                - Judul Modul yang menarik dan inspiratif di bagian atas.
                - 1. Identitas modul memuat: mapel, Kelas/Fase, elemen, Sekolah, Penulis, topik/Pokok Bahasan, Tujuan Pembelajaran, Kriteria Ketercapaian Pembelajaran, Dimensi profil lulusan, Alokasi Waktu (Buat rapi di dalam tabel HTML)
                - 2. Desain Pembelajaran memuat: Tujuan pembelajaran, Kriteria ketercapaian pembelajaran, Praktik pedagogis, Lingkungan Pembelajaran, Kemitraan Pembelajaran, Pemanfaatan Digital, Persiapan Pembelajaran dan  Tahapan Pembelajaran: Tentukan dibagian pembelajaran mana yang merupakan tahapan pembaelajaran mendalam tentang Memahami (Understanding) Mengaplikasi (Applying)dan Merefleksi (Reflecting)  (Buat rapi di dalam tabel HTML).
                - 3. Langkah Pembelajaran: Wajib dijabarkan detail satu per satu dari Pertemuan 1 sampai Pertemuan ke-{waktu}. Setiap pertemuan memuat Kegiatan Awal (sesuaikan waktu: memuat Apersepsi, Motivasi, Asesmen Diagnostik, Tujuan Pembelajaran, Manfaat Pembelajaran), Kegiatan Inti (sesuikan waktu: gunakan tahapan atau sintak pembelajaran pada praktik pedagogis yang dipilih, perhatikan aspek Meaningful, Eksplorasi Mendalam, dan Diskusi/Kolaborasi), Kegiatan Akhir (sesuaikan: berisi rangkuman atau kesimpulan pembelajaran dan refleksi, Fokus pada aspek Joyful, Refleksi, dan Apresiasi).
                - 4. Asesmen Formatif & Lembar Kerja Murid (LKM) untuk tiap pertemuan.
                - 5. Asesmen Sumatif (5 soal menjodohkan, 5 soal benar salah, 10 soal pilihan ganda HOTS dengan 4 obsi jawaban dan Kunci Jawaban sesuai dengan KKTP).
                - 6. Referensi / Daftar Pustaka (penulisan harus sesuai dengan kaedah penulisan daftar pustaka.
                
                Di akhir halaman dokumen, buatlah layout tanda tangan kiri-kanan menggunakan tabel HTML transparan:
                Sebelah kiri: Mengetahui, Kepala Sekolah {kepala_sekolah} (NIP: {nip_kepala_sekolah})
                Sebelah kanan: Metro, Penulis {penulis} (NIP: {nip_penulis})
                """
                
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
    
    # Membuat tombol-tombol aksi berdampingan yang berfungsi 100%
    btn_col1, btn_col2, btn_col3 = st.columns(3)
    
    with btn_col1:
        st.info("💡 **Untuk Cetak ke A4 / Simpan ke PDF:** Tekan kombinasi tombol **Ctrl + P** (Windows) atau **Cmd + P** (Mac) di keyboard Anda saat berada di halaman ini.")
        
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
            label="🌐 Unduh File Kode HTML Resmi",
            data=st.session_state['rpp_html'],
            file_name=f"RPP_{mapel.replace(' ', '_')}.html",
            mime="text/html"
        )
    
    # Tampilkan halaman cetak yang bersih
    html_content = f"""
    <div style="padding: 30px; border: 1px solid #ccc; background-color: white; color: black; font-family: Arial, sans-serif; line-height: 1.6; max-width: 800px; margin: 0 auto;">
        {st.session_state['rpp_html']}
    </div>
    """
    st.markdown(html_content, unsafe_allow_html=True)
