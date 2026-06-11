export interface Kelas {
  id: string;
  namaKelas: string;
  waliKelas: string;
  nuptkNgty: string;
  thAjaran?: string;
}

export interface Siswa {
  id: string;
  namaSiswa: string;
  nisn: string;
  alamat: string;
  namaAyah: string;
  pekerjaanAyah: string;
  namaIbu: string;
  pekerjaanIbu: string;
  noHp: string;
  tglLahir: string;
  tempatLahir: string;
  jenisKelamin: string;
  anakKe: number;
  tb: string; // Tinggi Badan
  bb: string; // Berat Badan
  agama: string;
  idKelas: string;
  thAjaran?: string;
}

export interface KategoriIntrakurikuler {
  id: string;
  namaKategori: string;
}

export interface TujuanPembelajaran {
  id: string;
  idKategori: string;
  deskripsi: string;
  idKelas: string; // Filter per class
  aktivitasMetode?: string; // Aktivitas / Metode
  thAjaran?: string;
}

export interface LabelP5 {
  id: string;
  namaLabel: string;
  order: number;
}

export interface SubdimensiKokurikuler {
  id: string;
  namaSubdimensi: string;
  idKelas: string; // Filter per class
  // Map of labelId -> description
  capaian?: Record<string, string>;
  // Legacy fields (optional)
  descBerkembang?: string;
  descCakap?: string;
  descMahir?: string;
  thAjaran?: string;
}

export interface NilaiIntrakurikuler {
  idSiswa: string;
  idTp: string;
  nilai: string; // Dinamis berdasarkan label
  deskripsi?: string; // Qualitative assessment per TP
  thAjaran?: string;
  semester?: string;
}

export interface NilaiKokurikuler {
  idSiswa: string;
  idSubdimensi: string;
  nilai: string; // Dinamis berdasarkan label
  deskripsi?: string; // Qualitative assessment per Subdimensi
  thAjaran?: string;
  semester?: string;
}

export interface CatatanAnak {
  idSiswa: string;
  catatan: string;
  thAjaran?: string;
  semester?: string;
}

export interface Kehadiran {
  idSiswa: string;
  sakit: number;
  ijin: number;
  tanpaKet: number;
  thAjaran?: string;
  semester?: string;
}

export interface DataSekolah {
  namaSekolah: string;
  alamat: string;
  kepalaSekolah: string;
  logo: string; // Base64 or URL
  logoSidebar?: string; // Optional
  logoLogin?: string; // Optional
  semester: string; // e.g., "1 (Ganjil)" atau "2 (Genap)"
  thAjaran: string; // e.g., "2025/2026"
  tglRaport: string; // e.g., "2026-06-20"
  npsn?: string;
  aiProvider?: "gemini" | "groq";
  aiModel?: string;
}

export interface AppState {
  kelas: Kelas[];
  siswa: Siswa[];
  kategoriIntrakurikuler: KategoriIntrakurikuler[];
  tujuanPembelajaran: TujuanPembelajaran[];
  subdimensiKokurikuler: SubdimensiKokurikuler[];
  labelP5: LabelP5[];
  nilaiIntrakurikuler: NilaiIntrakurikuler[];
  nilaiKokurikuler: NilaiKokurikuler[];
  catatanAnak: CatatanAnak[];
  kehadiran: Kehadiran[];
  dataSekolah: DataSekolah;
}

export interface UserProfile {
  uid: string;
  email: string;
  role: "admin" | "guru";
  nama: string;
  username?: string;
  password?: string;
  isCreatedInAuth?: boolean;
}

