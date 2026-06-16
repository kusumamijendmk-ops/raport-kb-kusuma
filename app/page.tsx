"use client";

import React, { useState, useEffect, useMemo } from "react";
import { 
  Home, Users, GraduationCap, BookOpen, PenTool, Printer, Settings, Plus, Trash2, 
  Sparkles, Download, Upload, Check, Search, Award, CheckCircle, 
  MapPin, UserCheck, ChevronRight, ChevronDown, FileText, Calendar, Edit3, ShieldAlert,
  Folder, Save, Zap, LayoutDashboard, Menu, X
} from "lucide-react";
import { PieChart, Pie, Cell, ResponsiveContainer, Legend, Tooltip } from "recharts";

import { AppState, Kelas, Siswa, TujuanPembelajaran, SubdimensiKokurikuler, NilaiIntrakurikuler, NilaiKokurikuler, CatatanAnak, Kehadiran, DataSekolah, UserProfile, KategoriIntrakurikuler, LabelP5 } from "../lib/types";
import { seedData, initialKategoriIntrakurikuler } from "../lib/seed";

// Firebase Integration
import { auth, db, googleProvider, OperationType, handleFirestoreError } from "../lib/firebase";
import { 
  onAuthStateChanged, signInWithPopup, signOut, User as FirebaseUser,
  signInWithEmailAndPassword, createUserWithEmailAndPassword
} from "firebase/auth";
import { 
  collection, doc, getDoc, setDoc, deleteDoc, updateDoc, onSnapshot, writeBatch, query, where, getDocs 
} from "firebase/firestore";

import * as XLSX from "xlsx";
import dynamic from "next/dynamic";

const ReactQuill = dynamic(() => import("react-quill-new"), { ssr: false });

const excelDateToIsoString = (val: any): string => {
  if (!val) return "";
  const str = String(val).trim();
  if (!str) return "";

  // If it's an Excel serial date number
  if (/^\d+(\.\d+)?$/.test(str)) {
    const serial = parseFloat(str);
    if (serial > 10000 && serial < 60000) {
      try {
        const msSinceEpoch = (serial - 25569) * 86400 * 1000;
        const date = new Date(msSinceEpoch);
        const yyyy = date.getUTCFullYear();
        const mm = String(date.getUTCMonth() + 1).padStart(2, "0");
        const dd = String(date.getUTCDate()).padStart(2, "0");
        return `${yyyy}-${mm}-${dd}`;
      } catch (e) {
        // Fallback
      }
    }
  }

  // If it is already standard YYYY-MM-DD
  if (/^\d{4}-\d{2}-\d{2}$/.test(str)) {
    return str;
  }

  // If DD-MM-YYYY or DD/MM/YYYY
  const slashRegex = /^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/;
  if (slashRegex.test(str)) {
    const parts = str.match(slashRegex);
    if (parts) {
      const d = String(parts[1]).padStart(2, "0");
      const m = String(parts[2]).padStart(2, "0");
      const y = parts[3];
      return `${y}-${m}-${d}`;
    }
  }

  // Try standard JS parser
  const ts = Date.parse(str);
  if (!isNaN(ts)) {
    try {
      const date = new Date(ts);
      const yyyy = date.getFullYear();
      const mm = String(date.getMonth() + 1).padStart(2, "0");
      const dd = String(date.getDate()).padStart(2, "0");
      return `${yyyy}-${mm}-${dd}`;
    } catch (e) {}
  }

  return str;
};

const cleanName = (name: string): string => {
  if (!name) return "";
  return name
    .replace(/\bS\.PD\.?$/gi, "S.Pd.")
    .replace(/\bS\.PD\.?\s/gi, "S.Pd. ")
    .replace(/\bS\.P\.?\s/gi, "S.P. ")
    .replace(/\bS\.TH\.I\.?$/gi, "S.Th.I.");
};

const formatIndonesianDate = (dateStr: string, fallbackDefault?: string): string => {
  if (!dateStr) return fallbackDefault !== undefined ? fallbackDefault : "19 Desember 2025";
  
  // Extract just the date part if it is a full ISO timestamp
  const trimmed = dateStr.trim().split("T")[0].replace(/\//g, "-");
  if (!trimmed) return fallbackDefault !== undefined ? fallbackDefault : "19 Desember 2025";

  // Months name array
  const months = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni", 
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];

  // Check if it's an Excel serial date number
  if (/^\d+(\.\d+)?$/.test(trimmed)) {
    const serial = parseFloat(trimmed);
    if (serial > 10000 && serial < 60000) {
      try {
        const msSinceEpoch = (serial - 25569) * 86400 * 1000;
        const date = new Date(msSinceEpoch);
        const day = date.getUTCDate();
        const monthIdx = date.getUTCMonth();
        const year = date.getUTCFullYear();
        
        if (monthIdx >= 0 && monthIdx < 12) {
          return `${day} ${months[monthIdx]} ${year}`;
        }
      } catch (e) {}
    }
  }

  // Check YYYY-MM-DD (Standard)
  const ymdRegex = /^(\d{4})-(\d{2})-(\d{2})$/;
  if (ymdRegex.test(trimmed)) {
    const parts = trimmed.split("-");
    const year = parseInt(parts[0], 10);
    const monthIdx = parseInt(parts[1], 10) - 1;
    const day = parseInt(parts[2], 10);
    
    if (monthIdx >= 0 && monthIdx < 12) {
      return `${day} ${months[monthIdx]} ${year}`;
    }
  }

  // Check DD-MM-YYYY
  const dmyRegex = /^(\d{1,2})-(\d{1,2})-(\d{4})$/;
  if (dmyRegex.test(trimmed)) {
    const parts = trimmed.split("-");
    const day = parseInt(parts[0], 10);
    const monthIdx = parseInt(parts[1], 10) - 1;
    const year = parseInt(parts[2], 10);
    
    if (monthIdx >= 0 && monthIdx < 12) {
      return `${day} ${months[monthIdx]} ${year}`;
    }
  }

  // Standard date parser fallback
  const parsedTimestamp = Date.parse(trimmed);
  if (!isNaN(parsedTimestamp)) {
    const date = new Date(parsedTimestamp);
    const day = date.getDate();
    const monthIdx = date.getMonth();
    const year = date.getFullYear();
    
    if (monthIdx >= 0 && monthIdx < 12) {
      return `${day} ${months[monthIdx]} ${year}`;
    }
  }

  return trimmed;
};

const cleanHtmlText = (htmlString: string) => {
  if (!htmlString) return "";
  return htmlString
    .replace(/&nbsp;/g, " ")
    .replace(/\u00a0/g, " ");
};

export default function RaportPAUD() {
  const [mounted, setMounted] = useState(false);
  const [currentUser, setCurrentUser] = useState<FirebaseUser | null>(null);
  const [currentUserProfile, setCurrentUserProfile] = useState<UserProfile | null>(null);
  const [dbLoading, setDbLoading] = useState<boolean>(true);
  const [allUsersList, setAllUsersList] = useState<UserProfile[]>([]);

  // Username Login States
  const [loginUsername, setLoginUsername] = useState("");
  const [loginPassword, setLoginPassword] = useState("");
  const [loginError, setLoginError] = useState("");
  const [loginLoading, setLoginLoading] = useState(false);
  const [loginTab, setLoginTab] = useState<"username" | "google">("username");

  // Admin Create Teacher Account States
  const [newTeacherNama, setNewTeacherNama] = useState("");
  const [newTeacherUsername, setNewTeacherUsername] = useState("");
  const [newTeacherPassword, setNewTeacherPassword] = useState("");
  const [newTeacherRole, setNewTeacherRole] = useState<"admin" | "guru">("guru");

  const [rawState, setState] = useState<AppState>({
    kelas: [],
    siswa: [],
    kategoriIntrakurikuler: initialKategoriIntrakurikuler,
    tujuanPembelajaran: [],
    subdimensiKokurikuler: [],
    labelP5: [],
    nilaiIntrakurikuler: [],
    nilaiKokurikuler: [],
    catatanAnak: [],
    kehadiran: [],
    dataSekolah: {
      namaSekolah: "",
      alamat: "",
      kepalaSekolah: "",
      logo: "",
      semester: "2 (Genap)",
      thAjaran: "2025/2026",
      tglRaport: "",
      npsn: "",
      aiProvider: "groq",
      aiModel: "llama-3.3-70b-versatile"
    }
  });

  const state = useMemo(() => {
    const th = rawState.dataSekolah?.thAjaran || "2025/2026";
    const sem = rawState.dataSekolah?.semester || "1 (Ganjil)";

    const matchThAjaran = (itemTh?: string) => {
      return itemTh === th || (!itemTh && th === "2025/2026");
    };

    return {
      ...rawState,
      kelas: rawState.kelas.filter(k => matchThAjaran(k.thAjaran)),
      siswa: rawState.siswa.filter(s => matchThAjaran(s.thAjaran)),
      tujuanPembelajaran: rawState.tujuanPembelajaran.filter(tp => matchThAjaran(tp.thAjaran)),
      subdimensiKokurikuler: rawState.subdimensiKokurikuler.filter(sub => matchThAjaran(sub.thAjaran)),
      nilaiIntrakurikuler: rawState.nilaiIntrakurikuler.filter(n => n.thAjaran === th && n.semester === sem),
      nilaiKokurikuler: rawState.nilaiKokurikuler.filter(n => n.thAjaran === th && n.semester === sem),
      catatanAnak: rawState.catatanAnak.filter(c => c.thAjaran === th && c.semester === sem),
      kehadiran: rawState.kehadiran.filter(k => k.thAjaran === th && k.semester === sem),
    };
  }, [rawState]);

  const isAdmin = currentUserProfile?.role === "admin" || 
                  currentUser?.email === "kusumamijen.dmk@gmail.com" || 
                  currentUser?.email === "admin@paud.local";

  const [activeTab, setActiveTab] = useState<string>("dashboard"); 
  const [isSidebarOpenOnMobile, setIsSidebarOpenOnMobile] = useState<boolean>(false);
  const [nilaiSubTab, setNilaiSubTab] = useState<string>("intra"); // intra, kokuri, catatan, kehadiran

  // Collapsible Sub-menu States
  const [isMasterDataExpanded, setIsMasterDataExpanded] = useState<boolean>(false);
  const [isPenilaianExpanded, setIsPenilaianExpanded] = useState<boolean>(false);
  const [isLaporanExpanded, setIsLaporanExpanded] = useState<boolean>(false);

  // Reference checklist / rubric guidance toggle for co-curricular assessment
  const [showKokuriRubrikRef, setShowKokuriRubrikRef] = useState<boolean>(true);

  // Master Data Search & Form temporary states
  const [searchSiswa, setSearchSiswa] = useState("");
  const [selectedKelasFilter, setSelectedKelasFilter] = useState<string>("");
  const [selectedKelasFilterIntra, setSelectedKelasFilterIntra] = useState<string>("");
  const [selectedSiswaIdIntra, setSelectedSiswaIdIntra] = useState<string>("");
  const [activeIntraCategoryTab, setActiveIntraCategoryTab] = useState<string>("");
  const [selectedKelasFilterKokuri, setSelectedKelasFilterKokuri] = useState<string>("");
  const [selectedSiswaIdKokuri, setSelectedSiswaIdKokuri] = useState<string>("");
  const [selectedKelasFilterKehadiran, setSelectedKelasFilterKehadiran] = useState<string>("");
  const [selectedSiswaIdKehadiran, setSelectedSiswaIdKehadiran] = useState<string>("");
  const [selectedSiswaIdCatatan, setSelectedSiswaIdCatatan] = useState<string>("");
  const [tahunSumberAmbil, setTahunSumberAmbil] = useState<string>("");

  // Modals / Input toggles
  const [showAddSiswa, setShowAddSiswa] = useState(false);
  const [showAddKelas, setShowAddKelas] = useState(false);
  const [showAddTp, setShowAddTp] = useState(false);
  const [showAddSub, setShowAddSub] = useState(false);

  // Edit trackers
  const [editingSiswa, setEditingSiswa] = useState<Siswa | null>(null);
  const [editingKelas, setEditingKelas] = useState<Kelas | null>(null);

  // AI draft helpers
  const [generatingAi, setGeneratingAi] = useState<string | null>(null); // student ID
  const [generatingAiItem, setGeneratingAiItem] = useState<string | null>(null); // item ID (TP or SubId)
  const [aiPromptCustom, setAiPromptCustom] = useState("");

  // Add/Edit Siswa state
  const [siswaForm, setSiswaForm] = useState<Partial<Siswa>>({
    namaSiswa: "", nisn: "", alamat: "", namaAyah: "", pekerjaanAyah: "", namaIbu: "", pekerjaanIbu: "",
    noHp: "", tglLahir: "", tempatLahir: "", jenisKelamin: "Laki-laki", anakKe: 1, tb: "100", bb: "15", agama: "Islam", idKelas: ""
  });

  // Add/Edit Kelas state
  const [kelasForm, setKelasForm] = useState({
    namaKelas: "", waliKelas: "", nuptkNgty: ""
  });

  // Add TP state
  const [tpForm, setTpForm] = useState({
    idKategori: "KAT-01", deskripsi: "", idKelas: "", aktivitasMetode: ""
  });

  // Add/Edit Kategori state
  const [showAddKategori, setShowAddKategori] = useState(false);
  const [editingKategori, setEditingKategori] = useState<KategoriIntrakurikuler | null>(null);
  const [kategoriForm, setKategoriForm] = useState<Partial<KategoriIntrakurikuler>>({
    namaKategori: ""
  });
  const [kategoriToDelete, setKategoriToDelete] = useState<KategoriIntrakurikuler | null>(null);
  const [editingTp, setEditingTp] = useState<TujuanPembelajaran | null>(null);
  const [tpToDelete, setTpToDelete] = useState<TujuanPembelajaran | null>(null);

  // Add Subdimensi state
  const [subForm, setSubForm] = useState({
    namaSubdimensi: "", idKelas: "", descBerkembang: "", descCakap: "", descMahir: "",
    capaian: {} as Record<string, string>
  });
  const [editingSub, setEditingSub] = useState<SubdimensiKokurikuler | null>(null);
  const [subToDelete, setSubToDelete] = useState<SubdimensiKokurikuler | null>(null);

  // Label P5 States
  const [showAddLabel, setShowAddLabel] = useState(false);
  const [editingLabel, setEditingLabel] = useState<LabelP5 | null>(null);
  const [labelForm, setLabelForm] = useState<Partial<LabelP5>>({ namaLabel: "", order: 0 });
  const [labelToDelete, setLabelToDelete] = useState<LabelP5 | null>(null);

  // Print selector
  const [printSiswaId, setPrintSiswaId] = useState<string>("");
  const [printKelasId, setPrintKelasId] = useState<string>("");

  // Custom Modals & Toasts for Admin User Management to bypass iframe prompt/alert/confirm blocking
  const [userToDelete, setUserToDelete] = useState<UserProfile | null>(null);
  const [kelasToDelete, setKelasToDelete] = useState<Kelas | null>(null);
  const [siswaToDelete, setSiswaToDelete] = useState<Siswa | null>(null);
  const [userToEditPassword, setUserToEditPassword] = useState<UserProfile | null>(null);
  const [newPasswordValue, setNewPasswordValue] = useState("");
  const [userToChangeRole, setUserToChangeRole] = useState<UserProfile | null>(null);
  const [targetRoleValue, setTargetRoleValue] = useState<"admin" | "guru">("guru");
  const [toastMessage, setToastMessage] = useState<{ text: string; type: "success" | "error" | "info" } | null>(null);

  const showToast = (text: string, type: "success" | "error" | "info" = "success") => {
    setToastMessage({ text, type });
    setTimeout(() => {
      setToastMessage((prev) => (prev?.text === text ? null : prev));
    }, 4000);
  };

  const fileInputRef = React.useRef<HTMLInputElement>(null);

  // Initialize auth listener
  useEffect(() => {
    setMounted(true);
    
    const unsubscribeAuth = onAuthStateChanged(auth, async (firebaseUser) => {
      if (firebaseUser) {
        setCurrentUser(firebaseUser);
        setDbLoading(true);

        try {
          const userDocRef = doc(db, "users", firebaseUser.uid);
          const snap = await getDoc(userDocRef);
          
          let role: "admin" | "guru" = "guru";
          let nama = firebaseUser.displayName || "Operator Guru";
          
          if (firebaseUser.email === "kusumamijen.dmk@gmail.com" || firebaseUser.email === "admin@paud.local") {
            role = "admin";
          }
          
          let profileData: UserProfile;
          if (snap.exists()) {
            profileData = snap.data() as UserProfile;
            setCurrentUserProfile(profileData);
          } else {
            // Check if there is a pre-created temporary account under this email in /users
            let foundTempProfile: UserProfile | null = null;
            let tempDocId: string | null = null;
            if (firebaseUser.email) {
              const q = query(collection(db, "users"), where("email", "==", firebaseUser.email));
              const qSnap = await getDocs(q);
              if (!qSnap.empty) {
                const docFound = qSnap.docs.find(d => d.id !== firebaseUser.uid);
                if (docFound) {
                  foundTempProfile = docFound.data() as UserProfile;
                  tempDocId = docFound.id;
                }
              }
            }

            if (foundTempProfile) {
              profileData = {
                ...foundTempProfile,
                uid: firebaseUser.uid,
                isCreatedInAuth: true
              };
              await setDoc(userDocRef, profileData);
              if (tempDocId) {
                try {
                  await deleteDoc(doc(db, "users", tempDocId));
                } catch (delErr) {
                  console.warn("Could not delete temporary user doc:", delErr);
                }
              }
              setCurrentUserProfile(profileData);
            } else {
              profileData = {
                uid: firebaseUser.uid,
                email: firebaseUser.email || "",
                role,
                nama
              };
              await setDoc(userDocRef, profileData);
              setCurrentUserProfile(profileData);
            }
          }

          if (profileData.role === "guru") {
            setActiveTab("dashboard");
          }
          setDbLoading(false);

        } catch (err) {
          console.error("Gagal inisialisasi profile pengguna:", err);
          setDbLoading(false);
        }
      } else {
        setCurrentUser(null);
        setCurrentUserProfile(null);
        setDbLoading(false);
      }
    });

    return () => unsubscribeAuth();
  }, []);

  // Auto-set and lock class filters for logged-in Guru/Teacher
  useEffect(() => {
    if (currentUserProfile?.role === "guru" && state.kelas.length > 0) {
      const profileName = (currentUserProfile?.nama || "").trim().toLowerCase();
      const profileUsername = (currentUserProfile?.username || "").trim().toLowerCase();
      
      const gK = state.kelas.find(k => {
        const wk = (k.waliKelas || "").trim().toLowerCase();
        return wk && (
          wk === profileName || 
          wk.includes(profileName) || 
          profileName.includes(wk) || 
          wk === profileUsername || 
          wk.includes(profileUsername) || 
          profileUsername.includes(wk)
        );
      }) || state.kelas[0];
      
      if (gK) {
        setSelectedKelasFilter(prev => prev !== gK.id ? gK.id : prev);
        setSelectedKelasFilterIntra(prev => prev !== gK.id ? gK.id : prev);
        setSelectedKelasFilterKokuri(prev => prev !== gK.id ? gK.id : prev);
        setSelectedKelasFilterKehadiran(prev => prev !== gK.id ? gK.id : prev);
        setPrintKelasId(prev => prev !== gK.id ? gK.id : prev);
      }
    }
  }, [state.kelas, currentUserProfile]);

  // 2. Real-time Firestore synchronized listeners
  useEffect(() => {
    if (!currentUser || !currentUserProfile) return;

    // Listen to users list if the user is an admin
    let unsubscribeUsers = () => {};
    if (currentUserProfile?.role === "admin") {
      unsubscribeUsers = onSnapshot(collection(db, "users"), (snapshot) => {
        const uList: UserProfile[] = [];
        snapshot.forEach((doc) => {
          uList.push(doc.data() as UserProfile);
        });
        setAllUsersList(uList);
      }, (error) => {
        console.error("Users list listener failed: ", error);
      });
    }

    // Real-time synchronization
    const unmKelas = onSnapshot(collection(db, "kelas"), (snap) => {
      const items: Kelas[] = [];
      snap.forEach((doc) => items.push(doc.data() as Kelas));
      setState(prev => ({ ...prev, kelas: items }));
      
      if (items.length > 0) {
        setSelectedKelasFilter(prev => prev || items[0].id);
        setSelectedKelasFilterIntra(prev => prev || items[0].id);
        setSelectedKelasFilterKokuri(prev => prev || items[0].id);
        setPrintKelasId(prev => prev || items[0].id);
      }
    }, (error) => handleFirestoreError(error, OperationType.LIST, "kelas"));

    const unmSiswa = onSnapshot(collection(db, "siswa"), (snap) => {
      const items: Siswa[] = [];
      snap.forEach((doc) => items.push(doc.data() as Siswa));
      setState(prev => ({ ...prev, siswa: items }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "siswa"));

    const unmKategoriIntra = onSnapshot(collection(db, "kategoriIntrakurikuler"), async (snap) => {
      const items: KategoriIntrakurikuler[] = [];
      snap.forEach((doc) => items.push(doc.data() as KategoriIntrakurikuler));
      
      if (items.length === 0 && !snap.metadata.fromCache) {
        // Auto-seed Categories for backward compatibility
        try {
          // Check if TP exists to ensure we aren't seeding a completely erased DB unnecessarily
          const tpQ = await getDocs(collection(db, "tujuanPembelajaran"));
          if (!tpQ.empty && currentUserProfile?.role === "admin") {
             for (const kat of initialKategoriIntrakurikuler) {
               await setDoc(doc(db, "kategoriIntrakurikuler", kat.id), kat);
             }
          } else if (!tpQ.empty) {
             setState(prev => ({ ...prev, kategoriIntrakurikuler: initialKategoriIntrakurikuler }));
             return;
          }
        } catch(e) {}
      }
      
      setState(prev => ({ ...prev, kategoriIntrakurikuler: items }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "kategoriIntrakurikuler"));

    const unmTp = onSnapshot(collection(db, "tujuanPembelajaran"), (snap) => {
      const items: TujuanPembelajaran[] = [];
      snap.forEach((doc) => items.push(doc.data() as TujuanPembelajaran));
      setState(prev => ({ ...prev, tujuanPembelajaran: items }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "tujuanPembelajaran"));

    const unmSub = onSnapshot(collection(db, "subdimensiKokurikuler"), (snap) => {
      const items: SubdimensiKokurikuler[] = [];
      snap.forEach((doc) => items.push(doc.data() as SubdimensiKokurikuler));
      setState(prev => ({ ...prev, subdimensiKokurikuler: items }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "subdimensiKokurikuler"));

    const unmLabelP5 = onSnapshot(collection(db, "labelP5"), async (snap) => {
      const items: LabelP5[] = [];
      snap.forEach((doc) => items.push(doc.data() as LabelP5));
      
      const isAdminByEmail = currentUser?.email === "kusumamijen.dmk@gmail.com" || currentUser?.email === "admin@paud.local";
      const isAdmin = currentUserProfile?.role === "admin" || isAdminByEmail;

      if (items.length === 0 && !snap.metadata.fromCache && isAdmin) {
        const defaultLabels = [
          { id: "L-01", namaLabel: "Berkembang", order: 1 },
          { id: "L-02", namaLabel: "Cakap", order: 2 },
          { id: "L-03", namaLabel: "Mahir", order: 3 },
        ];
        for (const lbl of defaultLabels) {
          await setDoc(doc(db, "labelP5", lbl.id), lbl);
        }
      }
      
      setState(prev => ({ ...prev, labelP5: items.sort((a,b) => a.order - b.order) }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "labelP5"));

    const unmNilaiIntra = onSnapshot(collection(db, "nilaiIntrakurikuler"), (snap) => {
      const items: NilaiIntrakurikuler[] = [];
      snap.forEach((doc) => items.push(doc.data() as NilaiIntrakurikuler));
      setState(prev => ({ ...prev, nilaiIntrakurikuler: items }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "nilaiIntrakurikuler"));

    const unmNilaiKokuri = onSnapshot(collection(db, "nilaiKokurikuler"), (snap) => {
      const items: NilaiKokurikuler[] = [];
      snap.forEach((doc) => items.push(doc.data() as NilaiKokurikuler));
      setState(prev => ({ ...prev, nilaiKokurikuler: items }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "nilaiKokurikuler"));

    const unmCatatan = onSnapshot(collection(db, "catatanAnak"), (snap) => {
      const items: CatatanAnak[] = [];
      snap.forEach((doc) => items.push(doc.data() as CatatanAnak));
      setState(prev => ({ ...prev, catatanAnak: items }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "catatanAnak"));

    const unmKehadiran = onSnapshot(collection(db, "kehadiran"), (snap) => {
      const items: Kehadiran[] = [];
      snap.forEach((doc) => items.push(doc.data() as Kehadiran));
      setState(prev => ({ ...prev, kehadiran: items }));
      setDbLoading(false);
    }, (error) => handleFirestoreError(error, OperationType.LIST, "kehadiran"));

    const unmDataSekolah = onSnapshot(collection(db, "dataSekolah"), (snap) => {
      let ds: DataSekolah = {
        namaSekolah: "",
        alamat: "",
        kepalaSekolah: "",
        logo: "",
        semester: "2 (Genap)",
        thAjaran: "2025/2026",
        tglRaport: ""
      };
      snap.forEach((doc) => {
        if (doc.id === "default") {
          ds = doc.data() as DataSekolah;
        }
      });
      if (!ds.namaSekolah) {
        ds = seedData.dataSekolah;
        setDoc(doc(db, "dataSekolah", "default"), seedData.dataSekolah).catch(e => console.error(e));
      }
      setState(prev => ({ ...prev, dataSekolah: ds }));
    }, (error) => handleFirestoreError(error, OperationType.LIST, "dataSekolah"));

    return () => {
      unsubscribeUsers();
      unmKelas();
      unmSiswa();
      unmKategoriIntra();
      unmTp();
      unmSub();
      unmLabelP5();
      unmNilaiIntra();
      unmNilaiKokuri();
      unmCatatan();
      unmKehadiran();
      unmDataSekolah();
    };
  }, [currentUser, currentUserProfile]);

  // Seeding Template data to initial empty firestore DB
  const handleSeedDatabase = async () => {
    if (!currentUserProfile || currentUserProfile.role !== "admin") {
      alert("Hanya Admin yang dapat menginisialisasi database template!");
      return;
    }
    if (!confirm("Apakah Anda yakin ingin mengisi database kosong ini dengan data contoh template Kurikulum Merdeka PAUD? Ini akan mempercepat proses uji coba.")) {
      return;
    }

    setDbLoading(true);
    try {
      await setDoc(doc(db, "dataSekolah", "default"), seedData.dataSekolah);

      for (const k of seedData.kelas) {
        await setDoc(doc(db, "kelas", k.id), k);
      }
      for (const s of seedData.siswa) {
        await setDoc(doc(db, "siswa", s.id), s);
      }
      for (const tp of seedData.tujuanPembelajaran) {
        await setDoc(doc(db, "tujuanPembelajaran", tp.id), tp);
      }
      for (const sub of seedData.subdimensiKokurikuler) {
        await setDoc(doc(db, "subdimensiKokurikuler", sub.id), sub);
      }
      for (const ni of seedData.nilaiIntrakurikuler) {
        await setDoc(doc(db, "nilaiIntrakurikuler", `${ni.idSiswa}_${ni.idTp}`), ni);
      }
      for (const nk of seedData.nilaiKokurikuler) {
        await setDoc(doc(db, "nilaiKokurikuler", `${nk.idSiswa}_${nk.idSubdimensi}`), nk);
      }
      for (const kat of seedData.kategoriIntrakurikuler) {
        await setDoc(doc(db, "kategoriIntrakurikuler", kat.id), kat);
      }
      for (const c of seedData.catatanAnak) {
        await setDoc(doc(db, "catatanAnak", c.idSiswa), c);
      }
      for (const kh of seedData.kehadiran) {
        await setDoc(doc(db, "kehadiran", kh.idSiswa), kh);
      }

      alert("Inisialisasi database template sukses!");
    } catch (err) {
      console.error(err);
      alert("Gagal melakukan semai template database: " + (err instanceof Error ? err.message : String(err)));
    } finally {
      setDbLoading(false);
    }
  };

  // --- ACTIONS: AUTHENTICATION & LOGIN ---
  const handleUsernameLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!loginUsername || !loginPassword) {
      setLoginError("Nama pengguna dan kata sandi harus diisi.");
      return;
    }
    if (loginPassword.length < 6) {
      setLoginError("Kata sandi minimal harus 6 karakter.");
      return;
    }

    setLoginLoading(true);
    setLoginError("");

    try {
      const formattedUsername = loginUsername.trim().toLowerCase();
      const email = `${formattedUsername}@paud.local`;

      // 1. Check if admin exists in database, otherwise auto-seed first account for ease of access
      let adminExists = false;
      try {
        const uDoc = await getDoc(doc(db, "users", "admin_anchor"));
        if (uDoc.exists()) {
          adminExists = true;
        }
      } catch (err) {
        console.log("Checking admin anchor failed, might be empty", err);
      }

      // Create initial administrator instantly on demand if not found yet
      if (formattedUsername === "admin" && !adminExists) {
        try {
          const authResult = await createUserWithEmailAndPassword(auth, "admin@paud.local", loginPassword);
          await setDoc(doc(db, "users", authResult.user.uid), {
            uid: authResult.user.uid,
            email: "admin@paud.local",
            role: "admin",
            nama: "Admin Utama",
            username: "admin",
            password: loginPassword,
            isCreatedInAuth: true
          });
          await setDoc(doc(db, "users", "admin_anchor"), { exists: true });
          setLoginLoading(false);
          return;
        } catch (authErr: any) {
          if (authErr.code !== "auth/email-already-in-use") {
            throw authErr;
          }
        }
      }

      // 2. Direct authentication first; fallback to specific temp doc check if account has not been lazy-created in Auth yet
      let loginSuccess = false;
      try {
        await signInWithEmailAndPassword(auth, email, loginPassword);
        loginSuccess = true;
      } catch (authErr: any) {
        // If the account has not been lazy-created inside Auth yet, we look for a specific 'temp_<username>' document
        if (authErr.code === "auth/user-not-found" || authErr.code === "auth/invalid-credential" || authErr.code === "auth/wrong-password") {
          const tempDocRef = doc(db, "users", `temp_${formattedUsername}`);
          const tempSnap = await getDoc(tempDocRef);

          if (tempSnap.exists()) {
            const userData = tempSnap.data() as UserProfile;
            if (userData.password === loginPassword) {
              // Lazy-create the user in Firebase Auth
              const authResult = await createUserWithEmailAndPassword(auth, email, loginPassword);
              const realUid = authResult.user.uid;

              await setDoc(doc(db, "users", realUid), {
                ...userData,
                uid: realUid,
                isCreatedInAuth: true
              });

              try {
                await deleteDoc(tempDocRef);
              } catch (delErr) {
                console.warn("Non-blocking delete temp user doc failed:", delErr);
              }

              loginSuccess = true;
            } else {
              setLoginError("Kata sandi salah. Silakan coba lagi.");
              setLoginLoading(false);
              return;
            }
          } else {
            setLoginError("Nama pengguna belum terdaftar. Silakan hubungi Admin Sekolah.");
            setLoginLoading(false);
            return;
          }
        } else {
          throw authErr;
        }
      }
    } catch (e: any) {
      console.error(e);
      let errMsg = "Masuk gagal: ";
      if (e.code === "auth/invalid-credential" || e.code === "auth/wrong-password") {
        errMsg += "Kombinasi nama pengguna/sandi salah.";
      } else {
        errMsg += e.message || String(e);
      }
      setLoginError(errMsg);
    } finally {
      setLoginLoading(false);
    }
  };

  const handleCreateTeacherAccount = async (e: React.FormEvent) => {
    e.preventDefault();
    if (currentUserProfile?.role !== "admin") {
      showToast("Akses Ditolak: Hanya Admin yang dapat menambahkan akun baru!", "error");
      return;
    }
    if (!newTeacherNama || !newTeacherUsername || !newTeacherPassword) {
      showToast("Mohon lengkapi seluruh kolom formulir!", "error");
      return;
    }
    if (newTeacherPassword.length < 6) {
      showToast("Eror: Kata sandi minimal harus terdiri dari 6 karakter!", "error");
      return;
    }

    const cleanUsername = newTeacherUsername.trim().toLowerCase();
    if (cleanUsername.includes(" ")) {
      showToast("Eror: Nama pengguna tidak boleh memiliki spasi!", "error");
      return;
    }

    // Check pre-existence
    const isDuplicate = allUsersList.some(u => u.username === cleanUsername || u.email === `${cleanUsername}@paud.local`);
    if (isDuplicate) {
      showToast("Eror: Nama pengguna tersebut sudah terdaftar! Pilih nama pengguna lain.", "error");
      return;
    }

    try {
      const tempId = `temp_${cleanUsername}`;
      await setDoc(doc(db, "users", tempId), {
        uid: tempId,
        nama: newTeacherNama,
        username: cleanUsername,
        password: newTeacherPassword,
        role: newTeacherRole,
        email: `${cleanUsername}@paud.local`,
        isCreatedInAuth: false
      });

      showToast(`Sukses menambahkan akun ${newTeacherNama} (${cleanUsername})!`, "success");
      
      setNewTeacherNama("");
      setNewTeacherUsername("");
      setNewTeacherPassword("");
      setNewTeacherRole("guru");
    } catch (err) {
      showToast("Gagal menambahkan pengguna baru: " + (err instanceof Error ? err.message : String(err)), "error");
    }
  };

  // Custom Modal/Dialog operations
  const handleConfirmDelete = async () => {
    if (!userToDelete) return;
    try {
      await deleteDoc(doc(db, "users", userToDelete.uid));
      showToast(`Berhasil menghapus akun ${userToDelete.nama}!`, "success");
      setUserToDelete(null);
    } catch (e) {
      showToast("Gagal menghapus pengguna: " + (e instanceof Error ? e.message : String(e)), "error");
    }
  };

  const handleConfirmEditPassword = async () => {
    if (!userToEditPassword) return;
    if (newPasswordValue.length < 6) {
      showToast("Gagal: Kata sandi minimal harus terdiri dari 6 karakter!", "error");
      return;
    }
    try {
      await setDoc(doc(db, "users", userToEditPassword.uid), { ...userToEditPassword, password: newPasswordValue });
      showToast(`Berhasil memperbarui kata sandi untuk ${userToEditPassword.nama}!`, "success");
      setUserToEditPassword(null);
      setNewPasswordValue("");
    } catch (e) {
      showToast("Gagal merubah password: " + (e instanceof Error ? e.message : String(e)), "error");
    }
  };

  const handleConfirmChangeRole = async () => {
    if (!userToChangeRole) return;
    try {
      await setDoc(doc(db, "users", userToChangeRole.uid), { ...userToChangeRole, role: targetRoleValue });
      showToast(`Berhasil memperbarui peran ${userToChangeRole.nama} menjadi ${targetRoleValue.toUpperCase()}!`, "success");
      setUserToChangeRole(null);
    } catch (e) {
      showToast("Gagal mengubah peran: " + (e instanceof Error ? e.message : String(e)), "error");
    }
  };

  if (!mounted) {
    return (
      <div className="flex items-center justify-center min-h-screen bg-neutral-50">
        <div className="text-center">
          <div className="w-12 h-12 border-4 border-emerald-600 border-t-transparent rounded-full animate-spin mx-auto mb-4"></div>
          <p className="text-neutral-500 font-medium">Memuat Aplikasi Raport PAUD...</p>
        </div>
      </div>
    );
  }


  // --- ACTIONS: KELAS ---
  const handleSaveKelas = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!kelasForm.namaKelas) return;

    if (currentUserProfile?.role !== "admin") {
      alert("Akses Ditolak: Hanya Admin yang dapat mengelola data Kelas!");
      return;
    }

    try {
      const activeTh = state.dataSekolah.thAjaran || "2025/2026";
      if (editingKelas) {
        const docRef = doc(db, "kelas", editingKelas.id);
        await setDoc(docRef, { ...editingKelas, ...kelasForm, thAjaran: activeTh });
        setEditingKelas(null);
      } else {
        const id = "K-" + Date.now();
        const docRef = doc(db, "kelas", id);
        const newK: Kelas = {
          id,
          namaKelas: kelasForm.namaKelas,
          waliKelas: kelasForm.waliKelas,
          nuptkNgty: kelasForm.nuptkNgty,
          thAjaran: activeTh
        };
        await setDoc(docRef, newK);
      }
      setKelasForm({ namaKelas: "", waliKelas: "", nuptkNgty: "" });
      setShowAddKelas(false);
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "kelas");
    }
  };

  const handleDeleteKelas = async (id: string) => {
    if (currentUserProfile?.role !== "admin") {
      showToast("Akses Ditolak: Hanya Admin yang dapat menghapus data Kelas!", "error");
      return;
    }
    const targetKelas = state.kelas.find(k => k.id === id);
    if (!targetKelas) return;
    setKelasToDelete(targetKelas);
  };

  const confirmDeleteKelas = async () => {
    if (!kelasToDelete) return;
    try {
      await deleteDoc(doc(db, 'kelas', kelasToDelete.id));
      const studentsToCleanup = state.siswa.filter(s => s.idKelas === kelasToDelete.id);
      for (const student of studentsToCleanup) {
        await updateDoc(doc(db, 'siswa', student.id), { idKelas: "" });
      }
      showToast(`Kelas ${kelasToDelete.namaKelas} berhasil dihapus!`, "success");
      setKelasToDelete(null);
    } catch (err: any) {
      handleFirestoreError(err, OperationType.DELETE, `kelas/${kelasToDelete.id}`);
      setKelasToDelete(null);
    }
  };

  // --- ACTIONS: SISWA ---
  const handleExportSiswa = () => {
    const dataToExport = state.siswa.map(s => {
      const kelas = state.kelas.find(k => k.id === s.idKelas);
      return {
        "ID Siswa (Biarkan kosong jika baru)": s.id,
        "Nama Siswa": s.namaSiswa,
        "NISN": s.nisn,
        "Alamat": s.alamat,
        "Nama Ayah": s.namaAyah,
        "Pekerjaan Ayah": s.pekerjaanAyah,
        "Nama Ibu": s.namaIbu,
        "Pekerjaan Ibu": s.pekerjaanIbu,
        "No HP": s.noHp,
        "Tgl Lahir (YYYY-MM-DD)": s.tglLahir,
        "Tempat Lahir": s.tempatLahir,
        "Jenis Kelamin": s.jenisKelamin,
        "Anak Ke": s.anakKe,
        "Tinggi Badan": s.tb,
        "Berat Badan": s.bb,
        "Agama": s.agama,
        "ID Kelas": s.idKelas,
        "Nama Kelas (Info)": kelas?.namaKelas || "",
      };
    });

    if (dataToExport.length === 0) {
      dataToExport.push({
        "ID Siswa (Biarkan kosong jika baru)": "",
        "Nama Siswa": "Contoh Nama",
        "NISN": "1234567890123456",
        "Alamat": "Jl. Contoh",
        "Nama Ayah": "Budi",
        "Pekerjaan Ayah": "Wiraswasta",
        "Nama Ibu": "Rina",
        "Pekerjaan Ibu": "Guru",
        "No HP": "08123456789",
        "Tgl Lahir (YYYY-MM-DD)": "2020-01-01",
        "Tempat Lahir": "Kota A",
        "Jenis Kelamin": "Laki-laki",
        "Anak Ke": 1,
        "Tinggi Badan": "100",
        "Berat Badan": "15",
        "Agama": "Islam",
        "ID Kelas": state.kelas[0]?.id || "",
        "Nama Kelas (Info)": state.kelas[0]?.namaKelas || "Nama Kelas",
      });
    }

    const ws = XLSX.utils.json_to_sheet(dataToExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "DataSiswa");
    XLSX.writeFile(wb, "Template_Data_Siswa.xlsx");
    
    showToast("Template Excel berhasil diunduh.", "success");
  };

  const handleImportSiswa = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data, { type: "array" });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      
      const jsonData = XLSX.utils.sheet_to_json(worksheet) as any[];

      let countNew = 0;
      let countUpdate = 0;

      for (const row of jsonData) {
        if (!row["Nama Siswa"]) continue;

        let idSiswa = row["ID Siswa (Biarkan kosong jika baru)"];
        let isUpdate = !!idSiswa && state.siswa.some(s => s.id === String(idSiswa));

        let idKelas = row["ID Kelas"];
        if (!idKelas && state.kelas.length > 0) {
           idKelas = state.kelas[0].id;
        }

        const newId = isUpdate ? String(idSiswa) : `siswa-${Date.now()}-${Math.floor(Math.random()*1000)}`;

        const activeTh = state.dataSekolah.thAjaran || "2025/2026";
        const activeSem = state.dataSekolah.semester || "1 (Ganjil)";

        const newS: Siswa = {
          id: newId,
          namaSiswa: String(row["Nama Siswa"] || ""),
          nisn: String(row["NISN"] || ""),
          alamat: String(row["Alamat"] || ""),
          namaAyah: String(row["Nama Ayah"] || ""),
          pekerjaanAyah: String(row["Pekerjaan Ayah"] || ""),
          namaIbu: String(row["Nama Ibu"] || ""),
          pekerjaanIbu: String(row["Pekerjaan Ibu"] || ""),
          noHp: String(row["No HP"] || ""),
          tglLahir: excelDateToIsoString(row["Tgl Lahir (YYYY-MM-DD)"] || row["Tgl Lahir"] || ""),
          tempatLahir: String(row["Tempat Lahir"] || ""),
          jenisKelamin: String(row["Jenis Kelamin"] || "Laki-laki"),
          anakKe: Number(row["Anak Ke"]) || 1,
          tb: String(row["Tinggi Badan"] || ""),
          bb: String(row["Berat Badan"] || ""),
          agama: String(row["Agama"] || ""),
          idKelas: String(idKelas || ""),
          thAjaran: activeTh
        };

        const docRef = doc(db, "siswa", newId);
        await setDoc(docRef, newS);

        if (!isUpdate) {
          const attId = `${newId}_${activeTh.replace(/\//g, "-")}_${activeSem}`;
          const attRef = doc(db, "kehadiran", attId);
          await setDoc(attRef, {
            idSiswa: newId,
            sakit: 0,
            ijin: 0,
            tanpaKet: 0,
            thAjaran: activeTh,
            semester: activeSem
          });
        }

        if (isUpdate) countUpdate++;
        else countNew++;
      }
      
      showToast(`Import selesai: ${countNew} data baru, ${countUpdate} diupdate.`, "success");
    } catch (err) {
      console.error(err);
      showToast("Gagal mengimpor file Excel", "error");
    }

    if (fileInputRef.current) {
        fileInputRef.current.value = "";
    }
  };

  const handleSaveSiswa = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!siswaForm.namaSiswa || !siswaForm.idKelas) {
      alert("Nama siswa dan Kelas wajib dipilih!");
      return;
    }

    try {
      const activeTh = state.dataSekolah.thAjaran || "2025/2026";
      const activeSem = state.dataSekolah.semester || "1 (Ganjil)";
      if (editingSiswa) {
        const docRef = doc(db, "siswa", editingSiswa.id);
        await setDoc(docRef, { ...editingSiswa, ...siswaForm as Siswa, thAjaran: activeTh });
        setEditingSiswa(null);
      } else {
        const id = "S-" + Date.now();
        const docRef = doc(db, "siswa", id);
        const newS: Siswa = {
          id,
          namaSiswa: siswaForm.namaSiswa || "",
          nisn: siswaForm.nisn || "-",
          alamat: siswaForm.alamat || "-",
          namaAyah: siswaForm.namaAyah || "-",
          pekerjaanAyah: siswaForm.pekerjaanAyah || "-",
          namaIbu: siswaForm.namaIbu || "-",
          pekerjaanIbu: siswaForm.pekerjaanIbu || "-",
          noHp: siswaForm.noHp || "-",
          tglLahir: siswaForm.tglLahir || "",
          tempatLahir: siswaForm.tempatLahir || "",
          jenisKelamin: siswaForm.jenisKelamin || "Laki-laki",
          anakKe: Number(siswaForm.anakKe) || 1,
          tb: siswaForm.tb || "100",
          bb: siswaForm.bb || "15",
          agama: siswaForm.agama || "Islam",
          idKelas: siswaForm.idKelas || "",
          thAjaran: activeTh
        };
        await setDoc(docRef, newS);
 
        const attId = `${id}_${activeTh.replace(/\//g, "-")}_${activeSem}`;
        const attRef = doc(db, "kehadiran", attId);
        await setDoc(attRef, {
          idSiswa: id,
          sakit: 0,
          ijin: 0,
          tanpaKet: 0,
          thAjaran: activeTh,
          semester: activeSem
        });
      }

      setSiswaForm({
        namaSiswa: "", nisn: "", alamat: "", namaAyah: "", pekerjaanAyah: "", namaIbu: "", pekerjaanIbu: "",
        noHp: "", tglLahir: "", tempatLahir: "", jenisKelamin: "Laki-laki", anakKe: 1, tb: "100", bb: "15", agama: "Islam", idKelas: ""
      });
      setShowAddSiswa(false);
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "siswa");
    }
  };

  const handleConfirmDeleteSiswa = async () => {
    if (!siswaToDelete) return;
    const id = siswaToDelete.id;
    try {
      await deleteDoc(doc(db, "siswa", id));
      const relatedIntra = state.nilaiIntrakurikuler.filter(n => n.idSiswa === id);
      for (const item of relatedIntra) {
        await deleteDoc(doc(db, "nilaiIntrakurikuler", `${id}_${item.idTp}`));
      }
      const relatedKokuri = state.nilaiKokurikuler.filter(n => n.idSiswa === id);
      for (const item of relatedKokuri) {
        await deleteDoc(doc(db, "nilaiKokurikuler", `${id}_${item.idSubdimensi}`));
      }
      await deleteDoc(doc(db, "catatanAnak", id));
      await deleteDoc(doc(db, "kehadiran", id));

      if (printSiswaId === id) setPrintSiswaId("");
      showToast(`Berhasil menghapus data siswa ${siswaToDelete.namaSiswa}!`, "success");
      setSiswaToDelete(null);
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, `siswa/${id}`);
      showToast("Gagal menghapus siswa: " + (err instanceof Error ? err.message : String(err)), "error");
    }
  };

  // --- ACTIONS: KATEGORI INTRA ---
  const handleSaveKategori = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!kategoriForm.namaKategori) {
      showToast("Nama Kategori wajib diisi", "error");
      return;
    }

    if (currentUserProfile?.role !== "admin") {
      alert("Akses Ditolak: Hanya Admin yang dapat mengelola Kategori Intrakurikuler.");
      return;
    }

    try {
      if (editingKategori) {
        const docRef = doc(db, "kategoriIntrakurikuler", editingKategori.id);
        await updateDoc(docRef, kategoriForm);
        showToast("Berhasil update kategori", "success");
      } else {
        const id = "KAT-" + Date.now();
        const docRef = doc(db, "kategoriIntrakurikuler", id);
        const newKat: KategoriIntrakurikuler = {
          id,
          namaKategori: kategoriForm.namaKategori,
        };
        await setDoc(docRef, newKat);
        showToast("Berhasil menambah kategori", "success");
      }
      setKategoriForm({ namaKategori: "" });
      setShowAddKategori(false);
      setEditingKategori(null);
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "kategoriIntrakurikuler");
    }
  };

  const handleConfirmDeleteKategori = async () => {
    if (!kategoriToDelete) return;
    try {
      const id = kategoriToDelete.id;
      await deleteDoc(doc(db, "kategoriIntrakurikuler", id));

      const relatedTp = state.tujuanPembelajaran.filter(tp => tp.idKategori === id);
      for (const item of relatedTp) {
        await deleteDoc(doc(db, "tujuanPembelajaran", item.id));
        const relatedNilai = state.nilaiIntrakurikuler.filter(n => n.idTp === item.id);
        for (const n of relatedNilai) {
          await deleteDoc(doc(db, "nilaiIntrakurikuler", `${n.idSiswa}_${item.id}`));
        }
      }

      showToast(`Berhasil menghapus kategori ${kategoriToDelete.namaKategori}!`, "success");
      setKategoriToDelete(null);
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, `kategoriIntrakurikuler/${kategoriToDelete?.id}`);
      showToast("Gagal menghapus kategori: " + (err instanceof Error ? err.message : String(err)), "error");
    }
  };

  // --- ACTIONS: INTRA ---
  const handleAddTp = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!tpForm.deskripsi || !tpForm.idKelas) return;

    if (currentUserProfile?.role !== "admin") {
      alert("Akses Ditolak: Hanya Admin yang dapat mengelola butir Tujuan Pembelajaran!");
      return;
    }

    try {
      const activeTh = state.dataSekolah.thAjaran || "2025/2026";
      if (editingTp) {
        const updatedTp: TujuanPembelajaran = {
          ...editingTp,
          idKategori: tpForm.idKategori,
          deskripsi: tpForm.deskripsi,
          idKelas: tpForm.idKelas,
          aktivitasMetode: tpForm.aktivitasMetode,
          thAjaran: activeTh
        };
        await setDoc(doc(db, "tujuanPembelajaran", editingTp.id), updatedTp);
        showToast("Berhasil mengubah Tujuan Pembelajaran (TP)!", "success");
        setEditingTp(null);
      } else {
        const id = "TP-" + Date.now();
        const newTp: TujuanPembelajaran = {
          id,
          idKategori: tpForm.idKategori,
          deskripsi: tpForm.deskripsi,
          idKelas: tpForm.idKelas,
          aktivitasMetode: tpForm.aktivitasMetode,
          thAjaran: activeTh
        };
        await setDoc(doc(db, "tujuanPembelajaran", id), newTp);
        showToast("Berhasil menambahkan Tujuan Pembelajaran (TP)!", "success");
      }
      setTpForm({ ...tpForm, deskripsi: "", aktivitasMetode: "" });
      setShowAddTp(false);
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "tujuanPembelajaran");
    }
  };

  const handleDeleteTp = (tp: TujuanPembelajaran) => {
    if (currentUserProfile?.role !== "admin") {
      alert("Akses Ditolak: Hanya Admin yang dapat menghapus butir Tujuan Pembelajaran!");
      return;
    }
    setTpToDelete(tp);
  };

  const handleConfirmDeleteTp = async () => {
    if (!tpToDelete) return;
    try {
      await deleteDoc(doc(db, "tujuanPembelajaran", tpToDelete.id));
      const relatedNilai = state.nilaiIntrakurikuler.filter(n => n.idTp === tpToDelete.id);
      for (const item of relatedNilai) {
        await deleteDoc(doc(db, "nilaiIntrakurikuler", `${item.idSiswa}_${tpToDelete.id}`));
      }
      showToast("Berhasil menghapus Tujuan Pembelajaran (TP)!", "success");
      setTpToDelete(null);
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, `tujuanPembelajaran/${tpToDelete.id}`);
      showToast("Gagal menghapus Tujuan Pembelajaran: " + (err instanceof Error ? err.message : String(err)), "error");
    }
  };

  // --- ACTIONS: KOKURI ---
  const handleAddSub = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!subForm.namaSubdimensi || !subForm.idKelas) return;

    if (currentUserProfile?.role !== "admin") {
      alert("Akses Ditolak: Hanya Admin yang dapat mengelola butir Subdimensi P5!");
      return;
    }

    try {
      const activeTh = state.dataSekolah.thAjaran || "2025/2026";
      const id = editingSub ? editingSub.id : ("SD-" + Date.now());
      const newSub: SubdimensiKokurikuler = {
        id,
        namaSubdimensi: subForm.namaSubdimensi,
        idKelas: subForm.idKelas,
        descBerkembang: subForm.descBerkembang || "",
        descCakap: subForm.descCakap || "",
        descMahir: subForm.descMahir || "",
        capaian: subForm.capaian || {},
        thAjaran: activeTh
      };
      await setDoc(doc(db, "subdimensiKokurikuler", id), newSub);
      setSubForm({ namaSubdimensi: "", idKelas: subForm.idKelas, descBerkembang: "", descCakap: "", descMahir: "", capaian: {} });
      setEditingSub(null);
      setShowAddSub(false);
      showToast(editingSub ? "Subdimensi berhasil diperbarui" : "Subdimensi berhasil ditambahkan");
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "subdimensiKokurikuler");
    }
  };

  const handleDeleteSub = async (id: string) => {
    if (currentUserProfile?.role !== "admin") {
      alert("Akses Ditolak: Hanya Admin yang dapat menghapus butir Subdimensi P5!");
      return;
    }

    if (confirm("Hapus Subdimensi Kokurikuler ini beserta nilainya?")) {
      try {
        await deleteDoc(doc(db, "subdimensiKokurikuler", id));
        const relatedNilai = state.nilaiKokurikuler.filter(n => n.idSubdimensi === id);
        for (const item of relatedNilai) {
          await deleteDoc(doc(db, "nilaiKokurikuler", `${item.idSiswa}_${id}`));
        }
      } catch (err) {
        handleFirestoreError(err, OperationType.DELETE, `subdimensiKokurikuler/${id}`);
      }
    }
  };

  const handleSaveLabel = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!labelForm.namaLabel) return;

    if (!isAdmin) {
      showToast("Gagal: Hanya Admin yang dapat mengelola label", "error");
      return;
    }

    try {
      const id = editingLabel ? editingLabel.id : ("L-" + Date.now());
      const newLabel: LabelP5 = {
        id,
        namaLabel: labelForm.namaLabel,
        order: Number(labelForm.order) || (state.labelP5.length + 1)
      };
      await setDoc(doc(db, "labelP5", id), newLabel);
      setLabelForm({ namaLabel: "", order: 0 });
      setEditingLabel(null);
      showToast(editingLabel ? "Label berhasil diperbarui" : "Label berhasil ditambahkan");
    } catch (err) {
      console.error("Error saving label:", err);
      handleFirestoreError(err, OperationType.WRITE, "labelP5");
    }
  };

  const handleDeleteLabel = async (label: LabelP5) => {
    if (!isAdmin) return;
    try {
      await deleteDoc(doc(db, "labelP5", label.id));
      showToast("Label berhasil dihapus");
    } catch (err) {
      handleFirestoreError(err, OperationType.DELETE, "labelP5");
    }
  };

  // --- ACTIONS: SAVE GRADES ---
  const handleSetGradeIntra = async (idSiswa: string, idTp: string, val: string) => {
    try {
      const th = state.dataSekolah.thAjaran || "2025/2026";
      const sem = state.dataSekolah.semester || "1 (Ganjil)";
      const docId = `${idSiswa}_${idTp}_${th.replace(/\//g, "-")}_${sem}`;
      const docRef = doc(db, "nilaiIntrakurikuler", docId);
      if (val === "") {
        // If we clear the grade, we might still want to keep the description, 
        // but usually, clearing grade means clearing the assessment. 
        // For safety, let's just clear the "nilai" field but keep the document if deskripsi exists.
        const snap = await getDoc(docRef);
        if (snap.exists() && snap.data().deskripsi) {
          await updateDoc(docRef, { nilai: "" });
        } else {
          await deleteDoc(docRef);
        }
      } else {
        await setDoc(docRef, { idSiswa, idTp, nilai: val, thAjaran: th, semester: sem }, { merge: true });
      }
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "nilaiIntrakurikuler");
    }
  };

  const handleUpdateDescriptionIntra = async (idSiswa: string, idTp: string, text: string) => {
    try {
      const th = state.dataSekolah.thAjaran || "2025/2026";
      const sem = state.dataSekolah.semester || "1 (Ganjil)";
      const docId = `${idSiswa}_${idTp}_${th.replace(/\//g, "-")}_${sem}`;
      const docRef = doc(db, "nilaiIntrakurikuler", docId);
      await setDoc(docRef, { idSiswa, idTp, deskripsi: text, thAjaran: th, semester: sem }, { merge: true });
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "nilaiIntrakurikuler/deskripsi");
    }
  };

  const handleSetGradeKokuri = async (idSiswa: string, idSubdimensi: string, val: string) => {
    try {
      const th = state.dataSekolah.thAjaran || "2025/2026";
      const sem = state.dataSekolah.semester || "1 (Ganjil)";
      const docId = `${idSiswa}_${idSubdimensi}_${th.replace(/\//g, "-")}_${sem}`;
      const docRef = doc(db, "nilaiKokurikuler", docId);
      if (val === "") {
        const snap = await getDoc(docRef);
        if (snap.exists() && snap.data().deskripsi) {
          await updateDoc(docRef, { nilai: "" });
        } else {
          await deleteDoc(docRef);
        }
      } else {
        await setDoc(docRef, { idSiswa, idSubdimensi, nilai: val, thAjaran: th, semester: sem }, { merge: true });
      }
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "nilaiKokurikuler");
    }
  };

  const handleUpdateDescriptionKokuri = async (idSiswa: string, idSubdimensi: string, text: string) => {
    try {
      const th = state.dataSekolah.thAjaran || "2025/2026";
      const sem = state.dataSekolah.semester || "1 (Ganjil)";
      const docId = `${idSiswa}_${idSubdimensi}_${th.replace(/\//g, "-")}_${sem}`;
      const docRef = doc(db, "nilaiKokurikuler", docId);
      await setDoc(docRef, { idSiswa, idSubdimensi, deskripsi: text, thAjaran: th, semester: sem }, { merge: true });
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "nilaiKokurikuler/deskripsi");
    }
  };

  // --- ACTIONS: SCHOOL & SETTINGS ---
  const handleUpdateSchool = async (field: keyof DataSekolah, val: string) => {
    if (currentUserProfile?.role !== "admin") {
      alert("Akses Ditolak: Hanya Admin yang dapat mengubah identitas data sekolah!");
      return;
    }

    try {
      const schoolRef = doc(db, "dataSekolah", "default");
      
      // Prevent Firestore 1MB limits by rejecting values that exceed safe margins
      if (field === "logo" && val.length > 600000) {
        alert("Ukuran file logo terlalu besar. Silakan upload ulang file gambar yang lebih kecil (di bawah 500 KB).");
        return;
      }

      // Safeguard: Check if any pre-existing logo in local state is too large
      // When a user updates another field (e.g. Alamat or Kepala Sekolah), the spread operator
      // { ...state.dataSekolah, [field]: val } will write the full logo string too.
      // If that logo was historically a massive base64, we safely clear it to prevent Firestore crashes.
      let logoToSave = state.dataSekolah.logo || "";
      if (logoToSave.length > 600000) {
        console.warn("Melindungi Firestore: Logo lama terlalu besar, dikosongkan pengaman.");
        logoToSave = "";
      }

      await setDoc(schoolRef, { 
        ...state.dataSekolah, 
        logo: field === "logo" ? val : logoToSave,
        [field]: val 
      });
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "dataSekolah/default");
    }
  };

  const handleLogoUpload = (e: React.ChangeEvent<HTMLInputElement>, field: "logo" | "logoSidebar" | "logoLogin" = "logo") => {
    const file = e.target.files?.[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onloadend = () => {
      const base64Str = reader.result as string;
      const img = new Image();
      img.onload = () => {
        const canvas = document.createElement("canvas");
        const maxW = 150;
        const maxH = 150;
        let width = img.width;
        let height = img.height;

        if (width > height) {
          if (width > maxW) {
            height = Math.round((height * maxW) / width);
            width = maxW;
          }
        } else {
          if (height > maxH) {
            width = Math.round((width * maxH) / height);
            height = maxH;
          }
        }

        canvas.width = width;
        canvas.height = height;

        const ctx = canvas.getContext("2d");
        if (ctx) {
          ctx.clearRect(0, 0, width, height);
          ctx.drawImage(img, 0, 0, width, height);
          
          const isPng = file.type === "image/png";
          const format = isPng ? "image/png" : "image/jpeg";
          const quality = isPng ? undefined : 0.75;
          const compressed = canvas.toDataURL(format, quality);
          handleUpdateSchool(field, compressed);
        } else {
          handleUpdateSchool(field, base64Str);
        }
      };
      img.onerror = () => {
        handleUpdateSchool(field, base64Str);
      };
      img.src = base64Str;
    };
    reader.readAsDataURL(file);
    if (e.target) e.target.value = "";
  };

  // --- ACTIONS: NOTES & AI ASSIST ---
  const handleSaveCatatan = async (idSiswa: string, text: string) => {
    try {
      const th = state.dataSekolah.thAjaran || "2025/2026";
      const sem = state.dataSekolah.semester || "1 (Ganjil)";
      const docId = `${idSiswa}_${th.replace(/\//g, "-")}_${sem}`;
      await setDoc(doc(db, "catatanAnak", docId), { idSiswa, catatan: text, thAjaran: th, semester: sem });
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "catatanAnak");
    }
  };

  const composeAiItemText = async (idSiswa: string, idItem: string, type: "intra" | "kokuri") => {
    const siswaItem = state.siswa.find(s => s.id === idSiswa);
    if (!siswaItem) return;

    let promptText = "";

    if (type === "intra") {
      const kat = state.kategoriIntrakurikuler.find(k => k.id === idItem);
      const catName = kat?.namaKategori || "";
      const catTps = state.tujuanPembelajaran.filter(tp => tp.idKelas === siswaItem.idKelas && tp.idKategori === idItem);
      const tpSummary = catTps.map((tp, idx) => {
        const score = state.nilaiIntrakurikuler.find(n => n.idSiswa === idSiswa && n.idTp === tp.id)?.nilai || "Berkembang";
        return `${idx + 1}. Tujuan Pembelajaran: "${tp.deskripsi}" (Capaian: ${score})`;
      }).join("\n");

      promptText = `Buatkan deskripsi/narasi perkembangan gabungan (paling banyak 2-3 kalimat) yang meresumekan (merekapitulasi) capaian belajar untuk kategori kompetensi "${catName}" bagi siswa bernama ${siswaItem.namaSiswa}. 
Berikut adalah rincian capaian setiap Tujuan Pembelajaran (TP) di kategori ini:
${tpSummary}

Tuliskan ulasan dalam bahasa Indonesia yang hangat, bersahabat, profesional, positif, dan santun. Fokus pada kelebihan, progres, dan apresiasi apa yang sudah dicapai anak, dengan menggunakan kata ganti 'ananda ${siswaItem.namaSiswa.split(" ")[0]}' atau '${siswaItem.namaSiswa.split(" ")[0]}'.`;
    } else {
      const sub = state.subdimensiKokurikuler.find(s => s.id === idItem);
      const itemDesc = sub?.namaSubdimensi || "";
      const itemScore = state.nilaiKokurikuler.find(n => n.idSiswa === idSiswa && n.idSubdimensi === idItem)?.nilai || "Berkembang";
      promptText = `Buatkan deskripsi ringkas, positif, dan santun (1-2 kalimat) untuk capaian "${itemDesc}" dengan predikat "${itemScore}" untuk siswa bernama ${siswaItem.namaSiswa}. Fokus pada apa yang sudah dicapai anak dengan bahasa yang hangat.`;
    }

    setGeneratingAiItem(idItem);
    try {
      const response = await fetch("/api/gemini/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          prompt: promptText,
          provider: state.dataSekolah.aiProvider || "groq",
          model: state.dataSekolah.aiModel,
          context: {
            namaSiswa: siswaItem.namaSiswa,
            namaKelas: state.kelas.find(k => k.id === siswaItem.idKelas)?.namaKelas || "PAUD",
            intrakurikuler: [], // Minimal context as we use manual prompt
            kokurikuler: []
          }
        })
      });

      const data = await response.json();
      if (data.text) {
        if (type === "intra") {
          handleUpdateDescriptionIntra(idSiswa, idItem, data.text);
        } else {
          handleUpdateDescriptionKokuri(idSiswa, idItem, data.text);
        }
      }
    } catch (e) {
      console.error(e);
    } finally {
      setGeneratingAiItem(null);
    }
  };

  const composeAiText = async (idSiswa: string) => {
    const siswaItem = state.siswa.find(s => s.id === idSiswa);
    if (!siswaItem) return;

    setGeneratingAi(idSiswa);
    try {
      const classItem = state.kelas.find(k => k.id === siswaItem.idKelas);
      const studentIntras = state.kategoriIntrakurikuler.map(kat => {
        const catDescDoc = state.nilaiIntrakurikuler.find(nn => nn.idSiswa === idSiswa && nn.idTp === kat.id);
        const catDesc = catDescDoc?.deskripsi ? ` - Catatan Capaian: ${catDescDoc.deskripsi}` : "";
        const childTps = state.tujuanPembelajaran
          .filter(tp => tp.idKelas === siswaItem.idKelas && tp.idKategori === kat.id)
          .map(tp => {
            const n = state.nilaiIntrakurikuler.find(nn => nn.idSiswa === idSiswa && nn.idTp === tp.id);
            return `${tp.deskripsi} (${n?.nilai || "Belum dinilai"})`;
          }).join(", ");
        return `Kategori ${kat.namaKategori}: [${childTps}].${catDesc}`;
      });

      const studentKokuris = state.subdimensiKokurikuler
        .filter(sub => sub.idKelas === siswaItem.idKelas)
        .map(sub => {
          const n = state.nilaiKokurikuler.find(nn => nn.idSiswa === idSiswa && nn.idSubdimensi === sub.id);
          const score = n?.nilai || "Belum dinilai";
          const desc = n?.deskripsi ? ` - Aktivitas projek: ${n.deskripsi}` : "";
          return `${sub.namaSubdimensi} (${score})${desc}`;
        });

      const response = await fetch("/api/gemini/generate", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          prompt: aiPromptCustom,
          provider: state.dataSekolah.aiProvider || "groq",
          model: state.dataSekolah.aiModel,
          context: {
            namaSiswa: siswaItem.namaSiswa,
            namaKelas: classItem?.namaKelas || "PAUD",
            intrakurikuler: studentIntras,
            kokurikuler: studentKokuris
          }
        })
      });

      const data = await response.json();
      if (data.text) {
        handleSaveCatatan(idSiswa, data.text);
        setAiPromptCustom("");
      } else {
        alert("Gagal memproses narasi AI: " + (data.error || "Umpan balik tidak ditemukan."));
      }
    } catch (e: any) {
      console.error(e);
      alert("Error memproses AI. Silakan periksa koneksi Anda.");
    } finally {
      setGeneratingAi(null);
    }
  };

  // --- ACTIONS: ATTENDANCE ---
  const handleUpdateKehadiran = async (idSiswa: string, field: "sakit" | "ijin" | "tanpaKet", val: number) => {
    try {
      const th = state.dataSekolah.thAjaran || "2025/2026";
      const sem = state.dataSekolah.semester || "1 (Ganjil)";
      const docId = `${idSiswa}_${th.replace(/\//g, "-")}_${sem}`;
      const current = state.kehadiran.find(k => k.idSiswa === idSiswa) || { sakit: 0, ijin: 0, tanpaKet: 0 };
      const nextVal = {
        idSiswa,
        sakit: field === "sakit" ? Math.max(0, val) : current.sakit,
        ijin: field === "ijin" ? Math.max(0, val) : current.ijin,
        tanpaKet: field === "tanpaKet" ? Math.max(0, val) : current.tanpaKet,
        thAjaran: th,
        semester: sem
      };
      await setDoc(doc(db, "kehadiran", docId), nextVal);
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, "kehadiran");
    }
  };

  // --- ACTIONS: EXPORT & IMPORT ---
  const handleExportData = () => {
    const blob = new Blob([JSON.stringify(state, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Backup-Data-Raport-PAUD-${state.dataSekolah.namaSekolah.replace(/\s+/g, "-")}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleImportData = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    if (currentUserProfile?.role !== "admin") {
      alert("Akses Ditolak: Hanya Admin yang dapat mengimpor data!");
      return;
    }

    const reader = new FileReader();
    reader.onload = async (event) => {
      try {
        const parsed = JSON.parse(event.target?.result as string);
        if (parsed.kelas && parsed.siswa && parsed.dataSekolah) {
          setDbLoading(true);
          await setDoc(doc(db, "dataSekolah", "default"), parsed.dataSekolah);
          
          for (const k of parsed.kelas) {
            await setDoc(doc(db, "kelas", k.id), k);
          }
          for (const s of parsed.siswa) {
            await setDoc(doc(db, "siswa", s.id), s);
          }
          if (parsed.tujuanPembelajaran) {
            for (const tp of parsed.tujuanPembelajaran) {
              await setDoc(doc(db, "tujuanPembelajaran", tp.id), tp);
            }
          }
          if (parsed.subdimensiKokurikuler) {
            for (const sub of parsed.subdimensiKokurikuler) {
              await setDoc(doc(db, "subdimensiKokurikuler", sub.id), sub);
            }
          }
          if (parsed.nilaiIntrakurikuler) {
            for (const ni of parsed.nilaiIntrakurikuler) {
              await setDoc(doc(db, "nilaiIntrakurikuler", `${ni.idSiswa}_${ni.idTp}`), ni);
            }
          }
          if (parsed.nilaiKokurikuler) {
            for (const nk of parsed.nilaiKokurikuler) {
              await setDoc(doc(db, "nilaiKokurikuler", `${nk.idSiswa}_${nk.idSubdimensi}`), nk);
            }
          }
          if (parsed.catatanAnak) {
            for (const c of parsed.catatanAnak) {
              await setDoc(doc(db, "catatanAnak", c.idSiswa), c);
            }
          }
          if (parsed.kehadiran) {
            for (const kh of parsed.kehadiran) {
              await setDoc(doc(db, "kehadiran", kh.idSiswa), kh);
            }
          }
          alert("Data sekolah berhasil dipulihkan (Import sukses)!");
        } else {
          alert("Format file tidak valid.");
        }
      } catch (err) {
        alert("Gagal membaca file JSON.");
      } finally {
        setDbLoading(false);
      }
    };
    reader.readAsText(file);
  };

  const handleAmbilSiswa = async (tahunSumber: string) => {
    if (!tahunSumber) {
      alert("Pilih tahun pelajaran sumber terlebih dahulu!");
      return;
    }
    const targetTh = state.dataSekolah.thAjaran || "2025/2026";
    const activeSem = state.dataSekolah.semester || "1 (Ganjil)";

    const oldKelasList = rawState.kelas.filter(k => k.thAjaran === tahunSumber);
    const oldSiswaList = rawState.siswa.filter(s => s.thAjaran === tahunSumber);

    if (oldSiswaList.length === 0) {
      alert(`Tidak ditemukan data siswa pada tahun pelajaran ${tahunSumber}.`);
      return;
    }

    try {
      setDbLoading(true);
      const classMap: Record<string, string> = {};
      for (const oldK of oldKelasList) {
        const existingK = state.kelas.find(k => k.namaKelas.toLowerCase() === oldK.namaKelas.toLowerCase());
        if (existingK) {
          classMap[oldK.id] = existingK.id;
        } else {
          const newKelasId = "K-" + Date.now() + Math.floor(Math.random() * 1000);
          const newK: Kelas = {
            id: newKelasId,
            namaKelas: oldK.namaKelas,
            waliKelas: oldK.waliKelas,
            nuptkNgty: oldK.nuptkNgty,
            thAjaran: targetTh
          };
          await setDoc(doc(db, "kelas", newKelasId), newK);
          classMap[oldK.id] = newKelasId;
        }
      }

      let countImport = 0;
      for (const oldS of oldSiswaList) {
        const targetClassId = classMap[oldS.idKelas] || (state.kelas[0]?.id || "");
        const exists = state.siswa.some(s => s.namaSiswa.toLowerCase() === oldS.namaSiswa.toLowerCase() && s.idKelas === targetClassId);
        if (exists) continue;

        const newSiswaId = "S-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
        const newS: Siswa = {
          ...oldS,
          id: newSiswaId,
          idKelas: targetClassId,
          thAjaran: targetTh
        };
        await setDoc(doc(db, "siswa", newSiswaId), newS);

        const attId = `${newSiswaId}_${targetTh.replace(/\//g, "-")}_${activeSem}`;
        await setDoc(doc(db, "kehadiran", attId), {
          idSiswa: newSiswaId,
          sakit: 0,
          ijin: 0,
          tanpaKet: 0,
          thAjaran: targetTh,
          semester: activeSem
        });

        countImport++;
      }

      showToast(`Berhasil menyalin ${countImport} data siswa dari tahun pelajaran ${tahunSumber}!`, "success");
    } catch (err) {
      console.error(err);
      alert("Gagal menyalin data siswa lama.");
    } finally {
      setDbLoading(false);
    }
  };

  const handleAmbilIntrakurikuler = async (tahunSumber: string) => {
    if (!tahunSumber) {
      alert("Pilih tahun pelajaran sumber terlebih dahulu!");
      return;
    }
    const targetTh = state.dataSekolah.thAjaran || "2025/2026";
    const oldTpList = rawState.tujuanPembelajaran.filter(tp => tp.thAjaran === tahunSumber);

    if (oldTpList.length === 0) {
      alert(`Tidak ditemukan data butir intrakurikuler (TP) pada tahun pelajaran ${tahunSumber}.`);
      return;
    }

    try {
      setDbLoading(true);
      let countImport = 0;

      for (const oldTp of oldTpList) {
        const oldClass = rawState.kelas.find(k => k.id === oldTp.idKelas);
        if (!oldClass) continue;

        const matchK = state.kelas.find(k => k.namaKelas.toLowerCase() === oldClass.namaKelas.toLowerCase());
        if (!matchK) continue;

        const exists = state.tujuanPembelajaran.some(tp => tp.idKelas === matchK.id && tp.deskripsi.toLowerCase() === oldTp.deskripsi.toLowerCase());
        if (exists) continue;

        const newTpId = "TP-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
        const newTp: TujuanPembelajaran = {
          ...oldTp,
          id: newTpId,
          idKelas: matchK.id,
          thAjaran: targetTh
        };
        await setDoc(doc(db, "tujuanPembelajaran", newTpId), newTp);
        countImport++;
      }

      showToast(`Berhasil menyalin ${countImport} butir intrakurikuler dari tahun pelajaran ${tahunSumber}!`, "success");
    } catch (err) {
      console.error(err);
      alert("Gagal menyalin data butir intrakurikuler.");
    } finally {
      setDbLoading(false);
    }
  };

  const handleAmbilKokurikuler = async (tahunSumber: string) => {
    if (!tahunSumber) {
      alert("Pilih tahun pelajaran sumber terlebih dahulu!");
      return;
    }
    const targetTh = state.dataSekolah.thAjaran || "2025/2026";
    const oldSubList = rawState.subdimensiKokurikuler.filter(sub => sub.thAjaran === tahunSumber);

    if (oldSubList.length === 0) {
      alert(`Tidak ditemukan data butir kokurikuler pada tahun pelajaran ${tahunSumber}.`);
      return;
    }

    try {
      setDbLoading(true);
      let countImport = 0;

      for (const oldSub of oldSubList) {
        const oldClass = rawState.kelas.find(k => k.id === oldSub.idKelas);
        if (!oldClass) continue;

        const matchK = state.kelas.find(k => k.namaKelas.toLowerCase() === oldClass.namaKelas.toLowerCase());
        if (!matchK) continue;

        const exists = state.subdimensiKokurikuler.some(sub => sub.idKelas === matchK.id && sub.namaSubdimensi.toLowerCase() === oldSub.namaSubdimensi.toLowerCase());
        if (exists) continue;

        const newSubId = "SD-" + Date.now() + "-" + Math.floor(Math.random() * 1000);
        const newSub: SubdimensiKokurikuler = {
          ...oldSub,
          id: newSubId,
          idKelas: matchK.id,
          thAjaran: targetTh
        };
        await setDoc(doc(db, "subdimensiKokurikuler", newSubId), newSub);
        countImport++;
      }

      showToast(`Berhasil menyalin ${countImport} butir kokurikuler dari tahun pelajaran ${tahunSumber}!`, "success");
    } catch (err) {
      console.error(err);
      alert("Gagal menyalin data butir kokurikuler.");
    } finally {
      setDbLoading(false);
    }
  };

  // --- FILTERED COMPUTATIONS ---
  const filteredSiswa = state.siswa.filter(s => {
    const matchClass = selectedKelasFilter ? s.idKelas === selectedKelasFilter : true;
    const matchSearch = s.namaSiswa.toLowerCase().includes(searchSiswa.toLowerCase()) || 
                        s.nisn.includes(searchSiswa) || 
                        s.namaAyah.toLowerCase().includes(searchSiswa.toLowerCase());
    return matchClass && matchSearch;
  });

  const getSiswaClassLabel = (idKelas?: string) => {
    if (!idKelas) return "Belum Masuk Kelas";
    const found = state.kelas.find(k => k.id === idKelas);
    return found ? found.namaKelas : "Kelas Tidak valid";
  };

  // Report Card Prep
  const printSiswa = state.siswa.find(s => s.id === printSiswaId);
  const printKelasItem = state.kelas.find(k => k.id === printSiswa?.idKelas);
  const printCatatanSiswa = state.catatanAnak.find(c => c.idSiswa === printSiswaId)?.catatan || "Ananda aktif bersosialisasi dan mengikuti instruksi dengan baik.";
  const printAbsensi = state.kehadiran.find(k => k.idSiswa === printSiswaId) || { sakit: 0, ijin: 0, tanpaKet: 0 };

  if (!currentUser) {
    return (
      <div className="flex flex-col items-center justify-center min-h-screen bg-slate-50 p-4">
        <div className="w-full max-w-md bg-white rounded-2xl border border-slate-200 shadow-xl overflow-hidden p-8 space-y-6 animate-fade-in">
          {state.dataSekolah?.logoLogin && (
            <div className="flex justify-center -mb-2">
              <img src={state.dataSekolah.logoLogin} alt="Logo Login" className="max-h-20 max-w-[200px] object-contain" />
            </div>
          )}
          <div className="text-center space-y-2">
            <h1 className="text-2xl font-bold text-slate-800 font-display">Aplikasi Raport PAUD</h1>
            <p className="text-sm text-neutral-500">Pendidikan Anak Usia Dini - Kurikulum Merdeka</p>
          </div>

          {loginError && (
            <div className="p-3 bg-red-50 border border-red-200 text-red-700 text-xs rounded-xl flex items-center gap-2">
              <ShieldAlert className="w-4 h-4 flex-shrink-0" />
              <span>{loginError}</span>
            </div>
          )}

          <form onSubmit={handleUsernameLogin} className="space-y-4">
            <div className="space-y-1">
              <label className="block text-xs font-bold text-slate-600 uppercase">Nama Pengguna (Username)</label>
              <input
                type="text"
                required
                value={loginUsername}
                onChange={(e) => setLoginUsername(e.target.value)}
                placeholder="Contoh: admin atau nama_guru"
                className="w-full text-sm p-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50"
                disabled={loginLoading}
              />
            </div>

            <div className="space-y-1">
              <label className="block text-xs font-bold text-slate-600 uppercase">Kata Sandi (Password)</label>
              <input
                type="password"
                required
                value={loginPassword}
                onChange={(e) => setLoginPassword(e.target.value)}
                placeholder="Masukkan kata sandi Anda"
                className="w-full text-sm p-3 border border-slate-200 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50"
                disabled={loginLoading}
              />
            </div>

            <button
              type="submit"
              disabled={loginLoading}
              className="w-full py-3 px-4 bg-indigo-600 hover:bg-indigo-700 disabled:bg-indigo-400 text-white font-bold rounded-xl text-sm flex items-center justify-center gap-2 shadow transition-all cursor-pointer"
            >
              {loginLoading ? (
                <>
                  <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                  Memvalidasi...
                </>
              ) : (
                <>
                  <CheckCircle className="w-4 h-4" /> Masuk ke Laporan
                </>
              )}
            </button>
          </form>
        </div>
      </div>
    );
  }

  if (dbLoading) {
    const loaderLogo = state.dataSekolah?.logo || state.dataSekolah?.logoLogin || state.dataSekolah?.logoSidebar;
    return (
      <div className="flex items-center justify-center min-h-screen bg-slate-50">
        <div className="text-center space-y-4">
          {loaderLogo ? (
            <div className="flex justify-center mb-1">
              <img src={loaderLogo} alt="Logo Loading" className="w-20 h-20 object-contain rounded-xl p-1 bg-white shadow-md border border-slate-100 animate-bounce" />
            </div>
          ) : (
            <div className="w-12 h-12 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mx-auto"></div>
          )}
          {loaderLogo && (
            <div className="w-6 h-6 border-2 border-indigo-600 border-t-transparent rounded-full animate-spin mx-auto"></div>
          )}
          <div>
            <p className="text-slate-850 font-bold font-display">{state.dataSekolah?.namaSekolah || "Menghubungkan ke Cloud Firestore..."}</p>
            <p className="text-xs text-neutral-400 font-medium">Sinkronisasi data raport real-time sedang berjalan.</p>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="flex flex-col min-h-screen bg-slate-50 md:flex-row antialiased relative">
      <style dangerouslySetInnerHTML={{ __html: `
        @media print {
          @page {
            size: A4 portrait;
            margin: 12mm !important;
          }
          body {
            font-family: Arial, sans-serif !important;
            margin: 0 !important;
            padding: 0 !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          .print-container {
            width: 100% !important;
            max-width: none !important;
            margin: 0 auto !important;
            display: block !important;
            border: none !important;
            box-shadow: none !important;
            position: relative !important;
            overflow: visible !important;
          }
          .print-arial-large {
            font-family: Arial, "Helvetica Neue", Helvetica, sans-serif !important;
          }
          .print-arial-large table {
            font-size: 11pt !important;
            font-family: Arial, sans-serif !important;
          }
          .print-arial-large th, .print-arial-large td {
            font-size: 10pt !important;
            padding: 6px 8px !important;
            font-family: Arial, sans-serif !important;
          }
          .print-arial-large p, .print-arial-large italic {
            font-size: 11pt !important;
            font-family: Arial, sans-serif !important;
          }
          .print-arial-large h3, .print-arial-large h4 {
            font-size: 13pt !important;
            font-family: Arial, sans-serif !important;
          }
          .keep-together {
            break-inside: avoid !important;
            page-break-inside: avoid !important;
          }
          .no-print {
            display: none !important;
          }
        }
      `}} />
      
      {/* MOBILE BACKDROP DRAWER EFFECT */}
      {isSidebarOpenOnMobile && (
        <div 
          className="fixed inset-0 bg-slate-950/60 backdrop-blur-xs z-40 md:hidden" 
          onClick={() => setIsSidebarOpenOnMobile(false)}
        />
      )}

      {/* SIDEBAR NAVIGATION - Hidden during printing */}
      <aside className={`fixed md:sticky md:top-0 md:h-screen w-64 bg-indigo-900 text-indigo-100 flex-shrink-0 flex flex-col no-print z-50 md:z-10 border-r border-indigo-950 shadow-md transition-transform duration-300 ${
        isSidebarOpenOnMobile ? "translate-x-0" : "-translate-x-full md:translate-x-0"
      }`}>
        <div className="p-5 border-b border-indigo-800/60 bg-indigo-950/40 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-white/20 flex items-center justify-center text-white border border-white/25 font-bold shadow-md overflow-hidden">
              {state.dataSekolah?.logoSidebar ? (
                <img src={state.dataSekolah.logoSidebar} alt="Logo" className="w-full h-full object-cover" />
              ) : (
                "📝"
              )}
            </div>
          <div className="px-3 py-1 bg-white/10 rounded-lg backdrop-blur-sm border border-white/20">
            <h1 className="font-bold text-lg tracking-tight leading-tight text-white font-display">Raport PAUD</h1>
            <p className="text-sm text-indigo-200/90 font-medium tracking-wide">Pendidikan Anak Usia Dini</p>
          </div>
          </div>
          <button
            type="button"
            onClick={() => setIsSidebarOpenOnMobile(false)}
            className="md:hidden p-1.5 text-indigo-200 hover:text-white hover:bg-white/10 rounded-lg transition-colors cursor-pointer"
            title="Tutup menu"
            aria-label="Tutup menu"
          >
            <X className="w-5 h-5" />
          </button>
        </div>
        
        <div className="px-5 py-3 mt-4 text-[14px] bg-indigo-950/40 mx-4 rounded-xl text-indigo-100 border border-white/5">
          <span className="font-bold block truncate text-base">{state.dataSekolah.namaSekolah || "Nama Sekolah"}</span>
          <div className="opacity-90 font-semibold mt-0.5">Tahun Pelajaran {state.dataSekolah.thAjaran}</div>
        </div>

        <nav className="flex-1 px-3 py-4 space-y-1 overflow-y-auto">
          <button
            onClick={() => { setActiveTab("dashboard"); setIsSidebarOpenOnMobile(false); }}
            className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-lg text-base font-medium transition-all ${
              activeTab === "dashboard" ? "bg-white/10 text-white shadow-inner font-semibold" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
            }`}
          >
            <Home className="w-4 h-4" /> Dashboard
          </button>

          {/* COLLAPSIBLE MASTER DATA ACCORDION */}
          <div className="space-y-1">
            <button
              type="button"
              onClick={() => setIsMasterDataExpanded(!isMasterDataExpanded)}
              className={`w-full flex items-center justify-between px-4 py-3 rounded-xl text-sm font-black uppercase tracking-widest hover:bg-white/5 transition-all text-left cursor-pointer font-sans ${
                isMasterDataExpanded ? "text-white" : "text-indigo-200/80"
              }`}
            >
              <span>Master Data</span>
              <ChevronDown className={`w-3.5 h-3.5 text-indigo-300 transition-transform duration-200 ${isMasterDataExpanded ? "rotate-180" : ""}`} />
            </button>
            
            {isMasterDataExpanded && (
              <div className="pl-2 space-y-1 border-l border-indigo-800/40 ml-4 animate-fadeIn">
                {currentUserProfile?.role === 'admin' && (
                  <button
                    onClick={() => { setActiveTab("kelas"); setIsSidebarOpenOnMobile(false); }}
                    className={`w-full flex items-center gap-3 px-4 py-2 rounded-lg text-xs md:text-sm font-medium transition-all cursor-pointer ${
                      activeTab === "kelas" ? "bg-white/10 text-white font-semibold shadow-inner" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
                    }`}
                  >
                    <Home className="w-4 h-4 rotate-45" /> Data Kelas
                  </button>
                )}
                
                <button
                  onClick={() => { setActiveTab("siswa"); setIsSidebarOpenOnMobile(false); }}
                  className={`w-full flex items-center gap-3 px-4 py-2 rounded-lg text-xs md:text-sm font-medium transition-all cursor-pointer ${
                    activeTab === "siswa" ? "bg-white/10 text-white font-semibold shadow-inner" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
                  }`}
                >
                  <Users className="w-4 h-4" /> Data Siswa
                </button>

                {currentUserProfile?.role === 'admin' && (
                  <>
                    <button
                      onClick={() => { setActiveTab("intra"); setIsSidebarOpenOnMobile(false); }}
                      className={`w-full flex items-center gap-3 px-4 py-2 rounded-lg text-xs md:text-sm font-medium transition-all cursor-pointer ${
                        activeTab === "intra" ? "bg-white/10 text-white font-semibold shadow-inner" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
                      }`}
                    >
                      <BookOpen className="w-4 h-4" /> Intrakurikuler
                    </button>

                    <button
                      onClick={() => { setActiveTab("kokuri"); setIsSidebarOpenOnMobile(false); }}
                      className={`w-full flex items-center gap-3 px-4 py-2 rounded-lg text-xs md:text-sm font-medium transition-all cursor-pointer ${
                        activeTab === "kokuri" ? "bg-white/10 text-white font-semibold shadow-inner" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
                      }`}
                    >
                      <Award className="w-4 h-4" /> Kokurikuler
                    </button>
                  </>
                )}
              </div>
            )}
          </div>

          {/* COLLAPSIBLE PENILAIAN ACCORDION */}
          <div className="space-y-1">
            <button
              type="button"
              onClick={() => setIsPenilaianExpanded(!isPenilaianExpanded)}
              className={`w-full flex items-center justify-between px-4 py-3 rounded-xl text-sm font-black uppercase tracking-widest hover:bg-white/5 transition-all text-left cursor-pointer font-sans ${
                isPenilaianExpanded ? "text-white" : "text-indigo-200/80"
              }`}
            >
              <span>Penilaian</span>
              <ChevronDown className={`w-3.5 h-3.5 text-indigo-300 transition-transform duration-200 ${isPenilaianExpanded ? "rotate-180" : ""}`} />
            </button>

            {isPenilaianExpanded && (
              <div className="pl-2 space-y-1 border-l border-indigo-800/40 ml-4 animate-fadeIn">
            <button
              onClick={() => { setActiveTab("nilai"); setNilaiSubTab("intra"); setIsSidebarOpenOnMobile(false); }}
              className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-lg text-sm font-medium transition-all ${
                activeTab === "nilai" && nilaiSubTab === "intra" ? "bg-white/10 text-white shadow-inner font-semibold" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
              }`}
            >
              <BookOpen className="w-4 h-4" /> Nilai Intrakurikuler
            </button>
            <button
              onClick={() => { setActiveTab("nilai"); setNilaiSubTab("kokuri"); setIsSidebarOpenOnMobile(false); }}
              className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-lg text-sm font-medium transition-all ${
                activeTab === "nilai" && nilaiSubTab === "kokuri" ? "bg-white/10 text-white shadow-inner font-semibold" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
              }`}
            >
              <Award className="w-4 h-4" /> Nilai Kokurikuler
            </button>
            <button
              onClick={() => { setActiveTab("nilai"); setNilaiSubTab("catatan"); setIsSidebarOpenOnMobile(false); }}
              className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-lg text-sm font-medium transition-all ${
                activeTab === "nilai" && nilaiSubTab === "catatan" ? "bg-white/10 text-white shadow-inner font-semibold" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
              }`}
            >
              <Sparkles className="w-4 h-4" /> Catatan Siswa (AI)
            </button>
            <button
              onClick={() => { setActiveTab("nilai"); setNilaiSubTab("kehadiran"); setIsSidebarOpenOnMobile(false); }}
              className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-lg text-sm font-medium transition-all ${
                activeTab === "nilai" && nilaiSubTab === "kehadiran" ? "bg-white/10 text-white shadow-inner font-semibold" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
              }`}
            >
              <UserCheck className="w-4 h-4" /> Kehadiran
            </button>
              </div>
            )}
          </div>

          {/* CORE FLAT MENUS FOR MAIN ACTIONS */}
          {currentUserProfile?.role === 'admin' && (
            <>
              <button
                onClick={() => { 
                  setActiveTab("cetak"); 
                  setIsSidebarOpenOnMobile(false);
                  // pre-select first student if available
                  if (state.siswa.length > 0 && !printSiswaId) {
                    setPrintSiswaId(state.siswa[0].id);
                  }
                }}
                className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-lg text-sm font-medium transition-all ${
                  activeTab === "cetak" ? "bg-white/10 text-white shadow-inner font-semibold" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
                }`}
              >
                <Printer className="w-4 h-4" /> Cetak Raport PAUD
              </button>

              <button
                onClick={() => { setActiveTab("pengaturan"); setIsSidebarOpenOnMobile(false); }}
                className={`w-full flex items-center gap-3 px-4 py-2.5 rounded-lg text-sm font-medium transition-all ${
                  activeTab === "pengaturan" ? "bg-white/10 text-white shadow-inner font-semibold" : "text-indigo-100/70 hover:bg-white/5 hover:text-white"
                }`}
              >
                <Settings className="w-4 h-4" /> Pengaturan Sekolah
              </button>
            </>
          )}
        </nav>

        {/* ACTIVE USER PROFILE & SIGN OUT */}
        <div className="p-4 mx-4 mb-4 bg-indigo-950 rounded-2xl text-xs text-indigo-300 font-medium border border-indigo-800/40 flex-shrink-0">
          <div className="text-indigo-400 text-[10px] font-bold uppercase mb-1">Pengguna Masuk</div>
          <div className="text-white text-xs font-bold truncate leading-tight">{currentUserProfile?.nama || currentUser?.displayName}</div>
          <div className="opacity-80 text-[11px] font-mono mt-0.5 text-indigo-200">
            Peran: <span className={`px-1 rounded-sm text-[10px] font-bold ${
              currentUserProfile?.role === "admin" ? "bg-purple-900 border border-purple-800 text-purple-200" : "bg-teal-900 border border-teal-800 text-teal-200"
            }`}>{currentUserProfile?.role?.toUpperCase()}</span>
          </div>
          <button
            onClick={async () => {
              try {
                await signOut(auth);
                setIsSidebarOpenOnMobile(false);
              } catch (signOutErr) {
                console.error("Gagal keluar:", signOutErr);
              }
            }}
            className="mt-3 w-full bg-red-600/30 hover:bg-red-600 border border-red-500/20 text-red-200 hover:text-white py-1.5 px-3 rounded-lg text-[10px] font-bold block text-center transition cursor-pointer"
          >
            Keluar
          </button>
        </div>
      </aside>

      {/* CORE WORKSPACE */}
      <main className="flex-1 flex flex-col min-w-0 print:p-0 print:bg-white">
        
        {/* TOP BAR - Hidden during printing */}
        <header className="bg-white border-b border-neutral-200/80 h-16 flex items-center justify-between px-4 md:px-6 no-print flex-shrink-0">
          <div className="flex items-center gap-3">
            <button
              type="button"
              onClick={() => setIsSidebarOpenOnMobile(true)}
              className="md:hidden p-2 -ml-1 text-slate-500 hover:text-slate-800 hover:bg-slate-100 rounded-lg transition-colors cursor-pointer"
              title="Buka menu"
              aria-label="Buka menu"
            >
              <Menu className="w-5 h-5" />
            </button>
            <div className="flex items-center gap-2">
              <span className="text-neutral-400 text-sm hidden sm:inline">Dashboard</span>
              <ChevronRight className="w-3.5 h-3.5 text-neutral-400 hidden sm:inline" />
              <span className="text-neutral-700 font-medium capitalize text-sm">{activeTab.replace("-", " ")}</span>
            </div>
          </div>

          <div className="flex items-center gap-4">
            <div className="text-right">
              <div className="text-base font-bold text-neutral-800">{currentUserProfile?.nama || currentUser?.displayName}</div>
              <div className="text-[13px] text-neutral-500 font-medium">{state.dataSekolah.namaSekolah || "TK Bunga Bangsa"}</div>
            </div>
          </div>
        </header>

        {/* WORKSPACE AREA CONTAINER */}
        <div className="flex-1 p-4 sm:p-6 overflow-y-auto print:p-0 print:overflow-visible">
          
          {/* TAB A: DASHBOARD */}
          {activeTab === "dashboard" && (
            <div className="space-y-6 animate-fade-in no-print font-sans">
              
              {/* WELCOME BANNER - DASHBOARD INFO */}
              <div className="bg-gradient-to-br from-indigo-600 to-indigo-800 rounded-2xl p-6 text-white shadow-lg relative overflow-hidden">
                <div className="absolute right-0 bottom-0 top-0 w-1/3 opacity-10 bg-[radial-gradient(ellipse_at_center,_var(--tw-gradient-stops))] from-white via-indigo-900 to-indigo-950"></div>
                <div className="relative z-10 font-sans flex flex-col md:flex-row md:items-center justify-center gap-10">
                  <div className="flex items-center gap-5">
                    <div className="w-16 h-16 rounded-2xl bg-white/20 backdrop-blur-md flex items-center justify-center shadow-lg border border-white/30">
                      <Calendar className="w-8 h-8 text-white" />
                    </div>
                    <div>
                      <span className="text-indigo-100/80 text-xs font-bold uppercase tracking-wider block mb-1">Tahun Pelajaran</span>
                      <h2 className="text-2xl font-bold font-display tracking-tight text-white leading-none">
                        {state.dataSekolah.thAjaran || "2024/2025"}
                      </h2>
                    </div>
                  </div>

                  <div className="hidden md:block w-px h-12 bg-white/20"></div>

                  <div className="flex items-center gap-5">
                    <div className="w-16 h-16 rounded-2xl bg-white/20 backdrop-blur-md flex items-center justify-center shadow-lg border border-white/30">
                      <BookOpen className="w-8 h-8 text-white" />
                    </div>
                    <div>
                      <span className="text-indigo-100/80 text-xs font-bold uppercase tracking-wider block mb-1">Semester Aktif</span>
                      <h2 className="text-2xl font-bold font-display tracking-tight text-white leading-none">
                        {state.dataSekolah.semester || "I (Ganjil)"}
                      </h2>
                    </div>
                  </div>
                </div>
              </div>


              {/* STATS COUNT */}
              {(() => {
                const activeKelasId = selectedKelasFilter || state.kelas[0]?.id || "";
                const classStudents = state.siswa.filter(s => s.idKelas === activeKelasId);
                const totalStudents = classStudents.length;

                // Status computed helper functions
                const getIntraStatus = (studentId: string, idKelas: string) => {
                  const classTps = state.tujuanPembelajaran.filter(tp => tp.idKelas === idKelas);
                  const classTpsCount = classTps.length;
                  if (classTpsCount === 0) return { label: "Belum Dinilai", code: "belum", evaluated: 0, total: 0 };
                  
                  const evaluatedCount = classTps.filter(tp => 
                    state.nilaiIntrakurikuler.some(n => n.idSiswa === studentId && n.idTp === tp.id && n.nilai)
                  ).length;

                  if (evaluatedCount === 0) {
                    return { label: "Belum Dinilai", code: "belum", evaluated: 0, total: classTpsCount };
                  } else if (evaluatedCount === classTpsCount) {
                    return { label: `Lengkap (${evaluatedCount}/${classTpsCount})`, code: "lengkap", evaluated: evaluatedCount, total: classTpsCount };
                  } else {
                    return { label: `Kurang (${evaluatedCount}/${classTpsCount})`, code: "kurang", evaluated: evaluatedCount, total: classTpsCount };
                  }
                };

                const getKokuriStatus = (studentId: string, idKelas: string) => {
                  const classSubs = state.subdimensiKokurikuler.filter(sub => sub.idKelas === idKelas);
                  const classSubsCount = classSubs.length;
                  if (classSubsCount === 0) return { label: "Belum Dinilai", code: "belum", evaluated: 0, total: 0 };

                  const evaluatedCount = classSubs.filter(sub => 
                    state.nilaiKokurikuler.some(n => n.idSiswa === studentId && n.idSubdimensi === sub.id && n.nilai)
                  ).length;

                  if (evaluatedCount === 0) {
                    return { label: "Belum Dinilai", code: "belum", evaluated: 0, total: classSubsCount };
                  } else if (evaluatedCount === classSubsCount) {
                    return { label: `Lengkap (${evaluatedCount}/${classSubsCount})`, code: "lengkap", evaluated: evaluatedCount, total: classSubsCount };
                  } else {
                    return { label: `Kurang (${evaluatedCount}/${classSubsCount})`, code: "kurang", evaluated: evaluatedCount, total: classSubsCount };
                  }
                };

                const intraStats = classStudents.map(s => getIntraStatus(s.id, activeKelasId));
                const intraLengkap = intraStats.filter(st => st.code === "lengkap").length;
                const intraKurang = intraStats.filter(st => st.code === "kurang").length;
                const intraBelum = intraStats.filter(st => st.code === "belum").length;

                const kokuriStats = classStudents.map(s => getKokuriStatus(s.id, activeKelasId));
                const kokuriLengkap = kokuriStats.filter(st => st.code === "lengkap").length;
                const kokuriKurang = kokuriStats.filter(st => st.code === "kurang").length;
                const kokuriBelum = kokuriStats.filter(st => st.code === "belum").length;

                return (
                  <div className="space-y-6">
                    {/* STATS BENTO ROW */}
                    <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
                      <div className="bg-white p-5 rounded-2xl border border-neutral-200 shadow-sm flex items-center gap-4 transition hover:shadow-md">
                        <div className="w-12 h-12 rounded-xl bg-blue-500/10 text-blue-600 flex items-center justify-center">
                          <GraduationCap className="w-6 h-6" />
                        </div>
                        <div>
                          <div className="text-neutral-500 text-xs font-black uppercase tracking-widest">Total Siswa</div>
                          <div className="text-2xl font-bold font-display text-slate-950 leading-none mt-1">{state.siswa.length}</div>
                        </div>
                      </div>
                      
                      <div className="bg-white p-5 rounded-2xl border border-neutral-200 shadow-sm flex items-center gap-4 transition hover:shadow-md">
                        <div className="w-12 h-12 rounded-xl bg-emerald-500/10 text-emerald-600 flex items-center justify-center">
                          <Home className="w-6 h-6 rotate-45" />
                        </div>
                        <div>
                          <div className="text-neutral-500 text-xs font-black uppercase tracking-widest">Total Kelas</div>
                          <div className="text-2xl font-bold font-display text-slate-950 leading-none mt-1">{state.kelas.length}</div>
                        </div>
                      </div>

                      <div className="bg-white p-5 rounded-2xl border border-neutral-200 shadow-sm flex items-center gap-4 transition hover:shadow-md">
                        <div className="w-12 h-12 rounded-xl bg-orange-500/10 text-orange-600 flex items-center justify-center">
                          <BookOpen className="w-6 h-6" />
                        </div>
                        <div>
                          <div className="text-neutral-500 text-xs font-black uppercase tracking-widest">Total TP</div>
                          <div className="text-2xl font-bold font-display text-slate-950 leading-none mt-1">{state.tujuanPembelajaran.length}</div>
                        </div>
                      </div>

                      <div className="bg-white p-5 rounded-2xl border border-neutral-200 shadow-sm flex items-center gap-4 transition hover:shadow-md">
                        <div className="w-12 h-12 rounded-xl bg-purple-500/10 text-purple-600 flex items-center justify-center">
                          <Award className="w-6 h-6" />
                        </div>
                        <div>
                          <div className="text-neutral-500 text-xs font-black uppercase tracking-widest">Indikator Kokuri</div>
                          <div className="text-2xl font-bold font-display text-slate-950 leading-none mt-1">{state.subdimensiKokurikuler.length}</div>
                        </div>
                      </div>
                    </div>

                    {/* MONITORING HEADER SECTION */}
                    <div className="flex flex-col sm:flex-row sm:items-center sm:justify-between gap-4 mt-8 border-t border-slate-100 pt-6">
                      <div>
                        <h2 className="text-xl font-bold font-display text-slate-900 leading-tight">Monitoring Penilaian</h2>
                        <p className="text-sm font-medium text-neutral-500 mt-0.5">Pantau kelengkapan nilai Intrakurikuler dan Kokurikuler siswa secara real-time.</p>
                      </div>
                      {state.kelas.length > 0 && (
                        <div className="relative">
                          <select
                            value={activeKelasId}
                            onChange={(e) => setSelectedKelasFilter(e.target.value)}
                            className="appearance-none font-sans font-bold text-sm text-slate-750 bg-white border border-slate-300 pl-8 pr-12 py-2.5 rounded-xl hover:bg-slate-50 transition cursor-pointer focus:outline-none focus:ring-2 focus:ring-indigo-500 shadow-md"
                          >
                            {state.kelas.map(k => (
                              <option key={k.id} value={k.id}>{k.namaKelas}</option>
                            ))}
                          </select>
                          <Search className="absolute left-2.5 top-2.5 w-3.5 h-3.5 text-slate-400 pointer-events-none" />
                          <ChevronDown className="absolute right-2.5 top-2.5 w-3.5 h-3.5 text-slate-400 pointer-events-none" />
                        </div>
                      )}
                    </div>

                    {/* DONUT CHARTS PANEL */}
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      {/* DONUT A: INTRAKURIKULER */}
                      <div className="bg-white rounded-2xl border border-neutral-200 p-6 shadow-sm flex flex-col justify-between">
                        <div>
                          <div className="flex items-center gap-2 mb-4">
                            <BookOpen className="w-4 h-4 text-orange-500" />
                            <h3 className="font-black text-slate-900 text-base tracking-tight font-display mb-1">Status Intrakurikuler</h3>
                            <p className="text-xs text-slate-500 font-medium tracking-tight mb-4 whitespace-nowrap">Dashboard perolehan nilai berdasarkan Kelompok Kelas aktif.</p>
                          </div>
                          
                          <div className="flex flex-col sm:flex-row items-center gap-4 py-2">
                            {/* Donut graphic */}
                            <div className="relative w-40 h-40 flex-shrink-0">
                              <ResponsiveContainer width="100%" height="100%">
                                <PieChart>
                                  <Pie
                                    data={[
                                      { name: 'Lengkap', value: intraLengkap, color: '#10b981' },
                                      { name: 'Kurang', value: intraKurang, color: '#f59e0b' },
                                      { name: 'Belum Dinilai', value: intraBelum, color: '#ef4444' }
                                    ]}
                                    dataKey="value"
                                    cx="50%"
                                    cy="50%"
                                    innerRadius={48}
                                    outerRadius={68}
                                    paddingAngle={totalStudents > 0 ? 3 : 0}
                                  >
                                    {[
                                      { color: '#10b981' },
                                      { color: '#f59e0b' },
                                      { color: '#ef4444' }
                                    ].map((entry, idx) => (
                                      <Cell key={`cell-i-${idx}`} fill={entry.color} />
                                    ))}
                                  </Pie>
                                  <Tooltip formatter={(value) => [`${value} Siswa`]} />
                                </PieChart>
                              </ResponsiveContainer>
                              <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                                <span className="text-xl font-black text-slate-800 leading-none">{totalStudents}</span>
                                <span className="text-[9px] tracking-widest font-black text-slate-400 mt-1 uppercase">SISWA</span>
                              </div>
                            </div>

                            {/* Legend detail list */}
                            <div className="flex-1 space-y-2.5 w-full">
                              {[
                                { name: 'Lengkap', val: intraLengkap, color: 'bg-emerald-500' },
                                { name: 'Kurang', val: intraKurang, color: 'bg-amber-500' },
                                { name: 'Belum Dinilai', val: intraBelum, color: 'bg-red-500' }
                              ].map((item, idx) => {
                                const pct = totalStudents > 0 ? Math.round((item.val / totalStudents) * 100) : 0;
                                return (
                                  <div key={`leg-i-${idx}`} className="flex justify-between items-center text-xs text-slate-600 font-medium">
                                    <div className="flex items-center gap-2">
                                      <span className={`w-3 h-3 rounded-full ${item.color} flex-shrink-0`} />
                                      <span>{item.name}</span>
                                    </div>
                                    <span className="font-black text-slate-800">
                                      {item.val} <span className="text-neutral-400 font-normal font-mono">({pct}%)</span>
                                    </span>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* DONUT B: KOKURIKULER */}
                      <div className="bg-white rounded-2xl border border-neutral-200 p-6 shadow-sm flex flex-col justify-between">
                        <div>
                          <div className="flex items-center gap-2 mb-4">
                            <Award className="w-4 h-4 text-purple-500" />
                            <h3 className="font-black text-slate-900 text-base tracking-tight font-display mb-1">Status Kokurikuler</h3>
                            <p className="text-xs text-slate-500 font-medium tracking-tight mb-4 whitespace-nowrap">Dashboard perolehan nilai untuk indikator Kokurikuler aktif.</p>
                          </div>
                          
                          <div className="flex flex-col sm:flex-row items-center gap-4 py-2">
                            {/* Donut graphic */}
                            <div className="relative w-40 h-40 flex-shrink-0">
                              <ResponsiveContainer width="100%" height="100%">
                                <PieChart>
                                  <Pie
                                    data={[
                                      { name: 'Lengkap', value: kokuriLengkap, color: '#10b981' },
                                      { name: 'Kurang', value: kokuriKurang, color: '#f59e0b' },
                                      { name: 'Belum Dinilai', value: kokuriBelum, color: '#ef4444' }
                                    ]}
                                    dataKey="value"
                                    cx="50%"
                                    cy="50%"
                                    innerRadius={48}
                                    outerRadius={68}
                                    paddingAngle={totalStudents > 0 ? 3 : 0}
                                  >
                                    {[
                                      { color: '#10b981' },
                                      { color: '#f59e0b' },
                                      { color: '#ef4444' }
                                    ].map((entry, idx) => (
                                      <Cell key={`cell-k-${idx}`} fill={entry.color} />
                                    ))}
                                  </Pie>
                                  <Tooltip formatter={(value) => [`${value} Siswa`]} />
                                </PieChart>
                              </ResponsiveContainer>
                              <div className="absolute inset-0 flex flex-col items-center justify-center pointer-events-none">
                                <span className="text-xl font-black text-slate-800 leading-none">{totalStudents}</span>
                                <span className="text-[9px] tracking-widest font-black text-slate-400 mt-1 uppercase">SISWA</span>
                              </div>
                            </div>

                            {/* Legend detail list */}
                            <div className="flex-1 space-y-2.5 w-full">
                              {[
                                { name: 'Lengkap', val: kokuriLengkap, color: 'bg-emerald-500' },
                                { name: 'Kurang', val: kokuriKurang, color: 'bg-amber-500' },
                                { name: 'Belum Dinilai', val: kokuriBelum, color: 'bg-red-500' }
                              ].map((item, idx) => {
                                const pct = totalStudents > 0 ? Math.round((item.val / totalStudents) * 100) : 0;
                                return (
                                  <div key={`leg-k-${idx}`} className="flex justify-between items-center text-xs text-slate-600 font-medium">
                                    <div className="flex items-center gap-2">
                                      <span className={`w-3 h-3 rounded-full ${item.color} flex-shrink-0`} />
                                      <span>{item.name}</span>
                                    </div>
                                    <span className="font-black text-slate-800">
                                      {item.val} <span className="text-neutral-400 font-normal font-mono">({pct}%)</span>
                                    </span>
                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>

                    {/* STUDENT MONITORING TABLE CARD */}
                    <div className="bg-white rounded-2xl border border-neutral-200 shadow-sm overflow-hidden mt-6">
                      <div className="p-6 border-b border-neutral-100 flex justify-between items-center bg-slate-50/70">
                        <div className="flex items-center gap-3">
                          <Users className="w-5 h-5 text-indigo-600" />
                          <h3 className="font-bold text-slate-900 text-sm font-display tracking-widest uppercase">Rangkuman Aktivitas Kelas</h3>
                        </div>
                        <span className="bg-indigo-600 text-white font-black px-4 py-1.5 rounded-full text-xs tracking-widest border shadow-sm">
                          {totalStudents} SISWA TERDAFTAR
                        </span>
                      </div>
                      
                      <div className="overflow-x-auto">
                        <table className="w-full text-left text-sm border-collapse">
                          <thead>
                            <tr className="bg-slate-50/75 border-b border-slate-200 text-xs font-black text-slate-500 uppercase tracking-widest">
                              <th className="py-4 px-4 w-12 text-center text-slate-400">NO</th>
                              <th className="py-4 px-4">IDENTITAS SISWA</th>
                              <th className="py-4 px-4">KELAS</th>
                              <th className="py-4 px-4">PENILAIAN INTRAKURIKULER</th>
                              <th className="py-4 px-4">PENILAIAN KOKURIKULER</th>
                              <th className="py-4 px-4 text-center">AKSI CEPAT</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 bg-white">
                            {classStudents.map((s, idx) => {
                              const intraStatusObj = getIntraStatus(s.id, activeKelasId);
                              const kokuriStatusObj = getKokuriStatus(s.id, activeKelasId);

                              // Pill styles mapping
                              let intraPillStyle = "text-rose-600 bg-rose-50 border-rose-150";
                              if (intraStatusObj.code === "lengkap") intraPillStyle = "text-emerald-700 bg-emerald-50 border-emerald-150";
                              if (intraStatusObj.code === "kurang") intraPillStyle = "text-amber-700 bg-amber-50 border-amber-150";

                              let kokuriPillStyle = "text-rose-600 bg-rose-50 border-rose-150";
                              if (kokuriStatusObj.code === "lengkap") kokuriPillStyle = "text-emerald-700 bg-emerald-50 border-emerald-150";
                              if (kokuriStatusObj.code === "kurang") kokuriPillStyle = "text-amber-700 bg-amber-50 border-amber-150";

                              return (
                                <tr key={s.id} className="hover:bg-slate-50/50 transition duration-150">
                                  <td className="py-3 px-4 text-center font-bold text-slate-400 font-mono">{idx + 1}</td>
                                  <td className="py-4 px-4">
                                    <span className="font-extrabold text-slate-900 uppercase tracking-wide block text-base">{s.namaSiswa}</span>
                                    <span className="text-xs text-slate-500 font-bold font-mono">NISN: {s.nisn || "-"}</span>
                                  </td>
                                  <td className="py-4 px-4">
                                    <span className="font-black text-slate-700 bg-slate-100 px-3 py-1 rounded-lg border border-slate-200 uppercase text-xs">
                                      {state.kelas.find(k => k.id === s.idKelas)?.namaKelas || "PAUD"}
                                    </span>
                                  </td>
                                  <td className="py-4 px-4">
                                    <span className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full border text-xs font-black tracking-widest shadow-sm ${intraPillStyle}`}>
                                      {intraStatusObj.code === "belum" ? "○" : intraStatusObj.code === "lengkap" ? "✓" : "⚠"} {intraStatusObj.label.toUpperCase()}
                                    </span>
                                  </td>
                                  <td className="py-4 px-4">
                                    <span className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full border text-xs font-black tracking-widest shadow-sm ${kokuriPillStyle}`}>
                                      {kokuriStatusObj.code === "belum" ? "○" : kokuriStatusObj.code === "lengkap" ? "✓" : "⚠"} {kokuriStatusObj.label.toUpperCase()}
                                    </span>
                                  </td>
                                  <td className="py-3 px-4">
                                    <div className="flex justify-center items-center gap-2">
                                      <button
                                        onClick={() => {
                                          setActiveTab("nilai");
                                          setNilaiSubTab("intra");
                                          setSelectedKelasFilterIntra(s.idKelas);
                                          setSelectedSiswaIdIntra(s.id);
                                          if (!activeIntraCategoryTab && state.kategoriIntrakurikuler.length > 0) {
                                            setActiveIntraCategoryTab(state.kategoriIntrakurikuler[0].id);
                                          }
                                        }}
                                        className="w-8 h-8 rounded-lg bg-indigo-50 border border-indigo-150 flex items-center justify-center text-indigo-600 hover:bg-indigo-600 hover:text-white hover:border-indigo-600 hover:shadow-sm transition-all duration-150 cursor-pointer"
                                        title="Isi Nilai Intrakurikuler"
                                      >
                                        <Edit3 className="w-3.5 h-3.5" />
                                      </button>
                                      <button
                                        onClick={() => {
                                          setActiveTab("nilai");
                                          setNilaiSubTab("kokuri");
                                          setSelectedKelasFilterKokuri(s.idKelas);
                                          setSelectedSiswaIdKokuri(s.id);
                                        }}
                                        className="w-8 h-8 rounded-lg bg-purple-50 border border-purple-150 flex items-center justify-center text-purple-600 hover:bg-purple-600 hover:text-white hover:border-purple-600 hover:shadow-sm transition-all duration-150 cursor-pointer"
                                        title="Isi Nilai Kokurikuler"
                                      >
                                        <Award className="w-3.5 h-3.5" />
                                      </button>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                            {classStudents.length === 0 && (
                              <tr>
                                <td colSpan={6} className="py-10 text-center text-slate-400 italic font-semibold">
                                  Belum ada murid terdaftar dalam kelas ini. Daftarkan murid di menu Siswa.
                                </td>
                              </tr>
                            )}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </div>
                );
              })()}
            </div>
          )}

          {/* TAB B: DATA KELAS */}
          {activeTab === "kelas" && (
            <div className="space-y-6 animate-fade-in no-print">
              <div className="flex justify-between items-center">
                <div>
                  <h2 className="text-xl font-bold text-slate-800 font-display">Data Master Kelas PAUD</h2>
                  <p className="text-sm text-neutral-400">Atur database kelas dan identitas guru wali kelas.</p>
                </div>
                <button
                  onClick={() => {
                    setEditingKelas(null);
                    setKelasForm({ namaKelas: "", waliKelas: "", nuptkNgty: "" });
                    setShowAddKelas(!showAddKelas);
                  }}
                  className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1 shadow-sm transition"
                >
                  <Plus className="w-4 h-4" /> Tambah Kelas
                </button>
              </div>

              {/* ADD CLASS COLLAPSIBLE FORM */}
              {showAddKelas && (
                <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4" onClick={(e) => { if(e.target === e.currentTarget) setShowAddKelas(false); }}>
                  <form onSubmit={handleSaveKelas} className="bg-white p-6 rounded-xl border border-indigo-100 shadow-xl space-y-4 max-w-2xl w-full">
                    <div className="flex items-center gap-2 border-b border-indigo-50 pb-2 mb-2">
                      <span className="text-indigo-600 text-lg">🏫</span>
                      <strong className="text-slate-800 text-sm font-semibold">{editingKelas ? "Edit Kelompok Kelas" : "Tambah Kelompok Kelas Baru"}</strong>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div className="md:col-span-2">
                        <label className="block text-xs font-medium text-slate-600 mb-1">Nama Kelompok Kelas / Usia *</label>
                        <input
                          type="text"
                          placeholder="Contoh: Kelas Pelangi (A - Usia 4-5 Tahun)"
                          value={kelasForm.namaKelas}
                          onChange={(e) => setKelasForm({ ...kelasForm, namaKelas: e.target.value })}
                          required
                          className="w-full text-sm border border-slate-200 px-3 py-2 rounded-lg focus:outline-indigo-600"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-medium text-slate-600 mb-1">Nama Wali Kelas *</label>
                        <input
                          type="text"
                          placeholder="Contoh: Ibu Rahma Nuraini, S.Pd."
                          value={kelasForm.waliKelas}
                          onChange={(e) => setKelasForm({ ...kelasForm, waliKelas: e.target.value })}
                          required
                          className="w-full text-sm border border-slate-200 px-3 py-2 rounded-lg focus:outline-indigo-600"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-medium text-slate-600 mb-1">NUPTK / NGTY Wali Kelas</label>
                        <input
                          type="text"
                          placeholder="Masukkan nomor identitas pendidik"
                          value={kelasForm.nuptkNgty}
                          onChange={(e) => setKelasForm({ ...kelasForm, nuptkNgty: e.target.value })}
                          className="w-full text-sm border border-slate-200 px-3 py-2 rounded-lg focus:outline-indigo-600"
                        />
                      </div>
                    </div>

                    <div className="flex gap-2 justify-end pt-2">
                      <button
                        type="button"
                        onClick={() => setShowAddKelas(false)}
                        className="px-4 py-2 text-xs border border-slate-200 text-slate-600 rounded-lg hover:bg-slate-50 font-medium"
                      >
                        Batal
                      </button>
                      <button
                        type="submit"
                        className="px-4 py-2 text-xs bg-indigo-600 text-white rounded-lg hover:bg-indigo-700 font-semibold"
                      >
                        Simpan Kelas
                      </button>
                    </div>
                  </form>
                </div>
              )}

              {/* LIST OF CLASSES */}
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
                {/* Mobile view cards - optimized for handphones */}
                <div className="block md:hidden divide-y divide-slate-100 bg-white">
                  {state.kelas.length > 0 ? (
                    state.kelas.map((k, idx) => {
                      const count = state.siswa.filter(s => s.idKelas === k.id).length;
                      return (
                        <div key={k.id || `k-card-${idx}`} className="p-4 space-y-3 bg-white">
                          <div className="flex justify-between items-start">
                            <div>
                              <strong className="text-slate-800 text-sm block font-bold leading-snug">{k.namaKelas}</strong>
                              <span className="text-[10px] text-slate-400 font-mono">ID: {k.id}</span>
                            </div>
                            <span className="px-2.5 py-1 bg-indigo-50 text-indigo-700 rounded-md font-bold text-xs">{count} Siswa</span>
                          </div>
                          
                          <div className="grid grid-cols-2 gap-3 text-xs border-t border-slate-150 pt-2 text-slate-600">
                            <div>
                              <span className="text-slate-400 block text-[9px] uppercase font-black tracking-wider">Wali Kelas</span>
                              <span className="font-semibold text-slate-800">{k.waliKelas}</span>
                            </div>
                            <div>
                              <span className="text-slate-400 block text-[9px] uppercase font-black tracking-wider">NUPTK/NGTY</span>
                              <span className="font-mono text-slate-800">{k.nuptkNgty || "-"}</span>
                            </div>
                          </div>

                          <div className="flex justify-end gap-2 border-t border-slate-150 pt-2">
                            <button
                              onClick={() => {
                                setEditingKelas(k);
                                setKelasForm({ namaKelas: k.namaKelas, waliKelas: k.waliKelas, nuptkNgty: k.nuptkNgty });
                                setShowAddKelas(true);
                              }}
                              className="text-indigo-700 hover:bg-indigo-50 border border-indigo-100 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                              title="Edit Kelas"
                            >
                              <Edit3 className="w-3.5 h-3.5" /> Edit
                            </button>
                            <button
                              onClick={() => handleDeleteKelas(k.id)}
                              className="text-red-700 hover:bg-red-50 border border-red-100 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                              title="Hapus Kelas"
                            >
                              <Trash2 className="w-3.5 h-3.5" /> Hapus
                            </button>
                          </div>
                        </div>
                      );
                    })
                  ) : (
                    <div className="text-center py-8 text-slate-400 text-xs font-medium italic">Belum ada kelompok kelas. Kelola kelas baru dengan tombol Tambah Kelas.</div>
                  )}
                </div>

                <div className="hidden md:block overflow-x-auto w-full">
                  <table className="w-full text-left text-sm whitespace-nowrap">
                    <thead className="bg-slate-50 text-xs font-semibold text-slate-600 border-b border-slate-200">
                      <tr>
                        <th className="px-6 py-4">Nama Kelompok & Usia</th>
                        <th className="px-6 py-4">Wali Kelas</th>
                        <th className="px-6 py-4">NUPTK / NGTY</th>
                        <th className="px-6 py-4 text-center">Jumlah Murid</th>
                        <th className="px-6 py-4 text-right">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {state.kelas.length > 0 ? (
                        state.kelas.map((k, idx) => {
                          const count = state.siswa.filter(s => s.idKelas === k.id).length;
                          return (
                            <tr key={k.id || `k-row-${idx}`} className="hover:bg-slate-50/70">
                            <td className="px-6 py-4">
                              <strong className="text-slate-800 block font-medium">{k.namaKelas}</strong>
                              <span className="text-[10px] text-slate-400">ID Kelas: {k.id}</span>
                            </td>
                            <td className="px-6 py-4 text-slate-700 font-medium">{k.waliKelas}</td>
                            <td className="px-6 py-4 text-slate-500 font-mono text-xs">{k.nuptkNgty || "-"}</td>
                            <td className="px-6 py-4 text-center">
                              <span className="px-2.5 py-1 bg-slate-100 text-slate-700 rounded-full font-bold text-xs">{count} Siswa</span>
                            </td>
                            <td className="px-6 py-4 text-right">
                              <div className="inline-flex gap-2">
                                <button
                                  onClick={() => {
                                    setEditingKelas(k);
                                    setKelasForm({ namaKelas: k.namaKelas, waliKelas: k.waliKelas, nuptkNgty: k.nuptkNgty });
                                    setShowAddKelas(true);
                                  }}
                                  className="text-slate-600 hover:text-indigo-600 p-1 rounded hover:bg-slate-100"
                                  title="Edit Kelas"
                                >
                                  <Edit3 className="w-4 h-4" />
                                </button>
                                <button
                                  onClick={() => handleDeleteKelas(k.id)}
                                  className="text-slate-400 hover:text-red-600 p-1 rounded hover:bg-red-50"
                                  title="Hapus Kelas"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        );
                      })
                    ) : (
                      <tr key="empty-kelas">
                        <td colSpan={5} className="text-center py-8 text-slate-400">Belum ada kelompok kelas. Kelola kelas baru dengan tombol Tambah Kelas.</td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            </div>
          )}

          {/* TAB C: DATA SISWA */}
          {activeTab === "siswa" && (
            <div className="space-y-6 animate-fade-in no-print">
              <div className="flex flex-col md:flex-row md:justify-between md:items-center gap-4">
                <div>
                  <h2 className="text-2xl font-bold text-slate-800 font-display">Data Master Siswa PAUD</h2>
                  <p className="text-sm text-neutral-400">Kelola identitas, biodata fisik anak didik, serta pendaftaran kelas.</p>
                </div>
                <div className="flex items-center gap-2 flex-wrap">
                  <button
                    onClick={() => handleExportSiswa()}
                    className="bg-indigo-100 hover:bg-indigo-200 text-indigo-700 px-3 py-2 rounded-lg text-sm font-semibold flex items-center gap-1 shadow-sm transition"
                  >
                    <Download className="w-4 h-4" /> Export Excel
                  </button>
                  <div className="relative">
                    <input
                      type="file"
                      accept=".xlsx, .xls"
                      ref={fileInputRef}
                      onChange={handleImportSiswa}
                      className="hidden"
                    />
                    <button
                      onClick={() => fileInputRef.current?.click()}
                      className="bg-emerald-100 hover:bg-emerald-200 text-emerald-700 px-3 py-2 rounded-lg text-sm font-semibold flex items-center gap-1 shadow-sm transition"
                    >
                      <Upload className="w-4 h-4" /> Import Excel
                    </button>
                  </div>
                  <button
                    onClick={() => {
                      setEditingSiswa(null);
                      setSiswaForm({
                        namaSiswa: "", nisn: "", alamat: "", namaAyah: "", pekerjaanAyah: "", namaIbu: "", pekerjaanIbu: "",
                        noHp: "", tglLahir: "", tempatLahir: "", jenisKelamin: "Laki-laki", anakKe: 1, tb: "100", bb: "15", agama: "Islam", idKelas: state.kelas[0]?.id || ""
                      });
                      setShowAddSiswa(!showAddSiswa);
                    }}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1 shadow-sm transition"
                  >
                    <Plus className="w-4 h-4" /> Registrasi Siswa Baru
                  </button>
                </div>
              </div>
                  {/* REGISTER / EDIT SISWA FORM (MODAL) */}
              {showAddSiswa && (
                <div className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm animate-in fade-in duration-200">
                  <div className="bg-white w-full max-w-4xl max-h-[90vh] overflow-y-auto rounded-2xl shadow-2xl border border-white/20 animate-in zoom-in-95 duration-200">
                    <form onSubmit={handleSaveSiswa} className="p-0">
                      {/* Header Modal */}
                      <div className="sticky top-0 z-10 bg-white border-b border-slate-100 p-6 flex items-center justify-between">
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 rounded-xl bg-indigo-500/10 text-indigo-600 flex items-center justify-center">
                            <Users className="w-6 h-6" />
                          </div>
                        <div>
                          <h3 className="text-xl font-bold text-slate-800 font-display leading-tight">
                            {editingSiswa ? "Ubah Biodata Siswa" : "Registrasi Siswa Baru"}
                          </h3>
                          <p className="text-sm text-neutral-400">Lengkapi formulir identitas anak didik secara akurat.</p>
                        </div>
                        </div>
                        <button 
                          type="button"
                          onClick={() => {
                            setShowAddSiswa(false);
                            setEditingSiswa(null);
                          }}
                          className="w-10 h-10 rounded-xl hover:bg-slate-100 flex items-center justify-center text-slate-400 hover:text-slate-600 transition"
                        >
                          <X className="w-6 h-6" />
                        </button>
                      </div>

                      <div className="p-6 space-y-6">
                        {/* IDENTITAS ANAK */}
                        <div className="space-y-4">
                          <div className="flex items-center gap-2">
                             <div className="px-2 py-0.5 bg-indigo-600 text-white text-[10px] font-black rounded uppercase tracking-wider">Bagian A</div>
                             <span className="text-sm font-bold text-slate-700">IDENTITAS ANAK</span>
                          </div>
                          
                          <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Nama Lengkap Murid *</label>
                              <input
                                type="text"
                                placeholder="Nama Lengkap"
                                value={siswaForm.namaSiswa}
                                onChange={(e) => setSiswaForm({ ...siswaForm, namaSiswa: e.target.value })}
                                required
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">NISN</label>
                              <input
                                type="text"
                                placeholder="Nomor Induk Siswa"
                                value={siswaForm.nisn}
                                onChange={(e) => setSiswaForm({ ...siswaForm, nisn: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Kelas Tingkat PAUD *</label>
                              <select
                                value={siswaForm.idKelas}
                                onChange={(e) => setSiswaForm({ ...siswaForm, idKelas: e.target.value })}
                                required
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl bg-slate-50/50 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition-all font-medium"
                              >
                                <option value="">-- Pilih Kelas --</option>
                                {state.kelas.map(k => (
                                  <option key={k.id} value={k.id}>{k.namaKelas}</option>
                                ))}
                              </select>
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Tempat Lahir</label>
                              <input
                                type="text"
                                placeholder="Kota Lahir"
                                value={siswaForm.tempatLahir}
                                onChange={(e) => setSiswaForm({ ...siswaForm, tempatLahir: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Tanggal Lahir</label>
                              <input
                                type="date"
                                value={siswaForm.tglLahir}
                                onChange={(e) => setSiswaForm({ ...siswaForm, tglLahir: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Jenis Kelamin</label>
                              <select
                                value={siswaForm.jenisKelamin}
                                onChange={(e) => setSiswaForm({ ...siswaForm, jenisKelamin: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl bg-slate-50/50 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition-all font-medium"
                              >
                                <option value="Laki-laki">Laki-laki</option>
                                <option value="Perempuan">Perempuan</option>
                              </select>
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Agama</label>
                              <select
                                value={siswaForm.agama}
                                onChange={(e) => setSiswaForm({ ...siswaForm, agama: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl bg-slate-50/50 focus:outline-none focus:ring-2 focus:ring-indigo-500 transition-all font-medium"
                              >
                                <option value="Islam">Islam</option>
                                <option value="Kristen">Kristen</option>
                                <option value="Katolik">Katolik</option>
                                <option value="Hindu">Hindu</option>
                                <option value="Budha">Buddha</option>
                                <option value="Konghucu">Konghucu</option>
                              </select>
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Anak Ke-</label>
                              <input
                                type="number"
                                min="1"
                                value={siswaForm.anakKe}
                                onChange={(e) => setSiswaForm({ ...siswaForm, anakKe: Number(e.target.value) })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div className="grid grid-cols-2 gap-3">
                              <div>
                                <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">TB (cm)</label>
                                <input
                                  type="text"
                                  placeholder="cm"
                                  value={siswaForm.tb}
                                  onChange={(e) => setSiswaForm({ ...siswaForm, tb: e.target.value })}
                                  className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                                />
                              </div>
                              <div>
                                <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">BB (kg)</label>
                                <input
                                  type="text"
                                  placeholder="kg"
                                  value={siswaForm.bb}
                                  onChange={(e) => setSiswaForm({ ...siswaForm, bb: e.target.value })}
                                  className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                                />
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* DATA ORANG TUA */}
                        <div className="space-y-4">
                          <div className="flex items-center gap-2">
                             <div className="px-2 py-0.5 bg-indigo-600 text-white text-[10px] font-black rounded uppercase tracking-wider">Bagian B</div>
                             <span className="text-sm font-bold text-slate-700">DATA ORANG TUA</span>
                          </div>

                          <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Nama Ayah</label>
                              <input
                                type="text"
                                placeholder="Nama Ayah"
                                value={siswaForm.namaAyah}
                                onChange={(e) => setSiswaForm({ ...siswaForm, namaAyah: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Pekerjaan Ayah</label>
                              <input
                                type="text"
                                placeholder="Pekerjaan Ayah"
                                value={siswaForm.pekerjaanAyah}
                                onChange={(e) => setSiswaForm({ ...siswaForm, pekerjaanAyah: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">No. Telp / HP</label>
                              <input
                                type="text"
                                placeholder="08123xxxx"
                                value={siswaForm.noHp}
                                onChange={(e) => setSiswaForm({ ...siswaForm, noHp: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>
                            
                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Nama Ibu</label>
                              <input
                                type="text"
                                placeholder="Nama Ibu"
                                value={siswaForm.namaIbu}
                                onChange={(e) => setSiswaForm({ ...siswaForm, namaIbu: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div>
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Pekerjaan Ibu</label>
                              <input
                                type="text"
                                placeholder="Pekerjaan Ibu"
                                value={siswaForm.pekerjaanIbu}
                                onChange={(e) => setSiswaForm({ ...siswaForm, pekerjaanIbu: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>

                            <div className="md:col-span-3">
                              <label className="block text-xs font-bold text-slate-500 uppercase tracking-wider mb-1.5 ml-0.5">Alamat Rumah Lengkap</label>
                              <textarea
                                rows={2}
                                placeholder="Tuliskan alamat tinggal lengkap beserta kelurahan, kecamatan, kota"
                                value={siswaForm.alamat}
                                onChange={(e) => setSiswaForm({ ...siswaForm, alamat: e.target.value })}
                                className="w-full text-sm border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 bg-slate-50/50 transition-all font-medium"
                              />
                            </div>
                          </div>
                        </div>
                      </div>

                      {/* Footer Modal Actions */}
                      <div className="sticky bottom-0 bg-white border-t border-slate-100 p-6 flex gap-3 justify-end">
                        <button
                          type="button"
                          onClick={() => {
                            setShowAddSiswa(false);
                            setEditingSiswa(null);
                          }}
                          className="px-6 py-2.5 text-sm font-bold text-slate-500 hover:bg-slate-50 rounded-xl transition-all"
                        >
                          Batalkan
                        </button>
                        <button
                          type="submit"
                          className="px-8 py-2.5 text-sm font-bold bg-indigo-600 text-white rounded-xl hover:bg-indigo-700 shadow-lg shadow-indigo-200 transition-all flex items-center gap-2"
                        >
                          <Save className="w-4 h-4" /> 
                          {editingSiswa ? "Update Biodata" : "Registrasi Sekarang"}
                        </button>
                      </div>
                    </form>
                  </div>
                </div>
              )}

              {/* FILTER BAR */}
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col md:flex-row gap-4 items-center justify-between">
                <div className="flex flex-col sm:flex-row gap-2 w-full md:w-auto">
                  
                  <div className="flex items-center gap-2 border border-slate-200 rounded-lg px-3 bg-white w-full sm:w-64">
                    <Search className="w-4 h-4 text-slate-400 flex-shrink-0" />
                    <input
                      type="text"
                      placeholder="Cari siswa, NIK, ortu..."
                      value={searchSiswa}
                      onChange={(e) => setSearchSiswa(e.target.value)}
                      className="text-xs bg-transparent py-2 focus:outline-none w-full"
                    />
                  </div>

                  {currentUserProfile?.role === "guru" ? (
                    <div className="text-xs bg-indigo-50 border border-indigo-150 px-3 py-2 rounded-lg text-indigo-700 font-semibold flex items-center gap-1.5 shadow-sm">
                      <GraduationCap className="w-3.5 h-3.5" />
                      Kelas Anda: {state.kelas.find(k => k.id === selectedKelasFilter)?.namaKelas || "KB KUSUMA"}
                    </div>
                  ) : (
                    <select
                      value={selectedKelasFilter}
                      onChange={(e) => setSelectedKelasFilter(e.target.value)}
                      className="text-xs border border-slate-200 px-3 py-2 rounded-lg bg-white font-medium shadow-sm transition focus:ring-1 focus:ring-indigo-500"
                    >
                      <option value="">Semua Kelompok Kelas</option>
                      {state.kelas.map((k, idx) => (
                        <option key={k.id || `k-opt2-${idx}`} value={k.id}>{k.namaKelas}</option>
                      ))}
                    </select>
                  )}
                </div>

                <div className="text-xs text-slate-500">
                  Menampilkan <strong>{filteredSiswa.length}</strong> dari <strong>{state.siswa.length}</strong> total murid.
                </div>
              </div>

              {/* LIST OF STUDENTS */}
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm font-sans">
                {/* Mobile view cards for students */}
                <div className="block md:hidden divide-y divide-slate-100 bg-white">
                  {filteredSiswa.length > 0 ? (
                    filteredSiswa.map((s, idx) => (
                      <div key={s.id || `s-card-${idx}`} className="p-4 space-y-3 bg-white">
                        <div className="flex justify-between items-start">
                          <div>
                            <strong className="text-slate-800 text-sm block font-bold leading-snug">{s.namaSiswa}</strong>
                            <span className="text-[10px] text-slate-500 font-semibold uppercase tracking-wider bg-indigo-50 border border-indigo-100/50 px-1.5 py-0.5 rounded text-[9px] mt-1 inline-block">
                              {getSiswaClassLabel(s.idKelas)}
                            </span>
                          </div>
                          <div>
                            <span className="text-neutral-400 block text-[9px] uppercase font-black tracking-wider text-right">NISN</span>
                            <span className="font-mono text-slate-800 text-xs font-bold">{s.nisn || "-"}</span>
                          </div>
                        </div>

                        <div className="flex justify-end gap-2 border-t border-slate-150 pt-2">
                          <button
                            onClick={() => {
                              setEditingSiswa(s);
                              setSiswaForm(s);
                              setShowAddSiswa(true);
                              window.scrollTo({ top: 0, behavior: "smooth" });
                            }}
                            className="text-indigo-700 hover:bg-indigo-50 border border-indigo-100 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                            title="Edit Biodata"
                          >
                            <Edit3 className="w-3.5 h-3.5" /> Edit
                          </button>
                          <button
                            onClick={() => setSiswaToDelete(s)}
                            className="text-red-700 hover:bg-red-50 border border-red-100 px-3 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1.5 transition"
                            title="Hapus Siswa"
                          >
                            <Trash2 className="w-3.5 h-3.5" /> Hapus
                          </button>
                        </div>
                      </div>
                    ))
                  ) : (
                    <div className="text-center py-8 text-slate-400 text-xs font-semibold italic">Tidak ada data siswa yang cocok dengan filter.</div>
                  )}
                </div>

                <div className="hidden md:block overflow-x-auto w-full">
                  <table className="w-full text-left text-sm whitespace-nowrap">
                    <thead className="bg-slate-50 text-xs font-semibold text-slate-600 border-b border-slate-200">
                      <tr>
                        <th className="px-6 py-4">NISN</th>
                        <th className="px-6 py-4">Nama Siswa</th>
                        <th className="px-6 py-4">Kelas</th>
                        <th className="px-6 py-4 text-right">Aksi</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-100">
                      {filteredSiswa.length > 0 ? (
                        filteredSiswa.map((s, idx) => (
                          <tr key={s.id || `s-row-${idx}`} className="hover:bg-slate-50/70">
                            <td className="px-6 py-4 font-mono text-sm">
                              {s.nisn || "-"}
                            </td>
                            <td className="px-6 py-4">
                              <strong className="text-slate-800 block font-medium">{s.namaSiswa}</strong>
                            </td>
                            <td className="px-6 py-4">
                              <span className="text-xs bg-indigo-50 text-indigo-800 px-2.5 py-1 rounded-md font-semibold">
                                {getSiswaClassLabel(s.idKelas)}
                              </span>
                            </td>
                            <td className="px-6 py-4 text-right">
                              <div className="inline-flex gap-2">
                                <button
                                  onClick={() => {
                                    setEditingSiswa(s);
                                    setSiswaForm(s);
                                    setShowAddSiswa(true);
                                    window.scrollTo({ top: 0, behavior: "smooth" });
                                  }}
                                  className="text-slate-600 hover:text-indigo-600 p-1.5 rounded hover:bg-slate-100"
                                  title="Edit Biodata"
                                >
                                  <Edit3 className="w-4 h-4" />
                                </button>
                                <button
                                  onClick={() => setSiswaToDelete(s)}
                                  className="text-slate-400 hover:text-red-600 p-1.5 rounded hover:bg-red-50"
                                  title="Hapus Siswa"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </div>
                            </td>
                          </tr>
                        ))
                      ) : (
                        <tr key="empty-siswa">
                          <td colSpan={5} className="text-center py-8 text-slate-400">Tidak ada data siswa yang cocok dengan filter.</td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </div>

            </div>
          )}

          {/* TAB D: INTRAKURIKULER CAPAIAN */}
          {activeTab === "intra" && (
            <div className="space-y-6 animate-fade-in no-print">
              <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-4">
                <div>
                  <h2 className="text-xl font-bold text-slate-800 font-display">Kurikulum Intrakurikuler PAUD</h2>
                  <p className="text-xs text-neutral-400 font-medium">Pengelolaan Capaian & Tujuan Pembelajaran (TP) per kelompok kelas usia.</p>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => {
                      setShowAddKategori(true);
                      setEditingKategori(null);
                      setKategoriForm({ namaKategori: "" });
                    }}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1 shadow-sm transition"
                  >
                    <Folder className="w-4 h-4" /> Kelola Kategori
                  </button>
                  <button
                    onClick={() => {
                      setTpForm({ idKategori: state.kategoriIntrakurikuler[0]?.id || "KAT-01", deskripsi: "", idKelas: selectedKelasFilterIntra || state.kelas[0]?.id || "", aktivitasMetode: "" });
                      setEditingTp(null);
                      setShowAddTp(true);
                    }}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1 shadow-sm transition"
                  >
                    <Plus className="w-4 h-4" /> Tambah TP
                  </button>
                </div>
              </div>

              {/* CLASS FILTER SECTION */}
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex items-center gap-3">
                <span className="text-xs font-semibold text-slate-600">Filter Kelompok Kelas:</span>
                {currentUserProfile?.role === "guru" ? (
                  <div className="text-xs bg-indigo-50 border border-indigo-150 px-3 py-2 rounded-lg text-indigo-700 font-semibold flex items-center gap-1.5 shadow-sm">
                    <GraduationCap className="w-3.5 h-3.5" />
                    Kelas Anda: {state.kelas.find(k => k.id === selectedKelasFilterIntra)?.namaKelas || "KB KUSUMA"}
                  </div>
                ) : (
                  <select
                    value={selectedKelasFilterIntra}
                    onChange={(e) => setSelectedKelasFilterIntra(e.target.value)}
                    className="text-xs border border-slate-200 px-3 py-2 rounded-lg bg-white font-medium"
                  >
                    {state.kelas.map((k, idx) => (
                      <option key={k.id || `k-opt3-${idx}`} value={k.id}>{k.namaKelas}</option>
                    ))}
                  </select>
                )}
              </div>

              {/* The TP and Category forms will be rendered at the bottom as modals */}


              {/* LIST BY CATEGORIES */}
              {state.kategoriIntrakurikuler.map((kat, katIdx) => {
                const classTps = state.tujuanPembelajaran.filter(tp => tp.idKategori === kat.id && tp.idKelas === selectedKelasFilterIntra);
                return (
                  <div key={kat.id || `kat-list-${katIdx}`} className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                    <div className="bg-slate-50 border-b border-slate-200 px-5 py-3 flex justify-between items-center">
                      <div className="flex items-center gap-3">
                        <strong className="text-slate-800 text-sm font-semibold">{kat.namaKategori}</strong>
                        <span className="text-[10px] bg-indigo-100 text-indigo-800 px-2 py-0.5 rounded-full font-bold uppercase">{classTps.length} TP</span>
                      </div>
                      <div className="flex items-center gap-1">
                        <button
                          onClick={() => {
                            setEditingKategori(kat);
                            setKategoriForm({ namaKategori: kat.namaKategori });
                            setShowAddKategori(true);
                            window.scrollTo({ top: 0, behavior: "smooth" });
                          }}
                          className="text-slate-400 hover:text-indigo-600 p-1.5 rounded-md hover:bg-white transition shadow-sm"
                          title="Ubah Kategori"
                        >
                          <Edit3 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => setKategoriToDelete(kat)}
                          className="text-slate-400 hover:text-red-600 p-1.5 rounded-md hover:bg-white transition shadow-sm"
                          title="Hapus Kategori"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>

                    <div className="p-4 overflow-x-auto">
                      {classTps.length > 0 ? (
                        <table className="w-full text-left text-xs border border-slate-150 rounded-lg overflow-hidden border-collapse">
                          <thead>
                            <tr className="bg-slate-50 text-slate-700 border-b border-slate-150">
                              <th className="px-4 py-3 w-12 text-center font-bold font-display">No.</th>
                              <th className="px-4 py-3 font-semibold font-display">Deskripsi Tujuan Pembelajaran</th>
                              <th className="px-4 py-3 font-semibold font-display">Aktivitas / Metode</th>
                              <th className="px-4 py-3 text-center font-semibold font-display w-24">Aksi</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-slate-100 text-slate-700 bg-white">
                            {classTps.map((tp, idx) => (
                              <tr key={tp.id || `tp-row-manage-${idx}`} className="hover:bg-slate-50/50 transition-colors">
                                <td className="px-4 py-3 text-center font-mono font-semibold text-slate-400">#{idx + 1}</td>
                                <td className="px-4 py-3 leading-relaxed font-sans text-slate-800 font-medium">{tp.deskripsi}</td>
                                <td className="px-4 py-3 leading-relaxed font-sans text-slate-600 bg-slate-50/30">{tp.aktivitasMetode || "-"}</td>
                                <td className="px-4 py-3 text-center">
                                  <div className="flex items-center justify-center gap-2">
                                    <button
                                      onClick={() => {
                                        setEditingTp(tp);
                                        setTpForm({
                                          idKategori: tp.idKategori,
                                          deskripsi: tp.deskripsi,
                                          idKelas: tp.idKelas,
                                          aktivitasMetode: tp.aktivitasMetode || ""
                                        });
                                        setShowAddTp(true);
                                        window.scrollTo({ top: 0, behavior: "smooth" });
                                      }}
                                      className="text-slate-400 hover:text-indigo-600 p-1.5 rounded-lg hover:bg-indigo-50 transition-colors"
                                      title="Ubah TP"
                                    >
                                      <Edit3 className="w-4 h-4" />
                                    </button>
                                    <button
                                      onClick={() => handleDeleteTp(tp)}
                                      className="text-slate-400 hover:text-red-500 p-1.5 rounded-lg hover:bg-rose-50 transition-colors"
                                      title="Hapus TP"
                                    >
                                      <Trash2 className="w-4 h-4" />
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      ) : (
                        <div className="text-center py-6 text-xs text-neutral-400">Belum ada tujuan pembelajaran terdaftar pada kategori ini khusus kelas terpilih.</div>
                      )}
                    </div>
                  </div>
                );
              })}

            </div>
          )}

          {/* TAB E: KOKURIKULER CAPAIAN */}
          {activeTab === "kokuri" && (
            <div className="space-y-6 animate-fade-in no-print">
              <div className="flex flex-col sm:flex-row justify-between sm:items-center gap-4">
                <div>
                  <h2 className="text-xl font-bold text-slate-800 font-display">Kurikulum Kokurikuler (Projek Penguatan Projek Profil Pancasila - P5)</h2>
                  <p className="text-xs text-neutral-400 font-medium">Kelola daftar Subdimensi projek berdasarkan Kelas yang diampu.</p>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => {
                      setShowAddLabel(true);
                      setEditingLabel(null);
                      setLabelForm({ namaLabel: "", order: state.labelP5.length + 1 });
                    }}
                    className="bg-slate-100 hover:bg-slate-200 text-slate-700 px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1 shadow-sm transition"
                  >
                    <Settings className="w-4 h-4" /> Kelola Label Capaian
                  </button>
                  <button
                    onClick={() => {
                      setSubForm({ namaSubdimensi: "", idKelas: selectedKelasFilterKokuri || state.kelas[0]?.id || "", descBerkembang: "", descCakap: "", descMahir: "", capaian: {} });
                      setEditingSub(null);
                      setShowAddSub(true);
                    }}
                    className="bg-indigo-600 hover:bg-indigo-700 text-white px-4 py-2 rounded-lg text-sm font-semibold flex items-center gap-1 shadow-sm transition"
                  >
                    <Plus className="w-4 h-4" /> Tambah Subdimensi Projek
                  </button>
                </div>
              </div>

              {/* CLASS SEGREGATOR */}
              <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex items-center gap-3">
                <span className="text-xs font-semibold text-slate-600">Filter Kelompok Kelas:</span>
                {currentUserProfile?.role === "guru" ? (
                  <div className="text-xs bg-indigo-50 border border-indigo-150 px-3 py-2 rounded-lg text-indigo-700 font-semibold flex items-center gap-1.5 shadow-sm">
                    <GraduationCap className="w-3.5 h-3.5" />
                    Kelas Anda: {state.kelas.find(k => k.id === selectedKelasFilterKokuri)?.namaKelas || "KB KUSUMA"}
                  </div>
                ) : (
                  <select
                    value={selectedKelasFilterKokuri}
                    onChange={(e) => setSelectedKelasFilterKokuri(e.target.value)}
                    className="text-xs border border-slate-200 px-3 py-2 rounded-lg bg-white font-medium"
                  >
                    {state.kelas.map((k, idx) => (
                      <option key={k.id || `k-opt5-${idx}`} value={k.id}>{k.namaKelas}</option>
                    ))}
                  </select>
                )}
              </div>

              {/* The Subdimensi form will be rendered at the bottom as a modal */}

              {/* LIST SUBDIMENSI */}
              <div className="bg-white rounded-xl border border-slate-200 overflow-hidden shadow-sm">
                <div className="bg-slate-50 border-b border-slate-200 px-5 py-3 flex justify-between items-center">
                  <strong className="text-slate-800 text-sm font-semibold">Subdimensi Kokurikuler Kelas - Indikator Capaian</strong>
                  <div className="flex gap-2">
                    {state.labelP5.map(lbl => (
                      <span key={lbl.id} className="text-[10px] font-bold bg-white border border-slate-200 px-2 py-0.5 rounded text-slate-500">{lbl.namaLabel}</span>
                    ))}
                  </div>
                </div>

                <div className="p-5">
                  {state.subdimensiKokurikuler.filter(s => s.idKelas === selectedKelasFilterKokuri).length > 0 ? (
                    <div className="space-y-4">
                      {state.subdimensiKokurikuler
                        .filter(s => s.idKelas === selectedKelasFilterKokuri)
                        .map((sub, idx) => (
                          <div key={sub.id || `sub-div-${idx}`} className="flex justify-between items-start gap-4 p-4 border border-slate-150 rounded-xl bg-white shadow-sm hover:shadow-md transition-shadow">
                            <div className="flex-1 space-y-3">
                              <div className="flex items-center gap-3">
                                <span className="font-mono text-xs font-bold text-indigo-600 bg-indigo-50 px-2 py-0.5 rounded">Sub #{idx + 1}</span>
                                <h4 className="text-sm font-bold text-slate-800 leading-relaxed">{sub.namaSubdimensi}</h4>
                              </div>
                              
                              <div className={`grid grid-cols-1 md:grid-cols-${state.labelP5.length > 0 ? Math.min(state.labelP5.length, 4) : 3} gap-3`}>
                                {state.labelP5.map((label) => {
                                  const colors = [
                                    { bg: "bg-emerald-500", text: "text-emerald-700", border: "border-emerald-100", light: "bg-emerald-50/50" },
                                    { bg: "bg-sky-500", text: "text-sky-700", border: "border-sky-100", light: "bg-sky-50/50" },
                                    { bg: "bg-indigo-500", text: "text-indigo-700", border: "border-indigo-100", light: "bg-indigo-50/50" },
                                    { bg: "bg-amber-500", text: "text-amber-700", border: "border-amber-100", light: "bg-amber-50/50" },
                                    { bg: "bg-rose-500", text: "text-rose-700", border: "border-rose-100", light: "bg-rose-50/50" },
                                  ];
                                  const color = colors[(label.order - 1) % colors.length];
                                  const desc = sub.capaian?.[label.id] || (label.namaLabel === "Berkembang" ? sub.descBerkembang : label.namaLabel === "Cakap" ? sub.descCakap : label.namaLabel === "Mahir" ? sub.descMahir : "");
                                  
                                  return (
                                    <div key={label.id} className={`${color.light} p-2.5 rounded-lg border ${color.border}`}>
                                      <span className={`block text-[10px] font-bold ${color.text} uppercase tracking-wider mb-1`}>{label.namaLabel}</span>
                                      <p className="text-[11px] text-slate-600 leading-relaxed italic">{desc || "-"}</p>
                                    </div>
                                  );
                                })}
                                {state.labelP5.length === 0 && (
                                  <div className="col-span-full py-4 text-center text-slate-400 text-xs italic">
                                    Belum ada label capaian terdaftar.
                                  </div>
                                )}
                              </div>
                            </div>

                            <div className="flex items-center gap-1.5 flex-shrink-0 pt-1">
                              <button
                                onClick={() => {
                                  setEditingSub(sub);
                                  setSubForm({ 
                                    namaSubdimensi: sub.namaSubdimensi, 
                                    idKelas: sub.idKelas,
                                    descBerkembang: sub.descBerkembang || "",
                                    descCakap: sub.descCakap || "",
                                    descMahir: sub.descMahir || "",
                                    capaian: sub.capaian || {}
                                  });
                                  setShowAddSub(true);
                                }}
                                className="text-slate-400 hover:text-indigo-600 p-2 rounded-lg hover:bg-indigo-50 transition-colors"
                                title="Ubah Subdimensi"
                              >
                                <Edit3 className="w-4 h-4" />
                              </button>
                              <button
                                onClick={() => setSubToDelete(sub)}
                                className="text-slate-400 hover:text-red-500 p-2 rounded-lg hover:bg-rose-50 transition-colors"
                                title="Hapus Subdimensi"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </div>
                          </div>
                      ))}
                    </div>
                  ) : (
                    <div className="text-center py-8 text-sm text-neutral-400">Belum ada deskripsi subdimensi Kokurikuler terdaftar untuk kelompok kelas terpilih.</div>
                  )}
                </div>
              </div>

            </div>
          )}

          {/* TAB F: INPUT NILAI & ATTENDANCE */}
          {activeTab === "nilai" && (
            <div className="space-y-6 animate-fade-in no-print">
              
              {/* CORE INPUT PANEL: 1. INTRA */}
              {nilaiSubTab === "intra" && (
                <div className="space-y-6">
                  <div>
                    <h3 className="font-black text-slate-900 text-xl font-display">1. Lembar Penilaian Intrakurikuler</h3>
                    <p className="text-sm font-medium text-neutral-500 font-sans mt-1">Input kriteria nilai perkembangan anak didik berdasarkan Tujuan Pembelajaran (TP).</p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {currentUserProfile?.role === "guru" ? (
                      <div className="bg-indigo-50/50 p-4 rounded-xl border border-indigo-100 shadow-sm flex flex-col gap-1.5 justify-center">
                        <span className="text-[10px] font-bold text-indigo-500 uppercase tracking-wider">Kelas yang Diampu</span>
                        <div className="text-sm font-extrabold text-indigo-900 flex items-center gap-1.5">
                          <GraduationCap className="w-4 h-4 text-indigo-600" />
                          {state.kelas.find(k => k.id === selectedKelasFilterIntra)?.namaKelas || "KB KUSUMA"}
                        </div>
                      </div>
                    ) : (
                    <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xl flex flex-col gap-2">
                        <span className="text-base font-black text-slate-700">Langkah 1: Pilih Kelompok Kelas</span>
                        <select
                          value={selectedKelasFilterIntra}
                          onChange={(e) => {
                            setSelectedKelasFilterIntra(e.target.value);
                            setSelectedSiswaIdIntra(""); // Reset student on class change
                          }}
                          className="text-base border border-slate-300 px-4 py-3 rounded-xl bg-white font-bold text-indigo-900 focus:outline-indigo-600 shadow-sm"
                        >
                          <option value="">-- Pilih Kelas --</option>
                          {state.kelas.map((k, idx) => (
                            <option key={k.id || `k-opt7-${idx}`} value={k.id}>{k.namaKelas}</option>
                          ))}
                        </select>
                      </div>
                    )}

                    <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xl flex flex-col gap-2">
                      <span className="text-base font-black text-slate-700">Langkah 2: Pilih Nama Siswa</span>
                      <select
                        value={selectedSiswaIdIntra}
                        onChange={(e) => setSelectedSiswaIdIntra(e.target.value)}
                        className="text-base border border-slate-300 px-4 py-3 rounded-xl bg-white font-bold text-indigo-900 focus:outline-indigo-600 shadow-sm"
                        disabled={!selectedKelasFilterIntra}
                      >
                        <option value="">-- Pilih Siswa --</option>
                        {state.siswa.filter(s => s.idKelas === selectedKelasFilterIntra).map((s, idx) => (
                          <option key={s.id || `s-opt-i-${idx}`} value={s.id}>{s.namaSiswa}</option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {/* ACTIVE STUDENT EVAL LIST */}
                  {selectedSiswaIdIntra ? (
                    <div className="space-y-4">
                      {(() => {
                        const s = state.siswa.find(siswa => siswa.id === selectedSiswaIdIntra);
                        if (!s) return null;
                        const classTps = state.tujuanPembelajaran.filter(tp => tp.idKelas === selectedKelasFilterIntra);

                        // Determine the first category that actually has TP data for this selected class
                        const firstKatWithTps = state.kategoriIntrakurikuler.find(kat => classTps.some(tp => tp.idKategori === kat.id));
                        const fallbackKatId = firstKatWithTps ? firstKatWithTps.id : (state.kategoriIntrakurikuler[0]?.id || "");
                        const hasTpsForActiveTab = classTps.some(tp => tp.idKategori === activeIntraCategoryTab);
                        const currentActiveKatId = (activeIntraCategoryTab && hasTpsForActiveTab) ? activeIntraCategoryTab : fallbackKatId;

                        return (
                          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6 font-sans">
                            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
                              <div>
                                <h4 className="text-base font-bold text-indigo-950 font-display">{s.namaSiswa}</h4>
                                <span className="text-[10px] text-slate-400 block uppercase font-bold tracking-widest">Penilaian Capaian Belajar</span>
                              </div>
                              <span className="text-xs font-mono font-medium text-slate-400 bg-slate-50 px-2 py-1 rounded border border-slate-100">NISN: {s.nisn}</span>
                            </div>

                            {classTps.length > 0 ? (
                              <div className="space-y-6 animate-fadeIn">
                                {/* CATEGORY TABS SELECTOR */}
                                <div className="flex flex-wrap gap-1 border-b border-slate-200">
                                  {state.kategoriIntrakurikuler.map((kat, katIdx) => {
                                    const catTps = classTps.filter(tp => tp.idKategori === kat.id);
                                    if (catTps.length === 0) return null;
                                    
                                    const isTabActive = currentActiveKatId === kat.id;
                                    const count = catTps.length;

                                    return (
                                      <button
                                        key={kat.id}
                                        type="button"
                                        onClick={() => setActiveIntraCategoryTab(kat.id)}
                                        className={`px-4 py-2.5 text-xs sm:text-sm font-semibold border-b-2 flex items-center gap-2 transition-all cursor-pointer ${
                                          isTabActive 
                                            ? "border-indigo-600 text-indigo-800 bg-indigo-50/20 font-bold" 
                                            : "border-transparent text-slate-500 hover:text-slate-800 hover:bg-slate-50"
                                        }`}
                                      >
                                        <span className={`w-5 h-5 rounded-full flex items-center justify-center text-[10px] font-bold font-mono transition-colors ${
                                          isTabActive ? "bg-indigo-600 text-white" : "bg-slate-100 text-slate-600"
                                        }`}>
                                          {katIdx + 1}
                                        </span>
                                        <span className="truncate max-w-[150px] sm:max-w-xs">{kat.namaKategori}</span>
                                        <span className={`text-[9.5px] font-bold px-1.5 py-0.5 rounded-full ${
                                          isTabActive ? "bg-indigo-600 text-white" : "bg-slate-200 text-slate-500"
                                        }`}>
                                          {count} TP
                                        </span>
                                      </button>
                                    );
                                  })}
                                </div>

                                {/* ACTIVE CATEGORY CONTENT COMPONENT */}
                                {(() => {
                                  const kat = state.kategoriIntrakurikuler.find(k => k.id === currentActiveKatId);
                                  if (!kat) return null;

                                  const catTps = classTps.filter(tp => tp.idKategori === kat.id);
                                  if (catTps.length === 0) {
                                    return (
                                      <div className="text-center py-8 text-xs text-slate-400">
                                        Tidak ada Tujuan Pembelajaran yang terdaftar untuk kategori ini pada kelompok kelas aktif.
                                      </div>
                                    );
                                  }

                                  const katIdx = state.kategoriIntrakurikuler.findIndex(k => k.id === kat.id);
                                  const catAssessment = state.nilaiIntrakurikuler.find(n => n.idSiswa === s.id && n.idTp === kat.id);
                                  const currentCatDesc = catAssessment?.deskripsi || "";
                                  const isGenerating = generatingAiItem === kat.id;
                                  const hasAnyGrades = catTps.some(tp => state.nilaiIntrakurikuler.find(n => n.idSiswa === s.id && n.idTp === tp.id)?.nilai);

                                  return (
                                    <div key={kat.id} className="bg-slate-50/40 p-5 rounded-2xl border border-slate-150 space-y-4 shadow-sm animate-fadeIn">
                                      <div className="flex items-center gap-2 pb-2 border-b border-slate-100">
                                        <span className="w-6 h-6 rounded-full bg-indigo-50 text-indigo-700 flex items-center justify-center text-xs font-bold font-mono shadow-sm">
                                          {katIdx + 1}
                                        </span>
                                        <h5 className="text-sm font-bold text-slate-800 font-display">
                                          {kat.namaKategori}
                                        </h5>
                                      </div>

                                      <div className="space-y-3 pl-1 font-sans">
                                        {catTps.map((tp, tpIdx) => {
                                          const assessment = state.nilaiIntrakurikuler.find(n => n.idSiswa === s.id && n.idTp === tp.id);
                                          const currentGrade = assessment?.nilai || "";

                                          return (
                                                <div key={tp.id} className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-slate-100 last:border-0 last:pb-0">
                                                  <div className="flex-1">
                                                    <p className="text-sm font-bold text-slate-800 leading-relaxed">
                                                      <span className="text-indigo-600 font-black mr-2 font-mono">TP-{tpIdx + 1}</span>
                                                      {tp.deskripsi}
                                                    </p>
                                                    {tp.aktivitasMetode && (
                                                      <div className="text-[11px] text-slate-500 font-medium italic mt-1 flex items-center gap-1">
                                                        <span className="w-1.5 h-1.5 rounded-full bg-slate-300"></span>
                                                        Metode: {tp.aktivitasMetode}
                                                      </div>
                                                    )}
                                                  </div>

                                                  <div className="flex-shrink-0">
                                                    <div className="bg-white p-1.5 rounded-xl border border-slate-300 inline-flex gap-1.5 shadow-md">
                                                        {state.labelP5.map((lbl) => (
                                                          <button
                                                            key={lbl.id}
                                                            type="button"
                                                            onClick={() => handleSetGradeIntra(s.id, tp.id, currentGrade === lbl.namaLabel ? "" : lbl.namaLabel)}
                                                            className={`px-4 py-2 text-[10px] font-bold rounded-lg transition-all cursor-pointer border-2 ${
                                                              currentGrade === lbl.namaLabel 
                                                                ? "bg-indigo-600 text-white border-indigo-700 shadow-md scale-105" 
                                                                : "text-slate-500 border-dashed border-slate-200 hover:text-slate-800 hover:bg-slate-50"
                                                            }`}
                                                          >
                                                            {lbl.namaLabel}
                                                          </button>
                                                        ))}
                                                    </div>
                                                  </div>
                                                </div>
                                          );
                                        })}
                                      </div>

                                      {/* Aggregated Description Textbox for Category */}
                                      <div className="space-y-2 pt-2 border-t border-slate-100">
                                        <div className="flex items-center justify-between font-sans">
                                          <label className="text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                                            Narasi Rekapitulasi - {kat.namaKategori}
                                          </label>
                                          <span className="text-[10px] text-slate-400 font-medium font-sans">Rekap aspek</span>
                                        </div>

                                        <div className="relative group">
                                          <div className="rich-text-editor-container bg-white rounded-xl border border-slate-200 overflow-hidden shadow-inner min-h-[150px]">
                                            <ReactQuill
                                              theme="snow"
                                              value={currentCatDesc}
                                              onChange={(content) => handleUpdateDescriptionIntra(s.id, kat.id, content)}
                                              placeholder={`Tuliskan deskripsi narasi perkembangan gabungan (rekap) untuk aspek ${kat.namaKategori}...`}
                                              modules={{
                                                toolbar: [
                                                  ['bold', 'italic', 'underline'],
                                                  [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                                                  ['clean']
                                                ],
                                              }}
                                              className="quill-editor"
                                            />
                                          </div>
                                          <button
                                            type="button"
                                            onClick={() => composeAiItemText(s.id, kat.id, "intra")}
                                            disabled={!hasAnyGrades || isGenerating}
                                            title={hasAnyGrades ? "Tulis rekap deskripsi otomatis dengan AI" : "Beri minimal satu predikat TP di atas untuk mengaktifkan AI"}
                                            className="absolute right-3 top-3 p-2 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-600 hover:text-white transition-all disabled:opacity-30 disabled:cursor-not-allowed shadow-sm border border-indigo-100 cursor-pointer"
                                          >
                                            {isGenerating ? (
                                              <div className="w-3.5 h-3.5 border-2 border-current border-t-transparent rounded-full animate-spin"></div>
                                            ) : (
                                              <Sparkles className="w-3.5 h-3.5" />
                                            )}
                                          </button>
                                        </div>
                                      </div>
                                    </div>
                                  );
                                })()}
                              </div>
                            ) : (
                              <div className="text-center py-12 text-slate-400 bg-slate-50 rounded-xl border border-dashed border-slate-200">
                                <Search className="w-8 h-8 mx-auto mb-2 opacity-20" />
                                <p className="text-sm font-medium">Belum ada TP terdaftar untuk kelas ini.</p>
                                <p className="text-xs">Atur di tab Intrakurikuler Manajemen.</p>
                              </div>
                            )}
                          </div>
                        );
                      })()}
                    </div>
                  ) : (
                    <div className="bg-white p-12 rounded-2xl border border-dashed border-slate-200 text-center space-y-3">
                      <div className="w-16 h-16 bg-slate-50 rounded-2xl flex items-center justify-center mx-auto text-slate-300">
                        <Users className="w-8 h-8" />
                      </div>
                      <div className="max-w-xs mx-auto">
                        <h4 className="font-bold text-slate-700">Pilih Siswa Terlebih Dahulu</h4>
                        <p className="text-xs text-slate-400 leading-relaxed">Silakan pilih kelompok kelas dan nama siswa untuk mulai memberikan penilaian terarah.</p>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* CORE INPUT PANEL: 2. KOKURI */}
              {nilaiSubTab === "kokuri" && (
                <div className="space-y-6">
                  <div>
                    <h3 className="font-black text-slate-900 text-xl font-display">2. Lembar Penilaian Projek K-Merdeka (P5 Kokurikuler)</h3>
                    <p className="text-sm font-medium text-neutral-500 font-sans mt-1">Input capaian murid pada komponen Projek P5 berdasarkan subdimensi kurikulum.</p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {currentUserProfile?.role === "guru" ? (
                      <div className="bg-indigo-50/50 p-4 rounded-xl border border-indigo-100 shadow-sm flex flex-col gap-1.5 justify-center">
                        <span className="text-[10px] font-bold text-indigo-500 uppercase tracking-wider">Kelas yang Diampu</span>
                        <div className="text-sm font-extrabold text-indigo-900 flex items-center gap-1.5">
                          <GraduationCap className="w-4 h-4 text-indigo-600" />
                          {state.kelas.find(k => k.id === selectedKelasFilterKokuri)?.namaKelas || "KB KUSUMA"}
                        </div>
                      </div>
                    ) : (
                    <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xl flex flex-col gap-2">
                        <span className="text-base font-black text-slate-700">Langkah 1: Pilih Kelompok Kelas</span>
                        <select
                          value={selectedKelasFilterKokuri}
                          onChange={(e) => {
                            setSelectedKelasFilterKokuri(e.target.value);
                            setSelectedSiswaIdKokuri(""); // Reset student
                          }}
                          className="text-base border border-slate-300 px-4 py-3 rounded-xl bg-white font-bold text-indigo-900 focus:outline-indigo-600 shadow-sm"
                        >
                          <option value="">-- Pilih Kelas --</option>
                          {state.kelas.map((k, idx) => (
                            <option key={k.id || `k-opt8-${idx}`} value={k.id}>{k.namaKelas}</option>
                          ))}
                        </select>
                      </div>
                    )}

                    <div className="bg-white p-5 rounded-2xl border border-slate-200 shadow-xl flex flex-col gap-2">
                      <span className="text-base font-black text-slate-700">Langkah 2: Pilih Nama Siswa</span>
                      <select
                        value={selectedSiswaIdKokuri}
                        onChange={(e) => setSelectedSiswaIdKokuri(e.target.value)}
                        className="text-base border border-slate-300 px-4 py-3 rounded-xl bg-white font-bold text-indigo-900 focus:outline-indigo-600 shadow-sm"
                        disabled={!selectedKelasFilterKokuri}
                      >
                        <option value="">-- Pilih Siswa --</option>
                        {state.siswa.filter(s => s.idKelas === selectedKelasFilterKokuri).map((s, idx) => (
                          <option key={s.id || `s-opt-k-${idx}`} value={s.id}>{s.namaSiswa}</option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {selectedSiswaIdKokuri ? (
                    <div className="space-y-4">
                      {(() => {
                        const s = state.siswa.find(siswa => siswa.id === selectedSiswaIdKokuri);
                        if (!s) return null;
                        const classSubs = state.subdimensiKokurikuler.filter(sub => sub.idKelas === selectedKelasFilterKokuri);
                        return (
                          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6 font-sans">
                            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
                              <div>
                                <h4 className="text-xl font-black text-indigo-950 font-display">{s.namaSiswa}</h4>
                                <span className="text-xs text-slate-500 block uppercase font-bold tracking-widest text-emerald-600">Capaian Projek Profil (P5)</span>
                              </div>
                              <div className="flex items-center gap-2">
                                <button
                                  type="button"
                                  onClick={() => setShowKokuriRubrikRef(!showKokuriRubrikRef)}
                                  className={`px-2.5 py-1.5 rounded-lg text-xs font-bold flex items-center gap-1 border transition-all cursor-pointer ${
                                    showKokuriRubrikRef 
                                      ? "bg-indigo-50 border-indigo-200 text-indigo-700 hover:bg-indigo-100/70"
                                      : "bg-slate-50 border-slate-200 text-slate-600 hover:bg-slate-100"
                                  }`}
                                >
                                  <BookOpen className="w-3.5 h-3.5" />
                                  {showKokuriRubrikRef ? "Sembunyikan Acuan" : "Tampilkan Acuan"}
                                </button>
                                <span className="text-xs font-semibold text-slate-450 hidden sm:inline ml-2">Projek Kelas</span>
                              </div>
                            </div>

                            {classSubs.length > 0 ? (
                              <div className="space-y-8">
                                {classSubs.map((sub, subIdx) => {
                                  const assessment = state.nilaiKokurikuler.find(n => n.idSiswa === s.id && n.idSubdimensi === sub.id);
                                  const currentGrade = assessment?.nilai || "";
                                  const currentDesc = assessment?.deskripsi || "";
                                  const isGenerating = generatingAiItem === sub.id;
                                  
                                  return (
                                    <div key={sub.id || `sub-card-${subIdx}`} className="space-y-3 pb-8 border-b border-slate-50 last:border-0 last:pb-0">
                                      <div className="flex flex-col md:flex-row md:items-start justify-between gap-5">
                                        <div className="flex-1">
                                          <div className="flex items-center gap-3 mb-2">
                                            <span className="bg-emerald-600 text-white text-[10px] font-black px-2.5 py-0.5 rounded-md uppercase font-sans shadow-sm tracking-widest">SUB-{subIdx + 1}</span>
                                          </div>
                                          <p className="text-sm font-bold text-slate-800 leading-relaxed tracking-tight">{sub.namaSubdimensi}</p>
                                        </div>

                                        <div className="flex-shrink-0 font-sans">
                                          <div className="bg-slate-100 p-2 rounded-2xl border border-slate-300 inline-flex gap-1.5 shadow-inner">
                                              {state.labelP5.map((lbl) => (
                                                <button
                                                  key={lbl.id}
                                                  type="button"
                                                  onClick={() => handleSetGradeKokuri(s.id, sub.id, currentGrade === lbl.namaLabel ? "" : lbl.namaLabel)}
                                                  className={`px-4 py-2 text-[10px] font-bold rounded-lg transition-all transform active:scale-95 cursor-pointer border-2 ${
                                                    currentGrade === lbl.namaLabel 
                                                      ? "bg-indigo-600 text-white border-indigo-700 shadow-md scale-105" 
                                                      : "text-slate-500 border-dashed border-slate-200 hover:text-slate-900 hover:bg-white"
                                                  }`}
                                                >
                                                  {lbl.namaLabel}
                                                </button>
                                              ))}
                                          </div>
                                        </div>
                                      </div>

                                      {/* REFERENCE/RUBRIC ACUAN FROM MASTER DATA */}
                                      {showKokuriRubrikRef && (
                                        <div className="bg-slate-50/70 p-3 rounded-xl border border-slate-100/90 space-y-2.5 font-sans">
                                          <div className="flex items-center gap-2 text-xs font-black text-indigo-700 uppercase tracking-widest pl-1">
                                            <div className="w-2 h-2 rounded-full bg-indigo-600 shadow-sm"></div>
                                            <span>Rubrik Acuan Penilaian P5 (Petunjuk Pengisian)</span>
                                          </div>
                                          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
                                            {state.labelP5.map((label) => {
                                              const colors = [
                                                { bg: "bg-emerald-500", text: "text-emerald-800", border: "border-emerald-200", light: "bg-emerald-50" },
                                                { bg: "bg-sky-500", text: "text-sky-800", border: "border-sky-200", light: "bg-sky-50" },
                                                { bg: "bg-indigo-500", text: "text-indigo-800", border: "border-indigo-200", light: "bg-indigo-50" },
                                                { bg: "bg-amber-500", text: "text-amber-800", border: "border-amber-200", light: "bg-amber-50" },
                                                { bg: "bg-rose-500", text: "text-rose-800", border: "border-rose-200", light: "bg-rose-50" },
                                              ];
                                              const color = colors[(label.order - 1) % colors.length] || colors[0];
                                              const desc = sub.capaian?.[label.id] || (label.namaLabel === "Berkembang" ? sub.descBerkembang : label.namaLabel === "Cakap" ? sub.descCakap : label.namaLabel === "Mahir" ? sub.descMahir : "");

                                              return (
                                                <div key={label.id} className={`${color.light} p-4 rounded-2xl border-2 ${color.border} flex flex-col justify-start gap-2 shadow-sm`}>
                                                  <span className={`block text-[11px] font-black ${color.text} uppercase tracking-widest border-b border-black/5 pb-1`}>{label.namaLabel}</span>
                                                  <p className="text-[13px] font-bold text-slate-800 leading-relaxed italic">{desc || "Deskripsi indikator tidak diatur."}</p>
                                                </div>
                                              );
                                            })}
                                          </div>
                                        </div>
                                      )}

                                      <div className="relative group">
                                          <div className="rich-text-editor-container bg-white rounded-xl border border-slate-300 overflow-hidden shadow-sm min-h-[150px]">
                                            <ReactQuill
                                              theme="snow"
                                              value={currentDesc}
                                              onChange={(content) => handleUpdateDescriptionKokuri(s.id, sub.id, content)}
                                              placeholder={`Tuliskan narasi perkembangan untuk subdimensi Projek ini...`}
                                              modules={{
                                                toolbar: [
                                                  ['bold', 'italic', 'underline'],
                                                  [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                                                  ['clean']
                                                ],
                                              }}
                                              className="quill-editor"
                                            />
                                          </div>
                                        <button
                                          onClick={() => composeAiItemText(s.id, sub.id, "kokuri")}
                                          disabled={!currentGrade || isGenerating}
                                          title="Tulis deskripsi dengan AI"
                                          className="absolute right-3 top-3 p-2 rounded-lg bg-indigo-50 text-indigo-600 hover:bg-indigo-600 hover:text-white transition-all disabled:opacity-30 disabled:cursor-not-allowed shadow-sm border border-indigo-100"
                                        >
                                          {isGenerating ? (
                                            <div className="w-3.5 h-3.5 border-2 border-current border-t-transparent rounded-full animate-spin"></div>
                                          ) : (
                                            <Sparkles className="w-3.5 h-3.5" />
                                          )}
                                        </button>
                                      </div>
                                    </div>
                                  );
                                })}
                              </div>
                            ) : (
                              <div className="text-center py-12 text-slate-400 bg-slate-50 rounded-xl border border-dashed border-slate-200">
                                <Search className="w-8 h-8 mx-auto mb-2 opacity-20" />
                                <p className="text-sm font-medium">Belum ada subdimensi P5 terdaftar untuk kelas ini.</p>
                                <p className="text-xs">Atur di tab Kokurikuler (P5) Manajemen.</p>
                              </div>
                            )}
                          </div>
                        );
                      })()}
                    </div>
                  ) : (
                    <div className="bg-white p-12 rounded-2xl border border-dashed border-slate-200 text-center space-y-3">
                      <div className="w-16 h-16 bg-slate-50 rounded-2xl flex items-center justify-center mx-auto text-slate-300">
                        <Award className="w-8 h-8" />
                      </div>
                      <div className="max-w-xs mx-auto">
                        <h4 className="font-bold text-slate-700">Pilih Siswa Terlebih Dahulu</h4>
                        <p className="text-xs text-slate-400 leading-relaxed">Silakan pilih kelompok kelas dan nama siswa untuk mulai memberikan penilaian Projek P5.</p>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* CORE INPUT PANEL: 3. CATATAN ANAK DENGAN AI CO-WRITER */}
              {nilaiSubTab === "catatan" && (
                <div className="space-y-6">
                  <div className="bg-indigo-50/40 border border-indigo-100 rounded-xl p-4 flex flex-col md:flex-row gap-4 items-start font-sans">
                    <div className="w-10 h-10 rounded-xl bg-indigo-100 flex items-center justify-center text-lg flex-shrink-0">
                      💡
                    </div>
                    <div>
                      <h4 className="font-bold text-slate-850 text-sm font-display">Asisten AI Guru PAUD Terintegrasi</h4>
                      <p className="text-xs text-slate-600 leading-relaxed mt-0.5">
                        Menyusun laporan deskripsi kualitatif raport PAUD membutuhkan dedikasi luar biasa. Sekarang Anda dapat mengetik instruksi pendek atau membiarkan **AI Asisten** kami membaca capaian nilai (Berkembang, Cakap, Mahir) anak dan menerjemahkannya menjadi 2 paragraf narasi lengkap yang bersahabat dan positif bagi orang tua secara instan!
                      </p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                    {currentUserProfile?.role === "guru" ? (
                      <div className="bg-indigo-50/50 p-4 rounded-xl border border-indigo-100 shadow-sm flex flex-col gap-1.5 justify-center">
                        <span className="text-[10px] font-bold text-indigo-500 uppercase tracking-wider">Kelas yang Diampu</span>
                        <div className="text-sm font-extrabold text-indigo-900 flex items-center gap-1.5">
                          <GraduationCap className="w-4 h-4 text-indigo-600" />
                          {state.kelas.find(k => k.id === selectedKelasFilter)?.namaKelas || "KB KUSUMA"}
                        </div>
                      </div>
                    ) : (
                      <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col gap-2">
                        <span className="text-xs font-semibold text-slate-600">Langkah 1: Pilih Kelompok Kelas</span>
                        <select
                          value={selectedKelasFilter}
                          onChange={(e) => {
                            setSelectedKelasFilter(e.target.value);
                            setSelectedSiswaIdCatatan(""); // Reset student
                          }}
                          className="text-xs border border-slate-200 px-3 py-2 rounded-lg bg-white font-medium focus:outline-indigo-600"
                        >
                          <option value="">-- Pilih Kelas --</option>
                          {state.kelas.map((k, idx) => (
                            <option key={k.id || `k-opt9-${idx}`} value={k.id}>{k.namaKelas}</option>
                          ))}
                        </select>
                      </div>
                    )}

                    <div className="bg-white p-4 rounded-xl border border-slate-200 shadow-sm flex flex-col gap-2">
                      <span className="text-xs font-semibold text-slate-600">Langkah 2: Pilih Nama Siswa</span>
                      <select
                        value={selectedSiswaIdCatatan}
                        onChange={(e) => setSelectedSiswaIdCatatan(e.target.value)}
                        className="text-xs border border-slate-200 px-3 py-2 rounded-lg bg-white font-medium focus:outline-indigo-600"
                        disabled={!selectedKelasFilter}
                      >
                        <option value="">-- Pilih Siswa --</option>
                        {state.siswa.filter(s => s.idKelas === selectedKelasFilter).map((s, idx) => (
                          <option key={s.id || `s-opt-c-${idx}`} value={s.id}>{s.namaSiswa}</option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {/* ACTIVE STUDENT EVAL LIST */}
                  {selectedSiswaIdCatatan ? (
                    <div className="space-y-4">
                      {(() => {
                        const s = state.siswa.find(siswa => siswa.id === selectedSiswaIdCatatan);
                        if (!s) return null;
                        const noteText = state.catatanAnak.find(c => c.idSiswa === s.id)?.catatan || "";
                        const matchesGenerating = generatingAi === s.id;
                        return (
                          <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6 font-sans">
                            <div className="flex justify-between items-center border-b border-slate-100 pb-3">
                              <div>
                                <h4 className="text-2xl font-black text-indigo-950 font-display">{s.namaSiswa}</h4>
                                <span className="text-xs text-slate-500 block mt-1.5 uppercase tracking-widest font-bold">Narasi Laporan Kualitatif (Evaluasi Deskriptif)</span>
                              </div>
                            </div>

                              <div className="space-y-4">
                                <div className="flex justify-between items-center bg-slate-50/50 p-2 rounded-lg border border-slate-100">
                                  <label className="block text-base font-black text-slate-800">Narasi Perkembangan Ananda</label>
                                  <span className="text-xs text-indigo-600 font-bold italic tracking-wide">Tulis narasi deskriptif atau gunakan asisten AI</span>
                                </div>
                                <div className="rich-text-editor-container bg-white rounded-2xl border border-slate-300 overflow-hidden shadow-md min-h-[300px]">
                                  <ReactQuill
                                    theme="snow"
                                    value={noteText}
                                    onChange={(content) => handleSaveCatatan(s.id, content)}
                                    placeholder="Ketik deskripsi perkembangan ananda di sini secara bebas..."
                                    modules={{
                                      toolbar: [
                                        [{ 'header': [1, 2, 3, false] }],
                                        ['bold', 'italic', 'underline'],
                                        [{ 'list': 'ordered'}, { 'list': 'bullet' }],
                                        ['clean']
                                      ],
                                    }}
                                    className="quill-editor"
                                    style={{ height: '250px' }}
                                  />
                                </div>
                              </div>

                            {/* AI ACTION PANEL FOR THIS STUDENT */}
                            <div className="p-5 bg-gradient-to-br from-indigo-50 to-white border border-indigo-100 rounded-2xl space-y-4 shadow-sm">
                              <div className="flex items-center gap-2">
                                <div className="p-1.5 bg-indigo-600 rounded-lg text-white">
                                  <Sparkles className="w-4 h-4" />
                                </div>
                                <div>
                                  <span className="text-xs font-bold text-slate-800 tracking-tight">Generate Narasi Raport dengan AI</span>
                                  <p className="text-[10px] text-slate-500">AI akan merangkum semua nilai TP dan P5 menjadi paragraf yang rapi.</p>
                                </div>
                              </div>
                              
                              <div className="space-y-2">
                                <span className="text-[10px] font-bold text-gray-500 tracking-wider block uppercase">Petunjuk Tambahan (Opsional)</span>
                                <input
                                  type="text"
                                  placeholder="Contoh: Tekankan pada perkembangan sosialnya yang luar biasa..."
                                  value={aiPromptCustom}
                                  onChange={(e) => setAiPromptCustom(e.target.value)}
                                  className="w-full bg-white text-xs border border-slate-200 px-4 py-2.5 rounded-xl focus:outline-indigo-600 placeholder:text-neutral-300 font-medium"
                                />
                              </div>

                              <button
                                type="button"
                                disabled={matchesGenerating}
                                onClick={() => composeAiText(s.id)}
                                className="w-full bg-indigo-600 hover:bg-slate-900 text-white py-3 rounded-xl text-xs font-bold flex items-center justify-center gap-2 shadow-lg shadow-indigo-100 transition-all font-display active:scale-95 disabled:opacity-50"
                              >
                                {matchesGenerating ? (
                                  <>
                                    <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
                                    Memproses Narasi...
                                  </>
                                ) : (
                                  <>
                                    <Zap className="w-4 h-4 text-amber-300 fill-amber-300" />
                                    Susun Otomatis dengan AI Assistant
                                  </>
                                )}
                              </button>
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  ) : (
                    <div className="bg-white p-12 rounded-2xl border border-dashed border-slate-200 text-center space-y-3">
                      <div className="w-16 h-16 bg-slate-50 rounded-2xl flex items-center justify-center mx-auto text-slate-300">
                        <BookOpen className="w-8 h-8" />
                      </div>
                      <div className="max-w-xs mx-auto">
                        <h4 className="font-bold text-slate-700">Pilih Siswa Terlebih Dahulu</h4>
                        <p className="text-xs text-slate-400 leading-relaxed">Silakan pilih kelompok kelas dan nama siswa untuk mulai menyusun narasi raport.</p>
                      </div>
                    </div>
                  )}
                </div>
              )}

              {/* CORE INPUT PANEL: 4. KEHADIRAN (ATTENDANCE) */}
              {nilaiSubTab === "kehadiran" && (
                <div className="space-y-6">
                  <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4">
                    <div>
                      <h3 className="font-bold text-slate-800 text-sm font-display">4. Lembar Pencatatan Kehadiran (Absensi)</h3>
                      <p className="text-xs text-neutral-400 font-sans">Isi jumlah hari ketidakhadiran murid dalam satu semester (Sakit, Ijin, Tanpa Keterangan).</p>
                    </div>

                    {currentUserProfile?.role === "guru" ? (
                      <div className="bg-indigo-50 border border-indigo-150 px-3 py-2 rounded-lg text-indigo-700 font-semibold flex items-center gap-1.5 shadow-sm">
                        <GraduationCap className="w-3.5 h-3.5" />
                        Kelas Anda: {state.kelas.find(k => k.id === selectedKelasFilterKehadiran)?.namaKelas || "KB KUSUMA"}
                      </div>
                    ) : (
                      <div className="bg-white px-3 py-2 rounded-xl border border-slate-200 shadow-sm flex items-center gap-2">
                        <span className="text-xs font-bold text-slate-600 whitespace-nowrap">Pilih Kelas:</span>
                        <select
                          value={selectedKelasFilterKehadiran}
                          onChange={(e) => {
                            setSelectedKelasFilterKehadiran(e.target.value);
                          }}
                          className="text-xs border border-slate-250 px-2 py-1.5 rounded-lg bg-white font-medium focus:outline-indigo-600 cursor-pointer"
                        >
                          <option value="">-- Pilih Kelas --</option>
                          {state.kelas.map((k, idx) => (
                            <option key={k.id || `k-opt10-${idx}`} value={k.id}>{k.namaKelas}</option>
                          ))}
                        </select>
                      </div>
                    )}
                  </div>

                  {selectedKelasFilterKehadiran ? (
                    <div className="space-y-4">
                      {(() => {
                        const classStudents = state.siswa.filter(s => s.idKelas === selectedKelasFilterKehadiran);
                        const kObj = state.kelas.find(k => k.id === selectedKelasFilterKehadiran);

                        if (classStudents.length === 0) {
                          return (
                            <div className="bg-white p-12 rounded-2xl border border-dashed border-slate-200 text-center space-y-3">
                              <div className="w-16 h-16 bg-slate-50 rounded-2xl flex items-center justify-center mx-auto text-slate-300">
                                <Users className="w-8 h-8 text-slate-300" />
                              </div>
                              <div className="max-w-xs mx-auto">
                                <h4 className="font-bold text-slate-700">Belum Ada Siswa</h4>
                                <p className="text-xs text-slate-400 leading-relaxed">Kelas {kObj?.namaKelas || ""} belum memiliki data siswa. Silakan tambahkan siswa terlebih dahulu di tab Master Data.</p>
                              </div>
                            </div>
                          );
                        }

                        return (
                          <div className="bg-white rounded-xl border border-slate-200/80 shadow-sm overflow-hidden font-sans">
                            {/* Header Panel */}
                            <div className="bg-slate-50 border-b border-slate-200 px-6 py-4 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                              <div>
                                <span className="text-[9px] font-bold text-indigo-600 uppercase tracking-widest block font-sans">INPUT REKAP ABSENSI</span>
                                <h4 className="text-sm font-bold text-slate-800">Kelas: {kObj?.namaKelas} (Wali Kelas: {kObj?.waliKelas || "-"})</h4>
                              </div>
                              <div className="text-xs text-slate-400 italic flex items-center gap-1 bg-white px-2.5 py-1 rounded-lg border border-slate-250/65">
                                <ShieldAlert className="w-3.5 h-3.5 text-indigo-500 opacity-80" />
                                <span>Perubahan tersimpan otomatis ke database.</span>
                              </div>
                            </div>

                            {/* Table Headers inside desktop list */}
                            <div className="hidden md:grid grid-cols-12 gap-4 px-6 py-3 bg-slate-50/50 border-b border-slate-100 text-xs font-bold text-slate-500 uppercase tracking-wider text-center">
                              <div className="col-span-1">No</div>
                              <div className="col-span-5 text-left">Nama Siswa</div>
                              <div className="col-span-2 text-center">Sakit (Hari)</div>
                              <div className="col-span-2 text-center">Ijin (Hari)</div>
                              <div className="col-span-2 text-center">Tanpa Keterangan (Hari)</div>
                            </div>

                            {/* Student Row Loop */}
                            <div className="divide-y divide-slate-100">
                              {classStudents.map((s, index) => {
                                const att = state.kehadiran.find(k => k.idSiswa === s.id) || { sakit: 0, ijin: 0, tanpaKet: 0 };
                                const currentSakit = att.sakit || 0;
                                const currentIjin = att.ijin || 0;
                                const currentTanpaKet = att.tanpaKet || 0;

                                return (
                                  <div key={s.id} className="grid grid-cols-12 gap-4 items-center px-6 py-4 hover:bg-slate-50/30 transition-colors">
                                    {/* Number & Name columns */}
                                    <div className="col-span-1 hidden md:flex items-center justify-center font-bold text-xs text-slate-400">
                                      {index + 1}
                                    </div>
                                    <div className="col-span-12 md:col-span-5 flex items-center gap-3">
                                      <div className="md:hidden w-6 h-6 rounded-full bg-slate-100 flex items-center justify-center text-[10px] font-bold text-slate-500">
                                        {index + 1}
                                      </div>
                                      <div>
                                        <h5 className="font-extrabold text-slate-900 text-base leading-snug">{s.namaSiswa}</h5>
                                        <span className="text-xs text-slate-500 block font-bold mt-1 tracking-wide">
                                          {s.nisn ? `NISN: ${s.nisn}` : "Tanpa NISN"} • {s.jenisKelamin === "L" ? "LAKI-LAKI" : "PEREMPUAN"}
                                        </span>
                                      </div>
                                    </div>

                                    {/* Input numeric fields on the right, neatly grouped */}
                                    <div className="col-span-12 md:contents grid grid-cols-3 gap-3 mt-2 md:mt-0">
                                      {/* Inputs Column: Sakit */}
                                      <div className="md:col-span-2 flex flex-col items-center justify-center gap-1.5 bg-amber-50/25 md:bg-transparent p-2 md:p-0 rounded-xl border border-amber-100/40 md:border-0">
                                        <span className="md:hidden text-[9px] font-extrabold text-amber-700 uppercase tracking-wider block">Sakit</span>
                                        <input
                                          type="number"
                                          min="0"
                                          value={currentSakit}
                                          onChange={(e) => handleUpdateKehadiran(s.id, "sakit", Math.max(0, Number(e.target.value)))}
                                          className="w-16 h-8 text-center font-bold text-sm border border-slate-250 rounded-lg focus:outline-indigo-600 focus:ring-1 focus:ring-indigo-600 bg-white shadow-sm"
                                        />
                                      </div>

                                      {/* Inputs Column: Ijin */}
                                      <div className="md:col-span-2 flex flex-col items-center justify-center gap-1.5 bg-sky-50/25 md:bg-transparent p-2 md:p-0 rounded-xl border border-sky-100/40 md:border-0">
                                        <span className="md:hidden text-[9px] font-extrabold text-sky-700 uppercase tracking-wider block">Ijin</span>
                                        <input
                                          type="number"
                                          min="0"
                                          value={currentIjin}
                                          onChange={(e) => handleUpdateKehadiran(s.id, "ijin", Math.max(0, Number(e.target.value)))}
                                          className="w-16 h-8 text-center font-bold text-sm border border-slate-250 rounded-lg focus:outline-indigo-600 focus:ring-1 focus:ring-indigo-600 bg-white shadow-sm"
                                        />
                                      </div>

                                      {/* Inputs Column: Tanpa Keterangan */}
                                      <div className="md:col-span-2 flex flex-col items-center justify-center gap-1.5 bg-rose-50/25 md:bg-transparent p-2 md:p-0 rounded-xl border border-rose-100/40 md:border-0">
                                        <span className="md:hidden text-[9px] font-extrabold text-rose-700 uppercase tracking-wider block">Alpa</span>
                                        <input
                                          type="number"
                                          min="0"
                                          value={currentTanpaKet}
                                          onChange={(e) => handleUpdateKehadiran(s.id, "tanpaKet", Math.max(0, Number(e.target.value)))}
                                          className="w-16 h-8 text-center font-bold text-sm border border-slate-250 rounded-lg focus:outline-indigo-600 focus:ring-1 focus:ring-indigo-600 bg-white shadow-sm"
                                        />
                                      </div>
                                    </div>

                                  </div>
                                );
                              })}
                            </div>
                          </div>
                        );
                      })()}
                    </div>
                  ) : (
                    <div className="bg-white p-12 rounded-2xl border border-dashed border-slate-200 text-center space-y-3">
                      <div className="w-16 h-16 bg-slate-50 rounded-2xl flex items-center justify-center mx-auto text-slate-300">
                        <UserCheck className="w-8 h-8 text-indigo-500/80" />
                      </div>
                      <div className="max-w-xs mx-auto">
                        <h4 className="font-bold text-slate-700">Pilih Kelas Terlebih Dahulu</h4>
                        <p className="text-xs text-slate-400 leading-relaxed">Silakan pilih kelompok kelas di atas untuk mulai mencatat kehadiran seluruh siswa secara terpadu.</p>
                      </div>
                    </div>
                  )}
                </div>
              )}

            </div>
          )}

          {/* TAB G: CETAK & PREVIEW RAPORT */}
          {activeTab === "cetak" && (
            <div className="space-y-6">
              
              {/* FILTERS & METADATA SELECTION PANEL - Hidden during printing */}
              <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4 no-print animate-fade-in">
                <div>
                  <h3 className="font-bold text-slate-900 text-lg font-display">Cetak Dokumen Raport PAUD</h3>
                  <p className="text-sm font-medium text-neutral-500 mt-1">Pilih siswa tujuan demi memuat preview lengkap secara utuh. Cetak langsung ke Printer Anda atau simpan sebagai dokumen digital PDF.</p>
                </div>

                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                   <div>
                    <label className="block text-base font-black text-slate-700 mb-1.5">Langkah 1: Pilih Kelompok Kelas</label>
                    <select
                      value={printKelasId}
                      onChange={(e) => {
                        setPrintKelasId(e.target.value);
                        const listSiswa = state.siswa.filter(s => s.idKelas === e.target.value);
                        if (listSiswa.length > 0) {
                          setPrintSiswaId(listSiswa[0].id);
                        } else {
                          setPrintSiswaId("");
                        }
                      }}
                      className="w-full text-base border border-slate-300 px-4 py-2.5 rounded-xl bg-white font-semibold focus:outline-indigo-600 shadow-sm"
                    >
                      <option value="">-- Pilih Kelas --</option>
                      {state.kelas.map((k, idx) => (
                        <option key={k.id || `k-opt11-${idx}`} value={k.id}>{k.namaKelas}</option>
                      ))}
                    </select>
                  </div>

                   <div>
                    <label className="block text-base font-black text-slate-700 mb-1.5">Langkah 2: Pilih Nama Siswa</label>
                    <select
                      value={printSiswaId}
                      onChange={(e) => setPrintSiswaId(e.target.value)}
                      className="w-full text-base border border-slate-300 px-4 py-2.5 rounded-xl bg-white font-semibold focus:outline-indigo-600 shadow-sm"
                      disabled={!printKelasId}
                    >
                      <option value="">-- Hubungkan Siswa --</option>
                      {state.siswa.filter(s => s.idKelas === printKelasId).map((s, idx) => (
                        <option key={s.id || `s-opt-${idx}`} value={s.id}>{s.namaSiswa}</option>
                      ))}
                    </select>
                  </div>
                </div>

                {printSiswaId ? (
                  <div className="flex gap-2 justify-end pt-2 border-t border-slate-100">
                    <button
                      onClick={() => window.print()}
                      className="bg-indigo-600 hover:bg-indigo-700 text-white px-5 py-2 rounded-lg text-xs font-bold flex items-center justify-center gap-1.5 shadow"
                    >
                      <Printer className="w-4 h-4" /> Cetak Raport Anak (Printer / PDF)
                    </button>
                  </div>
                ) : (
                  <div className="text-xs text-amber-600 bg-amber-50 p-2.5 rounded-lg border border-amber-100 font-medium">
                    ⓘ Pilih kelas dan identitas nama siswa untuk memunculkan preview lembar raport siap cetak.
                  </div>
                )}
              </div>

              {/* HIGH-FIDELITY PRINT PREVIEW DOCK - Styled with specific A4 letter bounds */}
              {printSiswaId && printSiswa ? (
                (() => {
                  const printCityRaw = (state.dataSekolah?.alamat || "").match(/(Demak|Semarang|Jepara|Kudus|Mijen)/i)?.[0] || "Demak";
                  const printCity = printCityRaw.charAt(0).toUpperCase() + printCityRaw.slice(1);
                  const activeTps = state.tujuanPembelajaran.filter(tp => tp.idKelas === printSiswa.idKelas);

                  // Reuse common header metadata block for pages 3, 4, 5, 6, and 7
                  const renderPageHeader = (pageTitle: string, subTitle?: string) => (
                    <div className="border-b-2 border-slate-900 pb-4 mb-5">
                      <div className="text-center space-y-0.5 mb-4">
                        <h2 className="text-[15px] font-black text-slate-950 uppercase tracking-widest font-sans">
                          {pageTitle}
                        </h2>
                        {subTitle && (
                          <p className="text-[12px] font-extrabold text-slate-500 uppercase tracking-wider font-mono">
                            {subTitle}
                          </p>
                        )}
                      </div>

                      <div className="grid grid-cols-2 gap-x-8 gap-y-1.5 text-[13px] leading-relaxed text-slate-900 font-medium">
                        <div className="space-y-0.5">
                          <table className="w-full">
                            <tbody>
                              <tr>
                                <td className="w-28 text-slate-500 font-bold">Nama Sekolah</td>
                                <td className="text-slate-950 font-bold">: {state.dataSekolah.namaSekolah || "KB KUSUMA"}</td>
                              </tr>
                              <tr>
                                <td className="text-slate-500 font-bold">Nama Peserta Didik</td>
                                <td className="text-slate-950 font-bold uppercase">: {printSiswa.namaSiswa}</td>
                              </tr>
                              <tr>
                                <td className="text-slate-500 font-bold">NIK / NISN</td>
                                <td className="text-slate-950 font-bold font-mono">: {printSiswa.nisn || "-"}</td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                        <div className="space-y-0.5">
                          <table className="w-full">
                            <tbody>
                              <tr>
                                <td className="w-28 text-slate-500 font-bold">Kelas</td>
                                <td className="text-slate-950 font-bold">: {printKelasItem?.namaKelas || "AL FARABI"}</td>
                              </tr>
                              <tr>
                                <td className="text-slate-500 font-bold">Sem / Th. Ajaran</td>
                                <td className="text-slate-950 font-semibold">: {state.dataSekolah.semester} / {state.dataSekolah.thAjaran}</td>
                              </tr>
                              <tr>
                                <td className="text-slate-500 font-bold">Tinggi / Berat Badan</td>
                                <td className="text-slate-950 font-semibold">: {printSiswa.tb || "-"} cm / {printSiswa.bb || "-"} kg</td>
                              </tr>
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  );

                  return (
                    <div className="space-y-12 no-print-gap print:space-y-0 text-slate-950 font-arial">
                      
                      {/* PAGE 1: COVER PAGE */}
                      <div className="print-page-break bg-white border border-slate-200 rounded-2xl shadow-lg p-12 max-w-3xl mx-auto flex flex-col items-center min-h-[960px] print:min-h-0 print:h-[273mm] print:w-full print:border-none print:shadow-none print:p-6 print:m-0 animate-fade-in font-sans relative text-center">
                        {/* Outer Frame - Elegant thin border */}
                        <div className="absolute inset-4 border border-slate-400 rounded-2xl pointer-events-none p-1 print:inset-4 print:border-slate-800 print:border-[3px] print:rounded-2xl"></div>

                        {/* Top Header Logo */}
                        <div className="z-10 mt-16 mb-8 flex flex-col items-center">
                          <div className="w-44 h-44 flex items-center justify-center">
                             {state.dataSekolah.logo ? (
                               <img src={state.dataSekolah.logo} alt="Logo" className="max-w-full max-h-full object-contain" />
                             ) : (
                               <div className="w-32 h-32 rounded-full border-2 border-slate-200 flex items-center justify-center text-slate-400 font-bold">LOGO</div>
                             )}
                          </div>
                        </div>

                        {/* Main Title Block */}
                        <div className="z-10 space-y-3 text-center my-10">
                          <h1 className="text-[26px] font-black text-slate-900 uppercase tracking-wider font-display">
                            LAPORAN HASIL
                          </h1>
                          <div className="space-y-1">
                            <h2 className="text-[16px] font-bold text-slate-800 uppercase tracking-wide px-4">
                              CAPAIAN PERKEMBANGAN PESERTA DIDIK
                            </h2>
                            <h2 className="text-[16px] font-bold text-slate-800 uppercase tracking-wide px-4">
                              PENDIDIKAN ANAK USIA DINI (PAUD)
                            </h2>
                          </div>
                        </div>

                        {/* Named Student Section */}
                        <div className="z-10 w-full flex flex-col items-center mt-10 mb-8 px-10">
                          <span className="text-[11px] font-bold text-slate-500 tracking-[0.15em] uppercase block mb-3">NAMA PESERTA DIDIK</span>
                          <div className="w-full border border-slate-750 rounded-[14px] py-3 px-6 bg-white shadow-sm flex items-center justify-center">
                            <h3 className="text-[22px] font-black text-slate-950 tracking-wider uppercase leading-tight font-display text-center">
                              {printSiswa.namaSiswa}
                            </h3>
                          </div>
                        </div>

                        <div className="z-10 flex flex-col items-center mt-12">
                          <span className="text-[11px] font-bold text-slate-500 tracking-[0.15em] uppercase block mb-2">NISN</span>
                          <div className="text-[22px] font-black text-slate-900 tracking-[0.1em]">
                            {printSiswa.nisn || "----------"}
                          </div>
                        </div>

                        {/* Bottom Metadata Block */}
                        <div className="z-10 mt-auto w-[100%] flex flex-col items-center pb-10">
                          <div className="w-full border-t border-slate-300 mb-8 opacity-60"></div>
                          <div className="space-y-1 text-center">
                            <h4 className="text-[20px] font-black text-slate-950 uppercase tracking-widest font-display">
                              {state.dataSekolah.namaSekolah || "KB KUSUMA"}
                            </h4>
                            {state.dataSekolah.npsn && (
                              <p className="text-[18px] font-bold text-slate-900 tracking-wide">NPSN: {state.dataSekolah.npsn}</p>
                            )}
                            <p className="text-[12px] text-slate-700 font-bold leading-relaxed uppercase max-w-[420px] mx-auto whitespace-pre-wrap mt-3">
                              {state.dataSekolah.alamat || "Alamat Lembaga Utama belum didokumentasikan."}
                            </p>
                          </div>
                        </div>
                      </div>

                      {/* PAGE 2: BIODATA PAGE (KETERANGAN DIRI) */}
                      <div className="print-page-break print-container bg-white border border-slate-200 rounded-2xl shadow-lg p-12 max-w-3xl mx-auto flex flex-col justify-between min-h-[960px] print:min-h-0 print:h-[273mm] print:w-full print:border-none print:shadow-none print:p-6 print:m-0 animate-fade-in font-sans relative">
                        {/* Outer Border */}
                        <div className="absolute inset-4 border border-slate-200 rounded-xl pointer-events-none p-1 print:inset-4 print:border-slate-800 print:border-[3px] print:rounded-2xl"></div>

                        <div className="z-10 space-y-6">
                          <div className="text-center space-y-2 pb-2 border-b border-slate-100">
                            <h2 className="text-md font-black text-slate-950 uppercase tracking-widest block font-display">
                              Keterangan Diri Anak Didik
                            </h2>
                          </div>

                          {/* Data Table */}
                          <div className="overflow-x-auto">
                            <table className="w-full text-xs font-medium text-slate-900 leading-relaxed border-collapse">
                              <tbody className="divide-y divide-slate-100">
                                <tr>
                                  <td className="py-2 w-6 text-slate-450 font-bold">1.</td>
                                  <td className="py-2 w-48 text-slate-600 font-bold">Nama Lengkap Siswa</td>
                                  <td className="py-2 w-4 text-slate-400 font-medium">:</td>
                                  <td className="py-2 font-black text-slate-950 uppercase">{printSiswa.namaSiswa}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">2.</td>
                                  <td className="py-2 text-slate-600 font-bold">NISN</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 font-bold text-slate-950 font-mono">{printSiswa.nisn || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">3.</td>
                                  <td className="py-2 text-slate-600 font-bold">Tempat, Tanggal Lahir</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-900 font-semibold">
                                    {(printSiswa.tempatLahir || "-").toUpperCase()}, {formatIndonesianDate(printSiswa.tglLahir, "-")}
                                  </td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">4.</td>
                                  <td className="py-2 text-slate-600 font-bold">Jenis Kelamin</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-850 font-semibold">{printSiswa.jenisKelamin || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">5.</td>
                                  <td className="py-2 text-slate-600 font-bold">Agama</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-850 font-semibold uppercase">{printSiswa.agama || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">6.</td>
                                  <td className="py-2 text-slate-600 font-bold">Anak Ke-</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-950 font-bold">{printSiswa.anakKe || "-"}</td>
                                </tr>

                                <tr className="bg-slate-50/50 font-bold uppercase text-[9px] tracking-widest text-sky-800">
                                  <td colSpan={4} className="py-1.5 px-3 font-black">ORANG TUA / WALI</td>
                                </tr>

                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">7.</td>
                                  <td className="py-2 text-slate-600 font-bold">Nama Ayah Kandung</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-850 font-bold uppercase">{printSiswa.namaAyah || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">8.</td>
                                  <td className="py-2 text-slate-600 font-bold">Pekerjaan Ayah</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-800">{printSiswa.pekerjaanAyah || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">9.</td>
                                  <td className="py-2 text-slate-600 font-bold">Nama Ibu Kandung</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-850 font-bold uppercase">{printSiswa.namaIbu || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">10.</td>
                                  <td className="py-2 text-slate-600 font-bold">Pekerjaan Ibu</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-800">{printSiswa.pekerjaanIbu || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">11.</td>
                                  <td className="py-2 text-slate-600 font-bold">No. Telepon / HP Wali</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 font-bold text-slate-900 font-mono">{printSiswa.noHp || "-"}</td>
                                </tr>
                                <tr>
                                  <td className="py-2 text-slate-450 font-bold">12.</td>
                                  <td className="py-2 text-slate-600 font-bold">Alamat Rumah Lengkap</td>
                                  <td className="py-2 text-slate-400 font-medium">:</td>
                                  <td className="py-2 text-slate-800 leading-normal text-justify">{printSiswa.alamat || "-"}</td>
                                </tr>
                              </tbody>
                            </table>
                          </div>

                          {/* Sign Block & Picture Placement Area - Positioned closely inside the flow */}
                          <div className="z-10 mt-10 flex justify-between items-start px-2">
                            {/* Photo Placeholder */}
                            <div className="w-[100px] h-[140px] border-2 border-dashed border-slate-400 bg-white flex flex-col items-center justify-center text-slate-400 select-none">
                              <span className="text-[11px] font-medium text-slate-400 mb-1">Tempel Foto</span>
                              <span className="text-xs font-bold text-slate-500">3 &times; 4</span>
                            </div>

                            {/* Signature stamp mock */}
                            <div className="text-center flex flex-col items-center justify-between h-[150px] w-64 pr-2">
                              <div className="space-y-1">
                                <p className="text-sm text-slate-900 font-medium">
                                  {printCity}, {formatIndonesianDate(state.dataSekolah.tglRaport)}
                                </p>
                                <p className="text-sm font-bold text-slate-900">
                                  Kepala Sekolah
                                </p>
                              </div>
                              <div className="w-full">
                                <strong className="text-sm font-black text-slate-950 underline decoration-slate-900 underline-offset-4 block leading-none pb-1">
                                  {cleanName(state.dataSekolah.kepalaSekolah || "MIRAH TITIMIRANTI, S.P., S.Pd.")}
                                </strong>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>

                           {/* INTRAKURIKULER SECTION: CONTINUOUS FLOW */}
                      <div className="print-page-break print-container bg-white border border-slate-200 rounded-2xl shadow-lg p-8 max-w-3xl mx-auto block min-h-[960px] print:min-h-0 print:h-auto print:w-full print:border-none print:shadow-none print:px-8 print:py-6 print:m-0 animate-fade-in relative text-slate-950 print-arial-large">
                        <div className="absolute inset-4 border border-slate-150 rounded-xl pointer-events-none print:hidden"></div>

                        <div className="z-10 space-y-6">
                            <div>
                                {renderPageHeader("Laporan Hasil Perkembangan Peserta Didik", "Pendidikan Anak Usia Dini (PAUD)")}

                                <div className="space-y-6">
                                    {state.kategoriIntrakurikuler.map((kat, katIdx) => {
                                        const catTps = activeTps.filter(tp => tp.idKategori === kat.id);
                                        const catAssessment = state.nilaiIntrakurikuler.find(n => n.idSiswa === printSiswa.id && n.idTp === kat.id);
                                        const categoryDescription = catAssessment?.deskripsi || "";

                                        return (
                                            <div key={kat.id} className="space-y-3">
                                                <div className="border-l-4 border-indigo-600 pl-3 py-1 bg-slate-50 rounded-r-md keep-together">
                                                    <span className="text-xs font-black uppercase text-indigo-700 tracking-widest block leading-none mb-1">
                                                        I. Capaian Pembelajaran
                                                    </span>
                                                    <h3 className="text-sm font-black text-slate-950 uppercase tracking-wide leading-tight">
                                                        {katIdx + 1}. {kat.namaKategori}
                                                    </h3>
                                                </div>

                                                <div className="overflow-hidden border border-slate-950 rounded-lg shadow-sm bg-white">
                                                    <table className="w-full text-left text-[14px] border-collapse leading-normal font-arial">
                                                        <thead className="print:table-header-group">
                                                            <tr className="bg-slate-100 text-slate-950 font-black uppercase border-b border-slate-950 text-xs tracking-wide">
                                                                <th className="px-4 py-3 border-r border-slate-950 text-center w-[45%] font-black">TUJUAN PEMBELAJARAN</th>
                                                                <th className="px-4 py-3 border-r border-slate-950 text-center w-[35%] font-black">AKTIVITAS</th>
                                                                <th className="px-2 py-3 text-center w-[20%] font-black leading-tight text-[11px]">DIMENSI<br />KEMANDIRIAN</th>
                                                            </tr>
                                                        </thead>
                                                        <tbody className="divide-y divide-slate-400 bg-white print:table-row-group">
                                                            {catTps.length > 0 ? (
                                                                catTps.map((tp, idx) => {
                                                                    const assessment = state.nilaiIntrakurikuler.find(n => n.idSiswa === printSiswa.id && n.idTp === tp.id);
                                                                    const score = assessment?.nilai || "Cakap";

                                                                    const sUpper = (score || "").toUpperCase();
                                                                    let ratingStyle = "";
                                                                    if (sUpper === "MAHIR") {
                                                                        ratingStyle = "bg-[#C6F6D5] border-[#4ADE80] text-slate-950";
                                                                    } else if (sUpper === "CAKAP") {
                                                                        ratingStyle = "bg-[#BAE6FD] border-[#38BDF8] text-slate-950";
                                                                    } else { // "BERKEMBANG"
                                                                        ratingStyle = "bg-[#FEF08A] border-[#FACC15] text-slate-950";
                                                                    }

                                                                    return (
                                                                        <tr key={tp.id} className="align-top keep-together">
                                                                            <td className="px-4 py-3 border-r border-slate-950 font-bold text-slate-950 leading-relaxed text-justify text-[13px] w-[45%]">
                                                                                {tp.deskripsi}
                                                                            </td>
                                                                            <td className="px-4 py-3 border-r border-slate-950 italic text-slate-800 leading-relaxed text-justify font-medium text-[13px] w-[35%]">
                                                                                {tp.aktivitasMetode || "-"}
                                                                            </td>
                                                                            <td className="px-2 py-3 text-center align-middle w-[20%]">
                                                                                <span className={`px-3 py-1.5 rounded-full border text-[10px] font-black uppercase tracking-wider ${ratingStyle} inline-block leading-none shadow-xs`}>
                                                                                    {score}
                                                                                </span>
                                                                            </td>
                                                                        </tr>
                                                                    );
                                                                })
                                                            ) : (
                                                                <tr>
                                                                    <td colSpan={3} className="px-4 py-6 text-center text-slate-400 font-bold italic">
                                                                        Belum ada indikator yang diinput.
                                                                    </td>
                                                                </tr>
                                                            )}
                                                        </tbody>
                                                    </table>
                                                </div>

                                                <div className="space-y-1 keep-together">
                                                    <span className="text-xs font-black uppercase text-indigo-700 tracking-wider ml-1">Deskripsi Munculan Narasi:</span>
                                                    <div 
                                                      className="border border-slate-950 rounded-xl p-5 bg-slate-50/50 min-h-[140px] text-base font-semibold text-slate-900 leading-relaxed text-justify shadow-inner narrative-box rich-text-print"
                                                      dangerouslySetInnerHTML={{ __html: cleanHtmlText(categoryDescription) || `Ananda ${printSiswa.namaSiswa} telah menunjukkan penguasaan yang sangat baik pada bidang ${kat.namaKategori}. Ia mampu mengikuti instruksi dengan antusias and menunjukkan kemajuan yang membanggakan dalam setiap sub-pembelajaran.` }}
                                                    />
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>

                                {/* SECTION II: KOKURIKULER (P5 PANEL) */}
                                <div className="space-y-3 pt-6 keep-together">
                                    <div className="border-l-4 border-indigo-650 pl-3 bg-slate-50 py-1.5 rounded-r">
                                        <h3 className="text-sm font-black uppercase text-slate-950 tracking-wider font-display">
                                            II. Capaian Kokurikuler (Projek Penguatan Profil Pelajar Pancasila - P5)
                                        </h3>
                                    </div>

                                    <div className="overflow-hidden border border-slate-950 rounded-lg shadow bg-white">
                                        <table className="w-full text-left text-sm border-collapse leading-normal bg-white">
                                            <thead>
                                                <tr className="bg-slate-100 text-slate-950 font-black uppercase border-b border-slate-950 text-xs tracking-wide">
                                                    <th className="px-3.5 py-3 w-40 border-r border-slate-950 text-center font-black">Dimensi P5</th>
                                                    <th className="px-3.5 py-3 border-r border-slate-950 text-center font-black">Narasi / Deskripsi Capaian Projek</th>
                                                    <th className="px-3.5 py-3 text-center w-28 font-black">Predikat</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-400 text-slate-900 font-semibold">
                                                {state.subdimensiKokurikuler.filter(sub => sub.idKelas === printSiswa.idKelas).length > 0 ? (
                                                    state.subdimensiKokurikuler
                                                        .filter(sub => sub.idKelas === printSiswa.idKelas)
                                                        .map((sub) => {
                                                            const assessment = state.nilaiKokurikuler.find(n => n.idSiswa === printSiswa.id && n.idSubdimensi === sub.id);
                                                            const score = assessment?.nilai || "Cakap";
                                                            const deskripsiAss = assessment?.deskripsi || "";

                                                            const sUpper = (score || "").toUpperCase();
                                                            let badgeStyle = "";
                                                            if (sUpper === "MAHIR") {
                                                                badgeStyle = "bg-[#C6F6D5] border-[#4ADE80] text-slate-950";
                                                            } else if (sUpper === "CAKAP") {
                                                                badgeStyle = "bg-[#BAE6FD] border-[#38BDF8] text-slate-950";
                                                            } else { // "BERKEMBANG"
                                                                badgeStyle = "bg-[#FEF08A] border-[#FACC15] text-slate-950";
                                                            }

                                                            return (
                                                                <tr key={sub.id} className="hover:bg-slate-50/10 align-top">
                                                                    <td className="px-3.5 py-3 border-r border-slate-950 font-bold text-slate-950">
                                                                        {sub.namaSubdimensi.replace(/^Dimensi\s+/i, "")}
                                                                    </td>
                                                                    <td 
                                                                      className="px-3.5 py-3 border-r border-slate-950 leading-relaxed text-slate-800 font-medium text-justify text-[11px] rich-text-print"
                                                                      dangerouslySetInnerHTML={{ __html: cleanHtmlText(deskripsiAss) || `Ananda ${printSiswa.namaSiswa} menunjukkan performa keterlibatan yang positif and berkembang konsisten dalam mewujudkan projek profil pancasila bertemakan sub-elemen ${sub.namaSubdimensi}.` }}
                                                                    />
                                                                    <td className="px-3.5 py-3 text-center font-extrabold">
                                                                        <span className={`px-3 py-1.5 rounded-full border text-[10px] font-black uppercase tracking-wider ${badgeStyle} inline-block leading-none shadow-xs`}>
                                                                            {score}
                                                                        </span>
                                                                    </td>
                                                                </tr>
                                                            );
                                                        })
                                                ) : (
                                                    <tr>
                                                        <td colSpan={3} className="px-4 py-8 text-center text-slate-400 italic font-bold">
                                                            Belum ada dimensi atau nilai projek kokurikuler yang dirangkai untuk kelompok kelas ini.
                                                        </td>
                                                    </tr>
                                                )}
                                            </tbody>
                                        </table>
                                    </div>
                                </div>

                                {/* SECTION III: CATATAN GURU WALI KELAS */}
                                <div className="space-y-2 pt-6 keep-together">
                                    <div className="border-l-4 border-indigo-700 pl-3 bg-slate-50 py-1 rounded-r">
                                        <h3 className="text-sm font-black uppercase text-slate-950 tracking-wider font-display">
                                            III. Catatan Guru Wali Kelas
                                        </h3>
                                    </div>
                                    <div 
                                        className="border border-slate-950 rounded-xl p-5 bg-slate-50/50 min-h-[110px] text-base font-semibold text-slate-900 leading-relaxed text-justify shadow-inner narrative-box rich-text-print"
                                        dangerouslySetInnerHTML={{ __html: cleanHtmlText(printCatatanSiswa) }}
                                    />
                                </div>

                                {/* GROUP SECTION IV & V */}
                                <div className="keep-together pt-6">
                                  {/* SECTION IV: KEHADIRAN (REKAPITULASI ABSENSI SEMESTER) */}
                                  <div className="space-y-2.5">
                                      <div className="border-l-4 border-indigo-700 pl-3 bg-slate-50 py-1 rounded-r">
                                          <h3 className="text-sm font-black uppercase text-slate-950 tracking-wider font-display">
                                              IV. Kehadiran (Rekapitulasi Absensi Semester)
                                          </h3>
                                      </div>
                                    
                                    <div className="max-w-md overflow-hidden border border-slate-950 rounded-lg shadow-sm bg-white">
                                        <table className="w-full text-left text-sm border-collapse bg-white">
                                            <thead>
                                                <tr className="bg-slate-100 text-slate-950 font-black border-b border-slate-950 text-xs tracking-wider uppercase">
                                                    <th className="px-4 py-2 border-r border-slate-950 font-black text-[13px]">Keterangan Harian</th>
                                                    <th className="px-4 py-2 text-center w-32 font-black text-[13px]">Jumlah Akumulasi</th>
                                                </tr>
                                            </thead>
                                            <tbody className="divide-y divide-slate-350 text-slate-900 font-bold bg-white">
                                                <tr>
                                                    <td className="px-3.5 py-1.5 text-slate-700 border-r border-slate-950">1. Sakit</td>
                                                    <td className="px-3.5 py-1.5 text-center text-slate-950">{printAbsensi.sakit} Hari</td>
                                                </tr>
                                                <tr>
                                                    <td className="px-3.5 py-1.5 text-slate-700 border-r border-slate-950">2. Izin</td>
                                                    <td className="px-3.5 py-1.5 text-center text-slate-950">{printAbsensi.ijin} Hari</td>
                                                </tr>
                                                <tr>
                                                    <td className="px-3.5 py-1.5 text-slate-705 border-r border-slate-950">3. Tanpa Keterangan (Alpa)</td>
                                                    <td className="px-3.5 py-1.5 text-center text-slate-950">{printAbsensi.tanpaKet} Hari</td>
                                                </tr>
                                            </tbody>
                                        </table>
                                    </div>
                                  </div>

                                  {/* SECTION V: SIGNATURE BLOCK */}
                                  <div className="pt-6 border-t border-slate-150 grid grid-cols-2 gap-y-8 text-center text-sm text-slate-900 leading-relaxed max-w-2xl mx-auto w-full mt-6">
                                      {/* Top Left: Orang Tua */}
                                    <div className="flex flex-col items-center justify-between">
                                        <p className="mb-14">Mengetahui,<br />Orang Tua/Wali,</p>
                                        <div className="border-b border-slate-900 w-3/4"></div>
                                    </div>

                                    {/* Top Right: Wali Kelas */}
                                    <div className="flex flex-col items-center justify-between">
                                        <p className="mb-14">
                                            {printCity}, {formatIndonesianDate(state.dataSekolah.tglRaport)}<br />
                                            Wali Kelas,
                                        </p>
                                        <div>
                                            <p className="font-bold underline decoration-slate-900 underline-offset-2">
                                                {cleanName(printKelasItem?.waliKelas || "WALIKELAS PAUD")}
                                            </p>
                                            <p>NUPTK: {printKelasItem?.nuptkNgty || "6359752654300053"}</p>
                                        </div>
                                    </div>

                                    {/* Bottom Center: Kepala Sekolah */}
                                    <div className="col-span-2 flex flex-col items-center justify-between mt-2">
                                          <p className="mb-14">Mengetahui,<br />Kepala Sekolah</p>
                                          <p className="font-bold underline decoration-slate-900 underline-offset-2">
                                              {cleanName(state.dataSekolah.kepalaSekolah || "MIRAH TITIMIRANTI, S.P., S.Pd.")}
                                          </p>
                                      </div>
                                  </div>
                                </div>
                            </div>
                        </div>
                      </div>

                    </div>
                  );
                })()
              ) : (
                <div className="bg-white p-12 text-center border rounded-xl shadow-sm text-neutral-400 no-print font-medium">
                  Raport belum siap dipreview. Silakan pilih identitas Kelompok Kelas dan nama Siswa yang sesuai di atas.
                </div>
              )}

            </div>
          )}

          {/* TAB H: PENGATURAN SEKOLAH & BACKUP */}
          {activeTab === "pengaturan" && (
            <div className="space-y-6 animate-fade-in no-print">
              <div>
                <h2 className="text-xl font-bold text-slate-800 font-display">Pengaturan Satuan Pendidikan & Cadangan</h2>
                <p className="text-sm text-neutral-400">Sesuaikan data kop sekolah, tahun ajaran aktif, logo, dsn backup seluruh data lokal.</p>
              </div>

              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* SCHOOL DETAILS FORM */}
                <div className="lg:col-span-2 bg-white rounded-xl border border-slate-200 p-5 shadow-sm space-y-4">
                    <span className="text-sm font-bold text-indigo-900 tracking-wider block uppercase border-b border-slate-100 pb-2">Detail Identitas Sekolah</span>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-x-4 gap-y-3">
                      <div className="md:col-span-2">
                        <label className="block text-xs font-bold text-slate-700 mb-1">Nama Satuan PAUD / TK *</label>
                        <input
                          type="text"
                          value={state.dataSekolah.namaSekolah}
                          onChange={(e) => handleUpdateSchool("namaSekolah", e.target.value)}
                          className="w-full text-sm border border-slate-300 px-3 py-2 rounded-lg font-bold text-indigo-900 focus:ring-1 focus:ring-indigo-600 focus:border-indigo-600 focus:outline-none transition shadow-sm"
                        />
                      </div>

                      <div className="md:col-span-2">
                        <label className="block text-xs font-bold text-slate-700 mb-1">Alamat Lembaga Sekolah Lengkap *</label>
                        <textarea
                          rows={2}
                          value={state.dataSekolah.alamat}
                          onChange={(e) => handleUpdateSchool("alamat", e.target.value)}
                          className="w-full text-sm border border-slate-300 px-3 py-2 rounded-lg font-medium focus:ring-1 focus:ring-indigo-600 focus:border-indigo-600 focus:outline-none transition shadow-sm"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-slate-700 mb-1">Nama Kepala Sekolah *</label>
                        <input
                          type="text"
                          value={state.dataSekolah.kepalaSekolah}
                          onChange={(e) => handleUpdateSchool("kepalaSekolah", e.target.value)}
                          className="w-full text-sm border border-slate-300 px-3 py-2 rounded-lg font-medium focus:ring-1 focus:ring-indigo-600 focus:border-indigo-600 focus:outline-none transition shadow-sm"
                        />
                      </div>

                    <div>
                      <label className="block text-xs font-bold text-slate-700 mb-1">NPSN Sekolah / Lembaga</label>
                      <input
                        type="text"
                        value={state.dataSekolah.npsn || ""}
                        onChange={(e) => handleUpdateSchool("npsn", e.target.value)}
                        placeholder="Contoh: 69835194"
                        className="w-full text-sm border border-slate-300 px-3 py-2 rounded-lg font-medium focus:ring-1 focus:ring-indigo-600 focus:border-indigo-600 focus:outline-none transition shadow-sm"
                      />
                    </div>

                    <div className="md:col-span-2 pt-3 border-t border-slate-50">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[10px] font-bold text-slate-400 uppercase tracking-widest">Tahun Pelajaran & Semester</span>
                        <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-widest">Konfigurasi Aktif</span>
                      </div>
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                        <div className="flex flex-col">
                          <label className="text-[10px] font-bold text-slate-500 uppercase mb-1">Tahun Ajaran Aktif</label>
                          <input
                            type="text"
                            placeholder="Contoh: 2025/2026"
                            value={state.dataSekolah.thAjaran}
                            onChange={(e) => handleUpdateSchool("thAjaran", e.target.value)}
                            className="w-full text-xs border border-slate-300 px-3 py-1.5 rounded-lg font-bold text-indigo-900 focus:ring-1 focus:ring-indigo-600 focus:border-indigo-600 focus:outline-none bg-slate-50/30"
                          />
                        </div>

                        <div className="flex flex-col text-right">
                          <label className="text-[10px] font-bold text-slate-500 uppercase mb-1">Semester Laporan</label>
                          <select
                            value={state.dataSekolah.semester}
                            onChange={(e) => handleUpdateSchool("semester", e.target.value)}
                            className="w-full text-xs border border-slate-300 px-3 py-1.5 rounded-lg bg-white font-bold text-indigo-900 focus:ring-1 focus:ring-indigo-600 focus:border-indigo-600 focus:outline-none"
                          >
                            <option value="1 (Ganjil)">1 (Ganjil)</option>
                            <option value="2 (Genap)">2 (Genap)</option>
                          </select>
                        </div>

                        <div className="sm:col-span-2 pt-1">
                          <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Tanggal Cetak Raport</label>
                          <input
                            type="date"
                            value={state.dataSekolah.tglRaport}
                            onChange={(e) => handleUpdateSchool("tglRaport", e.target.value)}
                            className="w-full text-xs border border-slate-300 px-3 py-1.5 rounded-lg font-medium focus:ring-1 focus:ring-indigo-600 focus:border-indigo-600 focus:outline-none"
                          />
                        </div>
                      </div>
                    </div>

                    {/* AMBIL DATA DARI TAHUN AJARAN LAIN */}
                    <div className="md:col-span-2 pt-4 border-t border-slate-100">
                      <div className="flex items-center justify-between mb-2">
                        <span className="text-[10px] font-bold text-slate-500 uppercase tracking-widest">Ambil Data Pelajaran Lain</span>
                        <span className="text-[10px] font-bold text-indigo-500 uppercase tracking-widest">Import Internal</span>
                      </div>
                      <p className="text-[11px] text-neutral-400 mb-3">
                        Salin data master dari tahun ajaran sebelumnya agar tidak perlu menginput ulang dari awal.
                      </p>
                      
                      <div className="flex flex-col sm:flex-row gap-3 items-center">
                        <div className="w-full sm:w-1/3">
                          <label className="block text-[9px] font-bold text-slate-500 uppercase mb-1">Pilih Tahun Ajaran Sumber</label>
                          <select
                            value={tahunSumberAmbil}
                            onChange={(e) => setTahunSumberAmbil(e.target.value)}
                            className="w-full text-xs border border-slate-300 px-2.5 py-1.5 rounded-lg bg-white font-bold text-slate-700 focus:ring-1 focus:ring-indigo-600 focus:border-indigo-600 focus:outline-none"
                          >
                            <option value="">-- Pilih Tahun --</option>
                            {Array.from(new Set([
                              ...rawState.siswa.map(s => s.thAjaran),
                              ...rawState.tujuanPembelajaran.map(tp => tp.thAjaran),
                              ...rawState.subdimensiKokurikuler.map(sub => sub.thAjaran),
                              "2024/2025",
                              "2025/2026"
                            ].filter(Boolean) as string[]))
                              .filter(y => y !== (state.dataSekolah.thAjaran))
                              .map(year => (
                                <option key={year} value={year}>{year}</option>
                              ))
                            }
                          </select>
                        </div>
                        
                        <div className="w-full sm:w-2/3 grid grid-cols-1 sm:grid-cols-3 gap-2 pt-4 sm:pt-0">
                          <button
                            type="button"
                            onClick={() => handleAmbilSiswa(tahunSumberAmbil)}
                            disabled={!tahunSumberAmbil}
                            className="text-xs px-3 py-1.5 bg-indigo-50 hover:bg-indigo-100 disabled:opacity-50 disabled:cursor-not-allowed text-indigo-700 font-bold border border-indigo-200 rounded-lg shadow-2xs transition active:scale-95 flex items-center justify-center gap-1.5"
                          >
                            <Users className="w-3.5 h-3.5" />
                            Siswa Lama
                          </button>
                          
                          <button
                            type="button"
                            onClick={() => handleAmbilIntrakurikuler(tahunSumberAmbil)}
                            disabled={!tahunSumberAmbil}
                            className="text-xs px-3 py-1.5 bg-emerald-50 hover:bg-emerald-100 disabled:opacity-50 disabled:cursor-not-allowed text-emerald-700 font-bold border border-emerald-200 rounded-lg shadow-2xs transition active:scale-95 flex items-center justify-center gap-1.5"
                          >
                            <Folder className="w-3.5 h-3.5" />
                            Intrakurikuler
                          </button>
                          
                          <button
                            type="button"
                            onClick={() => handleAmbilKokurikuler(tahunSumberAmbil)}
                            disabled={!tahunSumberAmbil}
                            className="text-xs px-3 py-1.5 bg-sky-50 hover:bg-sky-100 disabled:opacity-50 disabled:cursor-not-allowed text-sky-700 font-bold border border-sky-200 rounded-lg shadow-2xs transition active:scale-95 flex items-center justify-center gap-1.5"
                          >
                            <Award className="w-3.5 h-3.5" />
                            Kokurikuler
                          </button>
                        </div>
                      </div>
                    </div>

                    <div className="md:col-span-2 pt-4 border-t border-slate-50 grid grid-cols-1 sm:grid-cols-3 gap-4">
                      <div>
                        <label className="block text-xs font-bold text-slate-700 mb-1">Logo Raport</label>
                        {state.dataSekolah.logo && (
                          <div className="mb-2 flex items-center gap-2 border border-slate-100 p-1.5 rounded-lg bg-slate-50/50 w-fit">
                            <img src={state.dataSekolah.logo} alt="Preview Logo" className="w-12 h-12 object-contain rounded border bg-white shadow-xs" />
                          </div>
                        )}
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) => handleLogoUpload(e, "logo")}
                          className="w-full text-[10px] border border-slate-200 p-1 rounded-md bg-white file:mr-2 file:py-1 file:px-2 file:border-0 file:text-[10px] file:font-bold file:bg-indigo-50 file:text-indigo-700"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-700 mb-1">Logo Sidebar</label>
                        {state.dataSekolah.logoSidebar && (
                          <div className="mb-2 flex items-center gap-2 border border-slate-100 p-1.5 rounded-lg bg-slate-50/50 w-fit">
                            <img src={state.dataSekolah.logoSidebar} alt="Preview Sidebar" className="w-8 h-8 object-contain rounded border bg-white shadow-xs" />
                          </div>
                        )}
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) => handleLogoUpload(e, "logoSidebar")}
                          className="w-full text-[10px] border border-slate-200 p-1 rounded-md bg-white file:mr-2 file:py-1 file:px-2 file:border-0 file:text-[10px] file:font-bold file:bg-slate-100 file:text-slate-600"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-slate-700 mb-1">Logo Login</label>
                        {state.dataSekolah.logoLogin && (
                          <div className="mb-2 flex items-center gap-2 border border-slate-100 p-1.5 rounded-lg bg-slate-50/50 w-fit">
                            <img src={state.dataSekolah.logoLogin} alt="Preview Login" className="w-8 h-8 object-contain rounded border bg-white shadow-xs" />
                          </div>
                        )}
                        <input
                          type="file"
                          accept="image/*"
                          onChange={(e) => handleLogoUpload(e, "logoLogin")}
                          className="w-full text-[10px] border border-slate-200 p-1 rounded-md bg-white file:mr-2 file:py-1 file:px-2 file:border-0 file:text-[10px] file:font-bold file:bg-slate-100 file:text-slate-600"
                        />
                      </div>
                    </div>
                  </div>

                  {/* AI CONFIGURATION SECTION */}
                  <div className="pt-6 mt-6 border-t border-slate-100 space-y-4">
                    <div className="flex items-center gap-2 text-indigo-900">
                      <Zap className="w-5 h-5" />
                      <span className="text-base font-black tracking-widest block uppercase">Pengaturan Kecerdasan Buatan (AI)</span>
                    </div>
                    <p className="text-sm text-neutral-500 font-medium italic leading-relaxed">
                      Pilih penyedia layanan AI untuk membantu menyusun narasi raport. Masukkan API Key yang sesuai melalui panel <strong className="text-indigo-700 underline underline-offset-2">Secrets</strong> di menu AI Studio.
                    </p>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      <div>
                        <label className="block text-sm font-bold text-slate-700 mb-1">Penyedia AI (AI Provider) *</label>
                        <select
                          value={state.dataSekolah.aiProvider || "groq"}
                          onChange={(e) => handleUpdateSchool("aiProvider", e.target.value)}
                          className="w-full text-sm border border-slate-300 px-4 py-2.5 rounded-xl bg-white font-bold text-indigo-900 focus:outline-indigo-600 shadow-sm"
                        >
                          <option value="groq">Groq (Sangat Cepat & Llama 3)</option>
                          <option value="gemini">Google Gemini (Stabil & Akurat)</option>
                        </select>
                      </div>

                      <div>
                        <label className="block text-sm font-bold text-slate-700 mb-1">Model AI</label>
                        <select
                          value={state.dataSekolah.aiModel || ""}
                          onChange={(e) => handleUpdateSchool("aiModel", e.target.value)}
                          className="w-full text-sm border border-slate-300 px-4 py-2.5 rounded-xl bg-white font-bold text-indigo-900 focus:outline-indigo-600 shadow-sm"
                        >
                          {state.dataSekolah.aiProvider === "gemini" ? (
                            <>
                              <option value="gemini-1.5-flash">Gemini 1.5 Flash (Tercepat)</option>
                              <option value="gemini-1.5-pro">Gemini 1.5 Pro (Paling Cerdas)</option>
                            </>
                          ) : (
                            <>
                              <option value="llama-3.3-70b-versatile">Llama 3.3 70B Versatile</option>
                              <option value="llama-3.1-8b-instant">Llama 3.1 8B Instant</option>
                              <option value="mixtral-8x7b-32768">Mixtral 8x7B</option>
                            </>
                          )}
                        </select>
                      </div>
                    </div>
                    <div className="p-3 bg-indigo-50 border border-indigo-100 rounded-xl">
                      <p className="text-[11px] text-indigo-700 leading-relaxed">
                        <strong>Tips:</strong> Penggunaan <strong>Groq</strong> seringkali memberikan waktu respon yang jauh lebih cepat untuk pembuatan narasi rapor secara otomatis. Pastikan <code>GROQ_API_KEY</code> telah diatur di konfigurasi proyek.
                      </p>
                    </div>
                  </div>
                </div>
                   {/* BACKUP & RESTORE SIDEBAR */}
                <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm flex flex-col justify-between h-fit space-y-4">
                  <div>
                    <span className="text-base font-black text-indigo-900 tracking-widest block uppercase border-b border-slate-100 pb-2">Amankan Cadangan Laporan</span>
                    <p className="text-sm font-medium text-neutral-500 leading-relaxed mt-2.5 italic">
                      Seluruh data Anda kini disimpan secara otomatis pada Cloud Firestore database secara real-time. Untuk keamanan luring atau migrasi instan, Anda dapat mencadangkan basis data kapan saja ke berkas .JSON lokal.
                    </p>
                  </div>

                  <div className="space-y-2.5 pt-2">
                    <button
                      onClick={handleExportData}
                      className="w-full bg-slate-900 hover:bg-slate-800 text-white py-2.5 px-4 rounded-lg text-xs font-bold flex items-center justify-center gap-2 shadow transition"
                    >
                      <Download className="w-4 h-4 text-indigo-400" /> Unduh Backup Database (.JSON)
                    </button>

                    {currentUserProfile?.role === "admin" && (
                      <div className="border-t border-slate-100 pt-3 flex flex-col gap-1">
                        <label className="block text-xs font-semibold text-slate-600 mb-1">Pulihkan Data dari Pencadangan</label>
                        <input
                          type="file"
                          accept=".json"
                          onChange={handleImportData}
                          className="w-full text-xs text-slate-500 file:mr-3 file:py-1 file:px-2 file:border-0 file:rounded file:bg-slate-100 hover:file:bg-slate-200"
                        />
                      </div>
                    )}
                  </div>
                </div>

              </div>

              {/* USER PROFILE & ROLE MANAGEMENT (Admins Only) */}
              {currentUserProfile?.role === "admin" && (
                <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6">
                    <div className="border-b border-slate-150 pb-2">
                    <span className="text-sm font-bold text-indigo-850 tracking-wider block uppercase mb-1">Kelola Akun Guru & Admin Sekolah</span>
                    <p className="text-[10px] text-neutral-400">Sebagai Admin Utama, Anda dapat memantau pengguna terdaftar, menambahkan guru baru dengan username/password, serta menyesuaikan peran.</p>
                  </div>

                  {/* FORM TAMBAH GURU */}
                  <form onSubmit={handleCreateTeacherAccount} className="p-4 bg-slate-50 border border-slate-100 rounded-xl space-y-3">
                    <span className="text-xs font-bold text-indigo-900 uppercase block">Tambah Akun Username Guru/Admin Baru</span>
                    <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                      <div>
                        <label className="block text-xs font-bold text-slate-500 uppercase mb-1">Nama Lengkap</label>
                        <input
                          type="text"
                          required
                          value={newTeacherNama}
                          onChange={(e) => setNewTeacherNama(e.target.value)}
                          placeholder="Contoh: Bu Indah Puspita, S.Pd"
                          className="w-full text-xs p-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-indigo-500 bg-white"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Nama Pengguna (Username)</label>
                        <input
                          type="text"
                          required
                          value={newTeacherUsername}
                          onChange={(e) => setNewTeacherUsername(e.target.value.toLowerCase().replace(/\s+/g, ""))}
                          placeholder="Contoh: indah"
                          className="w-full text-xs p-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-indigo-500 bg-white font-mono"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Kata Sandi (Password)</label>
                        <input
                          type="text"
                          required
                          value={newTeacherPassword}
                          onChange={(e) => setNewTeacherPassword(e.target.value)}
                          placeholder="Min. 6 karakter"
                          className="w-full text-xs p-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-indigo-500 bg-white"
                        />
                      </div>
                      <div>
                        <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Peran Akses</label>
                        <select
                          value={newTeacherRole}
                          onChange={(e) => setNewTeacherRole(e.target.value as "admin" | "guru")}
                          className="w-full text-xs p-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-indigo-500 bg-white"
                        >
                          <option value="guru">GURU (Akses Input Nilai & Siswa)</option>
                          <option value="admin">ADMIN (Akses Penuh Pengaturan)</option>
                        </select>
                      </div>
                    </div>
                    <div className="flex justify-end">
                      <button
                        type="submit"
                        className="py-1.5 px-4 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-lg text-xs transition flex items-center gap-1.5 shadow"
                      >
                        <Plus className="w-3.5 h-3.5" /> Tambahkan Pengguna
                      </button>
                    </div>
                  </form>
                  
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs text-left border-collapse">
                      <thead>
                        <tr className="border-b border-slate-100 text-slate-500 font-bold bg-slate-50/50 uppercase tracking-wider text-[10px]">
                          <th className="py-2.5 px-3">Nama Lengkap</th>
                          <th className="py-2.5 px-3">Username (Nama Pengguna)</th>
                          <th className="py-2.5 px-3">Kata Sandi</th>
                          <th className="py-2.5 px-3">Peran Saat Ini</th>
                          <th className="py-2.5 px-3 text-right">Tindakan</th>
                        </tr>
                      </thead>
                      <tbody>
                        {allUsersList.filter(u => u.username !== "admin_anchor" && u.email !== "admin_anchor").length > 0 ? (
                          allUsersList.filter(u => u.username !== "admin_anchor" && u.email !== "admin_anchor").map((usr, usrIdx) => (
                            <tr key={usr.uid || `user-${usrIdx}`} className="border-b border-slate-50 hover:bg-slate-50/50">
                              <td className="py-3 px-3 font-semibold text-slate-700">{usr.nama}</td>
                              <td className="py-3 px-3 text-slate-500 font-mono">{usr.username || (usr.email ? usr.email.split("@")[0] : "") || "—"}</td>
                              <td className="py-3 px-3 font-mono text-slate-550">
                                {usr.password ? usr.password : <span className="text-slate-350 italic text-[10px]">Kredensial Google</span>}
                              </td>
                              <td className="py-3 px-3">
                                <span className={`px-2 py-0.5 rounded-full text-[10px] font-bold ${
                                  usr.role === "admin" 
                                    ? "bg-purple-100 text-purple-700 border border-purple-200" 
                                    : "bg-teal-100 text-teal-700 border border-teal-200"
                                }`}>
                                  {(usr.role || "guru").toUpperCase()}
                                </span>
                              </td>
                              <td className="py-3 px-3 text-right">
                                <div className="flex items-center justify-end gap-1.5">
                                  {/* Edit Password Button for Username Users */}
                                  {usr.username && usr.uid !== currentUser?.uid && (
                                    <button
                                      onClick={() => {
                                        setUserToEditPassword(usr);
                                        setNewPasswordValue(usr.password || "");
                                      }}
                                      className="text-[10px] bg-slate-100 hover:bg-slate-200 border border-slate-200 text-slate-700 font-bold px-2 py-1 rounded"
                                    >
                                      Sandi
                                    </button>
                                  )}

                                  {usr.uid !== currentUser?.uid ? (
                                    <>
                                      <button
                                        onClick={() => {
                                          setUserToChangeRole(usr);
                                          setTargetRoleValue(usr.role === "admin" ? "guru" : "admin");
                                        }}
                                        className="text-[10px] bg-indigo-50 hover:bg-indigo-100 border border-indigo-200 text-indigo-700 font-bold px-2.5 py-1 rounded-lg transition"
                                      >
                                        Ubah Peran
                                      </button>
                                      <button
                                        onClick={() => {
                                          setUserToDelete(usr);
                                        }}
                                        className="text-[10px] bg-red-50 hover:bg-red-100 border border-red-200 text-red-650 font-bold p-1.5 rounded-lg transition flex items-center justify-center"
                                        title="Hapus Pengguna"
                                      >
                                        <Trash2 className="w-3.5 h-3.5" />
                                      </button>
                                    </>
                                  ) : (
                                    <span className="text-[10px] text-neutral-400 italic">Akun Anda</span>
                                  )}
                                </div>
                              </td>
                            </tr>
                          ))
                        ) : (
                          <tr key="empty-users">
                            <td colSpan={5} className="text-center py-6 text-slate-400 italic">Memuat daftar pengguna...</td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

            </div>
          )}

        </div>

      </main>

      {/* Toast Notification Container */}
      {toastMessage && (
        <div className="fixed top-5 right-5 z-[100] shadow-2xl rounded-xl border border-slate-200">
          <div className={`p-4 rounded-xl flex items-center gap-3 text-xs font-semibold ${
            toastMessage.type === "success" 
              ? "bg-emerald-50 text-emerald-800 border-l-4 border-emerald-500" 
              : toastMessage.type === "error"
              ? "bg-rose-50 text-rose-800 border-l-4 border-rose-500"
              : "bg-blue-50 text-blue-800 border-l-4 border-blue-500"
          }`}>
            <div className="flex-1">{toastMessage.text}</div>
            <button 
              onClick={() => setToastMessage(null)} 
              className="hover:bg-slate-200/50 p-1 rounded transition ml-2 text-slate-400 hover:text-slate-700 font-bold"
            >
              ✕
            </button>
          </div>
        </div>
      )}

      {/* Custom Delete User Confirmation Modal */}
      {userToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl space-y-4">
            <div className="flex gap-3 items-start">
              <div className="p-2.5 bg-rose-100 rounded-full text-rose-600 flex-shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-bold text-slate-800 uppercase tracking-tight">Hapus Akun Pengguna?</h3>
                <p className="text-xs text-neutral-500 leading-relaxed">
                  Apakah Anda yakin ingin menghapus akun guru/staf bernama <strong className="text-slate-800">{userToDelete.nama}</strong> ({userToDelete.username || userToDelete.email})?
                </p>
              </div>
            </div>
            
            <div className="p-3 bg-rose-50 border border-rose-100/50 rounded-xl">
              <p className="text-[10px] text-rose-700 font-medium">⚠️ Tindakan ini permanen. Semua pengaturan peran dan kredensial masuk pengguna ini akan segera dihapus dari data sistem.</p>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setUserToDelete(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmDelete}
                className="px-4 py-2 text-xs font-bold bg-rose-600 hover:bg-rose-700 text-white rounded-xl transition shadow-lg shadow-rose-600/20"
              >
                Ya, Hapus Akun
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Custom Delete Kelas Confirmation Modal */}
      {kelasToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl space-y-4 animate-fade-in">
            <div className="flex gap-3 items-start">
              <div className="p-2.5 bg-rose-100 rounded-full text-rose-600 flex-shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-bold text-slate-800 uppercase tracking-tight">Hapus Kelompok Kelas?</h3>
                <p className="text-xs text-neutral-500 leading-relaxed">
                  Apakah Anda yakin ingin menghapus kelas <strong className="text-slate-800">{kelasToDelete.namaKelas}</strong>? 
                  Semua siswa yang terdaftar di kelas ini akan kehilangan referensi kelas.
                </p>
              </div>
            </div>
            
            <div className="p-3 bg-rose-50 border border-rose-100/50 rounded-xl">
              <p className="text-[10px] text-rose-700 font-medium">⚠️ Tindakan ini tidak dapat dibatalkan.</p>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setKelasToDelete(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={confirmDeleteKelas}
                className="px-4 py-2 text-xs font-bold bg-rose-600 hover:bg-rose-700 text-white rounded-xl transition shadow-lg shadow-rose-600/20"
              >
                Hapus Kelas
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Custom Delete Siswa Confirmation Modal */}
      {siswaToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl space-y-4 animate-fade-in">
            <div className="flex gap-3 items-start">
              <div className="p-2.5 bg-rose-100 rounded-full text-rose-600 flex-shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-bold text-slate-800 uppercase tracking-tight">Hapus Data Siswa?</h3>
                <p className="text-xs text-neutral-500 leading-relaxed">
                  Apakah Anda yakin ingin menghapus data siswa <strong className="text-slate-800">{siswaToDelete.namaSiswa}</strong>?
                </p>
              </div>
            </div>
            
            <div className="p-3 bg-rose-50 border border-rose-100/50 rounded-xl">
              <p className="text-[10px] text-rose-700 font-medium">⚠️ Peringatan: Tindakan ini akan menghapus <strong>seluruh data siswa beserta nilai, absensi, dan catatan raportnya</strong> secara permanen.</p>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setSiswaToDelete(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmDeleteSiswa}
                className="px-4 py-2 text-xs font-bold bg-rose-600 hover:bg-rose-700 text-white rounded-xl transition shadow-lg shadow-rose-600/20"
              >
                Ya, Hapus Data
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Custom Delete Kategori Confirmation Modal */}
      {kategoriToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl space-y-4 animate-fade-in">
            <div className="flex gap-3 items-start">
              <div className="p-2.5 bg-rose-100 rounded-full text-rose-600 flex-shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-bold text-slate-800 uppercase tracking-tight">Hapus Kategori?</h3>
                <p className="text-xs text-neutral-500 leading-relaxed">
                  Apakah Anda yakin ingin menghapus kategori <strong className="text-slate-800">{kategoriToDelete.namaKategori}</strong>?
                </p>
              </div>
            </div>
            
            <div className="p-3 bg-rose-50 border border-rose-100/50 rounded-xl">
              <p className="text-[10px] text-rose-700 font-medium">⚠️ Peringatan: Tindakan ini akan menghapus <strong>seluruh Tujuan Pembelajaran (TP) dan nilai terkait</strong> dalam kategori ini pada semua kelas secara permanen.</p>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setKategoriToDelete(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmDeleteKategori}
                className="px-4 py-2 text-xs font-bold bg-rose-600 hover:bg-rose-700 text-white rounded-xl transition shadow-lg shadow-rose-600/20"
              >
                Ya, Hapus Kategori
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Custom Delete TP Confirmation Modal */}
      {tpToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl space-y-4 animate-fade-in">
            <div className="flex gap-3 items-start">
              <div className="p-2.5 bg-rose-100 rounded-full text-rose-600 flex-shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-bold text-slate-800 uppercase tracking-tight font-display">Hapus Tujuan Pembelajaran?</h3>
                <p className="text-xs text-neutral-500 leading-relaxed font-sans">
                  Apakah Anda yakin ingin menghapus Tujuan Pembelajaran: <span className="font-semibold text-slate-800">&ldquo;{tpToDelete.deskripsi.length > 60 ? tpToDelete.deskripsi.substring(0, 60) + "..." : tpToDelete.deskripsi}&rdquo;</span>?
                </p>
              </div>
            </div>
            
            <div className="p-3 bg-rose-50 border border-rose-100/50 rounded-xl font-sans">
              <p className="text-[10px] text-rose-700 font-medium">⚠️ Peringatan: Tindakan ini akan menghapus <strong>Tujuan Pembelajaran ini beserta data nilai terkait</strong> pada semua siswa secara permanen.</p>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setTpToDelete(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition font-sans"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmDeleteTp}
                className="px-4 py-2 text-xs font-bold bg-rose-600 hover:bg-rose-700 text-white rounded-xl transition shadow-lg shadow-rose-600/20 font-sans"
              >
                Ya, Hapus TP
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Custom Edit Password Modal */}
      {userToEditPassword && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl space-y-4">
            <div className="border-b border-slate-100 pb-2">
              <h3 className="text-sm font-bold text-indigo-900 uppercase tracking-tight">Ubah Kata Sandi</h3>
              <p className="text-xs text-neutral-400">Tentukan kata sandi baru untuk pengguna <strong className="text-slate-700">{userToEditPassword.nama}</strong>.</p>
            </div>

            <div className="space-y-3">
              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Nama Pengguna (Username)</label>
                <input
                  type="text"
                  disabled
                  value={userToEditPassword.username || ""}
                  className="w-full text-xs p-2 bg-slate-100 border border-slate-200 rounded-lg text-slate-500 font-mono focus:outline-none"
                />
              </div>
              <div>
                <label className="block text-[10px] font-bold text-slate-500 uppercase mb-1">Kata Sandi Baru (Min. 6 Karakter)</label>
                <input
                  type="text"
                  value={newPasswordValue}
                  onChange={(e) => setNewPasswordValue(e.target.value)}
                  placeholder="Masukkan kata sandi baru"
                  className="w-full text-xs p-2 border border-slate-200 rounded-lg focus:outline-none focus:ring-1 focus:ring-indigo-500 bg-white"
                />
              </div>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => {
                  setUserToEditPassword(null);
                  setNewPasswordValue("");
                }}
                className="px-4 py-2 text-xs font-semibold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmEditPassword}
                className="px-4 py-2 text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl transition shadow-lg shadow-indigo-600/20"
              >
                Simpan Kata Sandi
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Custom Change Role Modal */}
      {userToChangeRole && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl space-y-4">
            <div className="flex gap-3 items-start">
              <div className="p-2.5 bg-indigo-100 rounded-full text-indigo-600 flex-shrink-0 animate-pulse">
                <Settings className="w-5 h-5" />
              </div>
              <div className="space-y-1">
                <h3 className="text-sm font-bold text-indigo-900 uppercase tracking-tight">Ubah Peran Pengguna?</h3>
                <p className="text-xs text-neutral-500 leading-relaxed">
                  Apakah Anda yakin ingin mengganti peran <strong className="text-slate-800">{userToChangeRole.nama}</strong> menjadi:
                </p>
              </div>
            </div>

            <div className="flex justify-center py-2">
              <span className="text-xs font-bold uppercase tracking-wider px-3 py-1 bg-indigo-50 border border-indigo-100 text-indigo-700 rounded-full">
                {targetRoleValue === "admin" ? "ADMIN (Akses Penuh)" : "GURU (Akses Input Raport)"}
              </span>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setUserToChangeRole(null)}
                className="px-4 py-2 text-xs font-semibold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition"
              >
                Batal
              </button>
              <button
                type="button"
                onClick={handleConfirmChangeRole}
                className="px-4 py-2 text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl transition shadow-lg shadow-indigo-600/20"
              >
                Ubah Peran
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Tambah/Ubah Tujuan Pembelajaran */}
      {showAddTp && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-xl w-full p-6 shadow-2xl space-y-5 animate-fade-in">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-indigo-100 rounded-xl text-indigo-600">
                  <BookOpen className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-800 tracking-tight font-display">
                    {editingTp ? "Ubah Tujuan Pembelajaran" : "Tambah Tujuan Pembelajaran"}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium uppercase tracking-wider italic">
                    Intrakurikuler &bull; {editingTp ? "Update Data" : "Data Baru"}
                  </p>
                </div>
              </div>
              <button 
                onClick={() => {
                  setShowAddTp(false);
                  setEditingTp(null);
                }}
                className="p-1.5 hover:bg-slate-100 rounded-lg text-slate-400 transition"
              >
                <Plus className="w-5 h-5 rotate-45" />
              </button>
            </div>

            <form onSubmit={handleAddTp} className="space-y-4">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div>
                  <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1.5 tracking-wide">Target Kelas</label>
                  <select
                    value={tpForm.idKelas}
                    onChange={(e) => setTpForm({ ...tpForm, idKelas: e.target.value })}
                    required
                    className="w-full text-xs border border-slate-200 px-3 py-2.5 rounded-xl bg-slate-50 focus:bg-white transition outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium"
                  >
                    <option value="">-- Pilih Kelas --</option>
                    {state.kelas.map((k) => (
                      <option key={k.id} value={k.id}>{k.namaKelas}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1.5 tracking-wide">Kategori Capaian</label>
                  <select
                    value={tpForm.idKategori}
                    onChange={(e) => setTpForm({ ...tpForm, idKategori: e.target.value })}
                    required
                    className="w-full text-xs border border-slate-200 px-3 py-2.5 rounded-xl bg-slate-50 focus:bg-white transition outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium"
                  >
                    {state.kategoriIntrakurikuler.map((kat) => (
                      <option key={kat.id} value={kat.id}>{kat.namaKategori}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div>
                <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1.5 tracking-wide">Deskripsi Tujuan Pembelajaran (TP)</label>
                <textarea
                  rows={3}
                  placeholder="Contoh: Mengenal tata cara berdoa sebelum dan setelah melakukan aktivitas harian."
                  value={tpForm.deskripsi}
                  onChange={(e) => setTpForm({ ...tpForm, deskripsi: e.target.value })}
                  required
                  className="w-full text-xs border border-slate-200 px-3 py-2.5 rounded-xl bg-slate-50 focus:bg-white transition outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium leading-relaxed"
                />
              </div>

              <div>
                <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1.5 tracking-wide">Aktivitas / Metode</label>
                <textarea
                  rows={3}
                  placeholder="Contoh: Menggunakan metode bercerita, demonstrasi, dan praktik langsung."
                  value={tpForm.aktivitasMetode}
                  onChange={(e) => setTpForm({ ...tpForm, aktivitasMetode: e.target.value })}
                  className="w-full text-xs border border-slate-200 px-3 py-2.5 rounded-xl bg-slate-50 focus:bg-white transition outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium leading-relaxed"
                />
              </div>

              <div className="flex items-center justify-end gap-3 pt-2">
                <button
                  type="button"
                  onClick={() => {
                    setShowAddTp(false);
                    setEditingTp(null);
                  }}
                  className="px-5 py-2.5 text-xs font-bold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition font-sans"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-6 py-2.5 text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl transition shadow-lg shadow-indigo-600/20 flex items-center gap-2 font-sans"
                >
                  <Save className="w-4 h-4" />
                  {editingTp ? "Simpan Perubahan TP" : "Simpan TP Baru"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal Kelola Kategori Intrakurikuler */}
      {showAddKategori && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-2xl w-full p-6 shadow-2xl space-y-6 animate-fade-in flex flex-col max-h-[90vh]">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4 flex-shrink-0">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-slate-100 rounded-xl text-slate-600">
                  <Folder className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-800 tracking-tight font-display">Kelola Kategori Intrakurikuler</h3>
                  <p className="text-[10px] text-slate-400 font-medium uppercase tracking-wider italic">Pengaturan Master Kategori</p>
                </div>
              </div>
              <button 
                onClick={() => {
                  setShowAddKategori(false);
                  setEditingKategori(null);
                  setKategoriForm({ namaKategori: "" });
                }}
                className="p-1.5 hover:bg-slate-100 rounded-lg text-slate-400 transition"
              >
                <Plus className="w-5 h-5 rotate-45" />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto space-y-6 pr-1 custom-scrollbar">
              <div className="bg-slate-50/50 p-4 rounded-xl border border-dashed border-slate-200">
                <form onSubmit={handleSaveKategori} className="space-y-3">
                  <label className="block text-[10px] uppercase font-bold text-slate-500 tracking-wide">
                    {editingKategori ? "Ubah Nama Kategori" : "Tambah Kategori Baru"}
                  </label>
                  <div className="flex gap-2">
                    <input
                      type="text"
                      placeholder="Masukkan nama kategori baru..."
                      value={kategoriForm.namaKategori}
                      onChange={(e) => setKategoriForm({ ...kategoriForm, namaKategori: e.target.value })}
                      required
                      className="flex-1 text-xs border border-slate-200 px-4 py-2.5 rounded-xl bg-white outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium shadow-sm"
                    />
                    <button
                      type="submit"
                      className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition flex items-center gap-1 shadow-sm font-sans"
                    >
                      {editingKategori ? <Check className="w-4 h-4"/> : <Plus className="w-4 h-4"/>}
                      {editingKategori ? "Update" : "Tambah"}
                    </button>
                    {editingKategori && (
                      <button
                        type="button"
                        onClick={() => {
                          setEditingKategori(null);
                          setKategoriForm({ namaKategori: "" });
                        }}
                        className="p-2.5 bg-slate-200 hover:bg-slate-300 text-slate-600 rounded-xl transition"
                      >
                        <Plus className="w-4 h-4 rotate-45" />
                      </button>
                    )}
                  </div>
                </form>
              </div>

              <div className="space-y-2">
                <h4 className="text-[10px] uppercase font-bold text-slate-400 tracking-widest pl-1">Daftar Kategori Aktif ({state.kategoriIntrakurikuler.length})</h4>
                <div className="bg-white border border-slate-150 rounded-xl overflow-hidden divide-y divide-slate-100">
                  {state.kategoriIntrakurikuler.map((kat, idx) => (
                    <div key={kat.id || `kat-mgr-${idx}`} className="px-4 py-3 flex items-center justify-between hover:bg-slate-50/50 transition-colors">
                      <div className="flex items-center gap-3">
                        <span className="text-[10px] font-mono font-bold text-slate-300">{(idx + 1).toString().padStart(2, '0')}</span>
                        <span className="text-xs font-semibold text-slate-700 leading-snug">{kat.namaKategori}</span>
                      </div>
                      <div className="flex items-center gap-1">
                        <button
                          onClick={() => {
                            setEditingKategori(kat);
                            setKategoriForm({ namaKategori: kat.namaKategori });
                          }}
                          className={`p-1.5 rounded-lg transition-colors ${editingKategori?.id === kat.id ? 'bg-indigo-100 text-indigo-600' : 'text-slate-400 hover:text-indigo-600 hover:bg-indigo-50'}`}
                        >
                          <Edit3 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => setKategoriToDelete(kat)}
                          className="p-1.5 text-slate-400 hover:text-red-500 hover:bg-rose-50 rounded-lg transition-colors"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="pt-4 border-t border-slate-100 flex justify-end flex-shrink-0">
              <button
                onClick={() => {
                  setShowAddKategori(false);
                  setEditingKategori(null);
                  setKategoriForm({ namaKategori: "" });
                }}
                className="px-6 py-2.5 bg-slate-900 text-white rounded-xl text-xs font-bold hover:bg-slate-800 transition shadow-lg shadow-slate-900/10 font-sans"
              >
                Selesai
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Tambah/Ubah Subdimensi Kokurikuler (P5) */}
      {showAddSub && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[100] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-2xl w-full p-6 shadow-2xl space-y-5 animate-fade-in overflow-y-auto max-h-[95vh] custom-scrollbar">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-indigo-100 rounded-xl text-indigo-600">
                  <Award className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-800 tracking-tight font-display">
                    {editingSub ? "Ubah Subdimensi Projek" : "Tambah Subdimensi Projek"}
                  </h3>
                  <p className="text-[10px] text-slate-400 font-medium uppercase tracking-wider italic">
                    Kokurikuler (P5) &bull; {editingSub ? "Update Data" : "Data Baru"}
                  </p>
                </div>
              </div>
              <button 
                onClick={() => {
                  setShowAddSub(false);
                  setEditingSub(null);
                }}
                className="p-1.5 hover:bg-slate-100 rounded-lg text-slate-400 transition"
              >
                <Plus className="w-5 h-5 rotate-45" />
              </button>
            </div>

            <form onSubmit={handleAddSub} className="space-y-4">
              <div>
                <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1.5 tracking-wide">Target Kelompok Kelas</label>
                <select
                  value={subForm.idKelas}
                  onChange={(e) => setSubForm({ ...subForm, idKelas: e.target.value })}
                  required
                  className="w-full text-xs border border-slate-200 px-3 py-2.5 rounded-xl bg-slate-50 focus:bg-white transition outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium font-sans"
                >
                  <option value="">-- Pilih Kelas --</option>
                  {state.kelas.map((k) => (
                    <option key={k.id} value={k.id}>{k.namaKelas}</option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-[10px] uppercase font-bold text-slate-500 mb-1.5 tracking-wide">Nama Subdimensi / Elemen Projek *</label>
                <textarea
                  rows={2}
                  placeholder="Contoh: Dimensi Gotong Royong: Menunjukkan inisiatif untuk bekerjasama..."
                  value={subForm.namaSubdimensi}
                  onChange={(e) => setSubForm({ ...subForm, namaSubdimensi: e.target.value })}
                  required
                  className="w-full text-xs border border-slate-200 px-3 py-2.5 rounded-xl bg-slate-50 focus:bg-white transition outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium leading-relaxed font-sans"
                />
              </div>

              <div className="space-y-4 pt-2 border-t border-slate-50">
                <h4 className="text-[10px] uppercase font-bold text-indigo-600 tracking-widest pl-1">Deskripsi Capaian per Tingkatan</h4>
                
                <div className="grid grid-cols-1 gap-4">
                  {state.labelP5.map((label) => {
                    const colors = [
                      { bg: "bg-emerald-500", text: "text-emerald-600", border: "border-emerald-500", light: "bg-emerald-50/10" },
                      { bg: "bg-sky-500", text: "text-sky-600", border: "border-sky-500", light: "bg-sky-50/10" },
                      { bg: "bg-indigo-500", text: "text-indigo-600", border: "border-indigo-500", light: "bg-indigo-50/10" },
                      { bg: "bg-amber-500", text: "text-amber-600", border: "border-amber-500", light: "bg-amber-50/10" },
                      { bg: "bg-rose-500", text: "text-rose-600", border: "border-rose-500", light: "bg-rose-50/10" },
                    ];
                    const color = colors[(label.order - 1) % colors.length];
                    
                    return (
                      <div key={label.id} className="space-y-1.5">
                        <label className={`flex items-center gap-2 text-[10px] uppercase font-bold ${color.text} tracking-wide`}>
                          <div className={`w-2 h-2 rounded-full ${color.bg}`} /> {label.namaLabel}
                        </label>
                        <textarea
                          rows={2}
                          placeholder={`Deskripsi untuk kriteria ${label.namaLabel}...`}
                          value={subForm.capaian?.[label.id] || (label.namaLabel === "Berkembang" ? subForm.descBerkembang : label.namaLabel === "Cakap" ? subForm.descCakap : label.namaLabel === "Mahir" ? subForm.descMahir : "")}
                          onChange={(e) => setSubForm({ 
                            ...subForm, 
                            capaian: { ...subForm.capaian, [label.id]: e.target.value } 
                          })}
                          className={`w-full text-xs border border-slate-200 px-3 py-2.5 rounded-xl ${color.light} focus:bg-white transition outline-none focus:ring-2 focus:ring-opacity-20 focus:ring-${color.bg.split('-')[1]}-500 focus:border-${color.bg.split('-')[1]}-500 font-medium font-sans`}
                        />
                      </div>
                    );
                  })}
                </div>
              </div>

              <div className="flex items-center justify-end gap-3 pt-4 border-t border-slate-50">
                <button
                  type="button"
                  onClick={() => {
                    setShowAddSub(false);
                    setEditingSub(null);
                  }}
                  className="px-5 py-2.5 text-xs font-bold text-slate-500 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition font-sans"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-6 py-2.5 text-xs font-bold bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl transition shadow-lg shadow-indigo-600/20 flex items-center gap-2 font-sans"
                >
                  <Save className="w-4 h-4" />
                  {editingSub ? "Update Subdimensi" : "Simpan Subdimensi"}
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal Kelola Label P5 */}
      {showAddLabel && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[110] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-2xl w-full p-6 shadow-2xl space-y-6 animate-fade-in flex flex-col max-h-[90vh]">
            <div className="flex items-center justify-between border-b border-slate-100 pb-4 flex-shrink-0">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-slate-100 rounded-xl text-slate-600">
                  <Settings className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-slate-800 tracking-tight font-display">Kelola Label Capaian Kokurikuler</h3>
                  <p className="text-[10px] text-slate-400 font-medium uppercase tracking-wider italic">Pengaturan Tingkat Capaian P5</p>
                </div>
              </div>
              <button 
                onClick={() => {
                  setShowAddLabel(false);
                  setEditingLabel(null);
                  setLabelForm({ namaLabel: "", order: 0 });
                }}
                className="p-1.5 hover:bg-slate-100 rounded-lg text-slate-400 transition"
              >
                <Plus className="w-5 h-5 rotate-45" />
              </button>
            </div>

            <div className="flex-1 overflow-y-auto space-y-6 pr-1 custom-scrollbar">
              <div className="bg-slate-50/50 p-4 rounded-xl border border-dashed border-slate-200">
                <form onSubmit={handleSaveLabel} className="space-y-3">
                  <div className="grid grid-cols-1 md:grid-cols-4 gap-3">
                    <div className="md:col-span-3">
                      <label className="block text-[10px] uppercase font-bold text-slate-500 tracking-wide mb-1">
                        {editingLabel ? "Ubah Nama Label" : "Tambah Label Baru"}
                      </label>
                      <input
                        type="text"
                        placeholder="Contoh: Sangat Berkembang"
                        value={labelForm.namaLabel}
                        onChange={(e) => setLabelForm({ ...labelForm, namaLabel: e.target.value })}
                        required
                        className="w-full text-xs border border-slate-200 px-4 py-2.5 rounded-xl bg-white outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium shadow-sm"
                      />
                    </div>
                    <div>
                      <label className="block text-[10px] uppercase font-bold text-slate-500 tracking-wide mb-1">Urutan</label>
                      <input
                        type="number"
                        value={labelForm.order}
                        onChange={(e) => setLabelForm({ ...labelForm, order: Number(e.target.value) })}
                        className="w-full text-xs border border-slate-200 px-4 py-2.5 rounded-xl bg-white outline-none focus:ring-2 focus:ring-indigo-600/20 focus:border-indigo-600 font-medium shadow-sm"
                      />
                    </div>
                  </div>
                  <div className="flex justify-end gap-2">
                    <button
                      type="submit"
                      className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-xs font-bold transition flex items-center gap-1 shadow-sm font-sans"
                    >
                      {editingLabel ? <Check className="w-4 h-4"/> : <Plus className="w-4 h-4"/>}
                      {editingLabel ? "Update Label" : "Tambah Label"}
                    </button>
                    {editingLabel && (
                      <button
                        type="button"
                        onClick={() => {
                          setEditingLabel(null);
                          setLabelForm({ namaLabel: "", order: 0 });
                        }}
                        className="p-2.5 bg-slate-200 hover:bg-slate-300 text-slate-600 rounded-xl transition"
                      >
                        <Plus className="w-4 h-4 rotate-45" />
                      </button>
                    )}
                  </div>
                </form>
              </div>

              <div className="space-y-2">
                <h4 className="text-[10px] uppercase font-bold text-slate-400 tracking-widest pl-1">Daftar Label Aktif ({state.labelP5.length})</h4>
                <div className="bg-white border border-slate-150 rounded-xl overflow-hidden divide-y divide-slate-100">
                  {state.labelP5.map((lbl, idx) => (
                    <div key={lbl.id || `lbl-mgr-${idx}`} className="px-4 py-3 flex items-center justify-between hover:bg-slate-50/50 transition-colors">
                      <div className="flex items-center gap-3">
                        <span className="text-[10px] font-mono font-bold text-slate-300">{(lbl.order || idx + 1).toString().padStart(2, '0')}</span>
                        <span className="text-xs font-semibold text-slate-700 leading-snug">{lbl.namaLabel}</span>
                      </div>
                      <div className="flex items-center gap-1">
                        <button
                          onClick={() => {
                            setEditingLabel(lbl);
                            setLabelForm({ namaLabel: lbl.namaLabel, order: lbl.order });
                          }}
                          className={`p-1.5 rounded-lg transition-colors ${editingLabel?.id === lbl.id ? 'bg-indigo-100 text-indigo-600' : 'text-slate-400 hover:text-indigo-600 hover:bg-indigo-50'}`}
                        >
                          <Edit3 className="w-3.5 h-3.5" />
                        </button>
                        <button
                          onClick={() => setLabelToDelete(lbl)}
                          className="p-1.5 text-slate-400 hover:text-red-500 hover:bg-rose-50 rounded-lg transition-colors"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </div>

            <div className="pt-4 border-t border-slate-100 flex justify-end flex-shrink-0">
              <button
                onClick={() => {
                  setShowAddLabel(false);
                  setEditingLabel(null);
                  setLabelForm({ namaLabel: "", order: 0 });
                }}
                className="px-6 py-2.5 bg-slate-900 text-white rounded-xl text-xs font-bold hover:bg-slate-800 transition shadow-lg shadow-slate-900/10 font-sans"
              >
                Selesai
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Konfirmasi Hapus Label */}
      {labelToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[120] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl text-center space-y-4 animate-scale-in">
            <div className="w-16 h-16 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-2">
              <Trash2 className="w-8 h-8" />
            </div>
            <div>
              <h3 className="text-base font-bold text-slate-800 tracking-tight font-display">Hapus Label ini?</h3>
              <p className="text-xs text-slate-500 mt-1 leading-relaxed font-sans">
                Label <strong className="text-slate-700">&quot;{labelToDelete.namaLabel}&quot;</strong> akan dihapus. Deskripsi capaian yang terasosiasi dengan label ini pada butir subdimensi mungkin tidak akan ditampilkan lagi.
              </p>
            </div>
            <div className="flex gap-3 pt-2">
              <button
                onClick={() => setLabelToDelete(null)}
                className="flex-1 px-4 py-2.5 border border-slate-200 text-slate-600 rounded-xl text-xs font-bold hover:bg-slate-50 transition font-sans"
              >
                Batal
              </button>
              <button
                onClick={() => {
                  handleDeleteLabel(labelToDelete);
                  setLabelToDelete(null);
                }}
                className="flex-1 px-4 py-2.5 bg-rose-600 text-white rounded-xl text-xs font-bold hover:bg-rose-700 transition shadow-lg shadow-rose-600/20 font-sans"
              >
                Hapus
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Modal Konfirmasi Hapus Subdimensi */}
      {subToDelete && (
        <div className="fixed inset-0 bg-slate-900/60 backdrop-blur-sm z-[110] flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-slate-150 max-w-sm w-full p-6 shadow-2xl text-center space-y-4 animate-scale-in">
            <div className="w-16 h-16 bg-rose-100 text-rose-600 rounded-full flex items-center justify-center mx-auto mb-2">
              <Trash2 className="w-8 h-8" />
            </div>
            <div>
              <h3 className="text-base font-bold text-slate-800 tracking-tight font-display">Hapus Subdimensi ini?</h3>
              <p className="text-xs text-slate-500 mt-1 leading-relaxed font-sans">
                Data subdimensi <strong className="text-slate-700">&quot;{subToDelete.namaSubdimensi}&quot;</strong> akan dihapus secara permanen. Tindakan ini tidak dapat dibatalkan.
              </p>
            </div>
            <div className="flex gap-3 pt-2">
              <button
                onClick={() => setSubToDelete(null)}
                className="flex-1 px-4 py-2.5 border border-slate-200 text-slate-600 rounded-xl text-xs font-bold hover:bg-slate-50 transition font-sans"
              >
                Batal
              </button>
              <button
                onClick={async () => {
                   try {
                     await deleteDoc(doc(db, "subdimensiKokurikuler", subToDelete.id));
                     setSubToDelete(null);
                     showToast("Subdimensi berhasil dihapus", "success");
                   } catch (err) {
                     handleFirestoreError(err, OperationType.DELETE, "subdimensiKokurikuler");
                   }
                }}
                className="flex-1 px-4 py-2.5 bg-rose-600 text-white rounded-xl text-xs font-bold hover:bg-rose-700 transition shadow-lg shadow-rose-600/20 font-sans"
              >
                Hapus
              </button>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
