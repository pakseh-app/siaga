import React, { useState, useEffect, useMemo } from 'react';
import {
  Car, Users, Wallet, Settings, FileText, CheckCircle2, Clock,
  AlertTriangle, XCircle, Plus, Search, Filter, Printer, History,
  ShieldCheck, Eye, Lock, Unlock, Calendar, TrendingUp, ArrowUpRight,
  ArrowDownRight, Edit3, RotateCcw, Building2, UserCheck, CheckSquare,
  BarChart3, PieChart, Receipt, UserPlus, Phone, MapPin, Navigation,
  Award, Camera, KeyRound, Database, Copy, RefreshCw, Check, Cloud,
  Sun, Moon
} from 'lucide-react';

// Daftar Pilihan RT dan RW 01 s/d 15
const RT_RW_LIST = Array.from({ length: 15 }, (_, i) => String(i + 1).padStart(2, '0'));

// ============================================================================
// LINK GOOGLE APPS SCRIPT DESA KALIKURMO (SUDAH TERPASANG OTOMATIS)
// ============================================================================
const DEFAULT_SHEET_URL = 'https://script.google.com/macros/s/AKfycbyCxdTF-88mw3GVkKwXiiDecmz07ZgUcckd8TbzMGSpfSNPYNzthUDIyfU5IjpVutG2/exec';

const STORAGE_KEY = 'SIMOSD_DESA_KALIKURMO_MINGGUAN_V2';
const THEME_KEY = 'SIMOSD_THEME_MODE_V2';

// Kode Google Apps Script (Kode.gs) untuk Sistem Iuran Mingguan
const KODE_GS_SCRIPT = [
  '// ================================================================',
  '// KODE.GS - DATABASE REALTIME MOBIL SIAGA DESA KALIKURMO (MINGGUAN)',
  '// ================================================================',
  'var SHEET_NAMES = {',
  '  STATE: "Database_JSON",',
  '  WARGA: "Data_Warga",',
  '  TRANSAKSI: "Transaksi_Iuran",',
  '  ANTAR: "Riwayat_Antar_Warga",',
  '  KAS: "Buku_Kas_Siaga"',
  '};',
  '',
  'function ensureSheets() {',
  '  var ss = SpreadsheetApp.getActiveSpreadsheet();',
  '  for (var key in SHEET_NAMES) {',
  '    var name = SHEET_NAMES[key];',
  '    if (!ss.getSheetByName(name)) ss.insertSheet(name);',
  '  }',
  '  return ss;',
  '}',
  '',
  'function doGet() {',
  '  try {',
  '    var ss = ensureSheets();',
  '    var sheet = ss.getSheetByName(SHEET_NAMES.STATE);',
  '    var raw = sheet.getRange("A2").getValue();',
  '    var data = raw ? JSON.parse(raw) : null;',
  '    return ContentService.createTextOutput(JSON.stringify({',
  '      status: "success",',
  '      updatedAt: sheet.getRange("B2").getValue() || null,',
  '      data: data',
  '    })).setMimeType(ContentService.MimeType.JSON);',
  '  } catch (err) {',
  '    return ContentService.createTextOutput(JSON.stringify({',
  '      status: "error",',
  '      message: err.toString()',
  '    })).setMimeType(ContentService.MimeType.JSON);',
  '  }',
  '}',
  '',
  'function doPost(e) {',
  '  try {',
  '    var ss = ensureSheets();',
  '    var body = JSON.parse(e.postData.contents);',
  '    var now = new Date().toLocaleString("id-ID");',
  '',
  '    var sState = ss.getSheetByName(SHEET_NAMES.STATE);',
  '    sState.clear();',
  '    sState.getRange("A1:B1").setValues([["JSON_DATA_DESA_KALIKURMO", "TERAKHIR_UPDATE"]]);',
  '    sState.getRange("A2:B2").setValues([[JSON.stringify(body), now]]);',
  '',
  '    if (Array.isArray(body.residents)) {',
  '      var sWarga = ss.getSheetByName(SHEET_NAMES.WARGA);',
  '      sWarga.clear();',
  '      var rowsW = [["ID", "Nama Warga", "Dusun", "RT", "RW", "No KK", "No HP", "Terdaftar"]];',
  '      body.residents.forEach(function(w) {',
  '        rowsW.push([w.id, w.nama, w.dusun, "\'" + w.rt, "\'" + w.rw, "\'" + w.noKK, "\'" + w.noHP, w.terdaftarSejak]);',
  '      });',
  '      sWarga.getRange(1, 1, rowsW.length, rowsW[0].length).setValues(rowsW);',
  '    }',
  '',
  '    if (Array.isArray(body.transactions)) {',
  '      var sTrx = ss.getSheetByName(SHEET_NAMES.TRANSAKSI);',
  '      sTrx.clear();',
  '      var rowsT = [["No Trx", "Tanggal", "Nama Warga", "Dusun", "RT", "RW", "Mulai", "Lunas Sampai", "Jml Minggu", "Tarif/Minggu", "Nominal", "Paket", "Admin", "Status", "Keterangan"]];',
  '      body.transactions.forEach(function(t) {',
  '        rowsT.push([t.id, t.tanggalTransaksi, t.residentName, t.dusun, "\'" + t.rt, "\'" + (t.rw || "01"), t.periodeMulai, t.periodeAkhir, t.jumlahMinggu, t.tarifPerMinggu, t.nominal, t.paketLabel, t.adminName, t.status, t.keterangan]);',
  '      });',
  '      sTrx.getRange(1, 1, rowsT.length, rowsT[0].length).setValues(rowsT);',
  '    }',
  '',
  '    if (Array.isArray(body.tripLogs)) {',
  '      var sAntar = ss.getSheetByName(SHEET_NAMES.ANTAR);',
  '      sAntar.clear();',
  '      var rowsA = [["ID", "Tanggal", "Jam", "Nama Warga", "Dusun", "RT", "RW", "Tujuan RS", "Keperluan", "Sopir", "Uang Keluar", "Uang Masuk", "Keterangan"]];',
  '      body.tripLogs.forEach(function(a) {',
  '        rowsA.push([a.id, a.tanggal, a.waktu, a.residentName, a.dusun, "\'" + a.rt, "\'" + (a.rw || "01"), a.tujuan, a.keperluan, a.driverName, a.uangKeluar, a.uangMasuk, a.keterangan]);',
  '      });',
  '      sAntar.getRange(1, 1, rowsA.length, rowsA[0].length).setValues(rowsA);',
  '    }',
  '',
  '    return ContentService.createTextOutput(JSON.stringify({',
  '      status: "success",',
  '      updatedAt: now',
  '    })).setMimeType(ContentService.MimeType.JSON);',
  '  } catch (err) {',
  '    return ContentService.createTextOutput(JSON.stringify({',
  '      status: "error",',
  '      message: err.toString()',
  '    })).setMimeType(ContentService.MimeType.JSON);',
  '  }',
  '}'
].join('\n');

// Data Awal Tarif Iuran Mingguan
const INITIAL_RATES = [
  {
    id: 'RATE-001',
    nominal: 5000,
    effectiveDate: '2026-01-01',
    keterangan: 'Tarif awal iuran mingguan Mobil Siaga Desa Kalikurmo hasil Musdes Tahun 2026',
    changedBy: 'Admin Budi (Bendahara Desa)',
    createdAt: '2026-01-01 08:00:00'
  }
];

// Data Awal Profil Kepala Desa & Sopir Desa Kalikurmo
const INITIAL_OFFICIALS = {
  kades: {
    nama: 'H. Muhammad Yusuf, S.Sos.',
    jabatan: 'Kepala Desa Kalikurmo',
    periode: '2024 - 2030',
    noHP: '081234567800',
    alamat: 'Dusun Krajan RT 01 / RW 01, Desa Kalikurmo',
    pesanSiaga: 'Mobil Siaga Desa Kalikurmo beroperasi 24 Jam Gratis untuk layanan darurat kesehatan & sosial seluruh warga Desa Kalikurmo.',
    unitMobil: 'Suzuki APV Arena GX - Nopol H 9912 XA (Mobil Siaga Desa Kalikurmo)',
    fotoUrl: ''
  },
  drivers: [
    {
      id: 'DRV-01',
      nama: 'Slamet Riyadi',
      panggilan: 'Bang Slamet',
      noHP: '081299887711',
      shift: 'Shift Pagi - Sore (07:00 - 19:00 WIB) & Siaga Panggilan',
      dusun: 'Dusun Krajan',
      rt: '01',
      rw: '01',
      statusSiaga: 'SIAGA AKTIF',
      pengalaman: '8 Tahun (Sopir Ambulans & Layanan Medis Desa Kalikurmo)',
      platMobil: 'H 9912 XA',
      fotoUrl: ''
    },
    {
      id: 'DRV-02',
      nama: 'Wahyu Setiawan',
      panggilan: 'Mas Wahyu',
      noHP: '081377665522',
      shift: 'Shift Malam (19:00 - 07:00 WIB) & Darurat Luar Kota',
      dusun: 'Dusun Kalikurmo Lor',
      rt: '02',
      rw: '01',
      statusSiaga: 'SIAGA AKTIF',
      pengalaman: '5 Tahun (Sopir Siaga Rujukan RSUD & Provinsi)',
      platMobil: 'H 9912 XA',
      fotoUrl: ''
    }
  ]
};

// Data Awal Warga Desa Kalikurmo
const INITIAL_RESIDENTS = [
  { id: 'WRG-001', nama: 'Ahmad Sulaiman', dusun: 'Dusun Krajan', rt: '01', rw: '01', noKK: '3321010101260001', noHP: '081234567801', terdaftarSejak: '2026-01-01' },
  { id: 'WRG-002', nama: 'Budi Santoso', dusun: 'Dusun Kalikurmo Lor', rt: '02', rw: '01', noKK: '3321010101260002', noHP: '081234567802', terdaftarSejak: '2026-01-01' },
  { id: 'WRG-003', nama: 'Siti Aminah', dusun: 'Dusun Krajan', rt: '03', rw: '02', noKK: '3321010101260003', noHP: '081234567803', terdaftarSejak: '2026-01-01' },
  { id: 'WRG-004', nama: 'Hendra Wijaya', dusun: 'Dusun Kalikurmo Kidul', rt: '01', rw: '02', noKK: '3321010101260004', noHP: '081234567804', terdaftarSejak: '2026-01-01' },
  { id: 'WRG-005', nama: 'Dewi Sartika', dusun: 'Dusun Kalikurmo Lor', rt: '04', rw: '03', noKK: '3321010101260005', noHP: '081234567805', terdaftarSejak: '2026-01-01' },
  { id: 'WRG-006', nama: 'Mulyono Hadi', dusun: 'Dusun Kalikurmo Kidul', rt: '02', rw: '03', noKK: '3321010101260006', noHP: '081234567806', terdaftarSejak: '2026-01-01' },
  { id: 'WRG-007', nama: 'Rina Marlina', dusun: 'Dusun Krajan', rt: '01', rw: '01', noKK: '3321010101260007', noHP: '081234567807', terdaftarSejak: '2026-01-01' },
  { id: 'WRG-008', nama: 'Joko Susilo', dusun: 'Dusun Kalikurmo Lor', rt: '03', rw: '02', noKK: '3321010101260008', noHP: '081234567808', terdaftarSejak: '2026-01-01' }
];

// Data Awal Transaksi Iuran Mingguan
const INITIAL_TRANSACTIONS = [
  {
    id: 'TRX-2026-0001',
    residentId: 'WRG-001',
    residentName: 'Ahmad Sulaiman',
    dusun: 'Dusun Krajan',
    rt: '01',
    rw: '01',
    tanggalTransaksi: '2026-01-01',
    nominal: 20000,
    tarifPerMinggu: 5000,
    periodeMulai: '2026-01-01',
    periodeAkhir: '2026-01-28',
    jumlahMinggu: 4,
    paketLabel: 'Lunas 1 Bulan (4 Minggu)',
    adminName: 'Admin Budi',
    keterangan: 'Pembayaran iuran 4 minggu bulan Januari 2026',
    status: 'VALID'
  },
  {
    id: 'TRX-2026-0002',
    residentId: 'WRG-001',
    residentName: 'Ahmad Sulaiman',
    dusun: 'Dusun Krajan',
    rt: '01',
    rw: '01',
    tanggalTransaksi: '2026-02-01',
    nominal: 40000,
    tarifPerMinggu: 5000,
    periodeMulai: '2026-01-29',
    periodeAkhir: '2026-03-25',
    jumlahMinggu: 8,
    paketLabel: 'Lunas 2 Bulan (8 Minggu)',
    adminName: 'Admin Budi',
    keterangan: 'Pelunasan 8 minggu (Februari - Maret 2026)',
    status: 'VALID'
  },
  {
    id: 'TRX-2026-0003',
    residentId: 'WRG-001',
    residentName: 'Ahmad Sulaiman',
    dusun: 'Dusun Krajan',
    rt: '01',
    rw: '01',
    tanggalTransaksi: '2026-03-26',
    nominal: 260000,
    tarifPerMinggu: 5000,
    periodeMulai: '2026-03-26',
    periodeAkhir: '2027-03-24',
    jumlahMinggu: 52,
    paketLabel: 'Lunas 1 Tahun (52 Minggu)',
    adminName: 'Admin Budi',
    keterangan: 'Pelunasan 1 tahun penuh (52 minggu) di muka',
    status: 'VALID'
  },
  {
    id: 'TRX-2026-0004',
    residentId: 'WRG-002',
    residentName: 'Budi Santoso',
    dusun: 'Dusun Kalikurmo Lor',
    rt: '02',
    rw: '01',
    tanggalTransaksi: '2026-11-01',
    nominal: 25000,
    tarifPerMinggu: 5000,
    periodeMulai: '2026-11-01',
    periodeAkhir: '2026-12-05',
    jumlahMinggu: 5,
    paketLabel: 'Lunas 1 Bulan (5 Minggu)',
    adminName: 'Admin Siti',
    keterangan: 'Lunas penuh bulan November 2026',
    status: 'VALID'
  },
  {
    id: 'TRX-2026-0005',
    residentId: 'WRG-004',
    residentName: 'Hendra Wijaya',
    dusun: 'Dusun Kalikurmo Kidul',
    rt: '01',
    rw: '02',
    tanggalTransaksi: '2026-11-05',
    nominal: 10000,
    tarifPerMinggu: 5000,
    periodeMulai: '2026-11-01',
    periodeAkhir: '2026-11-14',
    jumlahMinggu: 2,
    paketLabel: 'Bayar 2 Minggu',
    adminName: 'Admin Budi',
    keterangan: 'Bayar sebagian (2 minggu pertama November)',
    status: 'VALID'
  },
  {
    id: 'TRX-2026-0006',
    residentId: 'WRG-005',
    residentName: 'Dewi Sartika',
    dusun: 'Dusun Kalikurmo Lor',
    rt: '04',
    rw: '03',
    tanggalTransaksi: '2026-10-25',
    nominal: 60000,
    tarifPerMinggu: 5000,
    periodeMulai: '2026-11-01',
    periodeAkhir: '2027-01-23',
    jumlahMinggu: 12,
    paketLabel: 'Lunas 3 Bulan (12 Minggu)',
    adminName: 'Admin Siti',
    keterangan: 'Pelunasan di muka 12 minggu (Nov 2026 - Jan 2027)',
    status: 'VALID'
  },
  {
    id: 'TRX-2026-0007',
    residentId: 'WRG-007',
    residentName: 'Rina Marlina',
    dusun: 'Dusun Krajan',
    rt: '01',
    rw: '01',
    tanggalTransaksi: '2026-09-01',
    nominal: 20000,
    tarifPerMinggu: 5000,
    periodeMulai: '2026-09-01',
    periodeAkhir: '2026-09-28',
    jumlahMinggu: 4,
    paketLabel: 'Lunas 1 Bulan (4 Minggu)',
    adminName: 'Admin Budi',
    keterangan: 'Pembayaran 4 minggu bulan September 2026',
    status: 'VALID'
  }
];

// Data Awal Riwayat Antar Warga
const INITIAL_TRIPS = [
  {
    id: 'TRP-2026-001',
    tanggal: '2026-10-14',
    waktu: '08:15 WIB',
    residentId: 'WRG-003',
    residentName: 'Siti Aminah',
    dusun: 'Dusun Krajan',
    rt: '03',
    rw: '02',
    tujuan: 'RSUD Kabupaten',
    keperluan: 'Rujukan Persalinan Darurat',
    driverId: 'DRV-01',
    driverName: 'Slamet Riyadi',
    uangKeluar: 150000,
    uangMasuk: 0,
    keterangan: 'Pengisian BBM Pertalite & makan sopir (Layanan Gratis Warga Kalikurmo)',
    adminName: 'Admin Budi'
  },
  {
    id: 'TRP-2026-002',
    tanggal: '2026-11-03',
    waktu: '21:40 WIB',
    residentId: 'WRG-006',
    residentName: 'Mulyono Hadi',
    dusun: 'Dusun Kalikurmo Kidul',
    rt: '02',
    rw: '03',
    tujuan: 'RS Islam Harapan Bunda',
    keperluan: 'Antar Pasien Rawat Inap & Kontrol Jantung',
    driverId: 'DRV-02',
    driverName: 'Wahyu Setiawan',
    uangKeluar: 185000,
    uangMasuk: 50000,
    keterangan: 'BBM Rp150.000 + Tol Rp35.000. Keluarga memberi infaq sukarela kas Rp50.000',
    adminName: 'Admin Siti'
  }
];

// Data Awal Buku Kas Mobil Siaga
const INITIAL_CASH = [
  { id: 'KAS-0001', tanggal: '2026-01-01', tipe: 'MASUK', kategori: 'IURAN_WARGA', refId: 'TRX-2026-0001', deskripsi: 'Iuran Mobil Siaga - Ahmad Sulaiman (4 minggu)', nominal: 20000, admin: 'Admin Budi' },
  { id: 'KAS-0002', tanggal: '2026-02-01', tipe: 'MASUK', kategori: 'IURAN_WARGA', refId: 'TRX-2026-0002', deskripsi: 'Iuran Mobil Siaga - Ahmad Sulaiman (8 minggu)', nominal: 40000, admin: 'Admin Budi' },
  { id: 'KAS-0003', tanggal: '2026-03-26', tipe: 'MASUK', kategori: 'IURAN_WARGA', refId: 'TRX-2026-0003', deskripsi: 'Iuran Mobil Siaga - Ahmad Sulaiman (52 minggu / 1 Tahun)', nominal: 260000, admin: 'Admin Budi' },
  { id: 'KAS-0004', tanggal: '2026-09-01', tipe: 'MASUK', kategori: 'IURAN_WARGA', refId: 'TRX-2026-0007', deskripsi: 'Iuran Mobil Siaga - Rina Marlina (4 minggu)', nominal: 20000, admin: 'Admin Budi' },
  { id: 'KAS-0005', tanggal: '2026-10-14', tipe: 'KELUAR', kategori: 'ANTAR_WARGA', refId: 'TRP-2026-001', deskripsi: 'Operasional Antar Warga: Siti Aminah ke RSUD Kabupaten (Sopir: Slamet Riyadi)', nominal: 150000, admin: 'Admin Budi' },
  { id: 'KAS-0006', tanggal: '2026-10-25', tipe: 'MASUK', kategori: 'IURAN_WARGA', refId: 'TRX-2026-0006', deskripsi: 'Iuran Mobil Siaga - Dewi Sartika (12 minggu)', nominal: 60000, admin: 'Admin Siti' },
  { id: 'KAS-0007', tanggal: '2026-11-01', tipe: 'MASUK', kategori: 'IURAN_WARGA', refId: 'TRX-2026-0004', deskripsi: 'Iuran Mobil Siaga - Budi Santoso (5 minggu)', nominal: 25000, admin: 'Admin Siti' },
  { id: 'KAS-0008', tanggal: '2026-11-03', tipe: 'KELUAR', kategori: 'ANTAR_WARGA', refId: 'TRP-2026-002', deskripsi: 'Operasional Antar Warga: Mulyono Hadi ke RS Islam (Sopir: Wahyu Setiawan)', nominal: 185000, admin: 'Admin Siti' },
  { id: 'KAS-0009', tanggal: '2026-11-03', tipe: 'MASUK', kategori: 'INFAQ_PERJALANAN', refId: 'TRP-2026-002', deskripsi: 'Infaq Sukarela Perjalanan Warga: Mulyono Hadi', nominal: 50000, admin: 'Admin Siti' },
  { id: 'KAS-0010', tanggal: '2026-11-05', tipe: 'MASUK', kategori: 'IURAN_WARGA', refId: 'TRX-2026-0005', deskripsi: 'Iuran Mobil Siaga - Hendra Wijaya (2 minggu)', nominal: 10000, admin: 'Admin Budi' }
];

const INITIAL_AUDIT = [
  { id: 'AUD-001', timestamp: '2026-01-01 08:00:00', admin: 'Admin Budi', aksi: 'SET_TARIF_AWAL', detail: 'Menetapkan nominal iuran mingguan awal Rp5.000/minggu mulai 01/01/2026' },
  { id: 'AUD-002', timestamp: '2026-11-03 21:45:00', admin: 'Admin Siti', aksi: 'CATAT_ANTAR_WARGA', detail: 'Mencatat perjalanan antar warga Mulyono Hadi ke RS Islam (BBM Rp185.000, Infaq Rp50.000)' }
];

// Fungsi Utilitas Tanggal & Rupiah
function formatRupiah(num) {
  const n = Number(num) || 0;
  return new Intl.NumberFormat('id-ID', {
    style: 'currency',
    currency: 'IDR',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0
  }).format(n);
}

function parseYMD(dateStr) {
  if (!dateStr) return new Date(2026, 0, 1);
  const parts = String(dateStr).split('-').map(Number);
  return new Date(parts[0] || 2026, (parts[1] || 1) - 1, parts[2] || 1);
}

function toYMD(dateObj) {
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return y + '-' + m + '-' + d;
}

function formatDateID(dateStr) {
  if (!dateStr || dateStr === '-') return '-';
  const parts = String(dateStr).split('-');
  if (parts.length !== 3) return dateStr;
  return parts[2] + '/' + parts[1] + '/' + parts[0];
}

function formatDateLongID(dateStr) {
  if (!dateStr || dateStr === '-') return '-';
  const months = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'];
  const d = parseYMD(dateStr);
  return d.getDate() + ' ' + months[d.getMonth()] + ' ' + d.getFullYear();
}

function addDays(dateStr, daysToAdd) {
  const d = parseYMD(dateStr);
  d.setDate(d.getDate() + Number(daysToAdd));
  return toYMD(d);
}

function diffDaysInclusive(startStr, endStr) {
  const s = parseYMD(startStr);
  const e = parseYMD(endStr);
  const diffTime = e.getTime() - s.getTime();
  return Math.floor(diffTime / (1000 * 60 * 60 * 24)) + 1;
}

function getEndOfMonthYMD(dateStr) {
  const d = parseYMD(dateStr);
  const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0);
  return toYMD(lastDay);
}

function getEndOfNextMonthYMD(dateStr) {
  const d = parseYMD(dateStr);
  const lastDayNext = new Date(d.getFullYear(), d.getMonth() + 2, 0);
  return toYMD(lastDayNext);
}

function getStartOfNextMonthYMD(dateStr) {
  const d = parseYMD(dateStr);
  const firstDayNext = new Date(d.getFullYear(), d.getMonth() + 1, 1);
  return toYMD(firstDayNext);
}

function getEffectiveRateOnDate(rates, dateStr) {
  const sorted = [...rates].sort((a, b) => {
    if (a.effectiveDate !== b.effectiveDate) {
      return b.effectiveDate.localeCompare(a.effectiveDate);
    }
    return b.createdAt.localeCompare(a.createdAt);
  });
  const match = sorted.find((r) => r.effectiveDate <= dateStr);
  return match ? match : sorted[sorted.length - 1];
}

// Perhitungan Lengkap Status Pelunasan & Minggu Nunggak Warga
function computeResidentStatus(resident, transactions, systemDate, currentTarifMingguan) {
  const validTrx = transactions
    .filter((t) => t.residentId === resident.id && t.status === 'VALID')
    .sort((a, b) => a.periodeMulai.localeCompare(b.periodeMulai));

  if (validTrx.length === 0) {
    const startObligation = resident.terdaftarSejak || '2026-01-01';
    const overdueDays = startObligation <= systemDate ? diffDaysInclusive(startObligation, systemDate) : 0;
    const overdueWeeks = overdueDays > 0 ? Math.ceil(overdueDays / 7) : 0;
    return {
      statusKey: 'BELUM',
      statusLabel: 'BELUM BAYAR',
      badgeColor: 'red',
      isChecked: false,
      lunasSampai: '-',
      mulaiPertama: '-',
      tanggalPembayaranTerakhir: '-',
      totalNominal: 0,
      totalMinggu: 0,
      mingguNunggak: overdueWeeks,
      nominalTunggakan: overdueWeeks * currentTarifMingguan,
      keteranganTunggakan: overdueWeeks > 0 ? ('Belum Pernah Bayar (Nunggak ' + overdueWeeks + ' Minggu)') : 'Belum Ada Pembayaran',
      riwayat: [],
      hasAdvance1Month: false,
      hasAdvance1Year: false,
      isLunasBulanIni: false,
      isLunasBulanDepan: false,
      isBelumBayar: true
    };
  }

  let totalNominal = 0;
  let totalMinggu = 0;
  let maxEndDate = validTrx[0].periodeAkhir;
  let minStartDate = validTrx[0].periodeMulai;
  let lastTransDate = validTrx[0].tanggalTransaksi;
  let hasAdvance1Month = false;
  let hasAdvance1Year = false;

  validTrx.forEach((t) => {
    const weeks = Number(t.jumlahMinggu) || Math.max(1, Math.round((Number(t.jumlahHari) || 7) / 7));
    totalNominal += Number(t.nominal) || 0;
    totalMinggu += weeks;
    if (t.periodeAkhir > maxEndDate) maxEndDate = t.periodeAkhir;
    if (t.periodeMulai < minStartDate) minStartDate = t.periodeMulai;
    if (t.tanggalTransaksi > lastTransDate) lastTransDate = t.tanggalTransaksi;
    if (weeks >= 4 && weeks < 52) hasAdvance1Month = true;
    if (weeks >= 52) hasAdvance1Year = true;
  });

  const endOfCurrentMonth = getEndOfMonthYMD(systemDate);
  const startOfNextMonth = getStartOfNextMonthYMD(systemDate);
  const endOfNextMonth = getEndOfNextMonthYMD(systemDate);

  const isLunasBulanIni = maxEndDate >= endOfCurrentMonth;
  const isLunasBulanDepan = maxEndDate >= endOfNextMonth;
  const isLewatMasa = maxEndDate < systemDate;

  let mingguNunggak = 0;
  if (isLewatMasa) {
    const firstUnpaidDay = addDays(maxEndDate, 1);
    const overdueDays = diffDaysInclusive(firstUnpaidDay, systemDate);
    mingguNunggak = Math.ceil(overdueDays / 7);
  }
  const nominalTunggakan = mingguNunggak * currentTarifMingguan;

  if (isLewatMasa) {
    return {
      statusKey: 'BELUM',
      statusLabel: 'LEWAT MASA / NUNGGAK',
      badgeColor: 'red',
      isChecked: false,
      lunasSampai: maxEndDate,
      mulaiPertama: minStartDate,
      tanggalPembayaranTerakhir: lastTransDate,
      totalNominal,
      totalMinggu,
      mingguNunggak,
      nominalTunggakan,
      keteranganTunggakan: 'Nunggak ' + mingguNunggak + ' Minggu (Terakhir Lunas: ' + formatDateID(maxEndDate) + ')',
      riwayat: validTrx,
      hasAdvance1Month,
      hasAdvance1Year,
      isLunasBulanIni: false,
      isLunasBulanDepan: false,
      isBelumBayar: true
    };
  }

  if (maxEndDate >= startOfNextMonth) {
    const sisaHariKedepan = diffDaysInclusive(systemDate, maxEndDate);
    const sisaMingguKedepan = Math.ceil(sisaHariKedepan / 7);
    return {
      statusKey: 'LUNAS_MENDATANG',
      statusLabel: 'LUNAS PERIODE MENDATANG',
      badgeColor: 'blue',
      isChecked: true,
      lunasSampai: maxEndDate,
      mulaiPertama: minStartDate,
      tanggalPembayaranTerakhir: lastTransDate,
      totalNominal,
      totalMinggu,
      mingguNunggak: 0,
      nominalTunggakan: 0,
      keteranganTunggakan: 'Bebas Kewajiban ' + sisaMingguKedepan + ' Minggu ke Depan',
      riwayat: validTrx,
      hasAdvance1Month,
      hasAdvance1Year,
      isLunasBulanIni: true,
      isLunasBulanDepan,
      isBelumBayar: false
    };
  }

  if (maxEndDate >= endOfCurrentMonth) {
    return {
      statusKey: 'LUNAS',
      statusLabel: 'LUNAS',
      badgeColor: 'green',
      isChecked: true,
      lunasSampai: maxEndDate,
      mulaiPertama: minStartDate,
      tanggalPembayaranTerakhir: lastTransDate,
      totalNominal,
      totalMinggu,
      mingguNunggak: 0,
      nominalTunggakan: 0,
      keteranganTunggakan: 'Lunas Penuh Bulan Ini',
      riwayat: validTrx,
      hasAdvance1Month,
      hasAdvance1Year,
      isLunasBulanIni: true,
      isLunasBulanDepan: false,
      isBelumBayar: false
    };
  }

  return {
    statusKey: 'SEBAGIAN',
    statusLabel: 'PEMBAYARAN SEBAGIAN',
    badgeColor: 'yellow',
    isChecked: false,
    lunasSampai: maxEndDate,
    mulaiPertama: minStartDate,
    tanggalPembayaranTerakhir: lastTransDate,
    totalNominal,
    totalMinggu,
    mingguNunggak: 0,
    nominalTunggakan: 0,
    keteranganTunggakan: 'Baru Lunas s/d ' + formatDateID(maxEndDate) + ' (Belum Penuh 1 Bulan)',
    riwayat: validTrx,
    hasAdvance1Month,
    hasAdvance1Year,
    isLunasBulanIni: false,
    isLunasBulanDepan: false,
    isBelumBayar: false
  };
}

// ============================================================================
// KOMPONEN UTAMA APLIKASI SIMOSD DESA KALIKURMO
// ============================================================================
export default function App() {
  const savedState = useMemo(() => {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      return raw ? JSON.parse(raw) : null;
    } catch (e) {
      return null;
    }
  }, []);

  // State Mode Terang / Mode Malam (Default langsung di Tema Terang / 'light')
  const [themeMode, setThemeMode] = useState(() => {
    try {
      return localStorage.getItem(THEME_KEY) || 'light';
    } catch (e) {
      return 'light';
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem(THEME_KEY, themeMode);
    } catch (e) {}
  }, [themeMode]);

  // Tampilan pertama kali dibuka langsung ke Halaman Publik
  const [viewMode, setViewMode] = useState('PUBLIC');
  const [isAdminUnlocked, setIsAdminUnlocked] = useState(false);
  const [passwordModalOpen, setPasswordModalOpen] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const [passwordError, setPasswordError] = useState('');

  const [activeTab, setActiveTab] = useState('DASHBOARD');
  const [currentAdmin, setCurrentAdmin] = useState('Admin Budi (Bendahara)');
  const [systemDate, setSystemDate] = useState('2026-11-11');

  // Database Utama
  const [rates, setRates] = useState(() => savedState?.rates || INITIAL_RATES);
  const [residents, setResidents] = useState(() => savedState?.residents || INITIAL_RESIDENTS);
  const [transactions, setTransactions] = useState(() => savedState?.transactions || INITIAL_TRANSACTIONS);
  const [cashLedger, setCashLedger] = useState(() => savedState?.cashLedger || INITIAL_CASH);
  const [auditLogs, setAuditLogs] = useState(() => savedState?.auditLogs || INITIAL_AUDIT);
  const [villageOfficials, setVillageOfficials] = useState(() => savedState?.villageOfficials || INITIAL_OFFICIALS);
  const [tripLogs, setTripLogs] = useState(() => savedState?.tripLogs || INITIAL_TRIPS);

  // Google Sheets Sync State (Otomatis menggunakan DEFAULT_SHEET_URL Anda)
  const [sheetUrl, setSheetUrl] = useState(() => {
    const savedUrl = localStorage.getItem('SIMOSD_SHEET_URL');
    return savedUrl && savedUrl.startsWith('https://script.google.com') ? savedUrl : DEFAULT_SHEET_URL;
  });
  const [syncStatus, setSyncStatus] = useState('IDLE');
  const [lastSyncTime, setLastSyncTime] = useState('-');
  const [copiedCode, setCopiedCode] = useState(false);

  // Filter Global & Spesifik
  const [searchQuery, setSearchQuery] = useState('');
  const [filterDusun, setFilterDusun] = useState('ALL');
  const [filterRT, setFilterRT] = useState('ALL');
  const [filterRW, setFilterRW] = useState('ALL');
  const [filterStatus, setFilterStatus] = useState('ALL');
  const [filterYear, setFilterYear] = useState('2026');
  const [filterMonth, setFilterMonth] = useState('ALL');
  const [filterDate, setFilterDate] = useState('');

  // Modal State
  const [selectedResidentId, setSelectedResidentId] = useState(null);
  const [paymentModal, setPaymentModal] = useState({ open: false, residentId: null, presetType: '1_BULAN' });
  const [receiptModalTrx, setReceiptModalTrx] = useState(null);
  const [correctionModalTrx, setCorrectionModalTrx] = useState(null);
  const [addResidentModalOpen, setAddResidentModalOpen] = useState(false);

  // Notifikasi Toast
  const [toast, setToast] = useState(null);
  const showToast = (message, type = 'success') => {
    setToast({ message, type });
    setTimeout(() => setToast(null), 4000);
  };

  // Simpan otomatis ke localStorage
  useEffect(() => {
    try {
      const payload = {
        rates,
        residents,
        transactions,
        cashLedger,
        auditLogs,
        villageOfficials,
        tripLogs
      };
      localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
    } catch (e) {
      console.error('Gagal menyimpan ke localStorage:', e);
    }
  }, [rates, residents, transactions, cashLedger, auditLogs, villageOfficials, tripLogs]);

  useEffect(() => {
    try {
      if (sheetUrl) localStorage.setItem('SIMOSD_SHEET_URL', sheetUrl);
    } catch (e) {}
  }, [sheetUrl]);

  // Fungsi Sinkronisasi ke Google Sheet
  const syncToGoogleSheet = async (customPayload = null) => {
    if (!sheetUrl || !sheetUrl.startsWith('https://script.google.com')) return;
    setSyncStatus('SYNCING');
    try {
      const payload = customPayload || {
        rates,
        residents,
        transactions,
        cashLedger,
        auditLogs,
        villageOfficials,
        tripLogs
      };
      await fetch(sheetUrl, {
        method: 'POST',
        mode: 'no-cors',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(payload)
      });
      const nowStr = new Date().toLocaleTimeString('id-ID');
      setLastSyncTime(nowStr);
      setSyncStatus('SUCCESS');
    } catch (err) {
      setSyncStatus('ERROR');
    }
  };

  const fetchFromGoogleSheet = async (silent = false) => {
    if (!sheetUrl || !sheetUrl.startsWith('https://script.google.com')) {
      if (!silent) showToast('Masukkan URL Web App Google Sheet terlebih dahulu di menu Google Sheet.', 'error');
      return;
    }
    setSyncStatus('SYNCING');
    try {
      const res = await fetch(sheetUrl + '?t=' + Date.now());
      const json = await res.json();
      if (json && json.status === 'success' && json.data) {
        const d = json.data;
        if (Array.isArray(d.rates)) setRates(d.rates);
        if (Array.isArray(d.residents)) setResidents(d.residents);
        if (Array.isArray(d.transactions)) setTransactions(d.transactions);
        if (Array.isArray(d.cashLedger)) setCashLedger(d.cashLedger);
        if (Array.isArray(d.auditLogs)) setAuditLogs(d.auditLogs);
        if (d.villageOfficials) setVillageOfficials(d.villageOfficials);
        if (Array.isArray(d.tripLogs)) setTripLogs(d.tripLogs);
        setLastSyncTime(json.updatedAt || new Date().toLocaleTimeString('id-ID'));
        setSyncStatus('SUCCESS');
        if (!silent) showToast('Data berhasil disinkronkan dari Google Sheet Desa Kalikurmo!');
      } else {
        setSyncStatus('IDLE');
        if (!silent) showToast('Database Google Sheet masih kosong, silakan klik Simpan ke Sheet.', 'warning');
      }
    } catch (err) {
      setSyncStatus('ERROR');
      if (!silent) showToast('Gagal memuat dari Google Sheet. Pastikan akses Web App diatur ke Anyone.', 'error');
    }
  };

  useEffect(() => {
    if (sheetUrl && sheetUrl.startsWith('https://script.google.com')) {
      fetchFromGoogleSheet(true);
    }
  }, []);

  // Tarif Mingguan Aktif
  const activeRateObj = useMemo(() => getEffectiveRateOnDate(rates, systemDate), [rates, systemDate]);

  // Daftar Dusun Dinamis
  const dusunOptions = useMemo(() => {
    const setDusun = new Set(['Dusun Krajan', 'Dusun Kalikurmo Lor', 'Dusun Kalikurmo Kidul']);
    residents.forEach((r) => {
      if (r.dusun && r.dusun.trim()) setDusun.add(r.dusun.trim());
    });
    villageOfficials.drivers.forEach((d) => {
      if (d.dusun && d.dusun.trim()) setDusun.add(d.dusun.trim());
    });
    return Array.from(setDusun);
  }, [residents, villageOfficials]);

  // Data Warga + Perhitungan Status Mingguan
  const enrichedResidents = useMemo(() => {
    return residents.map((r) => ({
      ...r,
      computed: computeResidentStatus(r, transactions, systemDate, activeRateObj.nominal)
    }));
  }, [residents, transactions, systemDate, activeRateObj]);

  // Filter Warga
  const filteredResidents = useMemo(() => {
    return enrichedResidents.filter((r) => {
      const q = searchQuery.toLowerCase();
      const matchName = r.nama.toLowerCase().includes(q) || r.id.toLowerCase().includes(q) || r.dusun.toLowerCase().includes(q);
      const matchDusun = filterDusun === 'ALL' || r.dusun === filterDusun;
      const matchRT = filterRT === 'ALL' || r.rt === filterRT;
      const matchRW = filterRW === 'ALL' || (r.rw || '01') === filterRW;

      let matchStatus = true;
      if (filterStatus === 'LUNAS_BULAN_INI') {
        matchStatus = r.computed.isLunasBulanIni;
      } else if (filterStatus === 'LUNAS_BULAN_DEPAN') {
        matchStatus = r.computed.isLunasBulanDepan;
      } else if (filterStatus === 'BELUM_BAYAR') {
        matchStatus = r.computed.isBelumBayar;
      } else if (filterStatus !== 'ALL') {
        matchStatus = r.computed.statusKey === filterStatus;
      }

      return matchName && matchDusun && matchRT && matchRW && matchStatus;
    });
  }, [enrichedResidents, searchQuery, filterDusun, filterRT, filterRW, filterStatus]);

  // Statistik 9 Kartu Utama Dashboard
  const dashboardStats = useMemo(() => {
    const totalWarga = enrichedResidents.length;
    const sudahLunas = enrichedResidents.filter((r) => r.computed.statusKey === 'LUNAS' || r.computed.statusKey === 'LUNAS_MENDATANG').length;
    const belumLunas = enrichedResidents.filter((r) => r.computed.statusKey === 'BELUM').length;
    const bayarSebagian = enrichedResidents.filter((r) => r.computed.statusKey === 'SEBAGIAN').length;
    const lunasBulanIniCount = enrichedResidents.filter((r) => r.computed.isLunasBulanIni).length;
    const lunasBulanDepanCount = enrichedResidents.filter((r) => r.computed.isLunasBulanDepan).length;
    const lunas1Bulan = enrichedResidents.filter((r) => r.computed.hasAdvance1Month).length;
    const lunas1Tahun = enrichedResidents.filter((r) => r.computed.hasAdvance1Year).length;

    const validTransactions = transactions.filter((t) => {
      if (t.status !== 'VALID') return false;
      const matchDusun = filterDusun === 'ALL' || t.dusun === filterDusun;
      const matchRT = filterRT === 'ALL' || t.rt === filterRT;
      const matchRW = filterRW === 'ALL' || (t.rw || '01') === filterRW;
      const matchYear = filterYear === 'ALL' || t.tanggalTransaksi.startsWith(filterYear);
      const matchMonth = filterMonth === 'ALL' || t.tanggalTransaksi.slice(5, 7) === filterMonth;
      const matchDate = !filterDate || t.tanggalTransaksi === filterDate;
      return matchDusun && matchRT && matchRW && matchYear && matchMonth && matchDate;
    });

    const totalPemasukanIuran = validTransactions.reduce((acc, t) => acc + Number(t.nominal), 0);
    const endOfThisMonth = getEndOfMonthYMD(systemDate);
    const totalPembayaranDiMuka = validTransactions
      .filter((t) => t.periodeAkhir > endOfThisMonth)
      .reduce((acc, t) => acc + Number(t.nominal), 0);

    const totalMasukKas = cashLedger.filter((c) => c.tipe === 'MASUK').reduce((a, b) => a + Number(b.nominal), 0);
    const totalKeluarKas = cashLedger.filter((c) => c.tipe === 'KELUAR').reduce((a, b) => a + Number(b.nominal), 0);

    return {
      totalWarga,
      sudahLunas,
      belumLunas,
      bayarSebagian,
      lunasBulanIniCount,
      lunasBulanDepanCount,
      lunas1Bulan,
      lunas1Tahun,
      totalPemasukanIuran,
      totalPembayaranDiMuka,
      nominalMingguanBerlaku: activeRateObj.nominal,
      saldoKas: totalMasukKas - totalKeluarKas,
      totalMasukKas,
      totalKeluarKas,
      validTransactions
    };
  }, [enrichedResidents, transactions, cashLedger, filterDusun, filterRT, filterRW, filterYear, filterMonth, filterDate, systemDate, activeRateObj]);

  // Handler Audit Log
  const addAuditLog = (aksi, detail) => {
    const nowStr = systemDate + ' ' + new Date().toTimeString().slice(0, 8);
    setAuditLogs((prev) => [
      { id: 'AUD-' + String(prev.length + 1).padStart(3, '0'), timestamp: nowStr, admin: currentAdmin, aksi, detail },
      ...prev
    ]);
  };

  // Handler Login Admin Password "siaga"
  const handleAdminSwitch = () => {
    if (isAdminUnlocked) {
      setViewMode('ADMIN');
    } else {
      setPasswordInput('');
      setPasswordError('');
      setPasswordModalOpen(true);
    }
  };

  const handleVerifyPassword = (e) => {
    e.preventDefault();
    if (passwordInput.trim().toLowerCase() === 'siaga') {
      setIsAdminUnlocked(true);
      setViewMode('ADMIN');
      setPasswordModalOpen(false);
      setPasswordInput('');
      setPasswordError('');
      showToast('Berhasil masuk ke Dashboard Admin Desa Kalikurmo.');
    } else {
      setPasswordError('Password salah! Masukkan kata kunci: siaga');
    }
  };

  const handleLockAdmin = () => {
    setIsAdminUnlocked(false);
    setViewMode('PUBLIC');
    showToast('Dashboard Admin telah dikunci kembali.', 'warning');
  };

  // Handler Ubah Tarif Iuran Mingguan
  const handleAddRate = ({ nominal, effectiveDate, keterangan }) => {
    const numNominal = Number(nominal);
    if (!numNominal || numNominal < 500) {
      showToast('Nominal iuran mingguan tidak valid (minimal Rp500).', 'error');
      return;
    }
    if (!effectiveDate) {
      showToast('Tanggal mulai berlaku wajib diisi.', 'error');
      return;
    }
    const newRate = {
      id: 'RATE-' + String(rates.length + 1).padStart(3, '0'),
      nominal: numNominal,
      effectiveDate,
      keterangan: keterangan || 'Penyesuaian tarif iuran mingguan Mobil Siaga Desa Kalikurmo',
      changedBy: currentAdmin,
      createdAt: systemDate + ' ' + new Date().toTimeString().slice(0, 8)
    };
    const nextRates = [newRate, ...rates];
    setRates(nextRates);
    addAuditLog('UBAH_NOMINAL_IURAN', 'Mengubah tarif iuran mingguan menjadi ' + formatRupiah(numNominal) + '/minggu mulai ' + formatDateID(effectiveDate));
    syncToGoogleSheet({ rates: nextRates, residents, transactions, cashLedger, auditLogs, villageOfficials, tripLogs });
    showToast('Tarif baru ' + formatRupiah(numNominal) + '/minggu berhasil disimpan!');
  };

  // Handler Simpan Pembayaran Iuran Mingguan Warga
  const handleSavePayment = ({ residentId, tanggalTransaksi, periodeMulai, periodeAkhir, jumlahMinggu, paketLabel, keterangan }) => {
    const resident = residents.find((r) => r.id === residentId);
    if (!resident) return;

    const overlapping = transactions.find(
      (t) =>
        t.residentId === residentId &&
        t.status === 'VALID' &&
        !(periodeAkhir < t.periodeMulai || periodeMulai > t.periodeAkhir)
    );

    if (overlapping) {
      showToast(
        'Gagal: Periode ' + formatDateID(periodeMulai) + ' - ' + formatDateID(periodeAkhir) + ' bertumpukan dengan transaksi ' + overlapping.id,
        'error'
      );
      return;
    }

    const rateAtTrx = getEffectiveRateOnDate(rates, tanggalTransaksi);
    const tarifPerMinggu = rateAtTrx.nominal;
    const nominal = Number(jumlahMinggu) * tarifPerMinggu;
    const trxId = 'TRX-' + tanggalTransaksi.slice(0, 4) + '-' + String(transactions.length + 1).padStart(4, '0');

    const newTrx = {
      id: trxId,
      residentId: resident.id,
      residentName: resident.nama,
      dusun: resident.dusun,
      rt: resident.rt,
      rw: resident.rw || '01',
      tanggalTransaksi,
      nominal,
      tarifPerMinggu,
      periodeMulai,
      periodeAkhir,
      jumlahMinggu: Number(jumlahMinggu),
      paketLabel,
      adminName: currentAdmin,
      keterangan: keterangan || (paketLabel + ' (' + jumlahMinggu + ' minggu)'),
      status: 'VALID'
    };

    const newKas = {
      id: 'KAS-' + String(cashLedger.length + 1).padStart(4, '0'),
      tanggal: tanggalTransaksi,
      tipe: 'MASUK',
      kategori: 'IURAN_WARGA',
      refId: trxId,
      deskripsi: 'Iuran Mobil Siaga - ' + resident.nama + ' (' + jumlahMinggu + ' minggu: ' + formatDateID(periodeMulai) + ' s/d ' + formatDateID(periodeAkhir) + ')',
      nominal,
      admin: currentAdmin
    };

    const nextTrx = [newTrx, ...transactions];
    const nextKas = [newKas, ...cashLedger];
    setTransactions(nextTrx);
    setCashLedger(nextKas);
    addAuditLog('INPUT_PEMBAYARAN', 'Mencatat pembayaran ' + trxId + ' an. ' + resident.nama + ' (' + jumlahMinggu + ' minggu) sebesar ' + formatRupiah(nominal));
    syncToGoogleSheet({ rates, residents, transactions: nextTrx, cashLedger: nextKas, auditLogs, villageOfficials, tripLogs });
    setPaymentModal({ open: false, residentId: null, presetType: '1_BULAN' });
    showToast('Pembayaran ' + resident.nama + ' (' + jumlahMinggu + ' minggu) sebesar ' + formatRupiah(nominal) + ' berhasil disimpan!');
  };

  // Handler Koreksi Transaksi
  const handleCorrectTransaction = (trxId, alasanKoreksi) => {
    if (!alasanKoreksi || alasanKoreksi.trim().length < 5) {
      showToast('Alasan koreksi wajib diisi minimal 5 karakter untuk catatan audit.', 'error');
      return;
    }
    const target = transactions.find((t) => t.id === trxId);
    if (!target || target.status === 'VOID') return;

    const nextTrx = transactions.map((t) =>
      t.id === trxId
        ? { ...t, status: 'VOID', keterangan: t.keterangan + ' [DIKOREKSI: ' + alasanKoreksi + ']' }
        : t
    );

    const newKas = {
      id: 'KAS-' + String(cashLedger.length + 1).padStart(4, '0'),
      tanggal: systemDate,
      tipe: 'KELUAR',
      kategori: 'KOREKSI_IURAN',
      refId: trxId,
      deskripsi: 'Koreksi/Pembatalan Transaksi ' + trxId + ' (' + target.residentName + ') - Alasan: ' + alasanKoreksi,
      nominal: target.nominal,
      admin: currentAdmin
    };

    const nextKas = [newKas, ...cashLedger];
    setTransactions(nextTrx);
    setCashLedger(nextKas);
    addAuditLog('KOREKSI_TRANSAKSI', 'Membatalkan transaksi ' + trxId + ' (' + formatRupiah(target.nominal) + '). Alasan: ' + alasanKoreksi);
    syncToGoogleSheet({ rates, residents, transactions: nextTrx, cashLedger: nextKas, auditLogs, villageOfficials, tripLogs });
    setCorrectionModalTrx(null);
    showToast('Transaksi ' + trxId + ' telah dikoreksi & dicatat di Audit Log.', 'warning');
  };

  // Handler Tambah Warga Baru
  const handleAddResident = (newWarga) => {
    const id = 'WRG-' + String(residents.length + 1).padStart(3, '0');
    const created = {
      id,
      nama: newWarga.nama,
      dusun: newWarga.dusun.trim(),
      rt: newWarga.rt,
      rw: newWarga.rw || '01',
      noKK: newWarga.noKK || '3321010000000000',
      noHP: newWarga.noHP || '-',
      terdaftarSejak: systemDate
    };
    const nextRes = [...residents, created];
    setResidents(nextRes);
    addAuditLog('TAMBAH_WARGA', 'Menambahkan data warga baru: ' + created.nama + ' (' + created.dusun + ' RT ' + created.rt + '/RW ' + created.rw + ')');
    syncToGoogleSheet({ rates, residents: nextRes, transactions, cashLedger, auditLogs, villageOfficials, tripLogs });
    setAddResidentModalOpen(false);
    showToast('Warga ' + created.nama + ' berhasil ditambahkan!');
  };

  // Handler Tambah Riwayat Antar Warga
  const handleAddTripLog = (tripData) => {
    const id = 'TRP-' + systemDate.slice(0, 4) + '-' + String(tripLogs.length + 1).padStart(3, '0');
    const uangKeluar = Number(tripData.uangKeluar) || 0;
    const uangMasuk = Number(tripData.uangMasuk) || 0;

    const newTrip = {
      id,
      tanggal: tripData.tanggal || systemDate,
      waktu: tripData.waktu || '09:00 WIB',
      residentId: tripData.residentId || '-',
      residentName: tripData.residentName,
      dusun: tripData.dusun,
      rt: tripData.rt,
      rw: tripData.rw || '01',
      tujuan: tripData.tujuan,
      keperluan: tripData.keperluan,
      driverId: tripData.driverId,
      driverName: tripData.driverName,
      uangKeluar,
      uangMasuk,
      keterangan: tripData.keterangan || 'Layanan antar warga Mobil Siaga Desa Kalikurmo',
      adminName: currentAdmin
    };

    const newKasEntries = [];
    if (uangKeluar > 0) {
      newKasEntries.push({
        id: 'KAS-' + String(cashLedger.length + 1).padStart(4, '0'),
        tanggal: newTrip.tanggal,
        tipe: 'KELUAR',
        kategori: 'ANTAR_WARGA',
        refId: id,
        deskripsi: 'Operasional Antar Warga: ' + newTrip.residentName + ' ke ' + newTrip.tujuan + ' (Sopir: ' + newTrip.driverName + ')',
        nominal: uangKeluar,
        admin: currentAdmin
      });
    }
    if (uangMasuk > 0) {
      newKasEntries.push({
        id: 'KAS-' + String(cashLedger.length + newKasEntries.length + 1).padStart(4, '0'),
        tanggal: newTrip.tanggal,
        tipe: 'MASUK',
        kategori: 'INFAQ_PERJALANAN',
        refId: id,
        deskripsi: 'Infaq Sukarela Perjalanan Warga: ' + newTrip.residentName + ' (' + newTrip.tujuan + ')',
        nominal: uangMasuk,
        admin: currentAdmin
      });
    }

    const nextTrips = [newTrip, ...tripLogs];
    const nextKas = [...newKasEntries, ...cashLedger];
    setTripLogs(nextTrips);
    if (newKasEntries.length > 0) setCashLedger(nextKas);
    addAuditLog('CATAT_ANTAR_WARGA', 'Mencatat perjalanan ' + id + ': ' + newTrip.residentName + ' ke ' + newTrip.tujuan);
    syncToGoogleSheet({ rates, residents, transactions, cashLedger: nextKas, auditLogs, villageOfficials, tripLogs: nextTrips });
    showToast('Riwayat pemakaian mobil mengantar ' + newTrip.residentName + ' berhasil disimpan!');
  };

  // Handler Update Profil Kades & Sopir
  const handleUpdateKades = (newKades) => {
    const nextOff = { ...villageOfficials, kades: newKades };
    setVillageOfficials(nextOff);
    addAuditLog('UPDATE_PROFIL_KADES', 'Memperbarui profil Kepala Desa Kalikurmo: ' + newKades.nama);
    syncToGoogleSheet({ rates, residents, transactions, cashLedger, auditLogs, villageOfficials: nextOff, tripLogs });
    showToast('Profil Kepala Desa Kalikurmo berhasil diperbarui!');
  };

  const handleSaveDriver = (driverData) => {
    let nextDrivers = [];
    if (driverData.id) {
      nextDrivers = villageOfficials.drivers.map((d) => (d.id === driverData.id ? driverData : d));
      addAuditLog('UPDATE_SOPIR', 'Memperbarui data sopir Mobil Siaga: ' + driverData.nama + ' (' + (driverData.panggilan || '-') + ')');
      showToast('Profil Sopir ' + driverData.nama + ' berhasil diperbarui!');
    } else {
      const newId = 'DRV-' + String(villageOfficials.drivers.length + 1).padStart(2, '0');
      nextDrivers = [...villageOfficials.drivers, { ...driverData, id: newId }];
      addAuditLog('TAMBAH_SOPIR', 'Menambahkan sopir baru Mobil Siaga: ' + driverData.nama + ' (' + (driverData.panggilan || '-') + ')');
      showToast('Sopir baru ' + driverData.nama + ' berhasil ditambahkan!');
    }
    const nextOff = { ...villageOfficials, drivers: nextDrivers };
    setVillageOfficials(nextOff);
    syncToGoogleSheet({ rates, residents, transactions, cashLedger, auditLogs, villageOfficials: nextOff, tripLogs });
  };

  const selectedResident = useMemo(
    () => enrichedResidents.find((r) => r.id === selectedResidentId) || null,
    [enrichedResidents, selectedResidentId]
  );

  return (
    <div className={(themeMode === 'light' ? 'theme-light ' : '') + 'min-h-screen bg-slate-950 text-slate-100 flex flex-col font-sans selection:bg-emerald-500 selection:text-white transition-colors duration-200'}>
      {/* CSS Otomatis untuk Mode Terang (Default Light Mode) */}
      {themeMode === 'light' && (
        <style>{`
          .theme-light.bg-slate-950, .theme-light .bg-slate-950 { background-color: #f8fafc !important; color: #0f172a !important; }
          .theme-light .bg-slate-900, .theme-light .bg-slate-900\\/95, .theme-light .bg-slate-900\\/90 { background-color: #ffffff !important; color: #0f172a !important; box-shadow: 0 1px 3px rgba(15, 23, 42, 0.06); }
          .theme-light .bg-slate-800, .theme-light .bg-slate-800\\/90, .theme-light .bg-slate-800\\/60, .theme-light .bg-slate-950\\/70 { background-color: #f1f5f9 !important; color: #1e293b !important; }
          .theme-light .border-slate-800, .theme-light .border-slate-800\\/90, .theme-light .border-slate-800\\/80, .theme-light .border-slate-700, .theme-light .border-slate-700\\/80 { border-color: #cbd5e1 !important; }
          .theme-light .divide-slate-800 > :not([hidden]) ~ :not([hidden]), .theme-light .divide-slate-800\\/70 > :not([hidden]) ~ :not([hidden]) { border-color: #e2e8f0 !important; }
          .theme-light .text-white, .theme-light .text-slate-100 { color: #0f172a !important; }
          .theme-light .text-slate-200, .theme-light .text-slate-300 { color: #334155 !important; }
          .theme-light .text-slate-400, .theme-light .text-slate-500 { color: #64748b !important; }
          .theme-light .text-emerald-400, .theme-light .text-emerald-300 { color: #047857 !important; }
          .theme-light .text-sky-400, .theme-light .text-sky-300 { color: #0369a1 !important; }
          .theme-light .text-amber-400, .theme-light .text-amber-300 { color: #b45309 !important; }
          .theme-light .text-rose-400, .theme-light .text-rose-300 { color: #be123c !important; }
          .theme-light .from-emerald-950\\/90 { background: linear-gradient(to right, #ecfdf5, #f0f9ff) !important; border-color: #6ee7b7 !important; }
          .theme-light .bg-emerald-600 .text-white, .theme-light .bg-emerald-600, .theme-light .bg-sky-600, .theme-light .bg-amber-600, .theme-light .bg-rose-600 { color: #ffffff !important; }
        `}</style>
      )}

      {/* Notifikasi Toast */}
      {toast && (
        <div className="fixed bottom-5 right-5 z-50 max-w-md animate-bounce">
          <div
            className={
              'flex items-center gap-3 px-4 py-3 rounded-xl shadow-2xl border ' +
              (toast.type === 'error'
                ? 'bg-rose-950/95 border-rose-500 text-rose-100'
                : toast.type === 'warning'
                ? 'bg-amber-950/95 border-amber-500 text-amber-100'
                : 'bg-emerald-950/95 border-emerald-500 text-emerald-100')
            }
          >
            {toast.type === 'error' ? (
              <XCircle className="w-5 h-5 text-rose-400 shrink-0" />
            ) : toast.type === 'warning' ? (
              <AlertTriangle className="w-5 h-5 text-amber-400 shrink-0" />
            ) : (
              <CheckCircle2 className="w-5 h-5 text-emerald-400 shrink-0" />
            )}
            <span className="text-sm font-medium">{toast.message}</span>
          </div>
        </div>
      )}

      {/* Header Utama */}
      <header className="sticky top-0 z-30 bg-slate-900/95 backdrop-blur border-b border-slate-800 px-4 lg:px-8 py-3">
        <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-4">
          <div className="flex items-center gap-3">
            <div className="w-11 h-11 rounded-xl bg-gradient-to-br from-emerald-500 to-teal-600 flex items-center justify-center shadow-lg shadow-emerald-900/40">
              <Car className="w-6 h-6 text-white" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h1 className="text-lg font-bold tracking-tight text-white">SIMOSD — DESA KALIKURMO</h1>
                <span className="text-[11px] font-semibold uppercase px-2 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                  Tarif Aktif: {formatRupiah(activeRateObj.nominal)} / minggu
                </span>
              </div>
              <p className="text-xs text-slate-400">
                Sistem Informasi Manajemen Iuran Mingguan & Operasional Mobil Siaga Desa Kalikurmo
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2.5">
            {/* Tombol Mode Malam / Mode Terang */}
            <button
              onClick={() => setThemeMode((prev) => (prev === 'dark' ? 'light' : 'dark'))}
              className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-slate-800 hover:bg-slate-700 border border-slate-700 text-xs font-bold text-slate-200 transition-all"
              title="Ganti Mode Tampilan Terang / Malam"
            >
              {themeMode === 'dark' ? (
                <>
                  <Sun className="w-4 h-4 text-amber-400" />
                  <span>Mode Terang</span>
                </>
              ) : (
                <>
                  <Moon className="w-4 h-4 text-indigo-600" />
                  <span>Mode Malam</span>
                </>
              )}
            </button>

            {/* Simulasi Tanggal Sistem */}
            <div className="flex items-center gap-2 bg-slate-800/90 border border-slate-700 rounded-lg px-3 py-1.5">
              <Calendar className="w-4 h-4 text-emerald-400" />
              <div className="text-xs">
                <span className="text-slate-400 mr-1.5">Tgl Sistem:</span>
                <input
                  type="date"
                  value={systemDate}
                  onChange={(e) => e.target.value && setSystemDate(e.target.value)}
                  className="bg-transparent text-white font-semibold focus:outline-none cursor-pointer"
                />
              </div>
            </div>

            {/* Indikator Cloud Google Sheet */}
            {sheetUrl && (
              <button
                onClick={() => fetchFromGoogleSheet(false)}
                className="flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 border border-slate-700 text-xs text-emerald-300"
                title="Klik untuk sinkronisasi ulang dari Google Sheet"
              >
                <Cloud className="w-3.5 h-3.5" />
                <span>{syncStatus === 'SYNCING' ? 'Sinkronisasi...' : 'Sheet Terhubung'}</span>
              </button>
            )}

            {/* Switch Halaman Publik vs Mode Admin */}
            <div className="flex items-center bg-slate-800 p-1 rounded-xl border border-slate-700">
              <button
                onClick={() => setViewMode('PUBLIC')}
                className={
                  'flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ' +
                  (viewMode === 'PUBLIC'
                    ? 'bg-sky-600 text-white shadow'
                    : 'text-slate-400 hover:text-white')
                }
              >
                <Eye className="w-3.5 h-3.5" />
                Beranda Publik
              </button>
              <button
                onClick={handleAdminSwitch}
                className={
                  'flex items-center gap-1.5 px-3 py-1.5 rounded-lg text-xs font-semibold transition-all ' +
                  (viewMode === 'ADMIN'
                    ? 'bg-emerald-600 text-white shadow'
                    : 'text-slate-400 hover:text-white')
                }
              >
                {isAdminUnlocked ? <Unlock className="w-3.5 h-3.5" /> : <Lock className="w-3.5 h-3.5" />}
                Mode Admin
              </button>
            </div>

            {isAdminUnlocked && (
              <button
                onClick={handleLockAdmin}
                className="flex items-center gap-1 px-2.5 py-1.5 rounded-lg bg-rose-500/20 hover:bg-rose-500/30 border border-rose-500/30 text-rose-300 text-xs font-semibold"
                title="Kunci Kembali Dashboard Admin"
              >
                <Lock className="w-3.5 h-3.5" />
                Kunci
              </button>
            )}
          </div>
        </div>

        {/* Navigasi Menu Admin */}
        {viewMode === 'ADMIN' && isAdminUnlocked && (
          <div className="max-w-7xl mx-auto mt-3 pt-2 border-t border-slate-800/80 flex items-center justify-between gap-2 overflow-x-auto">
            <nav className="flex items-center gap-1">
              {[
                { id: 'DASHBOARD', label: 'Dashboard Rekapitulasi', icon: BarChart3 },
                { id: 'WARGA', label: 'Data Iuran Warga', icon: CheckSquare },
                { id: 'ANTAR', label: 'Riwayat Antar Warga & BBM', icon: Navigation },
                { id: 'PROFIL', label: 'Profil Kades & Sopir', icon: Phone },
                { id: 'KAS', label: 'Kas Mobil Siaga', icon: Wallet },
                { id: 'RIWAYAT', label: 'Riwayat Pembayaran', icon: History },
                { id: 'PENGATURAN', label: 'Pengaturan Iuran Mingguan', icon: Settings },
                { id: 'DATABASE', label: 'Link Google Sheet & Kode.gs', icon: Database }
              ].map((item) => {
                const Icon = item.icon;
                const active = activeTab === item.id;
                return (
                  <button
                    key={item.id}
                    onClick={() => setActiveTab(item.id)}
                    className={
                      'flex items-center gap-2 px-3.5 py-2 rounded-lg text-xs font-semibold whitespace-nowrap transition-all ' +
                      (active
                        ? 'bg-emerald-500/15 text-emerald-300 border border-emerald-500/30'
                        : 'text-slate-400 hover:text-slate-200 hover:bg-slate-800/60')
                    }
                  >
                    <Icon className="w-4 h-4" />
                    {item.label}
                  </button>
                );
              })}
            </nav>

            <div className="hidden md:flex items-center gap-2 text-xs text-slate-400 shrink-0">
              <UserCheck className="w-3.5 h-3.5 text-emerald-400" />
              <span>Petugas:</span>
              <select
                value={currentAdmin}
                onChange={(e) => setCurrentAdmin(e.target.value)}
                className="bg-slate-800 text-slate-200 border border-slate-700 rounded px-2 py-1 text-xs focus:outline-none"
              >
                <option value="Admin Budi (Bendahara)">Admin Budi (Bendahara)</option>
                <option value="Admin Siti (Sekretaris)">Admin Siti (Sekretaris)</option>
                <option value="Kepala Desa Kalikurmo">Kepala Desa Kalikurmo</option>
              </select>
            </div>
          </div>
        )}
      </header>

      {/* Konten Utama */}
      <main className="flex-1 max-w-7xl w-full mx-auto px-4 lg:px-8 py-6">
        {viewMode === 'PUBLIC' || !isAdminUnlocked ? (
          <PublicPortalView
            enrichedResidents={enrichedResidents}
            dashboardStats={dashboardStats}
            activeRateObj={activeRateObj}
            dusunOptions={dusunOptions}
            filterDusun={filterDusun}
            setFilterDusun={setFilterDusun}
            filterRT={filterRT}
            setFilterRT={setFilterRT}
            filterRW={filterRW}
            setFilterRW={setFilterRW}
            filterStatus={filterStatus}
            setFilterStatus={setFilterStatus}
            searchQuery={searchQuery}
            setSearchQuery={setSearchQuery}
            systemDate={systemDate}
            villageOfficials={villageOfficials}
            tripLogs={tripLogs}
          />
        ) : (
          <>
            {activeTab === 'DASHBOARD' && (
              <DashboardView
                stats={dashboardStats}
                enrichedResidents={enrichedResidents}
                transactions={transactions}
                dusunOptions={dusunOptions}
                filterDusun={filterDusun}
                setFilterDusun={setFilterDusun}
                filterRT={filterRT}
                setFilterRT={setFilterRT}
                filterRW={filterRW}
                setFilterRW={setFilterRW}
                filterStatus={filterStatus}
                setFilterStatus={setFilterStatus}
                filterYear={filterYear}
                setFilterYear={setFilterYear}
                filterMonth={filterMonth}
                setFilterMonth={setFilterMonth}
                filterDate={filterDate}
                setFilterDate={setFilterDate}
                onQuickPay={(resId, preset) => setPaymentModal({ open: true, residentId: resId, presetType: preset })}
                onOpenDetail={(resId) => setSelectedResidentId(resId)}
                onSwitchToWarga={() => setActiveTab('WARGA')}
              />
            )}

            {activeTab === 'WARGA' && (
              <DataIuranWargaView
                residents={filteredResidents}
                allResidentsCount={enrichedResidents.length}
                dusunOptions={dusunOptions}
                searchQuery={searchQuery}
                setSearchQuery={setSearchQuery}
                filterDusun={filterDusun}
                setFilterDusun={setFilterDusun}
                filterRT={filterRT}
                setFilterRT={setFilterRT}
                filterRW={filterRW}
                setFilterRW={setFilterRW}
                filterStatus={filterStatus}
                setFilterStatus={setFilterStatus}
                onOpenDetail={(id) => setSelectedResidentId(id)}
                onOpenPayment={(id, preset) => setPaymentModal({ open: true, residentId: id, presetType: preset })}
                onOpenAddResident={() => setAddResidentModalOpen(true)}
                onProtectedChecklistClick={() =>
                  showToast(
                    'Proteksi Sistem: Status checklist tidak dapat diubah manual tanpa transaksi pembayaran yang sah!',
                    'warning'
                  )
                }
              />
            )}

            {activeTab === 'ANTAR' && (
              <RiwayatAntarWargaView
                tripLogs={tripLogs}
                residents={residents}
                drivers={villageOfficials.drivers}
                dusunOptions={dusunOptions}
                systemDate={systemDate}
                onAddTrip={handleAddTripLog}
              />
            )}

            {activeTab === 'PROFIL' && (
              <ProfilKadesSopirView
                villageOfficials={villageOfficials}
                tripLogs={tripLogs}
                dusunOptions={dusunOptions}
                onUpdateKades={handleUpdateKades}
                onSaveDriver={handleSaveDriver}
                showToast={showToast}
              />
            )}

            {activeTab === 'KAS' && (
              <KasMobilSiagaView
                cashLedger={cashLedger}
                stats={dashboardStats}
                onOpenReceipt={(refId) => {
                  const trx = transactions.find((t) => t.id === refId);
                  if (trx) setReceiptModalTrx(trx);
                }}
              />
            )}

            {activeTab === 'RIWAYAT' && (
              <RiwayatPembayaranView
                transactions={transactions}
                dusunOptions={dusunOptions}
                onPrintReceipt={(trx) => setReceiptModalTrx(trx)}
                onCorrectTrx={(trx) => setCorrectionModalTrx(trx)}
                onOpenResident={(resId) => setSelectedResidentId(resId)}
              />
            )}

            {activeTab === 'PENGATURAN' && (
              <PengaturanIuranView
                rates={rates}
                activeRate={activeRateObj}
                auditLogs={auditLogs}
                systemDate={systemDate}
                onAddRate={handleAddRate}
              />
            )}

            {activeTab === 'DATABASE' && (
              <DatabaseSheetView
                sheetUrl={sheetUrl}
                setSheetUrl={setSheetUrl}
                syncStatus={syncStatus}
                lastSyncTime={lastSyncTime}
                onSyncNow={() => syncToGoogleSheet()}
                onFetchNow={() => fetchFromGoogleSheet(false)}
                copiedCode={copiedCode}
                onCopyCode={() => {
                  navigator.clipboard.writeText(KODE_GS_SCRIPT);
                  setCopiedCode(true);
                  showToast('Kode.gs berhasil disalin!');
                  setTimeout(() => setCopiedCode(false), 3000);
                }}
              />
            )}
          </>
        )}
      </main>

      {/* Modal Password Login Admin ("siaga") */}
      {passwordModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
          <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-md w-full p-6 shadow-2xl">
            <div className="flex items-center justify-between mb-4">
              <div className="flex items-center gap-2.5">
                <div className="w-10 h-10 rounded-xl bg-emerald-500/20 border border-emerald-500/40 flex items-center justify-center text-emerald-400">
                  <KeyRound className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="text-base font-bold text-white">Kunci Akses Dashboard Admin</h3>
                  <p className="text-xs text-slate-400">Pemerintah Desa Kalikurmo</p>
                </div>
              </div>
              <button onClick={() => setPasswordModalOpen(false)} className="text-slate-400 hover:text-white text-sm">
                ✕
              </button>
            </div>

            <form onSubmit={handleVerifyPassword} className="space-y-4">
              <div>
                <label className="block text-xs text-slate-300 mb-1.5 font-medium">
                  Masukkan Password Admin Mobil Siaga:
                </label>
                <input
                  type="password"
                  autoFocus
                  value={passwordInput}
                  onChange={(e) => {
                    setPasswordInput(e.target.value);
                    setPasswordError('');
                  }}
                  placeholder="Ketik password admin..."
                  className="w-full bg-slate-950 border border-slate-700 rounded-xl px-3.5 py-2.5 text-sm text-white focus:border-emerald-500 focus:outline-none"
                />
                {passwordError && <p className="text-xs text-rose-400 mt-1.5 font-medium">{passwordError}</p>}
              </div>

              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setPasswordModalOpen(false)}
                  className="px-4 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold shadow-lg"
                >
                  Buka Dashboard Admin
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* Modal Detail Warga */}
      {selectedResident && (
        <ResidentDetailModal
          resident={selectedResident}
          activeRate={activeRateObj}
          onClose={() => setSelectedResidentId(null)}
          onOpenPayment={(presetType) => setPaymentModal({ open: true, residentId: selectedResident.id, presetType })}
          onPrintReceipt={(trx) => setReceiptModalTrx(trx)}
          onCorrectTrx={(trx) => setCorrectionModalTrx(trx)}
        />
      )}

      {/* Modal Input Pembayaran Iuran Mingguan */}
      {paymentModal.open && (
        <PaymentFormModal
          resident={enrichedResidents.find((r) => r.id === paymentModal.residentId) || enrichedResidents[0]}
          allResidents={enrichedResidents}
          presetType={paymentModal.presetType}
          rates={rates}
          systemDate={systemDate}
          onClose={() => setPaymentModal({ open: false, residentId: null, presetType: '1_BULAN' })}
          onSave={handleSavePayment}
        />
      )}

      {/* Modal Cetak Kuitansi */}
      {receiptModalTrx && <ReceiptPrintModal trx={receiptModalTrx} onClose={() => setReceiptModalTrx(null)} />}

      {/* Modal Koreksi Transaksi */}
      {correctionModalTrx && (
        <CorrectionDialogModal
          trx={correctionModalTrx}
          onClose={() => setCorrectionModalTrx(null)}
          onConfirm={handleCorrectTransaction}
        />
      )}

      {/* Modal Tambah Warga Baru */}
      {addResidentModalOpen && (
        <AddResidentModal
          dusunOptions={dusunOptions}
          onClose={() => setAddResidentModalOpen(false)}
          onSave={handleAddResident}
        />
      )}

      {/* Footer & Watermark Pembuat Aplikasi (powered by.pakseh) */}
      <footer className="relative overflow-hidden border-t border-slate-800/90 bg-slate-900/90 py-8 px-4 lg:px-8 mt-12">
        <div className="pointer-events-none select-none flex items-center justify-center opacity-10 mb-3">
          <span className="text-3xl sm:text-5xl md:text-6xl font-black tracking-widest text-emerald-400 uppercase text-center">
            PEMERINTAH DESA KALIKURMO
          </span>
        </div>
        <div className="max-w-7xl mx-auto flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-400">
          <div className="flex items-center gap-2">
            <span className="text-xs text-slate-400 tracking-wide">powered by.</span>
            <span
              style={{
                fontFamily: "'Brush Script MT', 'Segoe Script', 'Lucida Handwriting', 'Georgia', cursive, serif",
                color: '#9ca3af'
              }}
              className="text-2xl font-semibold italic tracking-wider select-none"
            >
              pakseh
            </span>
          </div>
          <div className="text-center sm:text-right">
            <span className="inline-block px-3 py-1 rounded-full bg-emerald-500/15 border border-emerald-500/30 text-emerald-300 font-bold text-[11px]">
              © {new Date().getFullYear()} Pemerintah Desa Kalikurmo
            </span>
          </div>
        </div>
      </footer>
    </div>
  );
}

// ============================================================================
// KOMPONEN BADGE STATUS WARNA (HIJAU, BIRU, KUNING, MERAH)
// ============================================================================
function StatusBadge({ statusKey, label }) {
  const styles = {
    green: 'bg-emerald-500/15 text-emerald-300 border-emerald-500/40',
    blue: 'bg-sky-500/15 text-sky-300 border-sky-500/40',
    yellow: 'bg-amber-500/15 text-amber-300 border-amber-500/40',
    red: 'bg-rose-500/15 text-rose-300 border-rose-500/40'
  };
  const dotStyles = {
    green: 'bg-emerald-400',
    blue: 'bg-sky-400',
    yellow: 'bg-amber-400',
    red: 'bg-rose-400'
  };
  const colorKey =
    statusKey === 'LUNAS'
      ? 'green'
      : statusKey === 'LUNAS_MENDATANG'
      ? 'blue'
      : statusKey === 'SEBAGIAN'
      ? 'yellow'
      : 'red';

  return (
    <span className={'inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-bold border ' + styles[colorKey]}>
      <span className={'w-2 h-2 rounded-full ' + dotStyles[colorKey]} />
      {label}
    </span>
  );
}

// ============================================================================
// 1. HALAMAN PUBLIK / BERANDA TRANSPARANSI WARGA DESA KALIKURMO
// ============================================================================
function PublicPortalView({
  enrichedResidents,
  dashboardStats,
  activeRateObj,
  dusunOptions,
  filterDusun,
  setFilterDusun,
  filterRT,
  setFilterRT,
  filterRW,
  setFilterRW,
  filterStatus,
  setFilterStatus,
  searchQuery,
  setSearchQuery,
  villageOfficials,
  tripLogs
}) {
  const filteredPublic = enrichedResidents.filter((r) => {
    const q = searchQuery.toLowerCase();
    const matchName = r.nama.toLowerCase().includes(q) || r.dusun.toLowerCase().includes(q);
    const matchDusun = filterDusun === 'ALL' || r.dusun === filterDusun;
    const matchRT = filterRT === 'ALL' || r.rt === filterRT;
    const matchRW = filterRW === 'ALL' || (r.rw || '01') === filterRW;

    let matchStatus = true;
    if (filterStatus === 'LUNAS_BULAN_INI') matchStatus = r.computed.isLunasBulanIni;
    else if (filterStatus === 'LUNAS_BULAN_DEPAN') matchStatus = r.computed.isLunasBulanDepan;
    else if (filterStatus === 'BELUM_BAYAR') matchStatus = r.computed.isBelumBayar;
    else if (filterStatus !== 'ALL') matchStatus = r.computed.statusKey === filterStatus;

    return matchName && matchDusun && matchRT && matchRW && matchStatus;
  });

  const rekapDusun = useMemo(() => {
    return dusunOptions.map((dName) => {
      const wargaDusun = enrichedResidents.filter((r) => r.dusun === dName);
      const lunas = wargaDusun.filter((r) => r.computed.isChecked).length;
      const totalUang = wargaDusun.reduce((acc, r) => acc + r.computed.totalNominal, 0);
      return { dusun: dName, totalWarga: wargaDusun.length, lunas, totalUang };
    });
  }, [enrichedResidents, dusunOptions]);

  return (
    <div className="space-y-6">
      {/* Banner Transparansi Publik Desa Kalikurmo */}
      <div className="bg-gradient-to-r from-emerald-950/90 via-slate-900 to-sky-950/90 border border-emerald-500/30 rounded-2xl p-6 shadow-xl">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-6">
          <div className="space-y-2">
            <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full text-xs font-bold bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
              <Eye className="w-3.5 h-3.5" />
              PORTAL TRANSPARANSI MASYARAKAT — DESA KALIKURMO
            </span>
            <h2 className="text-2xl font-extrabold text-white">
              Laporan Publik Iuran Mingguan & Layanan Mobil Siaga Desa Kalikurmo
            </h2>
            <p className="text-sm text-slate-300 max-w-2xl">
              Masyarakat Desa Kalikurmo dapat memeriksa status pelunasan iuran mingguan, menghubungi Sopir Siaga 24 Jam, serta melihat riwayat pemakaian mobil dan kas secara terbuka.
            </p>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-3 gap-3 shrink-0">
            <div className="bg-slate-900/90 border border-slate-700/80 rounded-xl p-3.5">
              <span className="text-xs text-slate-400 block">Saldo Kas Mobil Siaga</span>
              <span className="text-lg font-extrabold text-emerald-400">
                {formatRupiah(dashboardStats.saldoKas)}
              </span>
              <span className="text-[11px] text-slate-400 block mt-0.5">Kas Aktif Siap Pakai</span>
            </div>
            <div className="bg-slate-900/90 border border-slate-700/80 rounded-xl p-3.5">
              <span className="text-xs text-slate-400 block">Tarif Iuran Mingguan</span>
              <span className="text-lg font-extrabold text-sky-400">
                {formatRupiah(activeRateObj.nominal)} / minggu
              </span>
              <span className="text-[11px] text-slate-400 block mt-0.5">Keputusan Musdes</span>
            </div>
            <div className="bg-slate-900/90 border border-slate-700/80 rounded-xl p-3.5 col-span-2 sm:col-span-1">
              <span className="text-xs text-slate-400 block">Warga Sudah Lunas</span>
              <span className="text-lg font-extrabold text-white">
                {dashboardStats.sudahLunas} / {dashboardStats.totalWarga} Warga
              </span>
              <span className="text-[11px] text-emerald-400 block mt-0.5">Tercatat Otomatis</span>
            </div>
          </div>
        </div>
      </div>

      {/* Kontak Darurat Sopir & Kepala Desa Kalikurmo */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-4">
        {/* Kartu Kepala Desa */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between mb-3">
              <span className="text-[11px] font-bold uppercase px-2.5 py-0.5 rounded-full bg-amber-500/20 text-amber-300 border border-amber-500/30">
                Penanggung Jawab Utama
              </span>
              <Award className="w-5 h-5 text-amber-400" />
            </div>
            <div className="flex items-center gap-3.5 mb-3">
              {villageOfficials.kades.fotoUrl ? (
                <img
                  src={villageOfficials.kades.fotoUrl}
                  alt={villageOfficials.kades.nama}
                  className="w-16 h-16 rounded-2xl object-cover border-2 border-amber-400/60 shrink-0"
                />
              ) : (
                <div className="w-16 h-16 rounded-2xl bg-amber-500/20 border border-amber-500/40 flex items-center justify-center text-amber-300 font-extrabold text-xl shrink-0">
                  KD
                </div>
              )}
              <div>
                <h3 className="text-base font-bold text-white">{villageOfficials.kades.nama}</h3>
                <p className="text-xs text-amber-400 font-semibold">{villageOfficials.kades.jabatan}</p>
                <p className="text-[11px] text-slate-400">{villageOfficials.kades.unitMobil}</p>
              </div>
            </div>
            <p className="text-xs text-slate-300 italic bg-slate-950/70 p-3 rounded-xl border border-slate-800">
              "{villageOfficials.kades.pesanSiaga}"
            </p>
          </div>
          <div className="mt-4 pt-3 border-t border-slate-800 flex items-center justify-between">
            <span className="text-xs text-slate-400">Kontak Kades: {villageOfficials.kades.noHP}</span>
            <a
              href={'https://wa.me/' + villageOfficials.kades.noHP.replace(/^0/, '62')}
              target="_blank"
              rel="noreferrer"
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold"
            >
              <Phone className="w-3.5 h-3.5" />
              Hubungi Kades
            </a>
          </div>
        </div>

        {/* Kartu Sopir Mobil Siaga */}
        {villageOfficials.drivers.map((drv) => (
          <div key={drv.id} className="bg-slate-900 border border-slate-800 rounded-2xl p-5 flex flex-col justify-between">
            <div>
              <div className="flex items-center justify-between mb-3">
                <span className="text-[11px] font-bold uppercase px-2.5 py-0.5 rounded-full bg-emerald-500/20 text-emerald-300 border border-emerald-500/30">
                  Sopir Mobil Siaga ({drv.statusSiaga})
                </span>
                <span className="text-xs font-mono text-slate-400">{drv.platMobil}</span>
              </div>
              <div className="flex items-center gap-3.5 mb-3">
                {drv.fotoUrl ? (
                  <img
                    src={drv.fotoUrl}
                    alt={drv.nama}
                    className="w-16 h-16 rounded-2xl object-cover border-2 border-emerald-400/60 shrink-0"
                  />
                ) : (
                  <div className="w-16 h-16 rounded-2xl bg-emerald-500/20 border border-emerald-500/40 flex items-center justify-center text-emerald-300 font-extrabold text-xl shrink-0">
                    {drv.nama.slice(0, 2).toUpperCase()}
                  </div>
                )}
                <div>
                  <h3 className="text-base font-bold text-white">
                    {drv.nama}{' '}
                    {drv.panggilan ? (
                      <span className="text-xs text-slate-400">({drv.panggilan})</span>
                    ) : null}
                  </h3>
                  <p className="text-xs text-emerald-400 font-semibold">{drv.shift}</p>
                  <p className="text-[11px] text-slate-400">
                    Domisili: {drv.dusun} • RT {drv.rt} / RW {drv.rw || '01'}
                  </p>
                </div>
              </div>
              <p className="text-xs text-slate-400 bg-slate-950/70 p-2.5 rounded-xl border border-slate-800">
                Pengalaman: {drv.pengalaman}
              </p>
            </div>
            <div className="mt-4 pt-3 border-t border-slate-800 flex items-center justify-between">
              <span className="text-xs font-bold text-white">{drv.noHP}</span>
              <div className="flex items-center gap-2">
                <a
                  href={'tel:' + drv.noHP}
                  className="px-3 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-xs font-semibold"
                >
                  Telepon
                </a>
                <a
                  href={'https://wa.me/' + drv.noHP.replace(/^0/, '62')}
                  target="_blank"
                  rel="noreferrer"
                  className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold"
                >
                  <Phone className="w-3.5 h-3.5" />
                  WhatsApp Sopir
                </a>
              </div>
            </div>
          </div>
        ))}
      </div>

      {/* Tombol Filter Cepat Kategori Pelunasan + Dusun + RT + RW */}
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <h3 className="text-sm font-bold text-white flex items-center gap-2">
            <Filter className="w-4 h-4 text-emerald-400" />
            Filter Cepat Status Pelunasan Warga (Dusun / RT / RW)
          </h3>
          <div className="flex flex-wrap gap-2">
            {[
              { id: 'ALL', label: 'Semua Warga (' + enrichedResidents.length + ')' },
              { id: 'LUNAS_BULAN_INI', label: 'Sudah Lunas Bulan Ini (' + dashboardStats.lunasBulanIniCount + ')' },
              { id: 'LUNAS_BULAN_DEPAN', label: 'Sudah Lunas Bulan Depan (' + dashboardStats.lunasBulanDepanCount + ')' },
              { id: 'BELUM_BAYAR', label: 'Belum Bayar / Nunggak (' + dashboardStats.belumLunas + ')' }
            ].map((btn) => (
              <button
                key={btn.id}
                onClick={() => setFilterStatus(btn.id)}
                className={
                  'px-3 py-1.5 rounded-xl text-xs font-bold border transition-all ' +
                  (filterStatus === btn.id
                    ? 'bg-emerald-600 text-white border-emerald-500 shadow'
                    : 'bg-slate-950 text-slate-300 border-slate-800 hover:border-slate-700')
                }
              >
                {btn.label}
              </button>
            ))}
          </div>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3">
          <div className="relative lg:col-span-2">
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Cari nama warga Desa Kalikurmo..."
              className="w-full bg-slate-950 border border-slate-800 rounded-xl pl-10 pr-4 py-2 text-xs text-white focus:border-emerald-500 focus:outline-none"
            />
          </div>

          <select
            value={filterDusun}
            onChange={(e) => setFilterDusun(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua Dusun</option>
            {dusunOptions.map((d) => (
              <option key={d} value={d}>{d}</option>
            ))}
          </select>

          <select
            value={filterRT}
            onChange={(e) => setFilterRT(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua RT (01 - 15)</option>
            {RT_RW_LIST.map((rt) => (
              <option key={rt} value={rt}>RT {rt}</option>
            ))}
          </select>

          <select
            value={filterRW}
            onChange={(e) => setFilterRW(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua RW (01 - 15)</option>
            {RT_RW_LIST.map((rw) => (
              <option key={rw} value={rw}>RW {rw}</option>
            ))}
          </select>
        </div>

        {/* Tabel Publik Status Pembayaran Warga */}
        <div className="overflow-x-auto rounded-xl border border-slate-800">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-950 text-[11px] uppercase text-slate-400 border-b border-slate-800">
                <th className="py-3 px-4">No</th>
                <th className="py-3 px-4">Checklist</th>
                <th className="py-3 px-4">Nama Warga</th>
                <th className="py-3 px-4">Dusun</th>
                <th className="py-3 px-4">RT / RW</th>
                <th className="py-3 px-4">Status</th>
                <th className="py-3 px-4">Terakhir Lunas Sampai</th>
                <th className="py-3 px-4">Keterangan / Minggu Nunggak</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800/70 text-xs">
              {filteredPublic.map((w, idx) => (
                <tr key={w.id} className="hover:bg-slate-800/40">
                  <td className="py-3 px-4 font-mono text-slate-400">{idx + 1}</td>
                  <td className="py-3 px-4">
                    {w.computed.isChecked ? (
                      <span className="inline-flex items-center justify-center w-6 h-6 rounded-md bg-emerald-500/20 text-emerald-400 border border-emerald-500/40">
                        ✓
                      </span>
                    ) : (
                      <span className="inline-flex items-center justify-center w-6 h-6 rounded-md bg-slate-800 text-slate-500 border border-slate-700">
                        -
                      </span>
                    )}
                  </td>
                  <td className="py-3 px-4 font-bold text-white">{w.nama}</td>
                  <td className="py-3 px-4 text-slate-300">{w.dusun}</td>
                  <td className="py-3 px-4 text-slate-300">
                    RT {w.rt} / RW {w.rw || '01'}
                  </td>
                  <td className="py-3 px-4">
                    <StatusBadge statusKey={w.computed.statusKey} label={w.computed.statusLabel} />
                  </td>
                  <td className="py-3 px-4 font-semibold text-white">
                    {formatDateID(w.computed.lunasSampai)}
                  </td>
                  <td className="py-3 px-4">
                    {w.computed.mingguNunggak > 0 ? (
                      <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-md bg-rose-500/15 border border-rose-500/30 text-rose-300 font-bold text-[11px]">
                        Nunggak {w.computed.mingguNunggak} Minggu ({formatRupiah(w.computed.nominalTunggakan)})
                      </span>
                    ) : (
                      <span className="text-emerald-300 font-medium">{w.computed.keteranganTunggakan}</span>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Rekap Per Dusun & Riwayat Pemakaian Mobil Siaga */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5">
          <h3 className="text-sm font-bold text-white mb-3 flex items-center gap-2">
            <Building2 className="w-4 h-4 text-emerald-400" />
            Rekapitulasi Iuran Mingguan Berdasarkan Dusun Desa Kalikurmo
          </h3>
          <div className="space-y-3">
            {rekapDusun.map((d) => (
              <div key={d.dusun} className="bg-slate-950 border border-slate-800 rounded-xl p-3.5 flex items-center justify-between">
                <div>
                  <p className="text-sm font-bold text-white">{d.dusun}</p>
                  <p className="text-xs text-slate-400">
                    Warga Lunas: <strong className="text-emerald-400">{d.lunas}</strong> dari {d.totalWarga} Warga
                  </p>
                </div>
                <div className="text-right">
                  <p className="text-sm font-extrabold text-emerald-400">{formatRupiah(d.totalUang)}</p>
                  <p className="text-[11px] text-slate-500">Total Iuran Terkumpul</p>
                </div>
              </div>
            ))}
          </div>
        </div>

        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5">
          <h3 className="text-sm font-bold text-white mb-3 flex items-center gap-2">
            <Navigation className="w-4 h-4 text-sky-400" />
            Transparansi Pemakaian Mobil Siaga (Antar Warga)
          </h3>
          <div className="space-y-3 max-h-72 overflow-y-auto pr-1">
            {tripLogs.map((trip) => (
              <div key={trip.id} className="bg-slate-950 border border-slate-800 rounded-xl p-3.5 space-y-1.5 text-xs">
                <div className="flex items-center justify-between">
                  <span className="font-bold text-white">
                    {formatDateID(trip.tanggal)} • {trip.waktu}
                  </span>
                  <span className="px-2 py-0.5 rounded bg-sky-500/20 text-sky-300 font-semibold">
                    Sopir: {trip.driverName}
                  </span>
                </div>
                <p className="text-slate-200">
                  Mengantar: <strong>{trip.residentName}</strong> ({trip.dusun} RT {trip.rt}/RW {trip.rw || '01'}) →{' '}
                  <strong className="text-emerald-300">{trip.tujuan}</strong>
                </p>
                <div className="flex items-center justify-between pt-1 border-t border-slate-800/80 text-[11px]">
                  <span className="text-rose-400 font-semibold">
                    BBM/Operasional Keluar: {formatRupiah(trip.uangKeluar)}
                  </span>
                  <span className="text-emerald-400 font-semibold">
                    Infaq Masuk: {formatRupiah(trip.uangMasuk)}
                  </span>
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// 2. DASHBOARD REKAPITULASI IURAN MINGGUAN (ADMIN)
// ============================================================================
function DashboardView({
  stats,
  enrichedResidents,
  dusunOptions,
  filterDusun,
  setFilterDusun,
  filterRT,
  setFilterRT,
  filterRW,
  setFilterRW,
  setFilterStatus,
  filterYear,
  setFilterYear,
  filterMonth,
  setFilterMonth,
  filterDate,
  setFilterDate,
  onQuickPay,
  onOpenDetail,
  onSwitchToWarga
}) {
  const statCards = [
    { title: '1. Total Warga Terdaftar', value: stats.totalWarga + ' Warga', sub: 'Database Warga Kalikurmo', color: 'text-white' },
    { title: '2. Warga Sudah Lunas', value: stats.sudahLunas + ' Warga', sub: 'Hijau & Biru (Bebas Kewajiban)', color: 'text-emerald-400' },
    { title: '3. Warga Belum Lunas / Nunggak', value: stats.belumLunas + ' Warga', sub: 'Merah (Perlu Tindak Lanjut)', color: 'text-rose-400' },
    { title: '4. Membayar Sebagian', value: stats.bayarSebagian + ' Warga', sub: 'Kuning (1 - 3 Minggu)', color: 'text-amber-400' },
    { title: '5. Melunasi ≥ 1 Bulan (4 Minggu)', value: stats.lunas1Bulan + ' Warga', sub: 'Paket Bulanan / Kuartal', color: 'text-sky-400' },
    { title: '6. Melunasi 1 Tahun (52 Minggu)', value: stats.lunas1Tahun + ' Warga', sub: 'Pelunasan 52 Minggu Penuh', color: 'text-indigo-400' },
    { title: '7. Total Pemasukan Iuran', value: formatRupiah(stats.totalPemasukanIuran), sub: 'Sesuai Filter Aktif', color: 'text-emerald-400' },
    { title: '8. Total Pembayaran Di Muka', value: formatRupiah(stats.totalPembayaranDiMuka), sub: 'Periode Mendatang', color: 'text-sky-400' },
    { title: '9. Nominal Iuran Mingguan', value: formatRupiah(stats.nominalMingguanBerlaku) + '/minggu', sub: 'Tarif Mingguan Aktif', color: 'text-amber-300' }
  ];

  const chartPerDusun = useMemo(() => {
    return dusunOptions.map((d) => {
      const total = stats.validTransactions
        .filter((t) => t.dusun === d)
        .reduce((acc, t) => acc + Number(t.nominal), 0);
      return { label: d, total };
    });
  }, [dusunOptions, stats.validTransactions]);

  const maxDusunVal = Math.max(...chartPerDusun.map((c) => c.total), 1);

  return (
    <div className="space-y-6">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 space-y-3">
        <div className="flex flex-wrap items-center justify-between gap-2">
          <span className="text-xs font-bold uppercase tracking-wider text-emerald-400 flex items-center gap-1.5">
            <Filter className="w-4 h-4" />
            Filter Rekapitulasi Dashboard (Tanggal, Bulan, Tahun, Dusun, RT, RW, Kategori Lunas)
          </span>
          <div className="flex flex-wrap gap-1.5">
            {[
              { id: 'ALL', label: 'Semua Status' },
              { id: 'LUNAS_BULAN_INI', label: 'Lunas Bulan Ini (' + stats.lunasBulanIniCount + ')' },
              { id: 'LUNAS_BULAN_DEPAN', label: 'Lunas Bulan Depan (' + stats.lunasBulanDepanCount + ')' },
              { id: 'BELUM_BAYAR', label: 'Belum Bayar (' + stats.belumLunas + ')' }
            ].map((b) => (
              <button
                key={b.id}
                onClick={() => {
                  setFilterStatus(b.id);
                  onSwitchToWarga();
                }}
                className="px-2.5 py-1 rounded-lg text-[11px] font-bold bg-slate-800 hover:bg-emerald-600 text-slate-200 hover:text-white transition-all"
              >
                Lihat {b.label} →
              </button>
            ))}
          </div>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2.5">
          <input
            type="date"
            value={filterDate}
            onChange={(e) => setFilterDate(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          />
          <select
            value={filterMonth}
            onChange={(e) => setFilterMonth(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua Bulan</option>
            <option value="01">Januari</option>
            <option value="02">Februari</option>
            <option value="03">Maret</option>
            <option value="04">April</option>
            <option value="05">Mei</option>
            <option value="06">Juni</option>
            <option value="07">Juli</option>
            <option value="08">Agustus</option>
            <option value="09">September</option>
            <option value="10">Oktober</option>
            <option value="11">November</option>
            <option value="12">Desember</option>
          </select>
          <select
            value={filterYear}
            onChange={(e) => setFilterYear(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua Tahun</option>
            <option value="2026">2026</option>
            <option value="2027">2027</option>
            <option value="2028">2028</option>
          </select>
          <select
            value={filterDusun}
            onChange={(e) => setFilterDusun(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua Dusun</option>
            {dusunOptions.map((d) => (
              <option key={d} value={d}>{d}</option>
            ))}
          </select>
          <select
            value={filterRT}
            onChange={(e) => setFilterRT(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua RT (01-15)</option>
            {RT_RW_LIST.map((rt) => (
              <option key={rt} value={rt}>RT {rt}</option>
            ))}
          </select>
          <select
            value={filterRW}
            onChange={(e) => setFilterRW(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua RW (01-15)</option>
            {RT_RW_LIST.map((rw) => (
              <option key={rw} value={rw}>RW {rw}</option>
            ))}
          </select>
        </div>
      </div>

      {/* 9 Kartu Statistik Utama */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
        {statCards.map((c, idx) => (
          <div key={idx} className="bg-slate-900 border border-slate-800 rounded-2xl p-4 hover:border-slate-700 transition-all">
            <span className="text-xs text-slate-400 font-medium">{c.title}</span>
            <p className={'text-2xl font-extrabold mt-1 ' + c.color}>{c.value}</p>
            <span className="text-[11px] text-slate-500 mt-1 block">{c.sub}</span>
          </div>
        ))}
      </div>

      {/* Grafik Visualisasi Pemasukan */}
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
          <h3 className="text-sm font-bold text-white flex items-center gap-2">
            <BarChart3 className="w-4 h-4 text-emerald-400" />
            Grafik Pemasukan Iuran Mingguan Berdasarkan Dusun
          </h3>
          <div className="space-y-3">
            {chartPerDusun.map((item) => {
              const pct = Math.round((item.total / maxDusunVal) * 100);
              return (
                <div key={item.label} className="space-y-1">
                  <div className="flex justify-between text-xs">
                    <span className="font-semibold text-slate-200">{item.label}</span>
                    <span className="font-bold text-emerald-400">{formatRupiah(item.total)}</span>
                  </div>
                  <div className="w-full h-2.5 bg-slate-800 rounded-full overflow-hidden">
                    <div className="h-full bg-emerald-500 rounded-full" style={{ width: Math.max(pct, 4) + '%' }} />
                  </div>
                </div>
              );
            })}
          </div>
        </div>

        {/* Daftar Warga Menunggak untuk Aksi Cepat */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3">
          <h3 className="text-sm font-bold text-white flex items-center gap-2">
            <AlertTriangle className="w-4 h-4 text-rose-400" />
            Daftar Warga Menunggak Mingguan / Belum Lunas (Aksi Cepat)
          </h3>
          <div className="space-y-2.5 max-h-64 overflow-y-auto pr-1">
            {enrichedResidents
              .filter((r) => !r.computed.isChecked)
              .map((w) => (
                <div
                  key={w.id}
                  className="bg-slate-950 border border-slate-800 rounded-xl p-3 flex flex-wrap items-center justify-between gap-2"
                >
                  <div>
                    <p className="text-xs font-bold text-white">
                      {w.nama}{' '}
                      <span className="text-slate-400 font-normal">
                        ({w.dusun} • RT {w.rt}/RW {w.rw || '01'})
                      </span>
                    </p>
                    <p className="text-[11px] text-rose-400 font-semibold">
                      {w.computed.keteranganTunggakan}
                    </p>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <button
                      onClick={() => onQuickPay(w.id, '1_BULAN')}
                      className="px-2.5 py-1 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white text-[11px] font-bold"
                    >
                      + Lunasi
                    </button>
                    <button
                      onClick={() => onOpenDetail(w.id)}
                      className="px-2.5 py-1 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 text-[11px]"
                    >
                      Detail
                    </button>
                  </div>
                </div>
              ))}
          </div>
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// 3. DATA IURAN WARGA & CHECKLIST OTOMATIS (MINGGUAN)
// ============================================================================
function DataIuranWargaView({
  residents,
  dusunOptions,
  searchQuery,
  setSearchQuery,
  filterDusun,
  setFilterDusun,
  filterRT,
  setFilterRT,
  filterRW,
  setFilterRW,
  filterStatus,
  setFilterStatus,
  onOpenDetail,
  onOpenPayment,
  onOpenAddResident,
  onProtectedChecklistClick
}) {
  return (
    <div className="space-y-5">
      <div className="flex flex-wrap items-center justify-between gap-4">
        <div>
          <h2 className="text-xl font-bold text-white">DATA IURAN MINGGUAN WARGA — DESA KALIKURMO</h2>
          <p className="text-xs text-slate-400">
            Checklist terkunci otomatis berdasarkan transaksi sah di database. Klik nama warga untuk detail & cetak kuitansi.
          </p>
        </div>
        <div className="flex items-center gap-2">
          <button
            onClick={onOpenAddResident}
            className="flex items-center gap-1.5 px-4 py-2 rounded-xl bg-slate-800 hover:bg-slate-700 border border-slate-700 text-xs font-bold text-white"
          >
            <UserPlus className="w-4 h-4 text-emerald-400" />
            Tambah Warga
          </button>
          <button
            onClick={() => onOpenPayment(residents[0]?.id || 'WRG-001', '1_BULAN')}
            className="flex items-center gap-1.5 px-4 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-xs font-bold text-white shadow-lg"
          >
            <Plus className="w-4 h-4" />
            Catat Pembayaran Mingguan
          </button>
        </div>
      </div>

      {/* Filter Kategori Cepat & Wilayah */}
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4 space-y-3">
        <div className="flex flex-wrap gap-2">
          {[
            { id: 'ALL', label: 'Semua Status' },
            { id: 'LUNAS_BULAN_INI', label: 'Sudah Lunas Bulan Ini' },
            { id: 'LUNAS_BULAN_DEPAN', label: 'Sudah Lunas Bulan Depan' },
            { id: 'BELUM_BAYAR', label: 'Belum Bayar / Nunggak' },
            { id: 'SEBAGIAN', label: 'Bayar Sebagian (Kuning)' }
          ].map((tab) => (
            <button
              key={tab.id}
              onClick={() => setFilterStatus(tab.id)}
              className={
                'px-3 py-1.5 rounded-xl text-xs font-bold border transition-all ' +
                (filterStatus === tab.id
                  ? 'bg-emerald-600 text-white border-emerald-500'
                  : 'bg-slate-950 text-slate-300 border-slate-800 hover:border-slate-700')
              }
            >
              {tab.label}
            </button>
          ))}
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Cari nama warga / ID / Dusun..."
            className="bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-xs text-white"
          />
          <select
            value={filterDusun}
            onChange={(e) => setFilterDusun(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua Dusun</option>
            {dusunOptions.map((d) => (
              <option key={d} value={d}>{d}</option>
            ))}
          </select>
          <select
            value={filterRT}
            onChange={(e) => setFilterRT(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua RT (01 - 15)</option>
            {RT_RW_LIST.map((rt) => (
              <option key={rt} value={rt}>RT {rt}</option>
            ))}
          </select>
          <select
            value={filterRW}
            onChange={(e) => setFilterRW(e.target.value)}
            className="bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-slate-200"
          >
            <option value="ALL">Semua RW (01 - 15)</option>
            {RT_RW_LIST.map((rw) => (
              <option key={rw} value={rw}>RW {rw}</option>
            ))}
          </select>
        </div>
      </div>

      {/* Tabel Checklist Warga */}
      <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-950 text-[11px] uppercase text-slate-400 border-b border-slate-800">
                <th className="py-3.5 px-4">No</th>
                <th className="py-3.5 px-4">Checklist</th>
                <th className="py-3.5 px-4">Nama Warga</th>
                <th className="py-3.5 px-4">Dusun</th>
                <th className="py-3.5 px-4">RT / RW</th>
                <th className="py-3.5 px-4">Status</th>
                <th className="py-3.5 px-4">Lunas Sampai</th>
                <th className="py-3.5 px-4">Keterangan Tunggakan Mingguan</th>
                <th className="py-3.5 px-4 text-right">Aksi Cepat</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800/70 text-xs">
              {residents.map((w, index) => (
                <tr key={w.id} className="hover:bg-slate-800/40">
                  <td className="py-3.5 px-4 font-mono text-slate-400">{index + 1}</td>
                  <td className="py-3.5 px-4">
                    <button
                      type="button"
                      onClick={onProtectedChecklistClick}
                      className={
                        'w-6 h-6 rounded-md flex items-center justify-center border font-bold ' +
                        (w.computed.isChecked
                          ? 'bg-emerald-500/20 border-emerald-500/50 text-emerald-400'
                          : 'bg-slate-800 border-slate-700 text-slate-500')
                      }
                      title="Otomatis dari transaksi database"
                    >
                      {w.computed.isChecked ? '✓' : ''}
                    </button>
                  </td>
                  <td className="py-3.5 px-4">
                    <button
                      onClick={() => onOpenDetail(w.id)}
                      className="font-bold text-white hover:text-emerald-400 text-left"
                    >
                      {w.nama}
                    </button>
                    <span className="block text-[11px] text-slate-500">
                      Total Bayar: {formatRupiah(w.computed.totalNominal)} ({w.computed.totalMinggu} minggu)
                    </span>
                  </td>
                  <td className="py-3.5 px-4 text-slate-300">{w.dusun}</td>
                  <td className="py-3.5 px-4 font-mono text-slate-300">
                    RT {w.rt} / RW {w.rw || '01'}
                  </td>
                  <td className="py-3.5 px-4">
                    <StatusBadge statusKey={w.computed.statusKey} label={w.computed.statusLabel} />
                  </td>
                  <td className="py-3.5 px-4 font-bold text-white">
                    {formatDateID(w.computed.lunasSampai)}
                  </td>
                  <td className="py-3.5 px-4">
                    {w.computed.mingguNunggak > 0 ? (
                      <span className="inline-block px-2.5 py-1 rounded-md bg-rose-500/15 border border-rose-500/30 text-rose-300 font-bold text-[11px]">
                        Nunggak {w.computed.mingguNunggak} Minggu ({formatRupiah(w.computed.nominalTunggakan)})
                      </span>
                    ) : (
                      <span className="text-emerald-300 font-medium">{w.computed.keteranganTunggakan}</span>
                    )}
                  </td>
                  <td className="py-3.5 px-4 text-right space-x-1.5 whitespace-nowrap">
                    <button
                      onClick={() => onOpenPayment(w.id, '1_BULAN')}
                      className="px-2.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-500 text-white font-bold"
                    >
                      + Bayar
                    </button>
                    <button
                      onClick={() => onOpenDetail(w.id)}
                      className="px-2.5 py-1.5 rounded-lg bg-slate-800 hover:bg-slate-700 text-slate-200 font-semibold"
                    >
                      Detail
                    </button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// 4. RIWAYAT ANTAR WARGA & UANG KELUAR/MASUK BBM
// ============================================================================
function RiwayatAntarWargaView({ tripLogs, residents, drivers, dusunOptions, systemDate, onAddTrip }) {
  const [selectedResidentId, setSelectedResidentId] = useState(residents[0]?.id || '');
  const [manualName, setManualName] = useState('');
  const [dusun, setDusun] = useState(dusunOptions[0] || 'Dusun Krajan');
  const [rt, setRt] = useState('01');
  const [rw, setRw] = useState('01');
  const [tanggal, setTanggal] = useState(systemDate);
  const [waktu, setWaktu] = useState('08:30 WIB');
  const [tujuan, setTujuan] = useState('');
  const [keperluan, setKeperluan] = useState('');
  const [driverId, setDriverId] = useState(drivers[0]?.id || 'DRV-01');
  const [uangKeluar, setUangKeluar] = useState('150000');
  const [uangMasuk, setUangMasuk] = useState('0');
  const [keterangan, setKeterangan] = useState('');

  const handleSelectResident = (id) => {
    setSelectedResidentId(id);
    const found = residents.find((r) => r.id === id);
    if (found) {
      setManualName(found.nama);
      setDusun(found.dusun);
      setRt(found.rt);
      setRw(found.rw || '01');
    }
  };

  const handleSubmit = (e) => {
    e.preventDefault();
    const resObj = residents.find((r) => r.id === selectedResidentId);
    const drvObj = drivers.find((d) => d.id === driverId) || drivers[0];
    const finalName = manualName.trim() || resObj?.nama || 'Warga Desa Kalikurmo';
    if (!tujuan.trim()) return;

    onAddTrip({
      tanggal,
      waktu,
      residentId: resObj ? resObj.id : '-',
      residentName: finalName,
      dusun,
      rt,
      rw,
      tujuan,
      keperluan: keperluan || 'Layanan Kesehatan / Darurat Warga',
      driverId: drvObj?.id || 'DRV-01',
      driverName: drvObj?.nama || 'Sopir Siaga',
      uangKeluar,
      uangMasuk,
      keterangan
    });
    setTujuan('');
    setKeperluan('');
    setKeterangan('');
  };

  return (
    <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
      <form onSubmit={handleSubmit} className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3.5 h-fit">
        <h3 className="text-base font-bold text-white flex items-center gap-2">
          <Navigation className="w-4 h-4 text-emerald-400" />
          Catat Pemakaian Mobil Siaga (Antar Warga)
        </h3>

        <div>
          <label className="block text-xs text-slate-400 mb-1">Pilih Warga Terdaftar</label>
          <select
            value={selectedResidentId}
            onChange={(e) => handleSelectResident(e.target.value)}
            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
          >
            <option value="">-- Ketik Nama Manual di Bawah --</option>
            {residents.map((r) => (
              <option key={r.id} value={r.id}>
                {r.nama} ({r.dusun} RT {r.rt}/RW {r.rw || '01'})
              </option>
            ))}
          </select>
        </div>

        <div>
          <label className="block text-xs text-slate-400 mb-1">Nama Warga yang Diantar</label>
          <input
            type="text"
            value={manualName}
            onChange={(e) => setManualName(e.target.value)}
            placeholder="Ketik nama warga..."
            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
          />
        </div>

        <div className="grid grid-cols-3 gap-2">
          <div>
            <label className="block text-xs text-slate-400 mb-1">Dusun</label>
            <input
              type="text"
              value={dusun}
              onChange={(e) => setDusun(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-2.5 py-2 text-xs text-white"
            />
          </div>
          <div>
            <label className="block text-xs text-slate-400 mb-1">RT</label>
            <select
              value={rt}
              onChange={(e) => setRt(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-2 py-2 text-xs text-white"
            >
              {RT_RW_LIST.map((num) => (
                <option key={num} value={num}>RT {num}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-xs text-slate-400 mb-1">RW</label>
            <select
              value={rw}
              onChange={(e) => setRw(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-2 py-2 text-xs text-white"
            >
              {RT_RW_LIST.map((num) => (
                <option key={num} value={num}>RW {num}</option>
              ))}
            </select>
          </div>
        </div>

        <div className="grid grid-cols-2 gap-2">
          <div>
            <label className="block text-xs text-slate-400 mb-1">Tanggal Pakai</label>
            <input
              type="date"
              value={tanggal}
              onChange={(e) => setTanggal(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
            />
          </div>
          <div>
            <label className="block text-xs text-slate-400 mb-1">Jam Keberangkatan</label>
            <input
              type="text"
              value={waktu}
              onChange={(e) => setWaktu(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
            />
          </div>
        </div>

        <div>
          <label className="block text-xs text-slate-400 mb-1">Tujuan RS / Puskesmas / Lokasi</label>
          <input
            type="text"
            required
            value={tujuan}
            onChange={(e) => setTujuan(e.target.value)}
            placeholder="Contoh: RSUD Kabupaten / Puskesmas"
            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
          />
        </div>

        <div>
          <label className="block text-xs text-slate-400 mb-1">Keperluan Medis / Sosial</label>
          <input
            type="text"
            value={keperluan}
            onChange={(e) => setKeperluan(e.target.value)}
            placeholder="Contoh: Rujukan Darurat / Kontrol Pasien"
            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
          />
        </div>

        <div>
          <label className="block text-xs text-slate-400 mb-1">Sopir yang Bertugas</label>
          <select
            value={driverId}
            onChange={(e) => setDriverId(e.target.value)}
            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
          >
            {drivers.map((d) => (
              <option key={d.id} value={d.id}>
                {d.nama} {d.panggilan ? '(' + d.panggilan + ')' : ''} - {d.noHP}
              </option>
            ))}
          </select>
        </div>

        <div className="grid grid-cols-2 gap-2">
          <div>
            <label className="block text-xs text-rose-300 mb-1">Uang Keluar BBM (Rp)</label>
            <input
              type="number"
              min="0"
              value={uangKeluar}
              onChange={(e) => setUangKeluar(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
            />
          </div>
          <div>
            <label className="block text-xs text-emerald-300 mb-1">Uang Masuk / Infaq (Rp)</label>
            <input
              type="number"
              min="0"
              value={uangMasuk}
              onChange={(e) => setUangMasuk(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-xs text-white"
            />
          </div>
        </div>

        <button
          type="submit"
          className="w-full py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold shadow-lg"
        >
          Simpan Riwayat Perjalanan & Kas
        </button>
      </form>

      {/* Tabel Riwayat Pemakaian Mobil Siaga */}
      <div className="lg:col-span-2 bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
        <h3 className="text-base font-bold text-white">
          Daftar Riwayat Pemakaian Mobil Siaga & Rincian Uang Keluar/Masuk
        </h3>
        <div className="overflow-x-auto rounded-xl border border-slate-800">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-950 text-[11px] uppercase text-slate-400 border-b border-slate-800">
                <th className="py-3 px-3">Tanggal & Jam</th>
                <th className="py-3 px-3">Warga Diantar</th>
                <th className="py-3 px-3">Tujuan & Keperluan</th>
                <th className="py-3 px-3">Sopir</th>
                <th className="py-3 px-3">Uang Keluar (BBM)</th>
                <th className="py-3 px-3">Uang Masuk</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-800 text-xs">
              {tripLogs.map((t) => (
                <tr key={t.id} className="hover:bg-slate-800/40">
                  <td className="py-3 px-3 whitespace-nowrap">
                    <span className="font-bold text-white block">{formatDateID(t.tanggal)}</span>
                    <span className="text-[11px] text-slate-400">{t.waktu}</span>
                  </td>
                  <td className="py-3 px-3">
                    <span className="font-bold text-white block">{t.residentName}</span>
                    <span className="text-[11px] text-slate-400">
                      {t.dusun} • RT {t.rt}/RW {t.rw || '01'}
                    </span>
                  </td>
                  <td className="py-3 px-3">
                    <span className="font-bold text-emerald-300 block">{t.tujuan}</span>
                    <span className="text-[11px] text-slate-300">{t.keperluan}</span>
                  </td>
                  <td className="py-3 px-3 text-slate-200 font-semibold">{t.driverName}</td>
                  <td className="py-3 px-3 font-bold text-rose-400">{formatRupiah(t.uangKeluar)}</td>
                  <td className="py-3 px-3 font-bold text-emerald-400">{formatRupiah(t.uangMasuk)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// 5. PROFIL KEPALA DESA & SOPIR + EDIT NAMA SOPIR & NAMA DALAM TANDA KURUNG
// ============================================================================
function ProfilKadesSopirView({ villageOfficials, dusunOptions, onUpdateKades, onSaveDriver, showToast }) {
  const [kadesForm, setKadesForm] = useState(villageOfficials.kades);
  const [editingDriver, setEditingDriver] = useState(null);

  const handleImageUpload = (file, callback) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
      callback(ev.target.result);
    };
    reader.readAsDataURL(file);
  };

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Form Edit Profil Kepala Desa Kalikurmo */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
          <h3 className="text-base font-bold text-white flex items-center gap-2">
            <Award className="w-5 h-5 text-amber-400" />
            Profil Kepala Desa Kalikurmo (Upload Foto & Kontak)
          </h3>

          <div className="flex items-center gap-4">
            <div className="relative">
              {kadesForm.fotoUrl ? (
                <img
                  src={kadesForm.fotoUrl}
                  alt={kadesForm.nama}
                  className="w-20 h-20 rounded-2xl object-cover border-2 border-amber-400"
                />
              ) : (
                <div className="w-20 h-20 rounded-2xl bg-amber-500/20 border border-amber-500/40 flex items-center justify-center text-amber-300 font-extrabold text-2xl">
                  KD
                </div>
              )}
              <label className="absolute -bottom-1 -right-1 w-7 h-7 rounded-full bg-emerald-600 hover:bg-emerald-500 text-white flex items-center justify-center cursor-pointer shadow">
                <Camera className="w-3.5 h-3.5" />
                <input
                  type="file"
                  accept="image/*"
                  className="hidden"
                  onChange={(e) =>
                    handleImageUpload(e.target.files?.[0], (base64) => {
                      const updated = { ...kadesForm, fotoUrl: base64 };
                      setKadesForm(updated);
                      onUpdateKades(updated);
                    })
                  }
                />
              </label>
            </div>
            <div className="text-xs text-slate-400">
              <p className="font-bold text-white text-sm">{kadesForm.nama}</p>
              <p>{kadesForm.jabatan}</p>
              <p className="text-emerald-400 mt-1">Klik ikon kamera untuk mengunggah foto Kepala Desa</p>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
            <div>
              <label className="block text-slate-400 mb-1">Nama Lengkap Kepala Desa</label>
              <input
                type="text"
                value={kadesForm.nama}
                onChange={(e) => setKadesForm({ ...kadesForm, nama: e.target.value })}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
              />
            </div>
            <div>
              <label className="block text-slate-400 mb-1">Nomor HP / WhatsApp</label>
              <input
                type="text"
                value={kadesForm.noHP}
                onChange={(e) => setKadesForm({ ...kadesForm, noHP: e.target.value })}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
              />
            </div>
            <div className="sm:col-span-2">
              <label className="block text-slate-400 mb-1">Unit Armada Mobil Siaga</label>
              <input
                type="text"
                value={kadesForm.unitMobil}
                onChange={(e) => setKadesForm({ ...kadesForm, unitMobil: e.target.value })}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
              />
            </div>
            <div className="sm:col-span-2">
              <label className="block text-slate-400 mb-1">Pesan Layanan Mobil Siaga Desa Kalikurmo</label>
              <textarea
                rows={2}
                value={kadesForm.pesanSiaga}
                onChange={(e) => setKadesForm({ ...kadesForm, pesanSiaga: e.target.value })}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
              />
            </div>
          </div>

          <button
            type="button"
            onClick={() => onUpdateKades(kadesForm)}
            className="px-4 py-2 rounded-xl bg-amber-600 hover:bg-amber-500 text-white text-xs font-bold"
          >
            Simpan Profil Kepala Desa
          </button>
        </div>

        {/* Daftar Sopir & Edit Nama Lengkap + Nama Dalam Tanda Kurung */}
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
          <div className="flex items-center justify-between">
            <h3 className="text-base font-bold text-white flex items-center gap-2">
              <Phone className="w-4 h-4 text-emerald-400" />
              Daftar Sopir Mobil Siaga Desa Kalikurmo
            </h3>
            <button
              onClick={() =>
                setEditingDriver({
                  id: '',
                  nama: '',
                  panggilan: '',
                  noHP: '',
                  shift: 'Siaga 24 Jam',
                  dusun: dusunOptions[0] || 'Dusun Krajan',
                  rt: '01',
                  rw: '01',
                  statusSiaga: 'SIAGA AKTIF',
                  pengalaman: 'Sopir Mobil Siaga Desa Kalikurmo',
                  platMobil: 'H 9912 XA',
                  fotoUrl: ''
                })
              }
              className="px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold"
            >
              + Tambah Sopir
            </button>
          </div>

          <div className="space-y-3">
            {villageOfficials.drivers.map((drv) => (
              <div key={drv.id} className="bg-slate-950 border border-slate-800 rounded-xl p-4 flex items-center justify-between gap-3">
                <div className="flex items-center gap-3">
                  <div className="relative">
                    {drv.fotoUrl ? (
                      <img src={drv.fotoUrl} alt={drv.nama} className="w-14 h-14 rounded-xl object-cover border border-emerald-500" />
                    ) : (
                      <div className="w-14 h-14 rounded-xl bg-emerald-500/20 border border-emerald-500/40 flex items-center justify-center text-emerald-300 font-bold">
                        {drv.nama.slice(0, 2).toUpperCase()}
                      </div>
                    )}
                    <label className="absolute -bottom-1 -right-1 w-6 h-6 rounded-full bg-emerald-600 text-white flex items-center justify-center cursor-pointer">
                      <Camera className="w-3 h-3" />
                      <input
                        type="file"
                        accept="image/*"
                        className="hidden"
                        onChange={(e) =>
                          handleImageUpload(e.target.files?.[0], (base64) => {
                            onSaveDriver({ ...drv, fotoUrl: base64 });
                          })
                        }
                      />
                    </label>
                  </div>
                  <div className="text-xs">
                    <p className="font-bold text-white text-sm">
                      {drv.nama}{' '}
                      {drv.panggilan ? (
                        <span className="text-slate-400 font-normal">({drv.panggilan})</span>
                      ) : null}
                    </p>
                    <p className="text-emerald-400 font-semibold">{drv.noHP}</p>
                    <p className="text-slate-400">
                      {drv.dusun} • RT {drv.rt}/RW {drv.rw || '01'} • {drv.shift}
                    </p>
                  </div>
                </div>
                <button
                  onClick={() => setEditingDriver({ ...drv, panggilan: drv.panggilan || '' })}
                  className="px-3 py-1.5 rounded-lg bg-emerald-600/20 hover:bg-emerald-600 text-emerald-300 hover:text-white border border-emerald-500/30 text-xs font-bold transition-all"
                >
                  Edit Nama & Data
                </button>
              </div>
            ))}
          </div>

          {editingDriver && (
            <form
              onSubmit={(e) => {
                e.preventDefault();
                if (!editingDriver.nama || !editingDriver.noHP) {
                  showToast('Nama dan Nomor HP Sopir wajib diisi!', 'error');
                  return;
                }
                onSaveDriver(editingDriver);
                setEditingDriver(null);
              }}
              className="bg-slate-950 border-2 border-emerald-500/50 rounded-xl p-4 space-y-3 text-xs"
            >
              <h4 className="font-bold text-emerald-400 text-sm">
                {editingDriver.id ? 'Edit Nama & Profil Sopir' : 'Tambah Sopir Baru'}
              </h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                <div>
                  <label className="block text-slate-400 mb-1">Nama Lengkap Sopir</label>
                  <input
                    type="text"
                    placeholder="Contoh: Slamet Riyadi"
                    value={editingDriver.nama}
                    onChange={(e) => setEditingDriver({ ...editingDriver, nama: e.target.value })}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-white font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">
                    Nama di Dalam Kurung (...) / Panggilan
                  </label>
                  <input
                    type="text"
                    placeholder="Contoh: Bang Slamet (Kosongkan jika tanpa kurung)"
                    value={editingDriver.panggilan || ''}
                    onChange={(e) => setEditingDriver({ ...editingDriver, panggilan: e.target.value })}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-white font-semibold"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Nomor HP / WhatsApp</label>
                  <input
                    type="text"
                    placeholder="0812xxxxxxxx"
                    value={editingDriver.noHP}
                    onChange={(e) => setEditingDriver({ ...editingDriver, noHP: e.target.value })}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-white"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Jadwal Shift / Siaga</label>
                  <input
                    type="text"
                    placeholder="Contoh: Shift Pagi / Siaga 24 Jam"
                    value={editingDriver.shift || ''}
                    onChange={(e) => setEditingDriver({ ...editingDriver, shift: e.target.value })}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-white"
                  />
                </div>
                <div>
                  <label className="block text-slate-400 mb-1">Nama Dusun (Ketik Manual)</label>
                  <input
                    type="text"
                    placeholder="Ketik Nama Dusun"
                    value={editingDriver.dusun}
                    onChange={(e) => setEditingDriver({ ...editingDriver, dusun: e.target.value })}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-white"
                  />
                </div>
                <div className="grid grid-cols-2 gap-1.5">
                  <div>
                    <label className="block text-slate-400 mb-1">RT</label>
                    <select
                      value={editingDriver.rt}
                      onChange={(e) => setEditingDriver({ ...editingDriver, rt: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-2 py-2 text-white"
                    >
                      {RT_RW_LIST.map((r) => (
                        <option key={r} value={r}>RT {r}</option>
                      ))}
                    </select>
                  </div>
                  <div>
                    <label className="block text-slate-400 mb-1">RW</label>
                    <select
                      value={editingDriver.rw || '01'}
                      onChange={(e) => setEditingDriver({ ...editingDriver, rw: e.target.value })}
                      className="w-full bg-slate-900 border border-slate-800 rounded-lg px-2 py-2 text-white"
                    >
                      {RT_RW_LIST.map((r) => (
                        <option key={r} value={r}>RW {r}</option>
                      ))}
                    </select>
                  </div>
                </div>
                <div className="sm:col-span-2">
                  <label className="block text-slate-400 mb-1">Keterangan Pengalaman Sopir</label>
                  <input
                    type="text"
                    placeholder="Contoh: 8 Tahun Sopir Ambulans Desa Kalikurmo"
                    value={editingDriver.pengalaman || ''}
                    onChange={(e) => setEditingDriver({ ...editingDriver, pengalaman: e.target.value })}
                    className="w-full bg-slate-900 border border-slate-800 rounded-lg px-3 py-2 text-white"
                  />
                </div>
              </div>
              <div className="flex justify-end gap-2 pt-2">
                <button
                  type="button"
                  onClick={() => setEditingDriver(null)}
                  className="px-3 py-1.5 rounded-lg bg-slate-800 text-slate-300"
                >
                  Batal
                </button>
                <button
                  type="submit"
                  className="px-4 py-1.5 rounded-lg bg-emerald-600 text-white font-bold"
                >
                  Simpan Data Sopir
                </button>
              </div>
            </form>
          )}
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// 6. BUKU KAS UTAMA MOBIL SIAGA DESA KALIKURMO
// ============================================================================
function KasMobilSiagaView({ cashLedger, stats }) {
  return (
    <div className="space-y-5">
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4">
          <span className="text-xs text-slate-400">Total Pemasukan Kas</span>
          <p className="text-2xl font-extrabold text-emerald-400 mt-1">{formatRupiah(stats.totalMasukKas)}</p>
        </div>
        <div className="bg-slate-900 border border-slate-800 rounded-2xl p-4">
          <span className="text-xs text-slate-400">Total Pengeluaran (BBM / Koreksi)</span>
          <p className="text-2xl font-extrabold text-rose-400 mt-1">{formatRupiah(stats.totalKeluarKas)}</p>
        </div>
        <div className="bg-slate-900 border border-emerald-500/30 rounded-2xl p-4">
          <span className="text-xs text-emerald-300 font-semibold">Saldo Kas Aktif Mobil Siaga</span>
          <p className="text-2xl font-extrabold text-white mt-1">{formatRupiah(stats.saldoKas)}</p>
        </div>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-2xl overflow-hidden">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-slate-950 text-[11px] uppercase text-slate-400 border-b border-slate-800">
              <th className="py-3 px-4">ID Kas</th>
              <th className="py-3 px-4">Tanggal</th>
              <th className="py-3 px-4">Tipe</th>
              <th className="py-3 px-4">Deskripsi</th>
              <th className="py-3 px-4">Nominal</th>
              <th className="py-3 px-4">Admin</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800 text-xs">
            {cashLedger.map((c) => (
              <tr key={c.id} className="hover:bg-slate-800/40">
                <td className="py-3 px-4 font-mono text-slate-400">{c.id}</td>
                <td className="py-3 px-4 text-white">{formatDateID(c.tanggal)}</td>
                <td className="py-3 px-4">
                  <span
                    className={
                      'px-2 py-0.5 rounded font-bold text-[11px] ' +
                      (c.tipe === 'MASUK'
                        ? 'bg-emerald-500/20 text-emerald-300'
                        : 'bg-rose-500/20 text-rose-300')
                    }
                  >
                    {c.tipe}
                  </span>
                </td>
                <td className="py-3 px-4 text-slate-200">{c.deskripsi}</td>
                <td className={'py-3 px-4 font-bold ' + (c.tipe === 'MASUK' ? 'text-emerald-400' : 'text-rose-400')}>
                  {c.tipe === 'MASUK' ? '+' : '-'} {formatRupiah(c.nominal)}
                </td>
                <td className="py-3 px-4 text-slate-400">{c.admin}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================================
// 7. RIWAYAT PEMBAYARAN PERMANEN (MINGGUAN)
// ============================================================================
function RiwayatPembayaranView({ transactions, onPrintReceipt, onCorrectTrx }) {
  return (
    <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
      <h3 className="text-base font-bold text-white">
        Riwayat Transaksi Pembayaran Iuran Mingguan Warga (Tercatat Permanen)
      </h3>
      <div className="overflow-x-auto rounded-xl border border-slate-800">
        <table className="w-full text-left border-collapse">
          <thead>
            <tr className="bg-slate-950 text-[11px] uppercase text-slate-400 border-b border-slate-800">
              <th className="py-3 px-3">No. Transaksi</th>
              <th className="py-3 px-3">Tgl Bayar</th>
              <th className="py-3 px-3">Nama Warga</th>
              <th className="py-3 px-3">Periode Lunas</th>
              <th className="py-3 px-3">Minggu & Tarif</th>
              <th className="py-3 px-3">Nominal</th>
              <th className="py-3 px-3">Admin</th>
              <th className="py-3 px-3 text-right">Aksi</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-800 text-xs">
            {transactions.map((t) => {
              const weeks = t.jumlahMinggu || Math.max(1, Math.round((t.jumlahHari || 7) / 7));
              const rateW = t.tarifPerMinggu || t.tarifPerHari || 5000;
              return (
                <tr key={t.id} className={t.status === 'VOID' ? 'opacity-50 bg-rose-950/20' : 'hover:bg-slate-800/40'}>
                  <td className="py-3 px-3 font-mono text-slate-300">{t.id}</td>
                  <td className="py-3 px-3 text-white">{formatDateID(t.tanggalTransaksi)}</td>
                  <td className="py-3 px-3">
                    <span className="font-bold text-white block">{t.residentName}</span>
                    <span className="text-[11px] text-slate-400">
                      {t.dusun} • RT {t.rt}/RW {t.rw || '01'}
                    </span>
                  </td>
                  <td className="py-3 px-3 text-emerald-300 font-semibold">
                    {formatDateID(t.periodeMulai)} s/d {formatDateID(t.periodeAkhir)}
                  </td>
                  <td className="py-3 px-3 text-slate-300">
                    {weeks} minggu × {formatRupiah(rateW)}
                  </td>
                  <td className="py-3 px-3 font-bold text-emerald-400">{formatRupiah(t.nominal)}</td>
                  <td className="py-3 px-3 text-slate-400">{t.adminName}</td>
                  <td className="py-3 px-3 text-right space-x-1.5 whitespace-nowrap">
                    <button
                      onClick={() => onPrintReceipt(t)}
                      className="px-2.5 py-1 rounded bg-slate-800 hover:bg-slate-700 text-slate-200"
                    >
                      Kuitansi
                    </button>
                    {t.status === 'VALID' && (
                      <button
                        onClick={() => onCorrectTrx(t)}
                        className="px-2.5 py-1 rounded bg-rose-500/20 hover:bg-rose-500/30 text-rose-300"
                      >
                        Koreksi
                      </button>
                    )}
                  </td>
                </tr>
              );
            })}
          </tbody>
        </table>
      </div>
    </div>
  );
}

// ============================================================================
// 8. PENGATURAN NOMINAL IURAN MINGGUAN (DINAMIS & HISTORI TARIF)
// ============================================================================
function PengaturanIuranView({ rates, activeRate, systemDate, onAddRate }) {
  const [nominal, setNominal] = useState(String(activeRate.nominal));
  const [effectiveDate, setEffectiveDate] = useState(systemDate);
  const [keterangan, setKeterangan] = useState('');

  return (
    <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
        <h3 className="text-base font-bold text-white">PENGATURAN → PENGATURAN IURAN MINGGUAN</h3>
        <p className="text-xs text-slate-400">
          Perubahan nominal iuran mingguan tidak akan mengubah transaksi pembayaran lama yang sudah tersimpan di database.
        </p>

        <div className="flex flex-wrap gap-2">
          <button
            type="button"
            onClick={() => {
              setNominal('5000');
              setEffectiveDate('2026-01-01');
              setKeterangan('Keputusan Desa Kalikurmo Tahun 2026: Rp5.000/minggu');
            }}
            className="px-3 py-1.5 rounded-lg bg-slate-800 text-xs text-emerald-300 font-semibold"
          >
            Preset 2026: Rp5.000 / minggu
          </button>
          <button
            type="button"
            onClick={() => {
              setNominal('7000');
              setEffectiveDate('2027-01-01');
              setKeterangan('Keputusan Desa Kalikurmo Tahun 2027: Rp7.000/minggu');
            }}
            className="px-3 py-1.5 rounded-lg bg-slate-800 text-xs text-sky-300 font-semibold"
          >
            Preset 2027: Rp7.000 / minggu
          </button>
          <button
            type="button"
            onClick={() => {
              setNominal('10000');
              setEffectiveDate('2028-01-01');
              setKeterangan('Keputusan Desa Kalikurmo Tahun 2028: Rp10.000/minggu');
            }}
            className="px-3 py-1.5 rounded-lg bg-slate-800 text-xs text-amber-300 font-semibold"
          >
            Preset 2028: Rp10.000 / minggu
          </button>
        </div>

        <form
          onSubmit={(e) => {
            e.preventDefault();
            onAddRate({ nominal, effectiveDate, keterangan });
            setKeterangan('');
          }}
          className="space-y-3 text-xs"
        >
          <div>
            <label className="block text-slate-300 mb-1">Nominal Iuran Mingguan Baru (Rp / Minggu)</label>
            <input
              type="number"
              value={nominal}
              onChange={(e) => setNominal(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-white font-bold"
            />
          </div>
          <div>
            <label className="block text-slate-300 mb-1">Tanggal Mulai Berlaku</label>
            <input
              type="date"
              value={effectiveDate}
              onChange={(e) => setEffectiveDate(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-white"
            />
          </div>
          <div>
            <label className="block text-slate-300 mb-1">Keterangan Perubahan Tarif</label>
            <input
              type="text"
              value={keterangan}
              onChange={(e) => setKeterangan(e.target.value)}
              placeholder="Contoh: Hasil Musyawarah Desa Kalikurmo"
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3.5 py-2 text-white"
            />
          </div>
          <button
            type="submit"
            className="w-full py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold"
          >
            Simpan Tarif Mingguan Baru
          </button>
        </form>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-4">
        <h3 className="text-base font-bold text-white">Riwayat Perubahan Nominal Mingguan</h3>
        <div className="space-y-2.5">
          {rates.map((r) => (
            <div key={r.id} className="bg-slate-950 border border-slate-800 rounded-xl p-3 text-xs">
              <div className="flex justify-between font-bold text-white">
                <span>{formatRupiah(r.nominal)} / minggu</span>
                <span className="text-emerald-400">Mulai: {formatDateID(r.effectiveDate)}</span>
              </div>
              <p className="text-slate-300 mt-1">{r.keterangan}</p>
              <p className="text-[11px] text-slate-500 mt-1">Diubah oleh: {r.changedBy}</p>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ============================================================================
// 9. TEMPAT INPUT LINK GOOGLE SHEET & KODE.GS
// ============================================================================
function DatabaseSheetView({ sheetUrl, setSheetUrl, syncStatus, lastSyncTime, onSyncNow, onFetchNow, copiedCode, onCopyCode }) {
  return (
    <div className="space-y-6">
      <div className="bg-slate-900 border-2 border-emerald-500/40 rounded-2xl p-5 space-y-4">
        <div className="space-y-1">
          <h3 className="text-base font-bold text-white flex items-center gap-2">
            <Database className="w-5 h-5 text-emerald-400" />
            LINK WEB APP GOOGLE APPS SCRIPT (TERHUBUNG)
          </h3>
          <p className="text-xs text-slate-300">
            Link Google Apps Script Anda sudah terpasang otomatis. Klik tombol <strong>Simpan ke Google Sheet</strong> untuk mengirim data awal ke Spreadsheet Anda.
          </p>
        </div>

        <div className="flex flex-col sm:flex-row gap-3">
          <input
            type="text"
            value={sheetUrl}
            onChange={(e) => setSheetUrl(e.target.value)}
            placeholder="Tempel link di sini: https://script.google.com/macros/s/......./exec"
            className="flex-1 bg-slate-950 border border-emerald-500/50 rounded-xl px-4 py-2.5 text-xs text-white font-mono"
          />
          <button
            onClick={onSyncNow}
            className="px-4 py-2.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold shadow-lg"
          >
            Simpan ke Google Sheet
          </button>
          <button
            onClick={onFetchNow}
            className="px-4 py-2.5 rounded-xl bg-sky-600 hover:bg-sky-500 text-white text-xs font-bold"
          >
            Muat dari Google Sheet
          </button>
        </div>

        <div className="bg-slate-950 border border-slate-800 rounded-xl p-3 text-xs text-slate-300 space-y-1">
          <p>
            Status Sinkronisasi: <strong className="text-emerald-400">{syncStatus}</strong> • Terakhir Update: <strong>{lastSyncTime}</strong>
          </p>
        </div>
      </div>

      <div className="bg-slate-900 border border-slate-800 rounded-2xl p-5 space-y-3">
        <div className="flex items-center justify-between">
          <h4 className="text-sm font-bold text-white">Kode Lengkap Backend Google Apps Script (Kode.gs)</h4>
          <button
            onClick={onCopyCode}
            className="flex items-center gap-1.5 px-3 py-1.5 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-bold"
          >
            {copiedCode ? <Check className="w-4 h-4" /> : <Copy className="w-4 h-4" />}
            {copiedCode ? 'Tersalin!' : 'Salin Kode.gs'}
          </button>
        </div>
        <pre className="bg-slate-950 border border-slate-800 rounded-xl p-4 text-[11px] text-emerald-300 overflow-x-auto max-h-96">
          {KODE_GS_SCRIPT}
        </pre>
      </div>
    </div>
  );
}

// ============================================================================
// MODAL DETAIL WARGA, FORM PEMBAYARAN MINGGUAN, KUITANSI, KOREKSI, TAMBAH WARGA
// ============================================================================
function ResidentDetailModal({ resident, activeRate, onClose, onOpenPayment, onPrintReceipt, onCorrectTrx }) {
  return (
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
      <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-3xl w-full p-6 space-y-5 my-8">
        <div className="flex items-start justify-between">
          <div>
            <h3 className="text-lg font-extrabold text-white">{resident.nama}</h3>
            <p className="text-xs text-slate-400">
              {resident.dusun} • RT {resident.rt} / RW {resident.rw || '01'} • Tarif Aktif: {formatRupiah(activeRate.nominal)}/minggu
            </p>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-white text-sm font-bold">✕ Tutup</button>
        </div>

        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
          <div className="bg-slate-950 border border-slate-800 rounded-xl p-3">
            <span className="text-slate-400 block">Status Saat Ini</span>
            <div className="mt-1">
              <StatusBadge statusKey={resident.computed.statusKey} label={resident.computed.statusLabel} />
            </div>
          </div>
          <div className="bg-slate-950 border border-slate-800 rounded-xl p-3">
            <span className="text-slate-400 block">Terakhir Lunas Sampai</span>
            <span className="text-sm font-bold text-white mt-1 block">{formatDateID(resident.computed.lunasSampai)}</span>
          </div>
          <div className="bg-slate-950 border border-slate-800 rounded-xl p-3">
            <span className="text-slate-400 block">Total Sudah Dibayar</span>
            <span className="text-sm font-bold text-emerald-400 mt-1 block">
              {formatRupiah(resident.computed.totalNominal)} ({resident.computed.totalMinggu} minggu)
            </span>
          </div>
          <div className="bg-slate-950 border border-slate-800 rounded-xl p-3">
            <span className="text-slate-400 block">Info Tunggakan Mingguan</span>
            <span className="text-xs font-bold text-rose-400 mt-1 block">{resident.computed.keteranganTunggakan}</span>
          </div>
        </div>

        <div className="flex flex-wrap gap-2">
          <button onClick={() => onOpenPayment('1_MINGGU')} className="px-3.5 py-2 rounded-xl bg-slate-800 text-white text-xs font-bold">
            + Bayar 1 Minggu
          </button>
          <button onClick={() => onOpenPayment('1_BULAN')} className="px-3.5 py-2 rounded-xl bg-emerald-600 text-white text-xs font-bold">
            + Lunasi 1 Bulan (4 Minggu)
          </button>
          <button onClick={() => onOpenPayment('1_TAHUN')} className="px-3.5 py-2 rounded-xl bg-sky-600 text-white text-xs font-bold">
            + Lunasi 1 Tahun (52 Minggu)
          </button>
          <button onClick={() => onOpenPayment('CUSTOM')} className="px-3.5 py-2 rounded-xl bg-slate-800 text-slate-200 text-xs font-bold">
            Pilih Jumlah Minggu Sendiri
          </button>
        </div>

        <div className="space-y-2">
          <h4 className="text-xs font-bold uppercase text-slate-400">Riwayat Pembayaran Mingguan Warga</h4>
          <div className="overflow-x-auto rounded-xl border border-slate-800">
            <table className="w-full text-left border-collapse text-xs">
              <thead>
                <tr className="bg-slate-950 text-slate-400 border-b border-slate-800">
                  <th className="py-2.5 px-3">No Trx</th>
                  <th className="py-2.5 px-3">Tgl Bayar</th>
                  <th className="py-2.5 px-3">Periode</th>
                  <th className="py-2.5 px-3">Jml Minggu</th>
                  <th className="py-2.5 px-3">Nominal</th>
                  <th className="py-2.5 px-3 text-right">Kuitansi</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-800">
                {resident.computed.riwayat.map((t) => {
                  const weeks = t.jumlahMinggu || Math.max(1, Math.round((t.jumlahHari || 7) / 7));
                  return (
                    <tr key={t.id}>
                      <td className="py-2.5 px-3 font-mono">{t.id}</td>
                      <td className="py-2.5 px-3">{formatDateID(t.tanggalTransaksi)}</td>
                      <td className="py-2.5 px-3 text-emerald-300">
                        {formatDateID(t.periodeMulai)} - {formatDateID(t.periodeAkhir)}
                      </td>
                      <td className="py-2.5 px-3">{weeks} minggu</td>
                      <td className="py-2.5 px-3 font-bold text-emerald-400">{formatRupiah(t.nominal)}</td>
                      <td className="py-2.5 px-3 text-right space-x-1">
                        <button onClick={() => onPrintReceipt(t)} className="px-2 py-1 rounded bg-slate-800 text-white">
                          Cetak
                        </button>
                        <button onClick={() => onCorrectTrx(t)} className="px-2 py-1 rounded bg-rose-500/20 text-rose-300">
                          Koreksi
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </div>
  );
}

function PaymentFormModal({ resident, allResidents, presetType, rates, systemDate, onClose, onSave }) {
  const [selectedResId, setSelectedResId] = useState(resident?.id || allResidents[0]?.id);
  const currentRes = allResidents.find((r) => r.id === selectedResId) || allResidents[0];

  const defaultStart = useMemo(() => {
    if (currentRes?.computed?.lunasSampai && currentRes.computed.lunasSampai !== '-') {
      return addDays(currentRes.computed.lunasSampai, 1);
    }
    return '2026-11-01';
  }, [currentRes]);

  const [tanggalTransaksi, setTanggalTransaksi] = useState(systemDate);
  const [periodeMulai, setPeriodeMulai] = useState(defaultStart);
  const [modePaket, setModePaket] = useState(presetType || '1_BULAN');
  const [customWeeks, setCustomWeeks] = useState(4);
  const [keterangan, setKeterangan] = useState('');

  useEffect(() => {
    setPeriodeMulai(defaultStart);
  }, [defaultStart]);

  const calcResult = useMemo(() => {
    let weeks = 4;
    let label = 'Lunas 1 Bulan (4 Minggu)';
    if (modePaket === '1_MINGGU') {
      weeks = 1;
      label = 'Bayar 1 Minggu';
    } else if (modePaket === '2_MINGGU') {
      weeks = 2;
      label = 'Bayar 2 Minggu';
    } else if (modePaket === '3_MINGGU') {
      weeks = 3;
      label = 'Bayar 3 Minggu';
    } else if (modePaket === '1_BULAN') {
      weeks = 4;
      label = 'Lunas 1 Bulan (4 Minggu)';
    } else if (modePaket === '2_BULAN') {
      weeks = 8;
      label = 'Lunas 2 Bulan (8 Minggu)';
    } else if (modePaket === '3_BULAN') {
      weeks = 12;
      label = 'Lunas 3 Bulan (12 Minggu)';
    } else if (modePaket === '6_BULAN') {
      weeks = 26;
      label = 'Lunas 6 Bulan (26 Minggu)';
    } else if (modePaket === '1_TAHUN') {
      weeks = 52;
      label = 'Lunas 1 Tahun (52 Minggu)';
    } else {
      weeks = Math.max(1, Number(customWeeks) || 1);
      label = 'Pilih Sendiri (' + weeks + ' Minggu)';
    }

    const totalDays = weeks * 7;
    const end = addDays(periodeMulai, totalDays - 1);
    const rateObj = getEffectiveRateOnDate(rates, tanggalTransaksi);
    return {
      periodeAkhir: end,
      jumlahMinggu: weeks,
      tarifMingguan: rateObj.nominal,
      total: weeks * rateObj.nominal,
      paketLabel: label
    };
  }, [periodeMulai, modePaket, customWeeks, rates, tanggalTransaksi]);

  return (
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
      <div className="bg-slate-900 border border-slate-700 rounded-2xl max-w-lg w-full p-6 space-y-4">
        <div className="flex items-center justify-between">
          <h3 className="text-base font-bold text-white">Input Pembayaran Iuran Mingguan & Pelunasan</h3>
          <button onClick={onClose} className="text-slate-400 hover:text-white text-sm">✕</button>
        </div>

        <div className="space-y-3 text-xs">
          <div>
            <label className="block text-slate-400 mb-1">Pilih Nama Warga</label>
            <select
              value={selectedResId}
              onChange={(e) => setSelectedResId(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white font-bold"
            >
              {allResidents.map((r) => (
                <option key={r.id} value={r.id}>
                  {r.nama} ({r.dusun} RT {r.rt}/RW {r.rw || '01'})
                </option>
              ))}
            </select>
          </div>

          {currentRes?.computed?.mingguNunggak > 0 && (
            <div className="bg-rose-950/40 border border-rose-500/30 rounded-xl p-3 flex items-center justify-between">
              <span className="text-rose-300 font-semibold">
                Tunggakan: {currentRes.computed.mingguNunggak} Minggu ({formatRupiah(currentRes.computed.nominalTunggakan)})
              </span>
              <button
                type="button"
                onClick={() => {
                  setModePaket('CUSTOM');
                  setCustomWeeks(currentRes.computed.mingguNunggak);
                }}
                className="px-2.5 py-1 rounded-lg bg-rose-600 text-white font-bold text-[11px]"
              >
                Lunasi {currentRes.computed.mingguNunggak} Minggu
              </button>
            </div>
          )}

          <div className="grid grid-cols-2 gap-2">
            <div>
              <label className="block text-slate-400 mb-1">Tanggal Transaksi</label>
              <input
                type="date"
                value={tanggalTransaksi}
                onChange={(e) => setTanggalTransaksi(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
              />
            </div>
            <div>
              <label className="block text-slate-400 mb-1">Mulai Periode Dibayar</label>
              <input
                type="date"
                value={periodeMulai}
                onChange={(e) => setPeriodeMulai(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
              />
            </div>
          </div>

          <div>
            <label className="block text-slate-400 mb-1">Pilih Paket Iuran Mingguan</label>
            <div className="grid grid-cols-3 gap-1.5">
              {[
                { id: '1_MINGGU', label: '1 Minggu' },
                { id: '2_MINGGU', label: '2 Minggu' },
                { id: '1_BULAN', label: '1 Bulan (4 Mg)' },
                { id: '2_BULAN', label: '2 Bulan (8 Mg)' },
                { id: '3_BULAN', label: '3 Bulan (12 Mg)' },
                { id: '6_BULAN', label: '6 Bulan (26 Mg)' },
                { id: '1_TAHUN', label: '1 Tahun (52 Mg)' },
                { id: '3_MINGGU', label: '3 Minggu' },
                { id: 'CUSTOM', label: 'Tentukan Minggu' }
              ].map((p) => (
                <button
                  key={p.id}
                  type="button"
                  onClick={() => setModePaket(p.id)}
                  className={
                    'py-2 px-2 rounded-lg font-bold border text-[11px] ' +
                    (modePaket === p.id
                      ? 'bg-emerald-600 text-white border-emerald-500'
                      : 'bg-slate-950 text-slate-300 border-slate-800')
                  }
                >
                  {p.label}
                </button>
              ))}
            </div>
          </div>

          {modePaket === 'CUSTOM' && (
            <div>
              <label className="block text-slate-400 mb-1">Masukkan Jumlah Minggu</label>
              <input
                type="number"
                min="1"
                value={customWeeks}
                onChange={(e) => setCustomWeeks(e.target.value)}
                className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
              />
            </div>
          )}

          <div className="bg-emerald-950/40 border border-emerald-500/30 rounded-xl p-3.5 space-y-1">
            <div className="flex justify-between">
              <span className="text-slate-300">Periode Lunas:</span>
              <span className="font-bold text-white">
                {formatDateID(periodeMulai)} s/d {formatDateID(calcResult.periodeAkhir)}
              </span>
            </div>
            <div className="flex justify-between">
              <span className="text-slate-300">Perhitungan Mingguan:</span>
              <span className="font-bold text-emerald-300">
                {calcResult.jumlahMinggu} Minggu × {formatRupiah(calcResult.tarifMingguan)}
              </span>
            </div>
            <div className="flex justify-between pt-1 border-t border-emerald-500/20 text-sm">
              <span className="font-bold text-white">Total Bayar:</span>
              <span className="font-extrabold text-emerald-400">{formatRupiah(calcResult.total)}</span>
            </div>
          </div>

          <div>
            <label className="block text-slate-400 mb-1">Keterangan (Opsional)</label>
            <input
              type="text"
              value={keterangan}
              onChange={(e) => setKeterangan(e.target.value)}
              placeholder="Catatan pembayaran..."
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
            />
          </div>

          <div className="flex justify-end gap-2 pt-2">
            <button onClick={onClose} className="px-4 py-2 rounded-xl bg-slate-800 text-slate-300 font-semibold">
              Batal
            </button>
            <button
              onClick={() =>
                onSave({
                  residentId: selectedResId,
                  tanggalTransaksi,
                  periodeMulai,
                  periodeAkhir: calcResult.periodeAkhir,
                  jumlahMinggu: calcResult.jumlahMinggu,
                  paketLabel: calcResult.paketLabel,
                  keterangan
                })
              }
              className="px-5 py-2 rounded-xl bg-emerald-600 hover:bg-emerald-500 text-white font-bold shadow-lg"
            >
              Simpan Pembayaran
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

function ReceiptPrintModal({ trx, onClose }) {
  const weeks = trx.jumlahMinggu || Math.max(1, Math.round((trx.jumlahHari || 7) / 7));
  const rateW = trx.tarifPerMinggu || trx.tarifPerHari || 5000;
  return (
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white text-slate-900 rounded-2xl max-w-md w-full p-6 space-y-4 shadow-2xl">
        <div className="text-center border-b-2 border-dashed border-slate-300 pb-3">
          <h3 className="text-base font-extrabold uppercase">KUITANSI RESMI IURAN MINGGUAN MOBIL SIAGA</h3>
          <p className="text-xs font-bold text-emerald-700">PEMERINTAH DESA KALIKURMO</p>
          <p className="text-[11px] text-slate-500">Nomor: {trx.id}</p>
        </div>
        <div className="space-y-2 text-xs">
          <div className="flex justify-between">
            <span className="text-slate-500">Nama Warga:</span>
            <span className="font-bold">{trx.residentName}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-slate-500">Wilayah:</span>
            <span>{trx.dusun} • RT {trx.rt} / RW {trx.rw || '01'}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-slate-500">Tanggal Bayar:</span>
            <span>{formatDateLongID(trx.tanggalTransaksi)}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-slate-500">Periode Lunas:</span>
            <span className="font-bold">{formatDateID(trx.periodeMulai)} s/d {formatDateID(trx.periodeAkhir)}</span>
          </div>
          <div className="flex justify-between">
            <span className="text-slate-500">Jumlah Minggu & Tarif:</span>
            <span>{weeks} Minggu × {formatRupiah(rateW)}</span>
          </div>
          <div className="flex justify-between pt-2 border-t border-slate-200 text-sm font-extrabold">
            <span>TOTAL DITERIMA:</span>
            <span className="text-emerald-700">{formatRupiah(trx.nominal)}</span>
          </div>
          <p className="text-[11px] text-slate-500 pt-2">Petugas Pencatat: {trx.adminName}</p>
        </div>
        <div className="flex justify-end gap-2 pt-3 border-t border-slate-200">
          <button onClick={onClose} className="px-4 py-1.5 rounded-lg bg-slate-200 text-slate-700 text-xs font-bold">
            Tutup
          </button>
          <button onClick={() => window.print()} className="px-4 py-1.5 rounded-lg bg-emerald-600 text-white text-xs font-bold">
            Cetak Kuitansi
          </button>
        </div>
      </div>
    </div>
  );
}

function CorrectionDialogModal({ trx, onClose, onConfirm }) {
  const [reason, setReason] = useState('');
  return (
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-slate-900 border border-rose-500/40 rounded-2xl max-w-md w-full p-6 space-y-4">
        <h3 className="text-base font-bold text-white">Koreksi Transaksi {trx.id}</h3>
        <p className="text-xs text-slate-300">
          Membatalkan transaksi an. <strong>{trx.residentName}</strong> ({formatRupiah(trx.nominal)}). Wajib mengisi alasan untuk Audit Log.
        </p>
        <textarea
          rows={3}
          value={reason}
          onChange={(e) => setReason(e.target.value)}
          placeholder="Tulis alasan koreksi..."
          className="w-full bg-slate-950 border border-slate-800 rounded-xl p-3 text-xs text-white"
        />
        <div className="flex justify-end gap-2">
          <button onClick={onClose} className="px-4 py-2 rounded-xl bg-slate-800 text-slate-300 text-xs font-semibold">
            Batal
          </button>
          <button
            onClick={() => onConfirm(trx.id, reason)}
            className="px-4 py-2 rounded-xl bg-rose-600 hover:bg-rose-500 text-white text-xs font-bold"
          >
            Konfirmasi Koreksi
          </button>
        </div>
      </div>
    </div>
  );
}

function AddResidentModal({ dusunOptions, onClose, onSave }) {
  const [nama, setNama] = useState('');
  const [dusun, setDusun] = useState(dusunOptions[0] || 'Dusun Krajan');
  const [rt, setRt] = useState('01');
  const [rw, setRw] = useState('01');
  const [noKK, setNoKK] = useState('');
  const [noHP, setNoHP] = useState('');

  return (
    <div className="fixed inset-0 z-50 bg-black/80 backdrop-blur-sm flex items-center justify-center p-4">
      <form
        onSubmit={(e) => {
          e.preventDefault();
          if (!nama.trim() || !dusun.trim()) return;
          onSave({ nama, dusun, rt, rw, noKK, noHP });
        }}
        className="bg-slate-900 border border-slate-700 rounded-2xl max-w-md w-full p-6 space-y-3.5 text-xs"
      >
        <h3 className="text-base font-bold text-white">Tambah Warga Baru — Desa Kalikurmo</h3>
        <div>
          <label className="block text-slate-400 mb-1">Nama Lengkap Warga</label>
          <input
            type="text"
            required
            value={nama}
            onChange={(e) => setNama(e.target.value)}
            placeholder="Masukkan nama warga..."
            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
          />
        </div>
        <div>
          <label className="block text-slate-400 mb-1">Nama Dusun (Bisa Ketik Manual)</label>
          <input
            type="text"
            required
            list="dusun-list-kalikurmo"
            value={dusun}
            onChange={(e) => setDusun(e.target.value)}
            placeholder="Ketik nama dusun..."
            className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
          />
          <datalist id="dusun-list-kalikurmo">
            {dusunOptions.map((d) => (
              <option key={d} value={d} />
            ))}
          </datalist>
        </div>
        <div className="grid grid-cols-2 gap-2">
          <div>
            <label className="block text-slate-400 mb-1">Pilih RT (01 - 15)</label>
            <select
              value={rt}
              onChange={(e) => setRt(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
            >
              {RT_RW_LIST.map((r) => (
                <option key={r} value={r}>RT {r}</option>
              ))}
            </select>
          </div>
          <div>
            <label className="block text-slate-400 mb-1">Pilih RW (01 - 15)</label>
            <select
              value={rw}
              onChange={(e) => setRw(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
            >
              {RT_RW_LIST.map((r) => (
                <option key={r} value={r}>RW {r}</option>
              ))}
            </select>
          </div>
        </div>
        <div className="grid grid-cols-2 gap-2">
          <div>
            <label className="block text-slate-400 mb-1">No. KK (Disembunyikan di Publik)</label>
            <input
              type="text"
              value={noKK}
              onChange={(e) => setNoKK(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
            />
          </div>
          <div>
            <label className="block text-slate-400 mb-1">No. HP Warga</label>
            <input
              type="text"
              value={noHP}
              onChange={(e) => setNoHP(e.target.value)}
              className="w-full bg-slate-950 border border-slate-800 rounded-xl px-3 py-2 text-white"
            />
          </div>
        </div>
        <div className="flex justify-end gap-2 pt-2">
          <button type="button" onClick={onClose} className="px-4 py-2 rounded-xl bg-slate-800 text-slate-300 font-semibold">
            Batal
          </button>
          <button type="submit" className="px-5 py-2 rounded-xl bg-emerald-600 text-white font-bold">
            Simpan Warga
          </button>
        </div>
      </form>
    </div>
  );
}
