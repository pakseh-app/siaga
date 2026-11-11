import React, { useState, useEffect, useMemo } from 'react';
import {
  Car, Users, Wallet, Settings, FileText, CheckCircle2, Clock,
  AlertTriangle, XCircle, Plus, Search, Filter, Printer, History,
  ShieldCheck, Eye, Lock, Unlock, Calendar, TrendingUp, ArrowUpRight,
  ArrowDownRight, Edit3, RotateCcw, Info, Phone, MapPin, Navigation,
  Award, Camera, KeyRound, Database, Copy, RefreshCw, Check, Cloud,
  UserPlus, Building2, CheckSquare, BarChart3, Receipt
} from 'lucide-react';

// ============================================================================
// KONFIGURASI DATABASE GOOGLE SHEETS & LOCAL STORAGE DESA KALIKURMO
// ============================================================================
// Jika sudah punya URL Web App Google Script, tempel di antara tanda kutip di bawah ini:
const DEFAULT_GOOGLE_SHEET_WEB_APP_URL = "";
const LOCAL_STORAGE_KEY = "SIMOSD_DESA_KALIKURMO_DB_V2";

// Pilihan RT & RW 01 sampai 15
const RT_RW_OPTIONS = Array.from({ length: 15 }, (_, i) => String(i + 1).padStart(2, '0'));

// ============================================================================
// KODE.GS LENGKAP UNTUK DATABASE GOOGLE SHEETS REALTIME
// ============================================================================
const GOOGLE_APPS_SCRIPT_CODE = `/**
 * KODE.GS — SISTEM INFORMASI MOBIL SIAGA DESA KALIKURMO (SIMOSD)
 * CARA PASANG:
 * 1. Buat Google Sheet baru di Google Drive Anda.
 * 2. Klik menu Ekstensi (Extensions) -> Apps Script.
 * 3. Hapus semua kode di Code.gs, lalu tempel (Paste) seluruh kode ini.
 * 4. Klik Simpan (ikon disket), lalu klik "Terapkan" (Deploy) -> "Deployment baru".
 * 5. Pilih jenis: "Aplikasi Web" (Web App).
 *    - Jalankan sebagai: "Saya" (Me)
 *    - Siapa yang memiliki akses: "Siapa saja" (Anyone)
 * 6. Klik Terapkan (Deploy), salin URL Aplikasi Web, lalu tempel di menu "Database & Kode.gs".
 */

const SHEET_NAMES = {
  STATE: 'Database_SIMOSD_JSON',
  WARGA: 'Data_Warga',
  TRANSAKSI: 'Transaksi_Iuran',
  ANTAR: 'Riwayat_Antar_Warga',
  KAS: 'Buku_Kas_Siaga'
};

function ensureSheetsExist() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.values(SHEET_NAMES).forEach(function(name) {
    if (!ss.getSheetByName(name)) {
      ss.insertSheet(name);
    }
  });
  return ss;
}

function doGet(e) {
  try {
    const ss = ensureSheetsExist();
    const stateSheet = ss.getSheetByName(SHEET_NAMES.STATE);
    const raw = stateSheet.getRange('A2').getValue();
    let payload = null;
    if (raw) {
      payload = JSON.parse(raw);
    }
    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      updatedAt: stateSheet.getRange('B2').getValue() || null,
      data: payload
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const ss = ensureSheetsExist();
    const body = JSON.parse(e.postData.contents);
    const now = new Date().toLocaleString('id-ID');

    // 1. Simpan snapshot JSON utama
    const stateSheet = ss.getSheetByName(SHEET_NAMES.STATE);
    stateSheet.clear();
    stateSheet.getRange('A1:B1').setValues([['JSON_DATABASE_SIMOSD', 'LAST_UPDATED']]);
    stateSheet.getRange('A2:B2').setValues([[JSON.stringify(body), now]]);

    // 2. Tulis ke tabel Data_Warga
    if (Array.isArray(body.residents)) {
      const sWarga = ss.getSheetByName(SHEET_NAMES.WARGA);
      sWarga.clear();
      const rows = [['ID Warga', 'Nama Lengkap', 'Dusun', 'RT', 'RW', 'No KK', 'No HP', 'Terdaftar Sejak']];
      body.residents.forEach(function(w) {
        rows.push([w.id, w.nama, w.dusun, "'" + w.rt, "'" + w.rw, "'" + w.noKK, "'" + w.noHP, w.terdaftarSejak]);
      });
      sWarga.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    }

    // 3. Tulis ke tabel Transaksi_Iuran
    if (Array.isArray(body.transactions)) {
      const sTrx = ss.getSheetByName(SHEET_NAMES.TRANSAKSI);
      sTrx.clear();
      const rows = [['No Transaksi', 'Tgl Transaksi', 'ID Warga', 'Nama Warga', 'Dusun', 'RT', 'RW', 'Mulai Periode', 'Akhir Pelunasan', 'Jml Hari', 'Tarif/Hari', 'Nominal', 'Paket', 'Admin', 'Status', 'Keterangan']];
      body.transactions.forEach(function(t) {
        rows.push([t.id, t.tanggalTransaksi, t.residentId, t.residentName, t.dusun, "'" + t.rt, "'" + (t.rw || '01'), t.periodeMulai, t.periodeAkhir, t.jumlahHari, t.tarifPerHari, t.nominal, t.paketLabel, t.adminName, t.status, t.keterangan]);
      });
      sTrx.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    }

    // 4. Tulis ke tabel Riwayat_Antar_Warga
    if (Array.isArray(body.tripLogs)) {
      const sAntar = ss.getSheetByName(SHEET_NAMES.ANTAR);
      sAntar.clear();
      const rows = [['ID Perjalanan', 'Tanggal', 'Jam Berangkat', 'Jam Kembali', 'Nama Warga', 'Dusun', 'RT', 'RW', 'Tujuan RS/Klinik', 'Keperluan', 'Sopir', 'Uang Keluar (BBM)', 'Uang Masuk (Infaq)', 'Keterangan']];
      body.tripLogs.forEach(function(tr) {
        rows.push([tr.id, tr.tanggal, tr.waktuBerangkat, tr.waktuKembali, tr.residentName, tr.dusun, "'" + tr.rt, "'" + (tr.rw || '01'), tr.tujuan, tr.keperluan, tr.driverName, tr.uangKeluar, tr.uangMasuk, tr.keterangan]);
      });
      sAntar.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      updatedAt: now
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}`;

// ============================================================================
// FUNGSI UTILITAS TANGGAL & RUPIAH (KALENDER SEBENARNYA)
// ============================================================================
const formatRupiah = (amount) => {
  return new Intl.NumberFormat('id-ID', {
    style: 'currency',
    currency: 'IDR',
    minimumFractionDigits: 0,
    maximumFractionDigits: 0,
  }).format(amount || 0);
};

const parseDate = (dateStr) => {
  if (!dateStr) return null;
  const [y, m, d] = dateStr.split('-').map(Number);
  return new Date(y, m - 1, d);
};

const toDateStr = (dateObj) => {
  if (!dateObj) return '';
  const y = dateObj.getFullYear();
  const m = String(dateObj.getMonth() + 1).padStart(2, '0');
  const d = String(dateObj.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

const formatDateID = (dateStr) => {
  if (!dateStr) return '-';
  const [y, m, d] = dateStr.split('-');
  return `${d}/${m}/${y}`;
};

const formatDateLongID = (dateStr) => {
  if (!dateStr) return '-';
  const date = parseDate(dateStr);
  return date.toLocaleDateString('id-ID', {
    day: 'numeric',
    month: 'long',
    year: 'numeric',
  });
};

const addDays = (dateStr, daysToAdd) => {
  const d = parseDate(dateStr);
  d.setDate(d.getDate() + (Number(daysToAdd) - 1));
  return toDateStr(d);
};

const addNextDay = (dateStr) => {
  const d = parseDate(dateStr);
  d.setDate(d.getDate() + 1);
  return toDateStr(d);
};

const addCalendarMonths = (dateStr, monthsToAdd) => {
  const start = parseDate(dateStr);
  const end = new Date(start.getFullYear(), start.getMonth() + Number(monthsToAdd), start.getDate());
  end.setDate(end.getDate() - 1);
  return toDateStr(end);
};

const addCalendarYears = (dateStr, yearsToAdd) => {
  const start = parseDate(dateStr);
  const end = new Date(start.getFullYear() + Number(yearsToAdd), start.getMonth(), start.getDate());
  end.setDate(end.getDate() - 1);
  return toDateStr(end);
};

const diffCalendarDaysInclusive = (startStr, endStr) => {
  if (!startStr || !endStr) return 0;
  const s = parseDate(startStr);
  const e = parseDate(endStr);
  const diffTime = e.getTime() - s.getTime();
  const days = Math.round(diffTime / (1000 * 60 * 60 * 24)) + 1;
  return days > 0 ? days : 0;
};

const diffDaysExclusive = (fromDateStr, toDateStr) => {
  if (!fromDateStr || !toDateStr) return 0;
  const s = parseDate(fromDateStr);
  const e = parseDate(toDateStr);
  const diffTime = e.getTime() - s.getTime();
  const days = Math.round(diffTime / (1000 * 60 * 60 * 24));
  return days > 0 ? days : 0;
};

const getEndOfMonthStr = (dateStr) => {
  const d = parseDate(dateStr);
  const lastDay = new Date(d.getFullYear(), d.getMonth() + 1, 0);
  return toDateStr(lastDay);
};

const getEndOfNextMonthStr = (dateStr) => {
  const d = parseDate(dateStr);
  const lastDayNext = new Date(d.getFullYear(), d.getMonth() + 2, 0);
  return toDateStr(lastDayNext);
};

// ============================================================================
// DATA AWAL DEFAULT DESA KALIKURMO
// ============================================================================
const INITIAL_TARIFF_HISTORY = [
  {
    id: 'TRF-2026-01',
    nominal: 2000,
    effectiveDate: '2026-01-01',
    keterangan: 'Penetapan awal Musdes Desa Kalikurmo Tahun 2026 (Rp2.000/hari)',
    adminName: 'Bpk. Hendra Wijaya (Admin Desa Kalikurmo)',
    createdAt: '2026-01-01 08:00',
  },
];

const INITIAL_VILLAGE_OFFICIALS = {
  kades: {
    nama: 'H. Muhammad Yusuf, S.Sos.',
    jabatan: 'Kepala Desa Kalikurmo',
    periode: '2024 - 2030',
    noHP: '081234560001',
    alamat: 'Dusun Krajan RT 01 / RW 01, Desa Kalikurmo',
    pesanSiaga: 'Mobil Siaga Desa Kalikurmo gratis 24 jam untuk melayani darurat medis, ibu melahirkan, dan rujukan rumah sakit seluruh warga Desa Kalikurmo.',
    fotoUrl: '',
  },
  armada: {
    merk: 'Suzuki APV Arena GX — Mobil Siaga Desa Kalikurmo',
    nopol: 'H 9912 XA',
    status: 'SIAGA 24 JAM',
    lokasiParkir: 'Garasi Balai Desa Kalikurmo',
    kelengkapan: 'Brankar Pasien, Tabung Oksigen, Kotak P3K, Sirine Darurat',
  },
  drivers: [
    {
      id: 'DRV-01',
      nama: 'Kang Slamet Riyadi',
      panggilan: 'Bang Slamet',
      noHP: '081398765432',
      shift: 'Shift
