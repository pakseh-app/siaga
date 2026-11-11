import React, { useState, useMemo, useEffect } from 'react';
import {
  Car, Users, Wallet, Settings, FileText, CheckCircle2, Clock, AlertTriangle,
  XCircle, Plus, Search, Filter, Printer, History, ShieldCheck, Eye, Lock,
  Unlock, Calendar, TrendingUp, ArrowUpRight, ArrowDownRight, Edit3, RotateCcw,
  Phone, MapPin, Navigation, Camera, KeyRound, Database, Copy, RefreshCw,
  Check, Cloud, UserPlus, Sparkles, Building2, BarChart3, PieChart, Receipt
} from 'lucide-react';

// Pilihan RT & RW 01 sampai 15
const RT_RW_OPTIONS = Array.from({ length: 15 }, (_, i) => String(i + 1).padStart(2, '0'));
const LOCAL_STORAGE_KEY = 'SIMOSD_DESA_KALIKURMO_V2';
const DEFAULT_GOOGLE_SHEET_URL = '';

// Kode Google Apps Script (Kode.gs) untuk Database Google Sheet Real-Time
const KODE_GS_SCRIPT = [
  "// KODE.GS - SISTEM MANAJEMEN IURAN MOBIL SIAGA DESA KALIKURMO",
  "const SHEET_STATE = 'Database_JSON';",
  "const SHEET_WARGA = 'Rekap_Warga';",
  "const SHEET_TRX = 'Transaksi_Iuran';",
  "const SHEET_ANTAR = 'Riwayat_Antar';",
  "const SHEET_KAS = 'Buku_Kas';",
  "",
  "function ensureSheets() {",
  "  const ss = SpreadsheetApp.getActiveSpreadsheet();",
  "  [SHEET_STATE, SHEET_WARGA, SHEET_TRX, SHEET_ANTAR, SHEET_KAS].forEach(function(name) {",
  "    if (!ss.getSheetByName(name)) ss.insertSheet(name);",
  "  });",
  "  return ss;",
  "}",
  "",
  "function doGet() {",
  "  try {",
  "    const ss = ensureSheets();",
  "    const st = ss.getSheetByName(SHEET_STATE);",
  "    const raw = st.getRange('A2').getValue();",
  "    return ContentService.createTextOutput(JSON.stringify({",
  "      status: 'success',",
  "      updatedAt: st.getRange('B2').getValue() || '',",
  "      data: raw ? JSON.parse(raw) : null",
  "    })).setMimeType(ContentService.MimeType.JSON);",
  "  } catch (err) {",
  "    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))",
  "      .setMimeType(ContentService.MimeType.JSON);",
  "  }",
  "}",
  "",
  "function doPost(e) {",
  "  try {",
  "    const ss = ensureSheets();",
  "    const body = JSON.parse(e.postData.contents);",
  "    const now = new Date().toLocaleString('id-ID');",
  "    const st = ss.getSheetByName(SHEET_STATE);",
  "    st.clear();",
  "    st.getRange('A1:B1').setValues([['JSON_DATA', 'UPDATED_AT']]);",
  "    st.getRange('A2:B2').setValues([[JSON.stringify(body), now]]);",
  "",
  "    if (Array.isArray(body.residents)) {",
  "      const sw = ss.getSheetByName(SHEET_WARGA);",
  "      sw.clear();",
  "      const rows = [['ID', 'Nama Warga', 'Dusun', 'RT', 'RW', 'No KK', 'No HP']];",
  "      body.residents.forEach(function(w) {",
  "        rows.push([w.id, w.nama, w.dusun, \"'\" + w.rt, \"'\" + w.rw, \"'\" + w.noKK, \"'\" + w.noHP]);",
  "      });",
  "      sw.getRange(1, 1, rows.length, rows[0].length).setValues(rows);",
  "    }",
  "    if (Array.isArray(body.transactions)) {",
  "      const stx = ss.getSheetByName(SHEET_TRX);",
  "      stx.clear();",
  "      const rows = [['No Trx', 'Tanggal', 'Nama Warga', 'Dusun', 'RT', 'Mulai', 'Sampai', 'Hari', 'Tarif', 'Nominal', 'Paket', 'Admin', 'Status']];",
  "      body.transactions.forEach(function(t) {",
  "        rows.push([t.id, t.tanggalTransaksi, t.residentName, t.dusun, \"'\" + t.rt, t.periodeMulai, t.periodeAkhir, t.jumlahHari, t.tarifPerHari, t.nominal, t.paketLabel, t.adminName, t.status]);",
  "      });",
  "      stx.getRange(1, 1, rows.length, rows[0].length).setValues(rows);",
  "    }",
  "    return ContentService.createTextOutput(JSON.stringify({ status: 'success', updatedAt: now }))",
  "      .setMimeType(ContentService.MimeType.JSON);",
  "  } catch (err) {",
  "    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))",
  "      .setMimeType(ContentService.MimeType.JSON);",
  "  }",
  "}"
].join('\n');

// Data Awal
const INITIAL_TARIFF_HISTORY = [
  {
    id: 'TRF-001',
    nominal: 2000,
    effectiveDate: '2026-01-01',
    createdAt: '2026-01-01 08:00',
    adminName: 'H. Sudirman (Ketua Pengelola)',
    keterangan: 'Tarif awal kesepakatan Musdes Mobil Siaga Desa Kalikurmo Tahun 20
