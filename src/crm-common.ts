/**
 * CRM操作の共通モジュール。
 * 認証・シート検索・企業行検索のボイラープレートを一箇所にまとめる。
 */

import "dotenv/config";
import path from "node:path";
import { google } from "googleapis";
import type { drive_v3, sheets_v4 } from "googleapis";

const TRACKING_SHEET_NAME = "リード管理CRM";

export interface CRMConnection {
  sheets: sheets_v4.Sheets;
  drive: drive_v3.Drive;
  spreadsheetId: string;
  tabs: Map<string, { sheetId: number; title: string }>;
}

/** 認証して CRM スプレッドシートへの接続を返す */
export async function getCRMConnection(): Promise<CRMConnection> {
  const keyPath = process.env.GOOGLE_APPLICATION_CREDENTIALS;
  const folderId = process.env.GOOGLE_DRIVE_FOLDER_ID;
  if (!keyPath || !folderId) {
    throw new Error("GOOGLE_APPLICATION_CREDENTIALS と GOOGLE_DRIVE_FOLDER_ID を .env に設定してください");
  }

  const auth = new google.auth.GoogleAuth({
    keyFile: path.resolve(keyPath),
    scopes: [
      "https://www.googleapis.com/auth/spreadsheets",
      "https://www.googleapis.com/auth/drive",
    ],
    clientOptions: {
      subject: process.env.GOOGLE_IMPERSONATE_USER,
    },
  });

  const drive = google.drive({ version: "v3", auth });
  const sheets = google.sheets({ version: "v4", auth });

  // リード管理CRMを検索
  const listRes = await drive.files.list({
    q: `name='${TRACKING_SHEET_NAME}' and '${folderId}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false`,
    fields: "files(id)",
    pageSize: 1,
  });

  const file = listRes.data.files?.[0];
  if (!file?.id) {
    throw new Error("リード管理CRMが見つかりません");
  }

  const spreadsheetId = file.id;

  // 全タブ情報を取得
  const meta = await sheets.spreadsheets.get({ spreadsheetId, fields: "sheets.properties" });
  const tabs = new Map<string, { sheetId: number; title: string }>();
  for (const s of meta.data.sheets ?? []) {
    const title = s.properties?.title ?? "";
    const sheetId = s.properties?.sheetId ?? 0;
    tabs.set(title, { sheetId, title });
  }

  return { sheets, drive, spreadsheetId, tabs };
}

/** CRM内でB列（会社名）を検索し、最後にマッチした行のインデックス（1-indexed）とデータを返す */
export async function findCompanyRow(
  sheets: sheets_v4.Sheets,
  spreadsheetId: string,
  tabTitle: string,
  companyName: string,
): Promise<{ rowIndex: number; rowData: string[] } | null> {
  if (!companyName) return null;

  const dataRes = await sheets.spreadsheets.values.get({
    spreadsheetId,
    range: `'${tabTitle}'!A:V`,
  });

  const rows = dataRes.data.values ?? [];
  if (rows.length <= 1) return null;

  let lastMatch: { rowIndex: number; rowData: string[] } | null = null;
  for (let i = 1; i < rows.length; i++) {
    const cellValue = String(rows[i][1] ?? "");
    if (cellValue.includes(companyName) || companyName.includes(cellValue)) {
      lastMatch = { rowIndex: i + 1, rowData: rows[i].map(String) };
    }
  }

  return lastMatch;
}

/** 日付を YY_MM_DD 形式にフォーマット */
export function formatDateYYMMDD(date?: Date): string {
  const d = date ?? new Date();
  const yy = String(d.getFullYear()).slice(2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yy}_${mm}_${dd}`;
}

/** 日時を YY_MM_DD HH:MM 形式にフォーマット */
export function formatDateTimeStamp(date?: Date): string {
  const d = date ?? new Date();
  const yy = String(d.getFullYear()).slice(2);
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const min = String(d.getMinutes()).padStart(2, "0");
  return `${yy}_${mm}_${dd} ${hh}:${min}`;
}
