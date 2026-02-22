#!/usr/bin/env npx tsx
/**
 * CRMタスク管理スクリプト。
 * Usage:
 *   npx tsx scripts/manage-crm-tasks.ts --list                         # 未完了タスク一覧
 *   npx tsx scripts/manage-crm-tasks.ts --list --company "株式会社〇〇"  # 企業別タスク
 *   npx tsx scripts/manage-crm-tasks.ts --add --company "xxx" --task "電話する" --due "2026-03-01" --priority 高
 *   npx tsx scripts/manage-crm-tasks.ts --complete --task-id "T-001"
 *   npx tsx scripts/manage-crm-tasks.ts --overdue                      # 期限切れタスク一覧
 */

import "dotenv/config";
import { getCRMConnection, formatDateYYMMDD } from "../src/crm-common.js";

const TAB_TASKS = "タスク";

function parseArgs(argv: string[]): {
  list: boolean;
  add: boolean;
  complete: boolean;
  overdue: boolean;
  companyName: string;
  task: string;
  due: string;
  priority: string;
  taskId: string;
} {
  let list = false, add = false, complete = false, overdue = false;
  let companyName = "", task = "", due = "", priority = "中", taskId = "";

  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--list") list = true;
    else if (argv[i] === "--add") add = true;
    else if (argv[i] === "--complete") complete = true;
    else if (argv[i] === "--overdue") overdue = true;
    else if (argv[i] === "--company" && argv[i + 1]) companyName = argv[++i];
    else if (argv[i] === "--task" && argv[i + 1]) task = argv[++i];
    else if (argv[i] === "--due" && argv[i + 1]) due = argv[++i];
    else if (argv[i] === "--priority" && argv[i + 1]) priority = argv[++i];
    else if (argv[i] === "--task-id" && argv[i + 1]) taskId = argv[++i];
    else if (argv[i] === "-h" || argv[i] === "--help") {
      console.log(`
Usage: npx tsx scripts/manage-crm-tasks.ts [options]

  CRMタスクの管理（一覧表示・追加・完了・期限切れ確認）。

  --list                  未完了タスク一覧（--companyで企業絞り込み可）
  --add                   タスク追加（--company, --task, --due, --priority）
  --complete              タスク完了（--task-id）
  --overdue               期限切れタスク一覧
  --company NAME          会社名
  --task TEXT             タスク内容
  --due DATE              期限（YYYY-MM-DD形式）
  --priority 高|中|低      優先度（デフォルト: 中）
  --task-id ID            タスクID（T-001形式）
`);
      process.exit(0);
    }
  }
  return { list, add, complete, overdue, companyName, task, due, priority, taskId };
}

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  const conn = await getCRMConnection();
  const taskTab = conn.tabs.get(TAB_TASKS);

  if (!taskTab) {
    console.error("Error: タスクタブが見つかりません。gen-reportを一度実行してCRMタブを作成してください。");
    process.exit(1);
  }

  const dataRes = await conn.sheets.spreadsheets.values.get({
    spreadsheetId: conn.spreadsheetId,
    range: `'${TAB_TASKS}'!A:H`,
  });
  const rows = dataRes.data.values ?? [];

  if (args.add) {
    if (!args.companyName || !args.task) {
      console.error("Error: --company と --task は必須です");
      process.exit(2);
    }

    // タスクIDを自動採番
    let maxId = 0;
    for (let i = 1; i < rows.length; i++) {
      const id = String(rows[i][0] ?? "");
      const match = id.match(/^T-(\d+)$/);
      if (match) maxId = Math.max(maxId, parseInt(match[1], 10));
    }
    const newId = `T-${String(maxId + 1).padStart(3, "0")}`;
    const today = formatDateYYMMDD();

    await conn.sheets.spreadsheets.values.append({
      spreadsheetId: conn.spreadsheetId,
      range: `'${TAB_TASKS}'!A:H`,
      valueInputOption: "USER_ENTERED",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [[
          newId,
          args.companyName,
          args.task,
          args.due,
          args.priority,
          "未着手",
          today,
          "",
        ]],
      },
    });

    console.error(`✅ タスクを追加しました: ${newId} / ${args.companyName} / ${args.task}`);
    return;
  }

  if (args.complete) {
    if (!args.taskId) {
      console.error("Error: --task-id は必須です");
      process.exit(2);
    }

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0] ?? "") === args.taskId) {
        const rowNum = i + 1;
        const tabTitle = TAB_TASKS;
        const today = formatDateYYMMDD();

        await conn.sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: conn.spreadsheetId,
          requestBody: {
            valueInputOption: "USER_ENTERED",
            data: [
              { range: `'${tabTitle}'!F${rowNum}`, values: [["完了"]] },
              { range: `'${tabTitle}'!H${rowNum}`, values: [[today]] },
            ],
          },
        });

        console.error(`✅ タスク ${args.taskId} を完了にしました`);
        return;
      }
    }

    console.error(`Error: タスクID「${args.taskId}」が見つかりません`);
    process.exit(1);
  }

  // 一覧表示（--list or --overdue）
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const tasks = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!row || !row[0]) continue;

    const status = String(row[5] ?? "");
    const company = String(row[1] ?? "");

    // フィルタリング
    if (args.companyName && !company.includes(args.companyName) && !args.companyName.includes(company)) continue;
    if (!args.overdue && status === "完了") continue;

    const dueStr = String(row[3] ?? "");
    let isOverdue = false;
    if (dueStr) {
      const dueDate = new Date(dueStr);
      isOverdue = dueDate < today && status !== "完了";
    }

    if (args.overdue && !isOverdue) continue;

    tasks.push({
      taskId: row[0] ?? "",
      company: row[1] ?? "",
      task: row[2] ?? "",
      due: row[3] ?? "",
      priority: row[4] ?? "",
      status: row[5] ?? "",
      created: row[6] ?? "",
      completed: row[7] ?? "",
      overdue: isOverdue,
    });
  }

  if (tasks.length === 0) {
    console.error(args.overdue ? "期限切れタスクはありません" : "該当するタスクがありません");
    process.exit(0);
  }

  console.log(JSON.stringify(tasks, null, 2));
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
