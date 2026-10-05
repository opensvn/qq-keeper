#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""QQ 账号保活台账工具.

在腾讯《QQ号码规则》的回收窗口（普通号 3 个月未登录、靓号会员到期 30 天）
之上建立安全余量：默认每 30 天产生一次有效登录，剩余不足 8 天即告警。

设计边界：本工具只做「记账 + 扫描 + 到点提醒」，不做自动登录。
原因：腾讯明文禁止第三方协议库与自动化工具操作 QQ，一旦账号被判定异常
先冻结、冻结满 30 天未解冻即回收，反而比不保活更糟。登录动作由人工完成。

支持多载体场景：每个账号用 host 字段标记其常挂载体（如实体机 / 虚拟机 /
安卓模拟器），台账与提醒统一在持有 db 的宿主机上运行。
"""

from __future__ import annotations

import argparse
import csv
import sqlite3
import subprocess
import sys
import tempfile
import unicodedata
from datetime import date, datetime, timedelta
from pathlib import Path

DB = Path(__file__).resolve().parent / "qq_keeper.db"
DEFAULT_PERIOD = 30      # 保活周期（天），官方回收窗口为 90 天，留 60 天余量
WARN_DAYS = 8            # 剩余天数低于此值进入「即将到期」
VIP_WARN_DAYS = 40       # 靓号会员到期提前告警天数
TOAST_MS = 15000

SCHEMA = """
CREATE TABLE IF NOT EXISTS accounts (
    uin         TEXT PRIMARY KEY,
    alias       TEXT    DEFAULT '',
    acct_type   TEXT    DEFAULT 'normal',
    phone       TEXT    DEFAULT '',
    vip_expire  TEXT    DEFAULT '',
    last_login  TEXT    NOT NULL,
    period      INTEGER NOT NULL DEFAULT 30,
    host        TEXT    DEFAULT '',
    note        TEXT    DEFAULT '',
    created_at  TEXT    NOT NULL
);
"""


# ---------------------------------------------------------------- 基础工具

def connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    conn.execute(SCHEMA)
    _migrate(conn)
    return conn


def _migrate(conn: sqlite3.Connection) -> None:
    """老库向后兼容：缺 host 列则补上。"""
    cols = {r[1] for r in conn.execute("PRAGMA table_info(accounts)")}
    if "host" not in cols:
        conn.execute("ALTER TABLE accounts ADD COLUMN host TEXT DEFAULT ''")


def to_date(text: str) -> date:
    return datetime.strptime(text, "%Y-%m-%d").date()


def pad(text: str, width: int) -> str:
    """按东亚字符宽度补齐，避免中文换行错位。"""
    span = sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)
    return text + " " * max(0, width - span)


def today() -> date:
    return date.today()


# ---------------------------------------------------------------- 业务计算

def rows_with_status(conn: sqlite3.Connection) -> list[dict]:
    """读取全部账号并计算出到期状态，按剩余天数升序排列。"""
    now = today()
    result = []
    for row in conn.execute("SELECT * FROM accounts"):
        item = dict(row)
        due = to_date(item["last_login"]) + timedelta(days=int(item["period"]))
        left = (due - now).days
        if left <= 0:
            state = "已到期"
        elif left <= WARN_DAYS:
            state = "即将到期"
        else:
            state = "正常"
        item["next_due"] = due.isoformat()
        item["days_left"] = left
        item["state"] = state

        item["vip_alert"] = ""
        if item["acct_type"] == "liang" and item["vip_expire"]:
            try:
                vip_left = (to_date(item["vip_expire"]) - now).days
                if vip_left <= VIP_WARN_DAYS:
                    item["vip_alert"] = f"会员 {vip_left} 天后到期"
            except ValueError:
                item["vip_alert"] = "会员日期格式错误"
        result.append(item)
    result.sort(key=lambda x: x["days_left"])
    return result


def print_table(items: list[dict]) -> None:
    headers = ("UIN", "备注", "类型", "绑定手机", "载体", "上次登录", "下次到期", "剩余", "状态")
    widths = (14, 14, 8, 12, 12, 12, 12, 6, 10)
    print("  ".join(pad(h, w) for h, w in zip(headers, widths)))
    print("-" * (sum(widths) + 2 * len(widths)))
    for it in items:
        cells = (
            it["uin"], it["alias"] or "-", "靓号" if it["acct_type"] == "liang" else "普通",
            it["phone"] or "-", it["host"] or "-", it["last_login"], it["next_due"],
            str(it["days_left"]), it["state"],
        )
        line = "  ".join(pad(str(c), w) for c, w in zip(cells, widths))
        print(line)
        if it["vip_alert"]:
            print(f"    ! {it['vip_alert']}")


# ---------------------------------------------------------------- 提醒

def notify(title: str, body: str) -> bool:
    """弹出 Windows 气泡提醒。通过临时 .ps1 传递文本，规避命令行编码问题。"""
    ps = (
        "Add-Type -AssemblyName System.Windows.Forms\n"
        "$n = New-Object System.Windows.Forms.NotifyIcon\n"
        "$n.Icon = [System.Drawing.SystemIcons]::Warning\n"
        f"$n.BalloonTipTitle = '{title}'\n"
        f"$n.BalloonTipText = '{body}'\n"
        "$n.Visible = $true\n"
        f"$n.ShowBalloonTip({TOAST_MS})\n"
        "Start-Sleep -Seconds 10\n"
        "$n.Dispose()\n"
    )
    tmp = Path(tempfile.gettempdir()) / "qq_keeper_notify.ps1"
    try:
        tmp.write_text(ps, encoding="utf-8-sig")
        subprocess.run(
            ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", str(tmp)],
            timeout=30, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, check=False,
        )
        return True
    except Exception:
        return False


# ---------------------------------------------------------------- 子命令

def cmd_init(_: argparse.Namespace) -> int:
    with connect() as conn:
        pass
    print(f"台账已就绪：{DB}")
    return 0


def cmd_add(args: argparse.Namespace) -> int:
    uin = args.uin.strip()
    if not uin.isdigit():
        print("错误：QQ 号必须为纯数字")
        return 1
    last = args.last_login or today().isoformat()
    try:
        to_date(last)
    except ValueError:
        print("错误：日期格式须为 YYYY-MM-DD")
        return 1
    if args.vip_expire:
        try:
            to_date(args.vip_expire)
        except ValueError:
            print("错误：会员到期日期格式须为 YYYY-MM-DD")
            return 1

    with connect() as conn:
        exists = conn.execute("SELECT 1 FROM accounts WHERE uin = ?", (uin,)).fetchone()
        if exists:
            print(f"错误：{uin} 已在台账中，改字段请用 set，仅记录登录用 touch")
            return 1
        conn.execute(
            "INSERT INTO accounts (uin, alias, acct_type, phone, vip_expire, last_login, period, host, note, created_at)"
            " VALUES (?,?,?,?,?,?,?,?,?,?)",
            (uin, args.alias or "", args.type, args.phone or "", args.vip_expire or "",
             last, args.period, args.host or "", args.note or "", today().isoformat()),
        )
    print(f"已登记 {uin}（载体：{args.host or '未标注'}），保活周期 {args.period} 天，"
          f"下次到期 {(to_date(last) + timedelta(days=args.period)).isoformat()}")
    return 0


def cmd_set(args: argparse.Namespace) -> int:
    uin = args.uin.strip()
    updates: dict = {}
    if args.alias is not None:
        updates["alias"] = args.alias
    if args.type is not None:
        updates["acct_type"] = args.type
    if args.phone is not None:
        updates["phone"] = args.phone
    if args.vip_expire is not None:
        try:
            to_date(args.vip_expire)
        except ValueError:
            print("错误：会员到期日期格式须为 YYYY-MM-DD")
            return 1
        updates["vip_expire"] = args.vip_expire
    if args.host is not None:
        updates["host"] = args.host
    if args.period is not None:
        updates["period"] = args.period
    if args.note is not None:
        updates["note"] = args.note
    if args.last_login is not None:
        try:
            to_date(args.last_login)
        except ValueError:
            print("错误：日期格式须为 YYYY-MM-DD")
            return 1
        updates["last_login"] = args.last_login

    if not updates:
        print("未提供任何要更新的字段")
        return 1
    with connect() as conn:
        if not conn.execute("SELECT 1 FROM accounts WHERE uin = ?", (uin,)).fetchone():
            print(f"错误：{uin} 不在台账中，先用 add 登记")
            return 1
        clause = ", ".join(f"{k} = ?" for k in updates)
        conn.execute(f"UPDATE accounts SET {clause} WHERE uin = ?", (*updates.values(), uin))
    print(f"已更新 {uin}：{', '.join(updates.keys())}")
    return 0


def cmd_list(args: argparse.Namespace) -> int:
    if not DB.exists():
        print("错误：尚未初始化，请先执行 init")
        return 1
    with connect() as conn:
        items = rows_with_status(conn)
    if not items:
        print("台账为空，使用 add 登记账号")
        return 0
    print_table(items)
    need = [i for i in items if i["days_left"] <= WARN_DAYS]
    print(f"\n共 {len(items)} 个账号，其中 {len(need)} 个需要尽快登录")
    return 0


def cmd_check(args: argparse.Namespace) -> int:
    if not DB.exists():
        print("错误：尚未初始化，请先执行 init")
        return 1
    with connect() as conn:
        items = rows_with_status(conn)
    due = [i for i in items if i["days_left"] <= 0]
    soon = [i for i in items if 0 < i["days_left"] <= WARN_DAYS]
    vip = [i for i in items if i["vip_alert"]]

    if not due and not soon and not vip:
        print(f"全部安全，最早到期还剩 {items[0]['days_left']} 天（{items[0]['uin']}）")
        return 0
    if due:
        print("【立即登录】")
        for i in due:
            print(f"  {i['uin']}  {i['alias'] or '-'}  已超期 {abs(i['days_left'])} 天  载体：{i['host'] or '-'}")
    if soon:
        print("【本周内登录】")
        for i in soon:
            print(f"  {i['uin']}  {i['alias'] or '-'}  还剩 {i['days_left']} 天  载体：{i['host'] or '-'}")
    if vip:
        print("【靓号会员】")
        for i in vip:
            print(f"  {i['uin']}  {i['vip_alert']}")
    return 0


def cmd_remind(_: argparse.Namespace) -> int:
    """供计划任务调用：有事才打扰。"""
    if not DB.exists():
        return 0
    with connect() as conn:
        items = rows_with_status(conn)
    pending = [i for i in items if i["days_left"] <= WARN_DAYS or i["vip_alert"]]
    if not pending:
        return 0
    lines = []
    for i in pending[:5]:
        tail = i["vip_alert"] or (f"超期 {abs(i['days_left'])} 天" if i["days_left"] <= 0
                                  else f"剩 {i['days_left']} 天")
        lines.append(f"{i['uin']} {i['alias'] or '-'} {tail} [{i['host'] or '-'}]")
    more = f" ...另有 {len(pending) - 5} 个" if len(pending) > 5 else ""
    notify(f"{len(pending)} 个 QQ 号需要保活", "\n".join(lines) + more)
    print("\n".join(lines) + more)
    return 0


def cmd_touch(args: argparse.Namespace) -> int:
    day = args.date or today().isoformat()
    try:
        to_date(day)
    except ValueError:
        print("错误：日期格式须为 YYYY-MM-DD")
        return 1
    with connect() as conn:
        if args.target == "all":
            cur = conn.execute("UPDATE accounts SET last_login = ?", (day,))
            n = cur.rowcount
        else:
            exists = conn.execute("SELECT 1 FROM accounts WHERE uin = ?", (args.target,)).fetchone()
            if not exists:
                print(f"错误：{args.target} 不在台账中")
                return 1
            conn.execute("UPDATE accounts SET last_login = ? WHERE uin = ?", (day, args.target))
            n = 1
    print(f"已记录 {n} 个账号于 {day} 完成登录")
    return 0


def cmd_remove(args: argparse.Namespace) -> int:
    with connect() as conn:
        cur = conn.execute("DELETE FROM accounts WHERE uin = ?", (args.uin,))
    print("已删除" if cur.rowcount else "未找到该账号")
    return 0


def cmd_export(args: argparse.Namespace) -> int:
    with connect() as conn:
        items = rows_with_status(conn)
    out = Path(args.path)
    with out.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["UIN", "备注", "类型", "绑定手机", "载体", "上次登录", "会员到期", "备注说明"])
        for i in items:
            writer.writerow([
                i["uin"], i["alias"], i["acct_type"], i["phone"], i["host"], i["last_login"],
                i["vip_expire"], i["note"],
            ])
    print(f"已导出 {len(items)} 条到 {out}")
    return 0


def cmd_import(args: argparse.Namespace) -> int:
    src = Path(args.path)
    if not src.exists():
        print("错误：文件不存在")
        return 1
    added = 0
    with src.open(encoding="utf-8-sig", newline="") as fh, connect() as conn:
        for row in csv.DictReader(fh):
            uin = (row.get("UIN") or "").strip()
            if not uin.isdigit():
                continue
            if conn.execute("SELECT 1 FROM accounts WHERE uin = ?", (uin,)).fetchone():
                continue
            last = (row.get("上次登录") or "").strip() or today().isoformat()
            conn.execute(
                "INSERT INTO accounts (uin, alias, acct_type, phone, vip_expire, last_login, period, host, note, created_at)"
                " VALUES (?,?,?,?,?,?,?,?,?,?)",
                (uin, (row.get("备注") or "").strip(),
                 "liang" if (row.get("类型") or "") == "靓号" else "normal",
                 (row.get("绑定手机") or "").strip(), (row.get("会员到期") or "").strip(),
                 last, DEFAULT_PERIOD, (row.get("载体") or "").strip(),
                 (row.get("备注说明") or "").strip(), today().isoformat()),
            )
            added += 1
    print(f"已导入 {added} 个账号")
    return 0


def cmd_install_task(_: argparse.Namespace) -> int:
    py = Path(sys.executable).resolve()
    script = Path(__file__).resolve()
    if py.name.lower().startswith("pythonw"):
        py = py.with_name("python.exe")
    task = f'schtasks /Create /TN "QQKeeperRemind" /TR "\\"{py}\\" \\"{script}\\" remind" /SC DAILY /ST 10:00 /F'
    print("将创建每日 10:00 的计划任务：")
    print("  " + task)
    rc = subprocess.run(task, shell=True, check=False).returncode
    print("创建成功" if rc == 0 else f"创建失败，退出码 {rc}（若为权限不足请以管理员身份重试）")
    return rc


# ---------------------------------------------------------------- 入口

def build_parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(
        prog="qq_keeper",
        description="QQ 账号保活台账：登记、扫描、到期提醒（不做自动登录）",
        epilog="每个子命令可用 `qq_keeper.py <命令> --help` 查看完整参数说明。详见同目录 README.md。",
    )
    sub = p.add_subparsers(dest="cmd", required=True, metavar="<命令>")

    sub.add_parser(
        "init",
        help="初始化台账数据库",
        description="在脚本同目录创建 qq_keeper.db 台账数据库（已存在则跳过）。首次使用前必须执行一次。",
    ).set_defaults(func=cmd_init)

    a = sub.add_parser(
        "add",
        help="登记一个新账号",
        description="把一个 QQ 号登记进台账。相同号码重复 add 会被拒绝，改字段请用 set。",
    )
    a.add_argument("uin", metavar="QQ号", help="要登记的 QQ 号码（纯数字）")
    a.add_argument("--alias", help="备注名 / 昵称，如 主号、小号A")
    a.add_argument("--type", default="normal", choices=["normal", "liang"],
                   help="账号类型：normal=普通号（默认），liang=精品/顶级靓号")
    a.add_argument("--phone", help="绑定手机号，用于找回与风控评估")
    a.add_argument("--vip-expire", dest="vip_expire", metavar="YYYY-MM-DD", help="靓号会员到期日；临近 40 天内会告警")
    a.add_argument("--last-login", dest="last_login", metavar="YYYY-MM-DD", help="上次登录日，默认今天")
    a.add_argument("--period", type=int, default=DEFAULT_PERIOD, metavar="天数", help=f"保活周期，默认 {DEFAULT_PERIOD} 天")
    a.add_argument("--host", help="常挂载体标识，如 实体机 / VM-win / 模拟器-MuMu")
    a.add_argument("--note", help="任意备注说明")
    a.set_defaults(func=cmd_add)

    sub.add_parser(
        "list",
        help="列出全部账号与到期状态",
        description="以表格输出所有账号的下次到期日、剩余天数与状态，并统计需处理的账号数。",
    ).set_defaults(func=cmd_list)

    sub.add_parser(
        "check",
        help="只输出需要处理的账号",
        description="仅在存在到期 / 即将到期 / 靓号会员临期时才输出，适合计划任务静默调用。",
    ).set_defaults(func=cmd_check)

    sub.add_parser(
        "remind",
        help="弹窗提醒（供计划任务调用）",
        description="扫描台账，仅当有待办时弹出 Windows 右下角气泡提醒；无事不打扰。",
    ).set_defaults(func=cmd_remind)

    s = sub.add_parser(
        "set",
        help="更新已存在账号的字段",
        description="只修改显式提供的字段，其余保持不变。本工具不存密码；改密码请在 QQ 客户端操作。",
    )
    s.add_argument("uin", metavar="QQ号", help="要更新的 QQ 号码")
    s.add_argument("--alias", help="新的备注名 / 昵称")
    s.add_argument("--type", choices=["normal", "liang"], help="账号类型：normal=普通号，liang=靓号")
    s.add_argument("--phone", help="新的绑定手机号")
    s.add_argument("--vip-expire", dest="vip_expire", metavar="YYYY-MM-DD", help="新的靓号会员到期日")
    s.add_argument("--host", help="新的常挂载体标识")
    s.add_argument("--period", type=int, metavar="天数", help="新的保活周期")
    s.add_argument("--note", help="新的备注说明")
    s.add_argument("--last-login", dest="last_login", metavar="YYYY-MM-DD",
                   help="修正上次登录日（一般不改，记录登录请用 touch）")
    s.set_defaults(func=cmd_set)

    t = sub.add_parser(
        "touch",
        help="记录一次登录完成（计时归零）",
        description="更新账号的 last_login 为今天或指定日期，使到期计时重置。登录动作需自行在 QQ 客户端完成。",
    )
    t.add_argument("target", metavar="QQ号|all", help="指定 QQ 号，或 all 表示全部账号")
    t.add_argument("--date", metavar="YYYY-MM-DD", help="登录日期，默认今天")
    t.set_defaults(func=cmd_touch)

    r = sub.add_parser(
        "remove",
        help="删除一个账号",
        description="从台账中移除指定 QQ 号，不可恢复，请谨慎。",
    )
    r.add_argument("uin", metavar="QQ号", help="要删除的 QQ 号码")
    r.set_defaults(func=cmd_remove)

    e = sub.add_parser(
        "export",
        help="导出全部账号为 CSV",
        description="导出为带 BOM 的 UTF-8 CSV，可用 Excel 直接打开，便于备份或编辑后回导。",
    )
    e.add_argument("path", metavar="文件.csv", help="输出 CSV 文件路径")
    e.set_defaults(func=cmd_export)

    i = sub.add_parser(
        "import",
        help="从 CSV 批量导入账号",
        description="读取 CSV（表头含 UIN/备注/类型/绑定手机/载体/上次登录/会员到期/备注说明）。已存在的号码自动跳过。",
    )
    i.add_argument("path", metavar="文件.csv", help="输入 CSV 文件路径")
    i.set_defaults(func=cmd_import)

    sub.add_parser(
        "install-task",
        help="注册每日 10:00 的 Windows 计划任务",
        description="创建 QQKeeperRemind 计划任务，每日定时执行 remind；权限不足请以管理员身份运行。",
    ).set_defaults(func=cmd_install_task)
    return p


def main() -> int:
    args = build_parser().parse_args()
    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
