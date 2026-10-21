#!/usr/bin/env python3
"""Small desktop file-picker front end for store_recon.py."""

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from store_recon import load_roster, load_spend, reconcile, write_report


def main() -> None:
    root = tk.Tk()
    root.withdraw()
    root.update()
    roster_file = filedialog.askopenfilename(
        title="选择总部流量激励名单",
        filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv *.tsv"), ("所有文件", "*.*")],
    )
    if not roster_file:
        return
    spend_files = filedialog.askopenfilenames(
        title="选择一个或多个巨量本地推消耗导出表",
        filetypes=[("Excel/CSV", "*.xlsx *.xlsm *.csv *.tsv"), ("所有文件", "*.*")],
    )
    if not spend_files:
        return
    output = filedialog.asksaveasfilename(
        title="保存核对结果",
        initialdir=str(Path(roster_file).resolve().parent),
        initialfile="门店投流预算核对.xlsx",
        defaultextension=".xlsx",
        filetypes=[("Excel工作簿", "*.xlsx")],
    )
    if not output:
        return
    try:
        entries, excluded = load_roster([Path(roster_file)])
        stores = load_spend([Path(p) for p in spend_files])
        report = reconcile(entries, stores)
        write_report(Path(output), report, excluded)
    except Exception as exc:  # show import/format errors to the user in the dialog
        messagebox.showerror("对账未完成", str(exc))
        return
    messagebox.showinfo(
        "对账完成",
        f"结果已保存：\n{output}\n\n已匹配 {len(report['results'])} 家"
        f"\n待确认 {len(report['pending'])} 家"
        f"\n超预算 {len(report['overspend'])} 家"
        f"\n临近预算 {len(report['near'])} 家"
        f"\n已剔除 {len(excluded)} 家",
    )


if __name__ == "__main__":
    main()
