"""
╔═══════════════════════════════════════════════════════════════════════════════╗
║               GT MASS DUMP GENERATOR — v2.3                                  ║
║               Tkinter GUI Desktop Application                                ║
╠═══════════════════════════════════════════════════════════════════════════════╣
║  Author  : Agami AI / Vishal                                                ║
║  Version : 2.3 (Code quality refactor — zero output changes)                ║
║  Purpose : Reads GT-Mass / Everyday PO Excel files from distributors,       ║
║            extracts meta info (SO Number, Distributor, City, State,          ║
║            Location) and ordered items (BC Code, Qty, Tester Qty),          ║
║            generates ERP-importable Sales Order sheets.                      ║
║  Stack   : Python 3.13, Tkinter, pandas, openpyxl                           ║
╚═══════════════════════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════════════════════
  CHANGELOG
═══════════════════════════════════════════════════════════════════════════════

  v2.3 — Code Quality Refactor (zero output changes)
    ┌─────────┬──────────────────────────┬───────────────────────────────┐
    │  FIX #  │ Problem                  │ Solution                      │
    ├─────────┼──────────────────────────┼───────────────────────────────┤
    │  ❌ 3   │ God-class EmailService   │ Split → EmailBuilder +        │
    │         │ (build + send + data)    │ EmailSender (SRP)             │
    ├─────────┼──────────────────────────┼───────────────────────────────┤
    │  ❌ 4   │ parse() 100+ lines       │ Split → _find_header_row,     │
    │         │ (too long)               │ _resolve_so_number,           │
    │         │                          │ _extract_rows                 │
    ├─────────┼──────────────────────────┼───────────────────────────────┤
    │  ❌ 5   │ Inline color constants   │ → class Colors (centralized)  │
    ├─────────┼──────────────────────────┼───────────────────────────────┤
    │  ❌ 7   │ except Exception (broad) │ → specific exception types    │
    ├─────────┼──────────────────────────┼───────────────────────────────┤
    │  ❌ 9   │ Repeated pd.notna()      │ → safe_str_val() helper       │
    ├─────────┼──────────────────────────┼───────────────────────────────┤
    │  ❌ 10  │ Verbose naming           │ ProcessResult,                │
    │         │                          │ MetadataExtractor,            │
    │         │                          │ SONumberFormatter             │
    └─────────┴──────────────────────────┴───────────────────────────────┘

  v2.2 — File → SO Mapping, full SKU email, source_file traceability
  v2.1 — SO from PO Number field, D365 export, email reports
  v2.0 — Initial ERP import format
"""

from __future__ import annotations
import os, sys, platform, time, logging, re, smtplib
from email.message import EmailMessage
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple, Dict
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")

# ═══════════════════════════════════════════════════════════════════════════════
#  EXPIRY CHECK
# ═══════════════════════════════════════════════════════════════════════════════

EXPIRY_DATE = "30-06-2026"

def check_expiry():
    """Block if expired, warn if expiring within 7 days."""
    expiry = datetime.strptime(EXPIRY_DATE, "%d-%m-%Y").date()
    today = datetime.now().date()
    if today > expiry:
        root = tk.Tk(); root.withdraw()
        messagebox.showerror("Application Expired",
            f"This application expired on {EXPIRY_DATE}.\nPlease contact the administrator.")
        root.destroy(); sys.exit(0)
    days_remaining = (expiry - today).days
    if days_remaining <= 7:
        root = tk.Tk(); root.withdraw()
        messagebox.showwarning("Expiration Warning",
            f"⚠️ Expires in {days_remaining} day(s).\nExpiry: {EXPIRY_DATE}")
        root.destroy()


# ═══════════════════════════════════════════════════════════════════════════════
#  FIX ❌5: Centralized color palette (was inline constants in _build_html)
# ═══════════════════════════════════════════════════════════════════════════════

class Colors:
    """Centralized color palette — change once, updates email + UI everywhere."""
    NAVY   = '#1A237E'
    GREEN  = '#2E7D32'
    ORANGE = '#E65100'
    PURPLE = '#6A1B9A'
    GOLD   = '#FFD600'
    GRAY   = '#666666'
    LTGRAY = '#f5f5f5'


# ═══════════════════════════════════════════════════════════════════════════════
#  CONSTANTS
# ═══════════════════════════════════════════════════════════════════════════════

LOCATION_CODE_MAP: Dict[str, str] = {
    'AHD': 'PICK',
    'BLR': 'DS_BL_OFF1',
}

STATE_LIKE_VALUES = {
    "up", "mp", "ap", "hp", "uk", "jk", "wb", "tn", "kl", "ka",
    "gj", "rj", "hr", "pb", "br", "od", "as", "mh", "cg", "jh",
    "north", "south", "east", "west", "central",
    "uttar pradesh", "madhya pradesh", "rajasthan", "punjab",
    "maharashtra", "gujarat", "karnataka", "tamil nadu",
    "haryana", "delhi", "u.p", "u.p.", "m.p", "m.p."
}

EMAIL_CONFIG = {
    'EMAIL_SENDER': 'abhishekwagh420@gmail.com',
    'EMAIL_PASSWORD': 'bomn ktfx jhct xexy',
    'SMTP_SERVER': 'smtp.gmail.com',
    'SMTP_PORT': 587,
    'DEFAULT_RECIPIENT': 'abhishek.wagh@reneecosmetics.in',
    'CC_RECIPIENTS': ['offlineb2b@reneecosmetics.in',
                      'kirpalsinh.bihola@reneecosmetics.in',
                      'aritra.barmanray@reneecosmetics.in',
                      'milan.nayak@reneecosmetics.in'],
}


# ═══════════════════════════════════════════════════════════════════════════════
#  FIX ❌9: DRY helper — replaces 6× repeated pd.notna() extraction pattern
# ═══════════════════════════════════════════════════════════════════════════════

def safe_str_val(row_vals, idx: Optional[int], as_int_str: bool = False) -> str:
    """
    Safely extract a string from row_vals[idx].

    FIX ❌9: Before, every field had:
        val = ''
        if idx is not None and pd.notna(row_vals[idx]):
            val = str(row_vals[idx]).strip()
    Now: val = safe_str_val(row_vals, idx)

    Args:
        row_vals   : numpy array (one data row)
        idx        : column index (None if column wasn't detected)
        as_int_str : True → convert float→int→str (removes '.0' from EAN)
    """
    if idx is None:
        return ''
    val = row_vals[idx]
    if pd.isna(val):
        return ''
    if as_int_str and isinstance(val, (int, float)):
        return str(int(val))
    return str(val).strip()


def format_indian(number) -> str:
    """Format number in Indian system: 1,23,456."""
    try:
        number = float(number)
    except (ValueError, TypeError):
        return str(number)
    sign = '-' if number < 0 else ''
    number = abs(number)
    if number == int(number):
        int_part, dec_part = str(int(number)), ''
    else:
        parts = f"{number:.2f}".split('.')
        int_part, dec_part = parts[0], '.' + parts[1]
    if len(int_part) <= 3:
        return sign + int_part + dec_part
    result = int_part[-3:]
    remaining = int_part[:-3]
    while remaining:
        result = remaining[-2:] + ',' + result
        remaining = remaining[:-2]
    return sign + result + dec_part


# ═══════════════════════════════════════════════════════════════════════════════
#  FIX ❌3: God-class split → EmailBuilder (HTML) + EmailSender (SMTP)
# ═══════════════════════════════════════════════════════════════════════════════

class EmailBuilder:
    """Pure data → HTML transform. No network I/O. Testable without SMTP."""

    @staticmethod
    def _aggregate(result: 'ProcessResult') -> dict:
        """Aggregate OrderRows into email-ready summaries."""
        unique_sos = list({r.so_number: r for r in result.rows}.values())
        total_order = sum(r.qty for r in result.rows)
        total_tester = sum(r.tester_qty for r in result.rows)

        sku_groups: Dict[str, dict] = {}
        for r in result.rows:
            if r.item_no not in sku_groups:
                sku_groups[r.item_no] = {'desc': r.description, 'cat': r.category, 'order': 0, 'tester': 0}
            sku_groups[r.item_no]['order'] += r.qty
            sku_groups[r.item_no]['tester'] += r.tester_qty
            if not sku_groups[r.item_no]['desc'] and r.description:
                sku_groups[r.item_no]['desc'] = r.description

        so_groups: Dict[str, dict] = {}
        for r in result.rows:
            if r.so_number not in so_groups:
                so_groups[r.so_number] = {'order': 0, 'tester': 0}
            so_groups[r.so_number]['order'] += r.qty
            so_groups[r.so_number]['tester'] += r.tester_qty

        return {
            'unique_sos': unique_sos, 'so_groups': so_groups,
            'total_items': len(result.rows), 'total_order': total_order,
            'total_tester': total_tester,
            'sorted_skus': sorted(sku_groups.items(), key=lambda x: x[1]['order']+x[1]['tester'], reverse=True),
        }

    @staticmethod
    def build_subject(result: 'ProcessResult') -> str:
        ts = datetime.now().strftime('%d-%m-%Y %H:%M')
        return f"📊 GT Mass SO Report: {len({r.so_number for r in result.rows})} SOs, {len(result.rows)} Items — {ts}"

    @staticmethod
    def build_html(result: 'ProcessResult', elapsed_str: str) -> str:
        """Build complete HTML email. Uses Colors class (FIX ❌5)."""
        d = EmailBuilder._aggregate(result)
        C = Colors
        total_qty = d['total_order'] + d['total_tester']
        ts = datetime.now().strftime('%d-%m-%Y %H:%M:%S')

        html = f"""<html><body style="margin:0;padding:0;font-family:Arial,sans-serif;background:#f0f2f5;">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#f0f2f5;">
<tr><td align="center" style="padding:20px 10px;">
<table width="800" cellpadding="0" cellspacing="0" style="background:#fff;border-radius:8px;overflow:hidden;border:1px solid #ddd;">
<tr><td style="background:{C.NAVY};padding:25px 30px;text-align:center;">
    <p style="margin:0;font-size:22px;font-weight:bold;color:white;">📊 GT Mass — Sales Order Report</p>
    <p style="margin:8px 0 0;font-size:12px;color:#9fa8da;">Generated: {ts} | Processing: {elapsed_str}</p>
    <table style="margin:10px auto 0;"><tr><td style="background:#283593;padding:5px 15px;border-radius:15px;">
        <span style="font-size:10px;color:#9fa8da;letter-spacing:1px;">⚡ GT MASS DUMP GENERATOR v2.3</span>
    </td></tr></table>
</td></tr>
<tr><td style="height:4px;font-size:0;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td width="25%" style="background:{C.ORANGE};height:4px;"></td>
    <td width="25%" style="background:{C.GOLD};height:4px;"></td>
    <td width="25%" style="background:#00E676;height:4px;"></td>
    <td width="25%" style="background:#2979FF;height:4px;"></td>
</tr></table></td></tr>
<tr><td style="padding:0;border-bottom:1px solid #eee;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td width="25%" style="text-align:center;padding:20px 10px;border-right:1px solid #f0f0f0;">
        <p style="margin:0;font-size:32px;font-weight:bold;color:{C.NAVY};">{len(d['unique_sos'])}</p>
        <p style="margin:5px 0 0;font-size:10px;color:#999;text-transform:uppercase;letter-spacing:1px;">Sales Orders</p>
    </td>
    <td width="25%" style="text-align:center;padding:20px 10px;border-right:1px solid #f0f0f0;">
        <p style="margin:0;font-size:32px;font-weight:bold;color:{C.GREEN};">{format_indian(d['total_items'])}</p>
        <p style="margin:5px 0 0;font-size:10px;color:#999;text-transform:uppercase;letter-spacing:1px;">Line Items</p>
    </td>
    <td width="25%" style="text-align:center;padding:20px 10px;border-right:1px solid #f0f0f0;">
        <p style="margin:0;font-size:32px;font-weight:bold;color:{C.ORANGE};">{format_indian(d['total_order'])}</p>
        <p style="margin:5px 0 0;font-size:10px;color:#999;text-transform:uppercase;letter-spacing:1px;">Order Qty</p>
    </td>
    <td width="25%" style="text-align:center;padding:20px 10px;">
        <p style="margin:0;font-size:32px;font-weight:bold;color:{C.PURPLE};">{format_indian(d['total_tester'])}</p>
        <p style="margin:5px 0 0;font-size:10px;color:#999;text-transform:uppercase;letter-spacing:1px;">Tester Qty</p>
    </td>
</tr></table></td></tr>
<tr><td style="padding:12px 20px;background:#f8f9fa;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td width="33%" style="height:2px;background:{C.NAVY};font-size:0;">&nbsp;</td>
    <td width="34%" style="height:2px;background:{C.GREEN};font-size:0;">&nbsp;</td>
    <td width="33%" style="height:2px;background:{C.ORANGE};font-size:0;">&nbsp;</td>
</tr></table></td></tr>
<tr><td style="padding:14px 20px;font-weight:bold;font-size:14px;color:{C.NAVY};border-left:5px solid {C.NAVY};background:#E8EAF6;">📋 Sales Order Details</td></tr>
<tr><td style="padding:0;"><table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
<tr>
    <th style="background:{C.NAVY};color:white;padding:10px 8px;font-size:11px;text-transform:uppercase;">SO Number</th>
    <th style="background:{C.NAVY};color:white;padding:10px 8px;font-size:11px;text-transform:uppercase;">Distributor</th>
    <th style="background:{C.NAVY};color:white;padding:10px 8px;font-size:11px;text-transform:uppercase;">City</th>
    <th style="background:{C.NAVY};color:white;padding:10px 8px;font-size:11px;text-transform:uppercase;">State</th>
    <th style="background:{C.NAVY};color:white;padding:10px 8px;font-size:11px;text-transform:uppercase;">Location</th>
    <th style="background:{C.NAVY};color:white;padding:10px 8px;font-size:11px;text-transform:uppercase;">Order Qty</th>
    <th style="background:{C.NAVY};color:white;padding:10px 8px;font-size:11px;text-transform:uppercase;">Tester Qty</th>
    <th style="background:{C.NAVY};color:white;padding:10px 8px;font-size:11px;text-transform:uppercase;">Total</th>
</tr>"""

        for i, so_row in enumerate(d['unique_sos']):
            si = d['so_groups'].get(so_row.so_number, {'order':0,'tester':0})
            bg = '#f9f9f9' if i%2==1 else '#ffffff'
            html += f'''<tr style="background:{bg};">
    <td style="padding:9px 8px;text-align:center;font-size:12px;border-bottom:1px solid #eee;font-weight:bold;">{so_row.so_number}</td>
    <td style="padding:9px 8px;text-align:left;font-size:12px;border-bottom:1px solid #eee;">{so_row.distributor or "—"}</td>
    <td style="padding:9px 8px;text-align:center;font-size:12px;border-bottom:1px solid #eee;">{so_row.city or "—"}</td>
    <td style="padding:9px 8px;text-align:center;font-size:12px;border-bottom:1px solid #eee;">{so_row.state or "—"}</td>
    <td style="padding:9px 8px;text-align:center;font-size:12px;border-bottom:1px solid #eee;">{so_row.location_code or "—"}</td>
    <td style="padding:9px 8px;text-align:center;font-size:12px;border-bottom:1px solid #eee;">{format_indian(si['order'])}</td>
    <td style="padding:9px 8px;text-align:center;font-size:12px;border-bottom:1px solid #eee;">{format_indian(si['tester'])}</td>
    <td style="padding:9px 8px;text-align:center;font-size:12px;border-bottom:1px solid #eee;font-weight:bold;">{format_indian(si['order']+si['tester'])}</td>
</tr>'''

        html += f'''<tr style="background:#E8EAF6;font-weight:bold;">
    <td style="padding:10px 8px;text-align:center;font-size:12px;">TOTAL</td>
    <td colspan="4" style="padding:10px 8px;text-align:left;font-size:12px;">{len(d['unique_sos'])} Sales Orders</td>
    <td style="padding:10px 8px;text-align:center;font-size:12px;">{format_indian(d['total_order'])}</td>
    <td style="padding:10px 8px;text-align:center;font-size:12px;">{format_indian(d['total_tester'])}</td>
    <td style="padding:10px 8px;text-align:center;font-size:12px;">{format_indian(total_qty)}</td>
</tr></table></td></tr>
<tr><td style="padding:12px 20px;background:#f8f9fa;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td width="33%" style="height:2px;background:{C.NAVY};font-size:0;">&nbsp;</td>
    <td width="34%" style="height:2px;background:{C.GREEN};font-size:0;">&nbsp;</td>
    <td width="33%" style="height:2px;background:{C.ORANGE};font-size:0;">&nbsp;</td>
</tr></table></td></tr>
<tr><td style="padding:14px 20px;font-weight:bold;font-size:14px;color:{C.GREEN};border-left:5px solid {C.GREEN};background:#E8F5E9;">📦 SKU Demand Summary</td></tr>
<tr><td style="padding:0;"><table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
<tr>
    <th style="background:{C.GREEN};color:white;padding:10px 6px;font-size:11px;">#</th>
    <th style="background:{C.GREEN};color:white;padding:10px 6px;font-size:11px;">BC CODE</th>
    <th style="background:{C.GREEN};color:white;padding:10px 6px;font-size:11px;">DESCRIPTION</th>
    <th style="background:{C.GREEN};color:white;padding:10px 6px;font-size:11px;">CATEGORY</th>
    <th style="background:{C.GREEN};color:white;padding:10px 6px;font-size:11px;">ORDER</th>
    <th style="background:{C.GREEN};color:white;padding:10px 6px;font-size:11px;">TESTER</th>
    <th style="background:{C.GREEN};color:white;padding:10px 6px;font-size:11px;">TOTAL</th>
</tr>'''

        for rank, (item_no, info) in enumerate(d['sorted_skus'], 1):
            total = info['order']+info['tester']
            desc = info['desc'][:45]+'...' if len(info['desc'])>45 else info['desc']
            bg = '#f1f8e9' if rank%2==0 else '#ffffff'
            html += f'''<tr style="background:{bg};">
    <td style="padding:8px 6px;text-align:center;font-size:12px;color:#999;border-bottom:1px solid #eee;">{rank}</td>
    <td style="padding:8px 6px;text-align:center;font-size:12px;font-weight:bold;border-bottom:1px solid #eee;">{item_no}</td>
    <td style="padding:8px 6px;text-align:left;font-size:12px;border-bottom:1px solid #eee;">{desc or "—"}</td>
    <td style="padding:8px 6px;text-align:center;font-size:12px;border-bottom:1px solid #eee;">{info["cat"] or "—"}</td>
    <td style="padding:8px 6px;text-align:center;font-size:12px;border-bottom:1px solid #eee;">{format_indian(info["order"])}</td>
    <td style="padding:8px 6px;text-align:center;font-size:12px;border-bottom:1px solid #eee;">{format_indian(info["tester"])}</td>
    <td style="padding:8px 6px;text-align:center;font-size:12px;font-weight:bold;border-bottom:1px solid #eee;">{format_indian(total)}</td>
</tr>'''

        html += f'''<tr style="background:#E8F5E9;font-weight:bold;">
    <td style="padding:10px 6px;text-align:center;font-size:12px;"></td>
    <td style="padding:10px 6px;text-align:center;font-size:12px;">GRAND TOTAL</td>
    <td style="padding:10px 6px;text-align:left;font-size:12px;">{len(d['sorted_skus'])} unique SKUs</td>
    <td style="padding:10px 6px;text-align:center;font-size:12px;"></td>
    <td style="padding:10px 6px;text-align:center;font-size:12px;">{format_indian(d['total_order'])}</td>
    <td style="padding:10px 6px;text-align:center;font-size:12px;">{format_indian(d['total_tester'])}</td>
    <td style="padding:10px 6px;text-align:center;font-size:12px;">{format_indian(total_qty)}</td>
</tr></table></td></tr>
<tr><td style="padding:12px 20px;background:#f8f9fa;"><table width="100%" cellpadding="0" cellspacing="0"><tr>
    <td width="33%" style="height:2px;background:{C.NAVY};font-size:0;">&nbsp;</td>
    <td width="34%" style="height:2px;background:{C.GREEN};font-size:0;">&nbsp;</td>
    <td width="33%" style="height:2px;background:{C.ORANGE};font-size:0;">&nbsp;</td>
</tr></table></td></tr>
<tr><td style="background:{C.NAVY};padding:30px;text-align:center;">
    <p style="margin:0 0 5px;font-size:16px;font-weight:bold;color:{C.GOLD};letter-spacing:1px;">⚡ GT MASS DUMP GENERATOR v2.3</p>
    <p style="margin:0 0 18px;font-size:11px;color:#7986CB;">Warehouse Automation Suite — One-click PO Intelligence</p>
    <table style="margin:0 auto;max-width:400px;"><tr><td style="background:rgba(255,255,255,0.08);border:1px solid rgba(255,255,255,0.15);padding:18px;border-radius:10px;text-align:center;">
        <p style="margin:0 0 3px;font-size:10px;color:#7986CB;text-transform:uppercase;letter-spacing:2px;">🚀 Engineered by</p>
        <p style="margin:0 0 5px;font-size:18px;font-weight:bold;color:white;">Abhishek Wagh</p>
        <p style="margin:0 0 3px;font-size:11px;color:#9FA8DA;">Order Management and Automation</p>
        <p style="margin:0;font-size:10px;color:#7986CB;">📧 abhishek.wagh@reneecosmetics.in</p>
    </td></tr></table>
    <table style="margin:15px auto 0;max-width:450px;"><tr><td style="background:rgba(255,255,255,0.05);padding:12px 20px;border-radius:8px;border-left:3px solid {C.GOLD};text-align:left;">
        <p style="margin:0;font-size:11px;font-style:italic;color:#C5CAE9;">🏆 "Automation isn't just about saving time — it's about building systems that <span style="color:{C.GOLD};font-weight:bold;">sell while you sleep.</span>"</p>
    </td></tr></table>
    <p style="margin:18px 0 0;font-size:9px;color:#5C6BC0;">© 2026 RENEE Cosmetics Pvt. Ltd. | Warehouse Automation Division | Confidential</p>
</td></tr></table></td></tr></table></body></html>'''
        return html


class EmailSender:
    """Network I/O only — sends pre-built HTML via SMTP. FIX ❌3 (SRP)."""

    @staticmethod
    def send_report(result: 'ProcessResult', elapsed_str: str) -> Tuple[bool, str]:
        """FIX ❌7: Specific exception types for SMTP, auth, network."""
        config = EMAIL_CONFIG
        if not config['EMAIL_SENDER'] or not config['DEFAULT_RECIPIENT']:
            return False, "Email not configured."
        try:
            html = EmailBuilder.build_html(result, elapsed_str)
            subject = EmailBuilder.build_subject(result)
            msg = EmailMessage()
            msg['From'] = config['EMAIL_SENDER']
            msg['To'] = config['DEFAULT_RECIPIENT']
            if config['CC_RECIPIENTS']:
                msg['Cc'] = ', '.join(config['CC_RECIPIENTS'])
            msg['Subject'] = subject
            msg.set_content("Please view in HTML-compatible client.")
            msg.add_alternative(html, subtype='html')
            server = smtplib.SMTP(config['SMTP_SERVER'], config['SMTP_PORT'])
            server.starttls()
            server.login(config['EMAIL_SENDER'], config['EMAIL_PASSWORD'])
            server.send_message(msg, to_addrs=[config['DEFAULT_RECIPIENT']] + config['CC_RECIPIENTS'])
            server.quit()
            logging.info(f"Email sent to {config['DEFAULT_RECIPIENT']} + {len(config['CC_RECIPIENTS'])} CC")
            return True, ""
        except smtplib.SMTPAuthenticationError as e:
            logging.error(f"Email auth failed: {e}")
            return False, f"Auth failed — check EMAIL_PASSWORD. ({e})"
        except smtplib.SMTPException as e:
            logging.error(f"SMTP error: {e}")
            return False, f"SMTP error: {e}"
        except (ConnectionError, OSError) as e:
            logging.error(f"Network error: {e}")
            return False, f"Network error: {e}"
        except (ValueError, KeyError) as e:
            logging.error(f"Config error: {e}")
            return False, f"Config error: {e}"


# ═══════════════════════════════════════════════════════════════════════════════
#  DATA MODEL — FIX ❌10: ProcessingResult → ProcessResult
# ═══════════════════════════════════════════════════════════════════════════════

@dataclass
class OrderRow:
    """Single ordered item from a GT-Mass file."""
    so_number: str
    item_no: str
    ean: str
    category: str
    description: str
    qty: int
    tester_qty: int
    distributor: str
    city: str
    state: str
    location: str
    location_code: str
    source_file: str

@dataclass
class ProcessResult:
    """FIX ❌10: Renamed from ProcessingResult."""
    rows: List[OrderRow] = field(default_factory=list)
    failed_files: List[Tuple[str, str]] = field(default_factory=list)
    warned_files: List[Tuple[str, str]] = field(default_factory=list)
    output_path: Optional[Path] = None


# ═══════════════════════════════════════════════════════════════════════════════
#  FIX ❌10: SOFormatter → SONumberFormatter
# ═══════════════════════════════════════════════════════════════════════════════

class SONumberFormatter:
    """Extracts SO number from filename digits (fallback)."""
    @staticmethod
    def from_filename(filepath: Path) -> Optional[str]:
        match = re.search(r"\d+", filepath.stem)
        if not match:
            logging.warning(f"SO number not found in filename: {filepath.name}")
            return None
        return f"SO/GTM/{match.group()}"


# ═══════════════════════════════════════════════════════════════════════════════
#  FILE READER — FIX ❌7: Specific exceptions
# ═══════════════════════════════════════════════════════════════════════════════

class FileReader:
    """Reads Excel files into raw DataFrames (no header)."""
    @staticmethod
    def read(file_path: Path) -> pd.DataFrame:
        ext = file_path.suffix.lower()
        if ext in (".xlsx", ".xlsm"):
            try:
                df = pd.read_excel(file_path, header=None, engine="openpyxl")
                logging.info(f"{file_path.name} — openpyxl ({len(df)} rows)")
                return df
            except (ValueError, KeyError) as e:
                raise RuntimeError(f"Cannot read '{file_path.name}': {e}")
        if ext == ".xls":
            try:
                df = pd.read_excel(file_path, header=None, engine="xlrd")
                logging.info(f"{file_path.name} — xlrd ({len(df)} rows)")
                return df
            except ImportError:
                raise RuntimeError(f"Cannot read '{file_path.name}' — pip install xlrd")
            except (ValueError, KeyError) as e:
                raise RuntimeError(f"Cannot read '{file_path.name}': {e}")
        raise RuntimeError(f"Unsupported format: '{ext}'. Only .xlsx/.xlsm/.xls.")


# ═══════════════════════════════════════════════════════════════════════════════
#  FIX ❌10: MetaExtractor → MetadataExtractor
# ═══════════════════════════════════════════════════════════════════════════════

class MetadataExtractor:
    """Extracts meta fields from header rows above the data table."""
    @staticmethod
    def extract(raw_df: pd.DataFrame, header_row: int) -> Tuple[dict, List[str]]:
        distributor = city = location = so_number = ""
        state_values: List[str] = []
        warnings: List[str] = []
        meta_df = raw_df.iloc[:header_row]

        for _, row in meta_df.iterrows():
            label = str(row.iloc[0]).strip().lower() if pd.notna(row.iloc[0]) else ""
            value = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            if value.lower() in ("nan", ""):
                value = ""

            if label == "distributor name" and not distributor:
                distributor = value
            elif label == "city" and not city:
                city = value
            elif label == "state":
                state_values.append(value)

            for col_idx in range(min(len(row)-1, 10)):
                cell_val = str(row.iloc[col_idx]).strip().lower() if pd.notna(row.iloc[col_idx]) else ""
                if cell_val == "location":
                    for vi in range(col_idx+1, min(col_idx+3, len(row))):
                        lv = row.iloc[vi]
                        if pd.notna(lv) and str(lv).strip() and str(lv).strip().lower() != 'nan':
                            location = str(lv).strip(); break
                elif cell_val == "po number" and not so_number:
                    for vi in range(col_idx+1, min(col_idx+3, len(row))):
                        pv = row.iloc[vi]
                        if pd.notna(pv) and str(pv).strip() and str(pv).strip().lower() != 'nan':
                            so_number = str(pv).strip(); break

        state = next((s for s in reversed(state_values) if s), "")
        location_code = ""
        if location:
            lu = location.upper().strip()
            location_code = LOCATION_CODE_MAP.get(lu, location)

        if not distributor: warnings.append("Distributor Name is blank.")
        if not city: warnings.append("City is blank.")
        if not state: warnings.append("State is blank.")
        # ── CRITICAL: Location Code missing → ERP import will fail ──
        if not location_code:
            warnings.append(
                "❌ CRITICAL: Location Code is EMPTY — Location field is missing in source file. "
                "ERP import will fail without Location Code. Fix the source file immediately."
            )
        if distributor and distributor.strip().lower() in STATE_LIKE_VALUES:
            warnings.append(f"Distributor '{distributor}' looks like a state — verify.")

        return {"distributor": distributor, "city": city, "state": state,
                "location": location, "location_code": location_code, "so_number": so_number}, warnings


# ═══════════════════════════════════════════════════════════════════════════════
#  EXCEL PARSER — FIX ❌4: parse() split into focused methods
# ═══════════════════════════════════════════════════════════════════════════════

class ExcelParser:
    """Parses GT-Mass PO files. FIX ❌4: parse() delegates to sub-methods."""
    BC_COLUMN = "bc code"
    QTY_COLUMN = "order qty"
    TESTER_COLUMN = "tester qty"

    def parse(self, file_path: Path) -> Tuple[List[OrderRow], List[str]]:
        """Orchestrator: read → find header → extract meta → resolve SO → extract rows."""
        logging.info(f"Parsing: {file_path.name}")
        warnings: List[str] = []
        raw_df = FileReader.read(file_path)
        header_row = self._find_header_row(raw_df)
        meta, mw = MetadataExtractor.extract(raw_df, header_row)
        warnings.extend(mw)
        so_number, sw = self._resolve_so_number(meta, file_path)
        warnings.extend(sw)
        df = raw_df.iloc[header_row+1:].copy()
        df.columns = raw_df.iloc[header_row].values
        df = df.reset_index(drop=True)
        rows, ew = self._extract_rows(df, so_number, meta, file_path.name)
        warnings.extend(ew)
        return rows, warnings

    def _find_header_row(self, raw_df: pd.DataFrame) -> int:
        """FIX ❌4: Extracted — scan for row with 'BC Code' + 'Order Qty'."""
        for i, rv in enumerate(raw_df.values):
            vals = [str(v).lower() for v in rv]
            if "bc code" in vals and any("order qty" in v for v in vals):
                return i
        raise RuntimeError("Header row not found — no 'BC Code' + 'Order Qty'.")

    def _resolve_so_number(self, meta: dict, file_path: Path) -> Tuple[str, List[str]]:
        """FIX ❌4: Extracted — file PO Number → filename → UNKNOWN."""
        warnings: List[str] = []
        so = meta.get("so_number", "")
        if so:
            logging.info(f"SO from file: '{so}'")
            return so, warnings
        so = SONumberFormatter.from_filename(file_path)
        if so:
            warnings.append(f"SO from filename: '{so}'. Fill PO Number field.")
            return so, warnings
        warnings.append("SO not found — using 'SO/GTM/UNKNOWN'.")
        return "SO/GTM/UNKNOWN", warnings

    def _extract_rows(self, df: pd.DataFrame, so_number: str,
                      meta: dict, filename: str) -> Tuple[List[OrderRow], List[str]]:
        """FIX ❌4: Extracted. FIX ❌9: Uses safe_str_val()."""
        warnings: List[str] = []
        bc_col, qty_col, tester_col, ean_col, cat_col, desc_col = self._detect_columns(df)
        if bc_col is None: raise RuntimeError("'BC Code' column not found.")
        if qty_col is None: raise RuntimeError("'Order Qty' column not found.")
        if tester_col is None: warnings.append("'Tester Qty' not found — defaulting to 0.")

        rows: List[OrderRow] = []
        bc_idx = df.columns.get_loc(bc_col)
        qty_idx = df.columns.get_loc(qty_col)
        tester_idx = df.columns.get_loc(tester_col) if tester_col else None
        ean_idx = df.columns.get_loc(ean_col) if ean_col else None
        cat_idx = df.columns.get_loc(cat_col) if cat_col else None
        desc_idx = df.columns.get_loc(desc_col) if desc_col else None

        for rv in df.values:
            bc = rv[bc_idx]
            if pd.isna(bc): continue
            try: bc = int(bc)
            except (ValueError, TypeError): continue
            qty = self._clean_qty(rv[qty_idx])
            tqty = self._clean_qty(rv[tester_idx]) if tester_idx is not None else 0
            if qty <= 0 and tqty <= 0: continue

            rows.append(OrderRow(
                so_number=so_number, item_no=str(bc),
                ean=safe_str_val(rv, ean_idx, as_int_str=True),
                category=safe_str_val(rv, cat_idx),
                description=safe_str_val(rv, desc_idx),
                qty=qty, tester_qty=tqty,
                distributor=meta["distributor"], city=meta["city"], state=meta["state"],
                location=meta["location"], location_code=meta["location_code"],
                source_file=filename,
            ))
        if not rows: warnings.append("No ordered rows — all quantities are 0.")
        return rows, warnings

    def _detect_columns(self, df) -> Tuple[Optional[str], ...]:
        bc = qty = tester = ean = cat = desc = None
        for col in df.columns:
            n = str(col).strip().lower()
            if n == self.BC_COLUMN: bc = col
            if self.QTY_COLUMN in n: qty = col
            if self.TESTER_COLUMN in n: tester = col
            if n == 'ean' and not ean: ean = col
            if n == 'category' and not cat: cat = col
            if 'article description' in n or (n == 'description' and not desc): desc = col
        return bc, qty, tester, ean, cat, desc

    @staticmethod
    def _clean_qty(value) -> int:
        if pd.isna(value): return 0
        value = str(value).strip()
        if value in ("", "-"): return 0
        value = value.replace(",", "")
        try: return int(float(value))
        except (ValueError, TypeError): return 0


# ═══════════════════════════════════════════════════════════════════════════════
#  DUMP EXPORTER — all 7 sheets (unchanged output)
# ═══════════════════════════════════════════════════════════════════════════════

class DumpExporter:
    """Writes output Excel with 7 sheets."""
    HEADER_FILL = PatternFill('solid', fgColor='1A237E')
    HEADER_FONT = Font(bold=True, color='FFFFFF', name='Aptos Display', size=11)
    DATA_FONT = Font(name='Aptos Display', size=11)
    THIN_SIDE = Side(style='thin', color='CCCCCC')
    BORDER = Border(left=THIN_SIDE, right=THIN_SIDE, top=THIN_SIDE, bottom=THIN_SIDE)

    def _hdr_cell(self, ws, row, col, value):
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = self.HEADER_FONT; cell.fill = self.HEADER_FILL
        cell.alignment = Alignment(horizontal='center', vertical='center'); cell.border = self.BORDER
        return cell

    def _data_cell(self, ws, row, col, value, fmt=None):
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = self.DATA_FONT; cell.border = self.BORDER
        if fmt: cell.number_format = fmt
        return cell

    def _auto_width(self, ws, max_w=50):
        for col in ws.columns:
            letter = col[0].column_letter
            w = max((len(str(c.value or '')) for c in col), default=8)
            ws.column_dimensions[letter].width = min(w+3, max_w)

    def export(self, result: ProcessResult) -> Optional[Path]:
        if result.failed_files:
            msg = "Files skipped:\n\n"
            for f, r in result.failed_files: msg += f"  • {f}: {r}\n"
            messagebox.showerror("Files Failed", msg)
        if not result.rows:
            messagebox.showwarning("No Data", "No valid rows. Nothing to export.")
            return None

        Path("output").mkdir(exist_ok=True)
        ref_path = Path("output") / f"gt_mass_dump_{datetime.now().strftime('%d-%m-%Y_%H%M%S')}.xlsx"
        wb = Workbook(); wb.remove(wb.active)
        self._write_headers_so(wb, result); self._write_lines_so(wb, result)
        self._write_sales_lines(wb, result); self._write_sales_header(wb, result)
        self._write_sku_summary(wb, result); self._write_file_so_mapping(wb, result)
        self._write_warnings(wb, result)
        wb.save(str(ref_path))
        logging.info(f"Saved: {ref_path}")
        return ref_path

    def export_d365(self, result: ProcessResult, template_path: str) -> Optional[Path]:
        if not result.rows:
            messagebox.showwarning("No Data", "Generate dump first."); return None
        try:
            import shutil, zipfile
            import re as re_mod
            Path("output").mkdir(exist_ok=True)
            d365_path = Path("output") / f"d365_import_{datetime.now().strftime('%d-%m-%Y_%H%M%S')}.xlsx"
            shutil.copy2(template_path, str(d365_path))
            today_str = datetime.now().strftime("%d-%m-%Y")
            seen = set(); unique_sos = []
            for row in result.rows:
                if row.so_number not in seen: seen.add(row.so_number); unique_sos.append(row)
            zc = {}
            with zipfile.ZipFile(str(d365_path), 'r') as z:
                for item in z.namelist(): zc[item] = z.read(item)
            ss_xml = zc['xl/sharedStrings.xml'].decode('utf-8')
            existing = re_mod.findall(r'<t[^>]*>([^<]*)</t>', ss_xml)
            sm = {s: i for i, s in enumerate(existing)}
            ns = {'Order','Item','B2B',today_str}
            for r in unique_sos: ns.add(r.so_number); ns.add(r.location_code) if r.location_code else None
            for r in result.rows: ns.add(r.so_number); ns.add(r.location_code) if r.location_code else None
            ni = len(existing)
            for s in sorted(ns):
                if s not in sm: sm[s] = ni; ni += 1
            si = ['']*ni
            for s, idx in sm.items():
                esc = s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;')
                si[idx] = f'<si><t>{esc}</t></si>'
            zc['xl/sharedStrings.xml'] = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="{ni}" uniqueCount="{ni}">{"".join(si)}</sst>').encode('utf-8')
            def fc(xml, cl, rn, v, is_s=True):
                ref = f"{cl}{rn}"; pat = f'<c r="{ref}" s="(\\d+)"\\s*/>'
                rep = f'<c r="{ref}" s="\\1" t="s"><v>{sm.get(str(v),0)}</v></c>' if is_s else f'<c r="{ref}" s="\\1"><v>{v}</v></c>'
                return re_mod.sub(pat, rep, xml, count=1)

            def inject_empty_row(xml, row_num, columns, style_id, before_tag='</sheetData>'):
                """
                Inject a new empty <row> with pre-formatted <c> cells into sheet XML.

                ROOT CAUSE FIX: The D365 template has a fixed number of pre-formatted
                empty rows (e.g., 33 in Sales Header). When we have more SOs than
                template rows, fill_cell()'s regex finds no <c r="A37" .../> to replace
                and silently drops the data. This function injects the missing rows
                BEFORE we try to fill them, so every SO gets a slot.

                Args:
                    xml      : sheet XML string
                    row_num  : row number to inject (e.g., 37)
                    columns  : list of column letters (e.g., ['A','B','C',...])
                    style_id : style index to apply (copied from existing data rows)
                    before_tag: XML tag to inject before (default: closing </sheetData>)
                """
                cells = ''.join(f'<c r="{c}{row_num}" s="{style_id}"/>' for c in columns)
                new_row = f'<row r="{row_num}" spans="1:{len(columns)}" x14ac:dyDescent="0.3">{cells}</row>'
                return xml.replace(before_tag, new_row + before_tag)

            # ── Count template rows to detect overflow ──
            s1 = zc['xl/worksheets/sheet1.xml'].decode('utf-8')
            s1_template_rows = len(re_mod.findall(r'<row r="(\d+)"', s1)) - 2  # subtract header rows (1,3)
            s1_data_slots = s1_template_rows  # rows 4..N
            hdr_cols = ['A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R']

            # Inject extra rows in sheet1 if we have more SOs than template slots
            if len(unique_sos) > s1_data_slots:
                logging.info(f"D365: Template has {s1_data_slots} header rows, need {len(unique_sos)} — injecting extras")
                for extra_row in range(s1_data_slots + 4, len(unique_sos) + 4):
                    s1 = inject_empty_row(s1, extra_row, hdr_cols, '11')

            s2 = zc['xl/worksheets/sheet2.xml'].decode('utf-8')
            s2_template_rows = len(re_mod.findall(r'<row r="(\d+)"', s2)) - 3  # subtract header rows (1,2,3)
            line_cols = ['A','B','C','D','E','F','G','H']

            # Inject extra rows in sheet2 if we have more items than template slots
            if len(result.rows) > s2_template_rows:
                logging.info(f"D365: Template has {s2_template_rows} line rows, need {len(result.rows)} — injecting extras")
                for extra_row in range(s2_template_rows + 4, len(result.rows) + 4):
                    s2 = inject_empty_row(s2, extra_row, line_cols, '8')

            # Now fill cells as before — every SO/item has a row to fill into
            for i, row in enumerate(unique_sos):
                r = i+4; s1 = fc(s1,'A',r,'Order'); s1 = fc(s1,'B',r,row.so_number)
                for c in 'EFGHI': s1 = fc(s1,c,r,today_str)
                s1 = fc(s1,'J',r,row.so_number)
                if row.location_code: s1 = fc(s1,'K',r,row.location_code)
                s1 = fc(s1,'M',r,'B2B')
            zc['xl/worksheets/sheet1.xml'] = s1.encode('utf-8')
            # s2 already loaded and injected above
            cso = None; ln = 0
            for i, row in enumerate(result.rows):
                if row.so_number != cso: cso = row.so_number; ln = 0
                ln += 10000; r = i+4
                s2 = fc(s2,'A',r,'Order'); s2 = fc(s2,'B',r,row.so_number)
                s2 = fc(s2,'C',r,ln,is_s=False); s2 = fc(s2,'D',r,'Item')
                try: s2 = fc(s2,'E',r,int(row.item_no),is_s=False)
                except (ValueError,TypeError): s2 = fc(s2,'E',r,row.item_no)
                if row.location_code: s2 = fc(s2,'F',r,row.location_code)
                s2 = fc(s2,'G',r,row.qty,is_s=False)
            zc['xl/worksheets/sheet2.xml'] = s2.encode('utf-8')
            lhd = 3+len(unique_sos); lld = 3+len(result.rows)
            s1c = zc['xl/worksheets/sheet1.xml'].decode('utf-8')
            for r in range(lhd+1,37): s1c = re_mod.sub(rf'<row r="{r}"[^>]*>.*?</row>','',s1c,flags=re_mod.DOTALL)
            s1c = re_mod.sub(r'<dimension ref="[^"]*"/>',f'<dimension ref="A1:R{lhd}"/>',s1c)
            zc['xl/worksheets/sheet1.xml'] = s1c.encode('utf-8')
            s2c = zc['xl/worksheets/sheet2.xml'].decode('utf-8')
            for r in range(lld+1,500): s2c = re_mod.sub(rf'<row r="{r}"[^>]*>.*?</row>','',s2c,flags=re_mod.DOTALL)
            s2c = re_mod.sub(r'<dimension ref="[^"]*"/>',f'<dimension ref="A1:H{lld}"/>',s2c)
            zc['xl/worksheets/sheet2.xml'] = s2c.encode('utf-8')
            t1 = zc['xl/tables/table1.xml'].decode('utf-8')
            t1 = re_mod.sub(r'ref="A3:[A-Z]+\d+"',f'ref="A3:R{lhd}"',t1); zc['xl/tables/table1.xml'] = t1.encode('utf-8')
            t2 = zc['xl/tables/table2.xml'].decode('utf-8')
            t2 = re_mod.sub(r'ref="A3:[A-Z]+\d+"',f'ref="A3:H{lld}"',t2); zc['xl/tables/table2.xml'] = t2.encode('utf-8')
            with zipfile.ZipFile(str(d365_path),'w',zipfile.ZIP_DEFLATED) as zo:
                for n,d in zc.items(): zo.writestr(n,d)
            logging.info(f"D365 saved: {d365_path}"); return d365_path
        except (FileNotFoundError, PermissionError) as e:
            messagebox.showerror("D365 Error", f"File error: {e}"); return None
        except (KeyError, ValueError) as e:
            messagebox.showerror("D365 Error", f"Template error: {e}"); return None

    def _write_headers_so(self, wb, result: ProcessResult):
        ws = wb.create_sheet('Headers (SO)')
        headers = ['Document Type','No.','Sell-to Customer No.','Ship-to Code','Posting Date','Order Date','Document Date','Invoice From Date','Invoice To Date','External Document No.','Location Code','Dimension Set ID','Supply Type','Voucher Narration','Brand Code (Dimension)','Channel Code (Dimension)','Catagory (Dimension)','Geography Code (Dimension)']
        for c, h in enumerate(headers, 1): self._hdr_cell(ws, 1, c, h)
        today_str = datetime.now().strftime("%d-%m-%Y")
        seen = set(); unique_sos = []
        for row in result.rows:
            if row.so_number not in seen: seen.add(row.so_number); unique_sos.append(row)
        r = 2
        for row in unique_sos:
            self._data_cell(ws,r,1,'Order'); self._data_cell(ws,r,2,row.so_number)
            self._data_cell(ws,r,3,''); self._data_cell(ws,r,4,'')
            for c in range(5,10): self._data_cell(ws,r,c,today_str)
            self._data_cell(ws,r,10,row.so_number); self._data_cell(ws,r,11,row.location_code)
            self._data_cell(ws,r,12,''); self._data_cell(ws,r,13,'B2B'); r += 1
        self._auto_width(ws)

    def _write_lines_so(self, wb, result: ProcessResult):
        ws = wb.create_sheet('Lines (SO)')
        for c, h in enumerate(['Document Type','Document No.','Line No.','Type','No.','Location Code','Quantity','Unit Price'], 1):
            self._hdr_cell(ws, 1, c, h)
        r = 2; cso = None; ln = 0
        for row in result.rows:
            if row.so_number != cso: cso = row.so_number; ln = 0
            ln += 10000
            self._data_cell(ws,r,1,'Order'); self._data_cell(ws,r,2,row.so_number)
            self._data_cell(ws,r,3,ln); self._data_cell(ws,r,4,'Item')
            self._data_cell(ws,r,5,row.item_no); self._data_cell(ws,r,6,row.location_code)
            self._data_cell(ws,r,7,row.qty); self._data_cell(ws,r,8,''); r += 1
        self._auto_width(ws)

    def _write_sales_lines(self, wb, result: ProcessResult):
        ws = wb.create_sheet('Sales Lines')
        for c, h in enumerate(['SO Number','EAN','BC Code','Category','Article Description','Order Qty','Tester Qty'], 1):
            self._hdr_cell(ws, 1, c, h)
        for r, row in enumerate(result.rows, 2):
            self._data_cell(ws,r,1,row.so_number); self._data_cell(ws,r,2,row.ean)
            self._data_cell(ws,r,3,row.item_no); self._data_cell(ws,r,4,row.category)
            self._data_cell(ws,r,5,row.description); self._data_cell(ws,r,6,row.qty)
            self._data_cell(ws,r,7,row.tester_qty)
        self._auto_width(ws)

    def _write_sales_header(self, wb, result: ProcessResult):
        ws = wb.create_sheet('Sales Header')
        for c, h in enumerate(['SO Number','Order Qty','Tester Qty','Total Qty','Distributor','City','State','Location'], 1):
            self._hdr_cell(ws, 1, c, h)
        sg: Dict[str, dict] = {}
        for row in result.rows:
            if row.so_number not in sg:
                sg[row.so_number] = {'oq':0,'tq':0,'d':row.distributor,'c':row.city,'s':row.state,'l':row.location}
            sg[row.so_number]['oq'] += row.qty; sg[row.so_number]['tq'] += row.tester_qty
        r = 2
        for sn, i in sg.items():
            self._data_cell(ws,r,1,sn); self._data_cell(ws,r,2,i['oq']); self._data_cell(ws,r,3,i['tq'])
            self._data_cell(ws,r,4,i['oq']+i['tq']); self._data_cell(ws,r,5,i['d'])
            self._data_cell(ws,r,6,i['c']); self._data_cell(ws,r,7,i['s']); self._data_cell(ws,r,8,i['l']); r += 1
        self._auto_width(ws)

    def _write_sku_summary(self, wb, result: ProcessResult):
        ws = wb.create_sheet('SKU Summary')
        for c, h in enumerate(['BC Code','Description','Category','Order Qty','Tester Qty','Total Qty'], 1):
            self._hdr_cell(ws, 1, c, h)
        sg: Dict[str, dict] = {}
        for row in result.rows:
            if row.item_no not in sg:
                sg[row.item_no] = {'d':row.description,'c':row.category,'oq':0,'tq':0}
            sg[row.item_no]['oq'] += row.qty; sg[row.item_no]['tq'] += row.tester_qty
            if not sg[row.item_no]['d'] and row.description: sg[row.item_no]['d'] = row.description
            if not sg[row.item_no]['c'] and row.category: sg[row.item_no]['c'] = row.category
        ss = sorted(sg.items(), key=lambda x: x[1]['oq']+x[1]['tq'], reverse=True)
        r = 2; go = gt = 0
        for item, i in ss:
            t = i['oq']+i['tq']; go += i['oq']; gt += i['tq']
            self._data_cell(ws,r,1,item); self._data_cell(ws,r,2,i['d']); self._data_cell(ws,r,3,i['c'])
            self._data_cell(ws,r,4,i['oq']); self._data_cell(ws,r,5,i['tq']); self._data_cell(ws,r,6,t); r += 1
        bf = Font(name='Aptos Display', size=11, bold=True)
        ws.cell(row=r,column=1,value='GRAND TOTAL').font = bf
        ws.cell(row=r,column=2,value=f'{len(ss)} unique SKUs').font = bf
        ws.cell(row=r,column=4,value=go).font = bf; ws.cell(row=r,column=5,value=gt).font = bf
        ws.cell(row=r,column=6,value=go+gt).font = bf
        for c in range(1,7): ws.cell(row=r,column=c).border = self.BORDER
        self._auto_width(ws)

    def _write_file_so_mapping(self, wb, result: ProcessResult):
        """Sheet 6: File → SO Mapping — traceability for UNKNOWN/wrong SOs."""
        ws = wb.create_sheet('File → SO Mapping')
        for c, h in enumerate(['Sr No','Filename','SO Number'], 1): self._hdr_cell(ws, 1, c, h)
        seen = set(); mappings = []
        for row in result.rows:
            key = (row.source_file, row.so_number)
            if key not in seen: seen.add(key); mappings.append(key)
        for i, (fn, so) in enumerate(mappings, 1):
            r = i+1; self._data_cell(ws,r,1,i); self._data_cell(ws,r,2,fn); self._data_cell(ws,r,3,so)
        sr = len(mappings)+2; bf = Font(name='Aptos Display', size=11, bold=True)
        ws.cell(row=sr,column=1,value='TOTAL').font = bf
        ws.cell(row=sr,column=2,value=f'{len(mappings)} file(s)').font = bf
        ws.cell(row=sr,column=3,value=f'{len(set(s for _,s in mappings))} unique SO(s)').font = bf
        for c in range(1,4): ws.cell(row=sr,column=c).border = self.BORDER
        self._auto_width(ws)

    def _write_warnings(self, wb, result: ProcessResult):
        if not result.warned_files: return
        ws = wb.create_sheet('Warnings')
        for c, h in enumerate(['File','Warning'], 1): self._hdr_cell(ws, 1, c, h)
        # Red fill for critical warnings (Location Code missing)
        red_fill = PatternFill('solid', fgColor='FFCDD2')
        red_font = Font(name='Aptos Display', size=11, bold=True, color='D32F2F')
        for r, (f, w) in enumerate(result.warned_files, 2):
            is_critical = '❌ CRITICAL' in w
            self._data_cell(ws, r, 1, f)
            self._data_cell(ws, r, 2, w)
            if is_critical:
                # Red background + bold red text for critical warnings
                ws.cell(row=r, column=1).fill = red_fill
                ws.cell(row=r, column=1).font = red_font
                ws.cell(row=r, column=2).fill = red_fill
                ws.cell(row=r, column=2).font = red_font
        self._auto_width(ws)


# ═══════════════════════════════════════════════════════════════════════════════
#  FILE OPENER — FIX ❌7: Specific exceptions
# ═══════════════════════════════════════════════════════════════════════════════

def open_file(file_path: Path):
    try:
        s = platform.system()
        if s == "Windows": os.startfile(str(file_path))
        elif s == "Darwin":
            import subprocess as sp; sp.Popen(["open", str(file_path)])
        else:
            import subprocess as sp; sp.Popen(["xdg-open", str(file_path)])
    except (FileNotFoundError, OSError) as e:
        messagebox.showerror("Open File Error", f"Could not open:\n{e}")


# ═══════════════════════════════════════════════════════════════════════════════
#  ENGINE — FIX ❌7: Layered exception handling
# ═══════════════════════════════════════════════════════════════════════════════

class GTMassAutomation:
    def __init__(self):
        self.parser = ExcelParser()
        self.exporter = DumpExporter()

    def process_files(self, files: List[Path]) -> ProcessResult:
        result = ProcessResult()
        for file in files:
            fname = file.name
            try:
                rows, warnings = self.parser.parse(file)
                result.rows.extend(rows)
                for w in warnings: result.warned_files.append((fname, w)); logging.warning(f"{fname}: {w}")
            except RuntimeError as e:
                result.failed_files.append((fname, str(e))); logging.error(f"{fname} FAILED: {e}")
            except (ValueError, KeyError, TypeError) as e:
                result.failed_files.append((fname, f"Data error: {e}")); logging.error(f"{fname} DATA: {e}")
            except OSError as e:
                result.failed_files.append((fname, f"File error: {e}")); logging.error(f"{fname} FILE: {e}")
            except Exception as e:
                # Safety net — one bad file must not crash the batch
                result.failed_files.append((fname, f"Unexpected: {e}")); logging.error(f"{fname} UNEXPECTED: {e}")
        logging.info(f"Done — {len(result.rows)} rows | {len(set(r.so_number for r in result.rows))} SOs | {len(result.failed_files)} failed | {len(result.warned_files)} warnings")
        return result


# ═══════════════════════════════════════════════════════════════════════════════
#  UI
# ═══════════════════════════════════════════════════════════════════════════════

class AutomationUI:
    def __init__(self, automation: GTMassAutomation):
        self.automation = automation
        self.files: List[Path] = []
        self.last_output_path: Optional[Path] = None
        self.last_result: Optional[ProcessResult] = None
        self.last_elapsed: str = ""

        self.root = tk.Tk()
        self.root.title("GT Mass Dump Generator v2.3")
        self.root.geometry("460x520"); self.root.resizable(False, False)
        tk.Label(self.root, text="GT Mass Dump Generator", font=("Arial",14,"bold")).pack(pady=10)
        tk.Label(self.root, text="GT-Mass / Everyday PO Files → ERP Import", font=("Arial",9), fg="gray").pack(pady=0)
        self.label = tk.Label(self.root, text="Selected Files: 0", font=("Arial",10)); self.label.pack(pady=6)
        tk.Button(self.root, text="Select Excel Files", width=22, command=self.select_files).pack(pady=6)
        tk.Button(self.root, text="Generate Dump", width=22, command=self.generate_dump).pack(pady=6)
        self.open_button = tk.Button(self.root, text="Open Last Output File", width=22, state=tk.DISABLED, command=self.open_last_file)
        self.open_button.pack(pady=6)
        tk.Button(self.root, text="📋 Download PO Template", width=22, command=self._download_template).pack(pady=6)
        tk.Button(self.root, text="📤 Export D365 Package", width=22, command=self._export_d365).pack(pady=6)
        tk.Button(self.root, text="📧 Send Email Report", width=22, command=self._send_email).pack(pady=6)
        self.status = tk.Label(self.root, text="Status: Waiting", font=("Arial",10), fg="gray"); self.status.pack(pady=6)
        self.time_label = tk.Label(self.root, text="", font=("Arial",9), fg="darkgreen"); self.time_label.pack(pady=2)

    def select_files(self):
        files = filedialog.askopenfilenames(title="Select Sales Order Files", filetypes=[("Excel Files","*.xlsx *.xls *.xlsm")])
        self.files = [Path(f) for f in files]
        self.label.config(text=f"Selected Files: {len(self.files)}")
        self.time_label.config(text=""); self.status.config(text="Status: Files selected", fg="gray")

    def generate_dump(self):
        if not self.files: messagebox.showwarning("Warning","Select files first."); return
        start = time.time()
        self.status.config(text="Processing...", fg="blue"); self.time_label.config(text=""); self.root.update()
        result = self.automation.process_files(self.files); self.last_result = result
        output_path = self.automation.exporter.export(result)
        elapsed_str = f"{time.time()-start:.2f} seconds"; self.last_elapsed = elapsed_str
        failed, warned = len(result.failed_files), len(result.warned_files)
        rows = len(result.rows); sos = len(set(r.so_number for r in result.rows)) if result.rows else 0
        if output_path:
            self.last_output_path = output_path; self.open_button.config(state=tk.NORMAL)
            self.status.config(text=f"Done — {rows} rows | {failed} failed | {warned} warn" if (failed or warned) else f"Done — {rows} rows, {sos} SOs", fg="orange" if (failed or warned) else "darkgreen")
            self.time_label.config(text=f"⏱ {elapsed_str}")
            wn = f"\n⚠️ {warned} warning(s)" if warned else ""
            fn = f"\n❌ {failed} failed" if failed else ""
            # Count SOs with missing Location Code — critical for ERP
            missing_loc = len(set(r.so_number for r in result.rows if not r.location_code))
            loc_note = f"\n🔴 {missing_loc} SO(s) have EMPTY Location Code — fix source files!" if missing_loc else ""
            if messagebox.askyesno("Done", f"File: {output_path.name}\nRows: {rows} | SOs: {sos}\nTime: {elapsed_str}{wn}{fn}{loc_note}\n\n📋 Check 'File → SO Mapping' for traceability.\n\nOpen file?"):
                open_file(output_path)
        else:
            self.status.config(text="No data", fg="red"); self.time_label.config(text=f"⏱ {elapsed_str}")

    def open_last_file(self):
        if self.last_output_path and self.last_output_path.exists(): open_file(self.last_output_path)
        else: messagebox.showwarning("Not Found","File gone. Generate new dump.")

    def _export_d365(self):
        if not self.last_result or not self.last_result.rows:
            messagebox.showwarning("No Data","Generate dump first."); return
        # Warn about missing Location Codes before exporting
        missing_loc = [r.so_number for r in self.last_result.rows if not r.location_code]
        missing_loc_sos = sorted(set(missing_loc))
        if missing_loc_sos:
            proceed = messagebox.askyesno(
                "⚠️ Missing Location Codes",
                f"The following {len(missing_loc_sos)} SO(s) have EMPTY Location Code:\n\n"
                + "\n".join(f"  • {s}" for s in missing_loc_sos) +
                f"\n\nD365 import may fail for these SOs.\n\n"
                f"Continue with export anyway?"
            )
            if not proceed:
                return
        tp = filedialog.askopenfilename(title="Select D365 Template", filetypes=[("Excel","*.xlsx")])
        if not tp: return
        d365 = self.automation.exporter.export_d365(self.last_result, tp)
        if d365:
            sos = len(set(r.so_number for r in self.last_result.rows))
            items = len(self.last_result.rows)
            answer = messagebox.askyesno(
                "D365 Package Exported",
                f"D365 import file created successfully!\n\n"
                f"File   : {d365.name}\n"
                f"SOs    : {sos}\n"
                f"Items  : {items}\n\n"
                f"Do you want to open the exported file?"
            )
            if answer:
                open_file(d365)

    def _send_email(self):
        if not self.last_result or not self.last_result.rows:
            messagebox.showwarning("No Data","Generate dump first."); return
        self.status.config(text="Sending email...", fg="blue"); self.root.update()
        ok, err = EmailSender.send_report(self.last_result, self.last_elapsed)
        if ok:
            self.status.config(text="Email sent ✓", fg="darkgreen")
            cc_list = ', '.join(EMAIL_CONFIG['CC_RECIPIENTS']) or 'none'
            messagebox.showinfo("Email Sent",
                f"Report sent successfully!\n\n"
                f"To : {EMAIL_CONFIG['DEFAULT_RECIPIENT']}\n"
                f"CC : {cc_list}")
        else:
            self.status.config(text="Email failed ✗", fg="red")
            messagebox.showerror("Failed", f"{err}")

    def _download_template(self):
        sp = filedialog.asksaveasfilename(title="Save Template", defaultextension=".xlsx", initialfile="GT-Mass_PO_Template.xlsx", filetypes=[("Excel","*.xlsx")])
        if not sp: return
        try:
            wb = Workbook(); ws = wb.active; ws.title = 'PO Template'
            tf = Font(name='Aptos Display',size=14,bold=True,color='1A237E')
            lf = Font(name='Aptos Display',size=11,bold=True)
            vf = Font(name='Aptos Display',size=11,color='0000CC')
            nf = Font(name='Aptos Display',size=10,italic=True,color='FF6600')
            hfl = PatternFill('solid',fgColor='1A237E'); hfn = Font(bold=True,color='FFFFFF',name='Aptos Display',size=11)
            sf = Font(name='Aptos Display',size=11,color='888888',italic=True)
            mf = PatternFill('solid',fgColor='E3F2FD')
            cf = PatternFill('solid',fgColor='FFCDD2'); cfn = Font(name='Aptos Display',size=11,bold=True,color='D32F2F')
            ws.cell(row=1,column=1,value='Purchase Order GT-Mass (Template)').font = tf
            for r,pairs in [(2,[('Distributor Name',lf,mf),('<Enter Distributor Name>',vf,None)]),
                            (3,[('DB Code',lf,mf),('<DB Code>',vf,None)]),
                            (5,[('City',lf,mf),('<City>',vf,None)]),
                            (6,[('State',lf,mf),('<State>',vf,None)])]:
                for ci,(v,fn,fl) in enumerate(pairs,1):
                    c = ws.cell(row=r,column=ci,value=v); c.font = fn
                    if fl: c.fill = fl
            ws.cell(row=2,column=7,value='ASM').font = lf; ws.cell(row=2,column=7).fill = mf
            ws.cell(row=2,column=9,value='<ASM Name>').font = vf
            ws.cell(row=3,column=7,value='RSM').font = lf; ws.cell(row=3,column=7).fill = mf
            ws.cell(row=3,column=9,value='<RSM Name>').font = vf
            ws.cell(row=4,column=1,value='BDE Name').font = lf; ws.cell(row=4,column=1).fill = mf
            ws.cell(row=4,column=2,value='<BDE Name>').font = vf
            ws.cell(row=4,column=7,value='PO Number').font = cfn; ws.cell(row=4,column=7).fill = cf
            ws.cell(row=4,column=9,value='SO/GTM/0000').font = cfn; ws.cell(row=4,column=9).fill = cf
            ws.cell(row=5,column=7,value='Date of PO').font = lf; ws.cell(row=5,column=7).fill = mf
            ws.cell(row=5,column=9,value='DD.MM.YYYY').font = vf
            ws.cell(row=6,column=7,value='Location').font = cfn; ws.cell(row=6,column=7).fill = cf
            ws.cell(row=6,column=9,value='AHD').font = cfn; ws.cell(row=6,column=9).fill = cf
            dh = ['EAN','BC Code','Category','Article Description ','Nail Paint Shade Number ','Product Classification','HSN Code\n8 Digit','MRP','Retiler Margin','Trade & Display Scheme','Ullage','QPS','Qty In Case','Rate @ RLP','Amount @ RLP','Order Qty','Order Amount','Tester Qty']
            cc = {'EAN','BC Code','Order Qty','Tester Qty'}
            crf = PatternFill('solid',fgColor='D32F2F'); crfn = Font(bold=True,color='FFFFFF',name='Aptos Display',size=11)
            for ci,h in enumerate(dh,1):
                c = ws.cell(row=7,column=ci,value=h); c.alignment = Alignment(horizontal='center',vertical='center',wrap_text=True)
                if h.strip() in cc: c.font = crfn; c.fill = crf
                else: c.font = hfn; c.fill = hfl
            sd = [8904473104307,201238,'Eye','RENEE PURE BROWN KAJAL PEN WITH SHARPENER, 0.35GM','-','Cosmetics',33049990,199,1.2,'16.67% on RLP','1.66 % on RLP','4.81% on RLP','','','',72,'',6]
            for ci,v in enumerate(sd,1): ws.cell(row=8,column=ci,value=v).font = sf
            ws.cell(row=10,column=1,value='⚠ INSTRUCTIONS:').font = Font(name='Aptos Display',size=11,bold=True,color='D32F2F')
            for i,ins in enumerate(['1. Fill PO Number (Row 4, Col I) SO/GTM/####','2. Fill Location (Row 6, Col I) AHD/BLR','3. Fill Distributor, City, State','4. Data from Row 8, delete sample','5. BC Code numeric, Qty numeric','6. RED = critical fields','7. Save .xlsx → load into generator']):
                ws.cell(row=11+i,column=1,value=ins).font = nf
            for cl,w in {'A':16,'B':12,'C':12,'D':50,'E':12,'F':18,'G':14,'H':8,'I':14,'J':20,'K':16,'L':14,'M':12,'N':12,'O':14,'P':12,'Q':14,'R':12}.items():
                ws.column_dimensions[cl].width = w
            ws.freeze_panes = 'A8'; wb.save(sp)
            messagebox.showinfo("Saved", f"Template: {sp}")
        except (PermissionError, OSError) as e:
            messagebox.showerror("Error", f"Save failed: {e}")

    def run(self):
        self.root.mainloop()


def main():
    check_expiry()
    automation = GTMassAutomation()
    ui = AutomationUI(automation)
    ui.run()

if __name__ == "__main__":
    main()