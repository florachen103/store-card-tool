"""
Native frontend entry for the work-card tool.
Run with: python app.py
"""

from __future__ import annotations

import base64
import gc
import hmac
import io
import math
import os
import shutil
import subprocess
import tempfile
import time
import uuid
import zipfile
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pandas as pd
from flask import Flask, Response, jsonify, redirect, render_template_string, request, send_file, session, url_for
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from sqlalchemy import func, select
from werkzeug.exceptions import RequestEntityTooLarge

from analytics import analytics_events, get_engine, get_session_id, init_analytics, track_event
import browser_pdf_export as browser_export
from excel_image_extractor import extract_images_from_excel


app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or os.urandom(32)
init_analytics()

MAX_UPLOAD_BYTES = int(os.environ.get("STORE_CARD_MAX_UPLOAD_MB", "20")) * 1024 * 1024
MAX_CARDS_PER_UPLOAD = int(os.environ.get("STORE_CARD_MAX_CARDS", "50"))
ARTIFACT_TTL_SECONDS = int(os.environ.get("STORE_CARD_ARTIFACT_TTL_SECONDS", "1800"))
MAX_ARTIFACTS = int(os.environ.get("STORE_CARD_MAX_ARTIFACTS", "3"))

app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES

ARTIFACTS: dict[str, dict] = {}
TMP_PREVIEW_PATH = Path(tempfile.gettempdir()) / "store-card-latest-preview.html"
LOGO_B64 = base64.b64encode((Path(__file__).resolve().parent / "assets" / "logo_live.svg").read_bytes()).decode()
STAR_SRC = browser_export._load_star_b64()
CENTER_BG_B64 = base64.b64encode((Path(__file__).resolve().parent / "assets" / "card_center_bg.png").read_bytes()).decode()
SLOGAN_B64 = base64.b64encode((Path(__file__).resolve().parent / "assets" / "huyou_da_jiankang_calligraphy_vector.svg").read_bytes()).decode()


@app.errorhandler(RequestEntityTooLarge)
def handle_upload_too_large(_exc):
    max_mb = MAX_UPLOAD_BYTES // (1024 * 1024)
    return jsonify({"ok": False, "error": f"Excel 文件不能超过 {max_mb}MB"}), 413


INDEX_HTML = """
<!DOCTYPE html>
<html lang="zh-CN">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>高济药房 · 员工工牌制作系统</title>
  <style>
    * { box-sizing: border-box; }
    body {
      margin: 0;
      font-family: "Noto Sans SC", "PingFang SC", "Microsoft YaHei", sans-serif;
      background: #eff1f5;
      color: #1a1a1a;
    }
    .site-header {
      background: #fff;
      border-bottom: 1px solid #ececec;
    }
    .site-header-inner {
      max-width: 1160px;
      margin: 0 auto;
      padding: 16px 20px;
      display: flex;
      align-items: center;
      gap: 14px;
    }
    .brand-badge {
      width: 44px;
      height: 44px;
      border-radius: 8px;
      background: #e60036;
      color: #fff;
      display: flex;
      align-items: center;
      justify-content: center;
      font-size: 26px;
      font-weight: 700;
    }
    .brand-title {
      font-size: 18px;
      font-weight: 700;
      color: #1a1a1a;
      margin: 0;
    }
    .brand-subtitle {
      margin: 2px 0 0;
      color: #9ca3af;
      font-size: 12px;
      font-weight: 500;
    }
    .page {
      max-width: 1160px;
      margin: 0 auto;
      padding: 20px;
    }
    .step-grid {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 16px;
      margin-bottom: 20px;
    }
    .card {
      background: #fff;
      border-radius: 12px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    }
    .step-card {
      padding: 16px;
      display: flex;
      gap: 12px;
      align-items: flex-start;
    }
    .step-index {
      width: 26px;
      height: 26px;
      background: #fff0f3;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      font-weight: 700;
      color: #E60036;
      font-size: 13px;
      flex-shrink: 0;
    }
    .step-title {
      margin: 0;
      font-weight: 600;
      font-size: 13px;
      color: #1a1a1a;
    }
    .step-desc {
      margin: 4px 0 0;
      font-size: 11px;
      color: #999;
      line-height: 1.55;
    }
    .drop-card {
      margin-bottom: 20px;
      overflow: hidden;
      padding: 0;
    }
    .drop-zone {
      margin: 0;
      border: 2px dashed #E60036;
      border-radius: 16px;
      cursor: pointer;
      transition: background 0.2s, border-color 0.2s;
      background: #fff;
      padding: 48px 24px;
      text-align: center;
    }
    .drop-zone:hover,
    .drop-zone.drag-active {
      background: #fff5f5;
      border-color: #a00020;
    }
    .drop-icon-wrap {
      width: 60px;
      height: 60px;
      background: #fff0f3;
      border-radius: 50%;
      display: flex;
      align-items: center;
      justify-content: center;
      margin: 0 auto 14px;
    }
    .drop-title {
      font-size: 17px;
      font-weight: 600;
      color: #1a1a1a;
      margin: 0;
    }
    .drop-subtitle {
      font-size: 13px;
      color: #999;
      margin: 6px 0 0;
      line-height: 1.6;
    }
    .btn-outline,
    .btn-red {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      border-radius: 999px;
      font-size: 13px;
      font-weight: 500;
      padding: 8px 18px;
      cursor: pointer;
      text-decoration: none;
      border: 1.5px solid #E60036;
    }
    .btn-outline {
      background: #fff;
      color: #E60036;
    }
    .btn-red {
      background: #E60036;
      color: #fff;
    }
    .progress-wrap,
    .error-wrap {
      display: none;
      padding: 16px 28px 20px;
    }
    .progress-row {
      display: flex;
      justify-content: space-between;
      margin-bottom: 8px;
      font-size: 13px;
    }
    .progress-track {
      background: #f0f0f0;
      border-radius: 4px;
      height: 6px;
    }
    .progress-bar {
      background: #E60036;
      height: 6px;
      border-radius: 4px;
      width: 0;
      transition: width 0.2s;
    }
    .progress-detail {
      font-size: 11px;
      color: #aaa;
      margin-top: 6px;
    }
    .error-box {
      background: #fff0f3;
      border: 1px solid #fcc;
      border-radius: 10px;
      padding: 14px 16px;
      display: flex;
      gap: 10px;
    }
    .preview-card {
      display: none;
    }
    .preview-toolbar {
      padding: 14px 18px;
      margin-bottom: 14px;
      display: flex;
      align-items: center;
      justify-content: space-between;
      flex-wrap: wrap;
      gap: 12px;
    }
    .preview-title {
      font-weight: 700;
      font-size: 15px;
      color: #1a1a1a;
    }
    .preview-meta {
      font-size: 13px;
      color: #999;
    }
    .template-controls {
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
      align-items: center;
    }
    .template-control {
      display: inline-flex;
      align-items: center;
      gap: 6px;
      color: #555;
      font-size: 12px;
      font-weight: 600;
    }
    .choice-control {
      position: relative;
    }
    .choice-trigger {
      height: 32px;
      border: 1px solid #ddd;
      border-radius: 8px;
      background: #fff;
      color: #1a1a1a;
      font-size: 13px;
      font-weight: 600;
      padding: 0 30px 0 10px;
      cursor: pointer;
      min-width: 76px;
      text-align: left;
      position: relative;
    }
    .choice-trigger::after {
      content: '';
      position: absolute;
      right: 10px;
      top: 50%;
      width: 7px;
      height: 7px;
      border-right: 1.5px solid #555;
      border-bottom: 1.5px solid #555;
      transform: translateY(-65%) rotate(45deg);
    }
    .choice-menu {
      display: none;
      position: absolute;
      top: calc(100% + 6px);
      left: 0;
      z-index: 1000;
      min-width: 100%;
      padding: 6px;
      background: #fff;
      border: 1px solid #e5e7eb;
      border-radius: 12px;
      box-shadow: 0 12px 28px rgba(15, 23, 42, 0.16);
    }
    .choice-control.open .choice-menu {
      display: grid;
      gap: 3px;
    }
    .choice-option {
      border: 0;
      background: transparent;
      color: #1a1a1a;
      border-radius: 8px;
      height: 30px;
      padding: 0 28px 0 10px;
      text-align: left;
      font-size: 13px;
      font-weight: 600;
      white-space: nowrap;
      cursor: pointer;
      position: relative;
    }
    .choice-option:hover {
      background: #fff0f3;
      color: #E60036;
    }
    .choice-option.is-selected {
      background: #E60036;
      color: #fff;
    }
    .choice-option.is-selected::before {
      content: '';
      position: absolute;
      left: 10px;
      top: 50%;
      width: 6px;
      height: 11px;
      border-right: 2px solid #fff;
      border-bottom: 2px solid #fff;
      transform: translateY(-62%) rotate(45deg);
    }
    .choice-option.is-selected {
      padding-left: 28px;
    }
    .print-tip {
      background: #fffbeb;
      border: 1px solid #fde68a;
      border-radius: 10px;
      padding: 10px 14px;
      margin-bottom: 14px;
      font-size: 12px;
      color: #92400e;
      line-height: 1.6;
    }
    .preview-shell {
      overflow: hidden;
      padding: 18px;
      background: #eef0f4;
    }
    .badge-preview-grid {
      display: flex;
      flex-direction: column;
      gap: 18px;
      align-items: center;
      justify-content: center;
      overflow-x: auto;
      padding-bottom: 4px;
    }
    .preview-a4-wrap {
      width: max-content;
    }
    .preview-a4-label {
      margin: 0 0 8px;
      color: #6b7280;
      font-size: 12px;
      font-weight: 600;
    }
    .preview-a4-page {
      width: 210mm;
      height: 297mm;
      padding: 5mm 7mm;
      background: #fff;
      box-shadow: 0 10px 28px rgba(15, 23, 42, 0.12);
      box-sizing: border-box;
    }
    .preview-a4-grid {
      display: grid;
      grid-template-columns: 86mm 86mm;
      gap: 2mm 4mm;
      width: 176mm;
      margin: 0 auto;
    }
    .preview-a4-page .badge-card {
      box-shadow: none;
      border: 0.3mm solid #e0e0e0;
    }
    .actions {
      margin-top: 14px;
      display: flex;
      gap: 10px;
      flex-wrap: wrap;
    }
    .footer {
      text-align: center;
      padding: 16px 0 24px;
      color: #bbb;
      font-size: 12px;
    }
    .print-area { display: none; }
    .badge-card {
      width: 86mm;
      height: 54mm;
      background: #fff;
      overflow: hidden;
      display: flex;
      flex-direction: column;
      position: relative;
      flex-shrink: 0;
      box-shadow: 0 2px 8px rgba(0,0,0,0.18);
    }
    .badge-card.staff-card .bh-role {
      display: none;
    }
    .badge-card.staff-card .bh-logo {
      left: 3mm;
    }
    .badge-card-bg {
      position: absolute;
      left: 50%;
      top: 50%;
      width: 45.5mm;
      height: auto;
      transform: translate(calc(-50% - 5px), calc(-50% + 10px));
      opacity: 0.56;
      pointer-events: none;
      z-index: 0;
    }
    .bh {
      background: #E50036;
      height: 11.3mm;
      margin-top: -5px;
      display: flex;
      align-items: center;
      padding: 0 2mm 0 0;
      gap: 2mm;
      flex-shrink: 0;
      position: relative;
      overflow: visible;
    }
    .bh-role {
      display: flex;
      flex-direction: column;
      align-items: center;
      flex-shrink: 0;
      align-self: flex-start;
      width: 10.8mm;
      border-radius: 2mm 2mm 0 0;
      margin-left: 12px;
      margin-top: 10px;
      overflow: visible;
    }
    .bh-role-top {
      width: 10.8mm;
      height: 8.6mm;
      background: #fff;
      border-radius: 2mm 2mm 0 0;
      display: flex;
      flex-direction: column;
      align-items: center;
      justify-content: center;
      gap: 0.8mm;
      flex-shrink: 0;
    }
    .bh-role-bottom {
      width: 10.8mm;
      background: #E50036;
      border-radius: 0 0 2mm 2mm;
      display: flex;
      align-items: center;
      justify-content: center;
      padding: 0.8mm 0.5mm;
    }
    .bh-role-crown svg { width: 4mm; height: 2.8mm; display: block; }
    .bh-role-cn {
      font-size: 3.5mm;
      font-weight: 900;
      color: #E50036;
      line-height: 1;
      text-align: center;
      letter-spacing: 0.5mm;
    }
    .bh-role-en {
      font-size: 1.1mm;
      font-weight: 700;
      color: #fff;
      line-height: 1;
      text-align: center;
      white-space: nowrap;
      letter-spacing: 0.1mm;
    }
    .bh-logo {
      position: absolute;
      left: 14.2mm;
      top: 10px;
      height: 6.55mm;
      width: auto;
      object-fit: contain;
      z-index: 200;
    }
    .bh-deco {
      position: absolute;
      right: 0;
      top: 0;
      height: 11.3mm;
      width: 24mm;
      display: flex;
      align-items: center;
      justify-content: flex-end;
      gap: 2mm;
      overflow: visible;
      pointer-events: none;
    }
    .bh-deco-circle {
      width: calc(14.5mm + 58px);
      height: calc(17mm - 5px);
      background: #fff;
      border-radius: 999px;
      flex-shrink: 0;
      transform: translate(calc(4mm + 18px), 3px);
      position: relative;
      z-index: 10;
    }
    .bh-deco-circle::before {
      content: '';
      position: absolute;
      inset: calc(3.6mm - 3px);
      background: #E50036;
      border-radius: inherit;
    }
    .bh-deco-circle::after {
      content: '';
      position: absolute;
      top: 0;
      bottom: 0;
      left: 50%;
      transform: translateX(-50%);
      width: 11px;
      background: #fff;
    }
    .bh-deco-pill {
      width: 5.5mm;
      height: 17mm;
      background: #fff;
      border-radius: 3mm;
      flex-shrink: 0;
      transform: translateX(4mm);
      position: relative;
    }
    .bh-deco-pill::before {
      content: '';
      position: absolute;
      inset: 3.6mm;
      background: #E50036;
      border-radius: calc(3mm - 1mm);
    }
    .bb {
      flex: 1;
      display: grid;
      grid-template-columns: 25mm 1fr 22mm;
      grid-template-rows: 11.2mm 1fr;
      padding: 2.8mm 2.1mm 2.2mm 2.1mm;
      column-gap: 1.6mm;
      row-gap: 1.2mm;
      background: transparent;
      min-height: 0;
      position: relative;
      z-index: 2;
    }
    .bb-photo {
      grid-column: 1;
      grid-row: 1 / 3;
      padding-top: calc(2.15mm + 3px);
      padding-left: 4px;
      display: flex;
      flex-direction: column;
    }
    .bb-photo-frame {
      width: 25mm;
      height: 35mm;
      flex-shrink: 0;
      border: 0.5mm dashed #999;
      position: relative;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
    }
    .bb-photo-frame img {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      object-fit: cover;
      display: block;
    }
    .bb-qr-frame img {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
      object-fit: contain;
      display: block;
    }
    .bb-photo-frame svg.placeholder-x,
    .bb-qr-frame svg.placeholder-x {
      position: absolute;
      inset: 0;
      width: 100%;
      height: 100%;
    }
    .bb-photo-frame .ph-text {
      position: absolute;
      font-size: 2.8mm;
      color: #aaa;
      font-weight: 500;
      writing-mode: vertical-rl;
      text-orientation: mixed;
      letter-spacing: 1mm;
    }
    .bb-stars {
      grid-column: 2 / 4;
      grid-row: 1;
      display: flex;
      justify-content: center;
      align-items: center;
      gap: 2px;
    }
    .bb-stars.slogan-mode {
      justify-content: center;
      padding: 0 2mm;
    }
    .bb-star {
      width: 9.8mm;
      height: 9.8mm;
      object-fit: contain;
      flex-shrink: 0;
    }
    .bb-slogan-img {
      width: 45mm;
      max-height: 8.6mm;
      object-fit: contain;
      display: block;
    }
    .bb-fields {
      grid-column: 2;
      grid-row: 2;
      display: flex;
      flex-direction: column;
      gap: 0;
      align-self: end;
      padding-right: 4%;
      margin-left: 4px;
    }
    .bb-field {
      display: flex;
      align-items: flex-end;
      margin-bottom: 1.8mm;
      gap: 0;
    }
    .bb-field-label {
      font-size: 3.65mm;
      font-weight: 700;
      color: #1a1a1a;
      white-space: nowrap;
      flex-shrink: 0;
      line-height: 1.2;
    }
    .bb-field-line {
      flex: 1;
      border-bottom: 0.5mm solid #1a1a1a;
      min-width: 6mm;
      margin-left: 0.7mm;
      padding-bottom: 0.35mm;
      font-size: 3.45mm;
      font-weight: 600;
      color: #1a1a1a;
      line-height: 1.2;
      white-space: nowrap;
      overflow: hidden;
      text-overflow: ellipsis;
    }
    .bb-qr {
      grid-column: 3;
      grid-row: 2;
      align-self: end;
    }
    .bb-qr-frame {
      width: 22mm;
      height: 22mm;
      transform: translate(-4px, 3px);
      border: 0.5mm dashed #999;
      position: relative;
      display: flex;
      align-items: center;
      justify-content: center;
      overflow: hidden;
    }
    .bb-qr-frame .ph-text {
      position: absolute;
      font-size: 2.8mm;
      color: #aaa;
      font-weight: 500;
    }
    .bb-qr-label {
      position: absolute;
      bottom: 1.2mm;
      left: 0;
      right: 0;
      text-align: center;
      font-size: 2.2mm;
      color: #888;
      pointer-events: none;
    }
    .bt {
      height: 2.9mm;
      background: #E60036;
      flex-shrink: 0;
    }
    @media print {
      @page { size: A4 portrait; margin: 5mm 7mm 5mm 7mm; }
      html, body { margin: 0; padding: 0; background: white !important; }
      .no-print { display: none !important; }
      .print-area { display: block !important; }
      .site-header,
      .footer,
      .drop-card,
      .preview-card {
        display: none !important;
      }
      .page { display: none !important; }
      .badge-page {
        display: grid !important;
        grid-template-columns: 86mm 86mm;
        gap: 2mm 4mm;
        width: 176mm;
        margin: 0 auto;
        padding: 0;
      }
      .badge-card {
        box-shadow: none !important;
        border: 0.3mm solid #e0e0e0;
        break-inside: avoid;
      }
      * {
        -webkit-print-color-adjust: exact !important;
        print-color-adjust: exact !important;
        color-adjust: exact !important;
      }
      .bb-photo { padding-top: calc(2.15mm - 3px) !important; }
    }
    @media (max-width: 900px) {
      .step-grid { grid-template-columns: 1fr; }
      .badge-preview-grid {
        align-items: flex-start;
        justify-content: flex-start;
      }
    }
  </style>
</head>
<body>
  <div class="print-area" id="printArea"></div>
  <header class="site-header">
    <div class="site-header-inner">
      <div class="brand-badge">高</div>
      <div>
        <p class="brand-title">员工工牌制作系统</p>
        <p class="brand-subtitle">高济药房 · CBWELL PHARMACY</p>
      </div>
    </div>
  </header>

  <main class="page">
    <section class="step-grid">
      <div class="card step-card">
        <div class="step-index">1</div>
        <div>
          <p class="step-title">准备 Excel 表格</p>
          <p class="step-desc">列1姓名·列2职务·列3工号·列4嵌入照片·列5嵌入企微码</p>
        </div>
      </div>
      <div class="card step-card">
        <div class="step-index">2</div>
        <div>
          <p class="step-title">导入 .xlsx 文件</p>
          <p class="step-desc">拖拽或点击上传，自动解析文字 + 嵌入图片</p>
        </div>
      </div>
      <div class="card step-card">
        <div class="step-index">3</div>
        <div>
          <p class="step-title">打印输出</p>
          <p class="step-desc">A4纸·每页10张(2×5)·实际大小·关闭页眉页脚</p>
        </div>
      </div>
    </section>

    <section class="card drop-card" id="uploadCard">
      <div class="drop-zone" id="dropZone">
        <input type="file" id="fileInput" accept=".xlsx" style="display:none">
        <div class="drop-icon-wrap">
          <svg width="28" height="28" fill="none" stroke="#E60036" stroke-width="1.5" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/>
          </svg>
        </div>
        <p class="drop-title">点击或拖拽上传 Excel 文件</p>
        <p class="drop-subtitle">支持 .xlsx 格式，自动提取照片和企业微信二维码</p>
        <div style="margin-top:12px;display:inline-flex;gap:10px;flex-wrap:wrap;justify-content:center;">
          <a class="btn-outline" href="/template.xlsx" id="downloadTemplateBtn">
            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M4 16v2a2 2 0 002 2h12a2 2 0 002-2v-2M7 10l5 5 5-5M12 15V3"/>
            </svg>
            下载 Excel 模板
          </a>
          <button class="btn-red" id="uploadTemplateBtn" type="button">
            <svg width="15" height="15" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24">
              <path stroke-linecap="round" stroke-linejoin="round" d="M12 21V9m0 0l-5 5m5-5l5 5M5 3h14a2 2 0 012 2v2H3V5a2 2 0 012-2z"/>
            </svg>
            上传 Excel 模板
          </button>
        </div>
      </div>

      <div class="progress-wrap" id="progressWrap">
        <div class="progress-row">
          <span id="progText" style="font-weight:500;color:#333;">解析中...</span>
          <span id="progPct" style="color:#999;">0%</span>
        </div>
        <div class="progress-track"><div class="progress-bar" id="progBar"></div></div>
        <p class="progress-detail" id="progDetail"></p>
      </div>

      <div class="error-wrap" id="errorWrap">
        <div class="error-box">
          <svg width="18" height="18" fill="#ef4444" viewBox="0 0 20 20" style="flex-shrink:0;margin-top:1px;">
            <path fill-rule="evenodd" d="M10 18a8 8 0 100-16 8 8 0 000 16zM8.707 7.293a1 1 0 00-1.414 1.414L8.586 10l-1.293 1.293a1 1 0 101.414 1.414L10 11.414l1.293 1.293a1 1 0 001.414-1.414L11.414 10l1.293-1.293a1 1 0 00-1.414-1.414L10 8.586 8.707 7.293z" clip-rule="evenodd"/>
          </svg>
          <div>
            <p style="font-weight:600;font-size:13px;color:#c00;margin:0;">解析失败</p>
            <p id="errorMsg" style="font-size:12px;color:#e55;margin:2px 0 0;"></p>
          </div>
        </div>
      </div>

      <div id="previewCard" class="preview-card">
      <div class="card preview-toolbar">
        <div style="display:flex;align-items:center;gap:16px;flex-wrap:wrap;">
          <span class="preview-title">工牌预览</span>
          <span class="preview-meta">共 <b style="color:#E60036;" id="totalNum">0</b> 张 · <b style="color:#555;" id="pageNum">0</b> 页</span>
        </div>
        <div class="template-controls">
          <label class="template-control">
            角色
            <div class="choice-control" data-choice="role">
              <button class="choice-trigger" type="button" data-choice-trigger>店长</button>
              <div class="choice-menu">
                <button class="choice-option is-selected" type="button" data-value="manager">店长</button>
                <button class="choice-option" type="button" data-value="staff">店员</button>
              </div>
            </div>
          </label>
          <label class="template-control">
            星级
            <div class="choice-control" data-choice="stars">
              <button class="choice-trigger" type="button" data-choice-trigger>五星</button>
              <div class="choice-menu">
                <button class="choice-option is-selected" type="button" data-value="5">五星</button>
                <button class="choice-option" type="button" data-value="4">四星</button>
                <button class="choice-option" type="button" data-value="3">三星</button>
                <button class="choice-option" type="button" data-value="2">二星</button>
                <button class="choice-option" type="button" data-value="1">一星</button>
                <button class="choice-option" type="button" data-value="0">无星</button>
              </div>
            </div>
          </label>
          <button class="btn-outline" id="reuploadBtn" type="button">重新上传</button>
          <button class="btn-outline" id="downloadPdfBtn" type="button">下载 PDF</button>
          <button class="btn-red" id="printBtn" type="button">打印工牌</button>
        </div>
      </div>

      <div class="print-tip">
        <b>打印设置：</b>A4纸 · 实际大小(100%) · 关闭页眉页脚 · 不勾选“适应页面” · 每页自动排列 10 张工牌（2列×5行）· 挂卡尺寸 86×54mm
      </div>

      <div class="card preview-shell">
        <div class="badge-preview-grid" id="badgeGrid"></div>
      </div>

      </div>
    </section>
  </main>

  <div class="footer">© 2026 高济药房 · 员工工牌制作系统 · 保留所有权利</div>

  <script>
    const fileInput = document.getElementById('fileInput');
    const dropZone = document.getElementById('dropZone');
    const progressWrap = document.getElementById('progressWrap');
    const errorWrap = document.getElementById('errorWrap');
    const errorMsg = document.getElementById('errorMsg');
    const previewCard = document.getElementById('previewCard');
    const badgeGrid = document.getElementById('badgeGrid');
    const printArea = document.getElementById('printArea');
    const progText = document.getElementById('progText');
    const progPct = document.getElementById('progPct');
    const progBar = document.getElementById('progBar');
    const progDetail = document.getElementById('progDetail');
    const totalNum = document.getElementById('totalNum');
    const pageNum = document.getElementById('pageNum');
    const printBtn = document.getElementById('printBtn');
    const downloadPdfBtn = document.getElementById('downloadPdfBtn');
    const reuploadBtn = document.getElementById('reuploadBtn');
    const uploadTemplateBtn = document.getElementById('uploadTemplateBtn');

    let currentArtifact = null;
    let currentDownloadPdfUrl = '';
    let employees = [];
    let selectedRole = 'manager';
    let selectedStarCount = 5;
    const LOGO_B64 = "__LOGO_B64__";
    const STAR_SRC = "__STAR_SRC__";
    const CENTER_BG_B64 = "__CENTER_BG_B64__";
    const SLOGAN_SRC = "data:image/svg+xml;base64,__SLOGAN_B64__";
    const POS_EN = {
      '店长':'Store Manager','副店长':'Asst. Manager','店员':'Store Staff',
      '药师':'Pharmacist','执业药师':'Pharmacist','收银员':'Cashier',
      '主任':'Director','经理':'Manager','顾问':'Consultant'
    };

    function getAnalyticsSessionId() {
      let sessionId = localStorage.getItem('analytics_session_id');
      if (!sessionId) {
        sessionId = crypto.randomUUID ? crypto.randomUUID() : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
        localStorage.setItem('analytics_session_id', sessionId);
      }
      return sessionId;
    }

    function track(eventName, payload = {}) {
      const body = {
        event_name: eventName,
        session_id: getAnalyticsSessionId(),
        page_path: window.location.pathname,
        ...payload
      };
      fetch('/api/track', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body),
        keepalive: true
      }).catch(() => {});
    }

    track('page_view');

    function getPosEn(pos) { return POS_EN[pos] || 'Employee'; }

    function esc(s) {
      return String(s || '').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
    }

    function setProgress(pct, text, detail='') {
      progressWrap.style.display = 'block';
      progPct.textContent = pct + '%';
      progBar.style.width = pct + '%';
      progText.textContent = text;
      progDetail.textContent = detail;
    }

    function hideError() {
      errorWrap.style.display = 'none';
      errorMsg.textContent = '';
    }

    function showError(msg) {
      errorWrap.style.display = 'block';
      errorMsg.textContent = msg;
      progressWrap.style.display = 'none';
    }

    function resetPreview() {
      dropZone.style.display = 'block';
      previewCard.style.display = 'none';
      badgeGrid.innerHTML = '';
      printArea.innerHTML = '';
      employees = [];
      currentArtifact = null;
      currentDownloadPdfUrl = '';
    }

    function buildStarsHtml() {
      if (selectedStarCount <= 0) {
        return `<img class="bb-slogan-img" src="${SLOGAN_SRC}" alt="护佑大众健康">`;
      }
      return Array.from({ length: selectedStarCount }, () => `<img class="bb-star" src="${STAR_SRC}" alt="★">`).join('');
    }

    function makeBadge(emp) {
      const isManager = selectedRole === 'manager';
      const roleCnDisplay = '店长';
      const roleEnDisplay = 'Store Manager';
      const starsHtml = buildStarsHtml();
      const starsClass = selectedStarCount <= 0 ? 'bb-stars slogan-mode' : 'bb-stars';
      const photoHtml = emp.photo
        ? `<img src="${emp.photo}" alt="照片">`
        : `<svg class="placeholder-x" viewBox="0 0 100 100" preserveAspectRatio="none">
            <line x1="0" y1="0" x2="100" y2="100" stroke="#ccc" stroke-width="1.2"/>
            <line x1="100" y1="0" x2="0" y2="100" stroke="#ccc" stroke-width="1.2"/>
          </svg><span class="ph-text">一寸照</span>`;
      const qrHtml = emp.qrcode
        ? `<img src="${emp.qrcode}" alt="二维码">`
        : `<svg class="placeholder-x" viewBox="0 0 100 100" preserveAspectRatio="none">
            <line x1="0" y1="0" x2="100" y2="100" stroke="#ccc" stroke-width="1.2"/>
            <line x1="100" y1="0" x2="0" y2="100" stroke="#ccc" stroke-width="1.2"/>
          </svg><span class="ph-text">企微码</span>`;

      const div = document.createElement('div');
      div.className = isManager ? 'badge-card' : 'badge-card staff-card';
      div.innerHTML = `
        <img class="badge-card-bg" src="data:image/png;base64,${CENTER_BG_B64}" alt="">
        <div class="bh">
          <div class="bh-role">
            <div class="bh-role-top">
              <div class="bh-role-crown">
                <svg viewBox="0 0 56 40" xmlns="http://www.w3.org/2000/svg" fill="none" stroke="#E60036" stroke-linejoin="round" stroke-linecap="round">
                  <rect x="6" y="28" width="44" height="8" rx="2" stroke-width="2.5"/>
                  <path d="M6 28 L6 16 L18 24 L28 6 L38 24 L50 16 L50 28 Z" stroke-width="2.5"/>
                  <circle cx="28" cy="6" r="3.5" stroke-width="2"/>
                  <circle cx="6" cy="16" r="2.5" stroke-width="2"/>
                  <circle cx="50" cy="16" r="2.5" stroke-width="2"/>
                </svg>
              </div>
              <div class="bh-role-cn">${esc(roleCnDisplay)}</div>
            </div>
            <div class="bh-role-bottom"><div class="bh-role-en">${esc(roleEnDisplay)}</div></div>
          </div>
          <img class="bh-logo" src="data:image/svg+xml;base64,${LOGO_B64}" alt="logo">
          <div class="bh-deco">
            <div class="bh-deco-circle"></div>
            <div class="bh-deco-pill"></div>
          </div>
        </div>
        <div class="bb">
          <div class="bb-photo">
            <div class="bb-photo-frame" ${emp.photo ? 'style="border-color:transparent"' : ''}>${photoHtml}</div>
          </div>
          <div class="${starsClass}">${starsHtml}</div>
          <div class="bb-fields">
            <div class="bb-field">
              <span class="bb-field-label">姓名:</span>
              <span class="bb-field-line">${esc(emp.name)}</span>
            </div>
            <div class="bb-field">
              <span class="bb-field-label">职务:</span>
              <span class="bb-field-line">${esc(emp.position)}</span>
            </div>
            <div class="bb-field" style="margin-bottom:0;">
              <span class="bb-field-label">工号:</span>
              <span class="bb-field-line">${esc(emp.employeeId)}</span>
            </div>
          </div>
          <div class="bb-qr">
            <div class="bb-qr-frame" ${emp.qrcode ? 'style="border-color:transparent"' : ''}>
              ${qrHtml}
              ${!emp.qrcode ? '<div class="bb-qr-label">企微码</div>' : ''}
            </div>
          </div>
        </div>
        <div class="bt"></div>`;
      return div;
    }

    function renderAll(emps) {
      badgeGrid.innerHTML = '';
      const perPage = 10;
      const pages = Math.ceil(emps.length / perPage);
      for (let p = 0; p < pages; p++) {
        const pageWrap = document.createElement('div');
        pageWrap.className = 'preview-a4-wrap';

        const pageLabel = document.createElement('p');
        pageLabel.className = 'preview-a4-label';
        pageLabel.textContent = `第 ${p + 1} 页 / 共 ${pages} 页`;

        const pageSheet = document.createElement('div');
        pageSheet.className = 'preview-a4-page';

        const pageGrid = document.createElement('div');
        pageGrid.className = 'preview-a4-grid';

        const slice = emps.slice(p * perPage, (p + 1) * perPage);
        slice.forEach(e => pageGrid.appendChild(makeBadge(e)));
        for (let i = slice.length; i < perPage; i++) {
          const blank = document.createElement('div');
          blank.className = 'badge-card';
          blank.style.visibility = 'hidden';
          pageGrid.appendChild(blank);
        }

        pageSheet.appendChild(pageGrid);
        pageWrap.appendChild(pageLabel);
        pageWrap.appendChild(pageSheet);
        badgeGrid.appendChild(pageWrap);
      }
      totalNum.textContent = emps.length;
      pageNum.textContent = pages;
    }

    function doPrint() {
      track('print_click', { card_count: employees.length });
      printArea.innerHTML = '';
      const perPage = 10;
      const pages = Math.ceil(employees.length / perPage);
      for (let p = 0; p < pages; p++) {
        const pageDiv = document.createElement('div');
        pageDiv.className = 'badge-page';
        if (p < pages - 1) pageDiv.style.pageBreakAfter = 'always';
        const slice = employees.slice(p * perPage, (p + 1) * perPage);
        slice.forEach(e => pageDiv.appendChild(makeBadge(e)));
        for (let i = slice.length; i < perPage; i++) {
          const blank = document.createElement('div');
          blank.className = 'badge-card';
          blank.style.visibility = 'hidden';
          pageDiv.appendChild(blank);
        }
        printArea.appendChild(pageDiv);
      }
      window.print();
    }

    async function processFile(file) {
      if (!file || !file.name.toLowerCase().endsWith('.xlsx')) {
        showError('请上传 .xlsx 格式的 Excel 文件');
        return;
      }

      hideError();
      resetPreview();
      setProgress(10, '上传中...', '正在提交 Excel 文件');
      track('upload_click');

      const formData = new FormData();
      formData.append('file', file);

      try {
        setProgress(35, '解析中...', '正在提取文字和内嵌图片');
        const res = await fetch('/api/upload', { method: 'POST', body: formData });
        const data = await res.json();
        if (!res.ok || !data.ok) {
          throw new Error(data.error || '处理失败');
        }

        setProgress(100, '生成完成', '正在加载预览');
        currentArtifact = data.artifactId;
        currentDownloadPdfUrl = data.downloadPdfUrl || '';
        employees = data.employees || [];
        renderAll(employees);
        dropZone.style.display = 'none';
        printBtn.onclick = () => doPrint();
        previewCard.style.display = 'block';
        setTimeout(() => { progressWrap.style.display = 'none'; }, 400);
      } catch (err) {
        showError(err.message || '处理失败');
      }
    }

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => {
      e.preventDefault();
      dropZone.classList.add('drag-active');
    });
    dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-active'));
    dropZone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropZone.classList.remove('drag-active');
      const file = e.dataTransfer.files && e.dataTransfer.files[0];
      processFile(file);
    });
    fileInput.addEventListener('change', (e) => processFile(e.target.files[0]));
    reuploadBtn.addEventListener('click', () => fileInput.click());
    downloadPdfBtn.addEventListener('click', async () => {
      if (!currentDownloadPdfUrl) return;
      track('pdf_download_click', { card_count: employees.length });
      downloadPdfBtn.disabled = true;
      const originalText = downloadPdfBtn.textContent;
      downloadPdfBtn.textContent = '正在生成 PDF...';
      try {
        const res = await fetch(currentDownloadPdfUrl);
        if (!res.ok) {
          let message = 'PDF 下载失败，请稍后重试';
          try {
            const data = await res.json();
            message = data.error || message;
          } catch (_) {}
          throw new Error(message);
        }
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = '挂卡打印.pdf';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
      } catch (err) {
        showError(err.message || 'PDF 下载失败，请稍后重试');
      } finally {
        downloadPdfBtn.disabled = false;
        downloadPdfBtn.textContent = originalText;
      }
    });
    document.querySelectorAll('[data-choice]').forEach((choice) => {
      const trigger = choice.querySelector('[data-choice-trigger]');
      const options = choice.querySelectorAll('.choice-option');
      trigger.addEventListener('click', (event) => {
        event.stopPropagation();
        document.querySelectorAll('.choice-control.open').forEach((openChoice) => {
          if (openChoice !== choice) openChoice.classList.remove('open');
        });
        choice.classList.toggle('open');
      });
      options.forEach((option) => {
        option.addEventListener('click', (event) => {
          event.stopPropagation();
          options.forEach((item) => item.classList.remove('is-selected'));
          option.classList.add('is-selected');
          trigger.textContent = option.textContent;
          choice.classList.remove('open');
          if (choice.dataset.choice === 'role') {
            selectedRole = option.dataset.value;
          } else {
            selectedStarCount = Number(option.dataset.value);
          }
          if (employees.length) renderAll(employees);
        });
      });
    });
    document.addEventListener('click', () => {
      document.querySelectorAll('.choice-control.open').forEach((choice) => choice.classList.remove('open'));
    });
    document.getElementById('downloadTemplateBtn').addEventListener('click', (e) => {
      e.stopPropagation();
      track('template_download_click');
    });
    uploadTemplateBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      fileInput.click();
    });
  </script>
</body>
</html>
"""


def prune_artifacts(max_age_seconds: int = ARTIFACT_TTL_SECONDS) -> None:
    cutoff = time.time() - max_age_seconds
    stale = [key for key, value in ARTIFACTS.items() if value["created_at"] < cutoff]
    if len(ARTIFACTS) - len(stale) > MAX_ARTIFACTS:
        active = sorted(
            ((key, value["created_at"]) for key, value in ARTIFACTS.items() if key not in stale),
            key=lambda item: item[1],
        )
        stale.extend(key for key, _ in active[: max(0, len(ARTIFACTS) - len(stale) - MAX_ARTIFACTS)])
    for key in stale:
        artifact = ARTIFACTS.pop(key, None)
        if artifact and artifact.get("tmpdir"):
            shutil.rmtree(artifact["tmpdir"], ignore_errors=True)


def build_template_xlsx() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "员工工牌模板"

    headers = ["姓名", "职务", "工号", "照片", "企业微信二维码"]
    samples = [
        ["李美华", "店长", "GJ-2024001", "在该列单元格内插入照片", "在该列单元格内插入企微码"],
        ["张晨", "副店长", "GJ-2024002", "在该列单元格内插入照片", "在该列单元格内插入企微码"],
    ]

    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="E33143")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, row in enumerate(samples, start=2):
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.alignment = Alignment(vertical="center")

    ws.freeze_panes = "A2"
    ws.row_dimensions[1].height = 24
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 24
    ws.column_dimensions["E"].width = 24

    guide = wb.create_sheet("填写说明")
    guide["A1"] = "使用说明"
    guide["A1"].font = Font(bold=True, size=14)
    instructions = [
        "1. 第一行表头请保持不变：姓名、职务、工号、照片、企业微信二维码",
        "2. 照片请直接插入到“照片”列对应单元格内",
        "3. 企微二维码请直接插入到“企业微信二维码”列对应单元格内",
        "4. 保存为 .xlsx 后上传到系统",
    ]
    for idx, text in enumerate(instructions, start=3):
        guide[f"A{idx}"] = text
    guide.column_dimensions["A"].width = 72

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def pil_to_b64(img, quality: int = 82) -> str:
    if img is None:
        return ""
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format="JPEG", quality=quality, optimize=True)
    return base64.b64encode(buf.getvalue()).decode()


def parse_excel(file_bytes: bytes) -> tuple[list[dict], dict]:
    df = pd.read_excel(io.BytesIO(file_bytes), dtype=str, nrows=MAX_CARDS_PER_UPLOAD + 1).fillna("")
    if len(df) > MAX_CARDS_PER_UPLOAD:
        raise ValueError(f"单次最多生成 {MAX_CARDS_PER_UPLOAD} 张工牌，请拆分 Excel 后再上传")

    max_excel_row = len(df) + 1
    photo_by_row, qr_by_row = extract_images_from_excel(file_bytes, max_row_1based=max_excel_row)

    warnings = {"missing_photo": [], "missing_qr": []}
    staff_list = []
    for i, row in df.iterrows():
        excel_row = i + 2
        name = row.get("姓名", f"第{excel_row}行")
        if excel_row not in photo_by_row:
            warnings["missing_photo"].append(f"行{excel_row}（{name}）")
        if excel_row not in qr_by_row:
            warnings["missing_qr"].append(f"行{excel_row}（{name}）")
        staff_list.append(
            {
                "name": row.get("姓名", ""),
                "title": row.get("职务", ""),
                "staff_id": row.get("工号", ""),
                "photo": photo_by_row.get(excel_row),
                "qr_code": qr_by_row.get(excel_row),
            }
        )
    return staff_list, warnings


def build_browser_cards(staff_list: list[dict]) -> list[dict]:
    cards = []
    for staff in staff_list:
        photo_b64 = pil_to_b64(staff.get("photo"), quality=82)
        qr_b64 = pil_to_b64(staff.get("qr_code"), quality=88)
        cards.append(
            {
                "name": staff["name"],
                "title": staff["title"],
                "staff_id": staff["staff_id"],
                "photo_b64": photo_b64,
                "qr_b64": qr_b64,
            }
        )
    return cards


def build_client_employees(cards: list[dict]) -> list[dict]:
    items = []
    for card in cards:
        photo_b64 = card.get("photo_b64", "")
        qr_b64 = card.get("qr_b64", "")
        items.append(
            {
                "name": card["name"],
                "position": card["title"],
                "employeeId": card["staff_id"],
                "photo": f"data:image/jpeg;base64,{photo_b64}" if photo_b64 else "",
                "qrcode": f"data:image/jpeg;base64,{qr_b64}" if qr_b64 else "",
            }
        )
    return items


def build_png_zip(pdf_path: Path, total_pages: int, output_path: Path) -> None:
    pdftoppm = shutil.which("pdftoppm")
    if not pdftoppm:
        raise RuntimeError("未找到 pdftoppm，无法生成 PNG ZIP")

    with tempfile.TemporaryDirectory() as tmpdir:
        out_prefix = Path(tmpdir) / "page"
        cmd = [
            pdftoppm,
            "-png",
            "-r",
            "180",
            "-f",
            "1",
            "-l",
            str(total_pages),
            str(pdf_path),
            str(out_prefix),
        ]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=90)
        except subprocess.TimeoutExpired as exc:
            raise RuntimeError("PNG 导出超时，请减少单次工牌数量后重试") from exc
        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for idx in range(1, total_pages + 1):
                page_path = Path(f"{out_prefix}-{idx}.png")
                if not page_path.exists():
                    break
                zf.write(page_path, f"page_{idx:02d}.png")


def get_artifact(artifact_id: str) -> dict:
    artifact = ARTIFACTS.get(artifact_id)
    if not artifact:
        raise KeyError("未找到预览内容，请重新上传 Excel 文件")
    return artifact


def ensure_artifact_pdf(artifact: dict) -> Path:
    pdf_path = Path(artifact["pdf_path"])
    if artifact.get("pdf_ready") and pdf_path.exists():
        return pdf_path

    html_content = Path(artifact["html_path"]).read_text(encoding="utf-8")
    browser_export.export_html_to_pdf(html_content, str(pdf_path))
    artifact["pdf_ready"] = True
    return pdf_path


@app.get("/")
def index():
    html = (
        INDEX_HTML.replace("__LOGO_B64__", LOGO_B64)
        .replace("__STAR_SRC__", STAR_SRC)
        .replace("__CENTER_BG_B64__", CENTER_BG_B64)
        .replace("__SLOGAN_B64__", SLOGAN_B64)
    )
    return render_template_string(html)


@app.get("/template.xlsx")
def download_template():
    track_event("template_download", success=True, page_path=request.path)
    data = build_template_xlsx()
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name="工牌信息模板.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.post("/api/track")
def api_track():
    payload = request.get_json(silent=True) or {}
    event_name = str(payload.get("event_name") or payload.get("eventName") or "")[:80]
    if not event_name:
        return jsonify({"ok": False, "error": "missing event_name"}), 400

    session_id = get_session_id(request, payload)
    returned_session_id = track_event(
        event_name,
        success=payload.get("success"),
        page_path=payload.get("page_path") or payload.get("pagePath") or request.path,
        card_count=payload.get("card_count") or payload.get("cardCount"),
        row_count=payload.get("row_count") or payload.get("rowCount"),
        meta_json=payload.get("meta") if isinstance(payload.get("meta"), dict) else None,
        session_id=session_id,
    )
    response = jsonify({"ok": True, "sessionId": returned_session_id or session_id})
    response.set_cookie(
        "analytics_session_id",
        returned_session_id or session_id,
        max_age=60 * 60 * 24 * 365,
        httponly=True,
        samesite="Lax",
        secure=request.is_secure,
    )
    return response


def _today_start() -> datetime:
    now = datetime.now().astimezone()
    return now.replace(hour=0, minute=0, second=0, microsecond=0)


def _admin_authorized() -> bool:
    configured = os.environ.get("ADMIN_STATS_PASSWORD")
    return bool(configured and session.get("admin_stats_ok"))


def _admin_stats_html(error: str = "") -> str:
    configured = os.environ.get("ADMIN_STATS_PASSWORD")
    if not configured:
        return "统计页未启用", 404
    if not _admin_authorized():
        return render_template_string(
            """
<!doctype html>
<html lang="zh-CN">
<head><meta charset="utf-8"><title>统计登录</title></head>
<body style="font-family:Arial,'Microsoft YaHei',sans-serif;max-width:420px;margin:80px auto;">
  <h1>统计登录</h1>
  {% if error %}<p style="color:#c00;">{{ error }}</p>{% endif %}
  <form method="post">
    <input name="password" type="password" placeholder="密码" style="width:100%;padding:10px;margin-bottom:12px;">
    <button type="submit" style="padding:10px 16px;">进入</button>
  </form>
</body>
</html>
            """,
            error=error,
        )

    engine = get_engine()
    if engine is None:
        return Response("统计数据库暂不可用", status=503, mimetype="text/plain; charset=utf-8")

    start = _today_start()
    seven_days_start = start - timedelta(days=6)
    with engine.begin() as conn:
        today_sessions = conn.execute(
            select(func.count(func.distinct(analytics_events.c.session_id))).where(
                analytics_events.c.created_at >= start,
                analytics_events.c.event_name == "page_view",
                analytics_events.c.session_id.is_not(None),
            )
        ).scalar() or 0
        today_page_views = conn.execute(
            select(func.count()).where(
                analytics_events.c.created_at >= start,
                analytics_events.c.event_name == "page_view",
            )
        ).scalar() or 0
        today_uploads = conn.execute(
            select(func.count()).where(
                analytics_events.c.created_at >= start,
                analytics_events.c.event_name == "upload_start",
            )
        ).scalar() or 0
        today_upload_success = conn.execute(
            select(func.count()).where(
                analytics_events.c.created_at >= start,
                analytics_events.c.event_name == "upload_success",
                analytics_events.c.success.is_(True),
            )
        ).scalar() or 0
        today_card_success = conn.execute(
            select(func.count()).where(
                analytics_events.c.created_at >= start,
                analytics_events.c.event_name == "card_generate_success",
                analytics_events.c.success.is_(True),
            )
        ).scalar() or 0
        today_pdf_download = conn.execute(
            select(func.count()).where(
                analytics_events.c.created_at >= start,
                analytics_events.c.event_name == "pdf_download",
            )
        ).scalar() or 0
        today_template_download = conn.execute(
            select(func.count()).where(
                analytics_events.c.created_at >= start,
                analytics_events.c.event_name == "template_download",
            )
        ).scalar() or 0
        recent_rows = conn.execute(
            select(analytics_events)
            .where(analytics_events.c.created_at >= seven_days_start)
            .order_by(analytics_events.c.created_at.asc())
        ).mappings().all()
        failures = conn.execute(
            select(
                analytics_events.c.created_at,
                analytics_events.c.event_name,
                analytics_events.c.page_path,
                analytics_events.c.error_message,
                analytics_events.c.session_id,
            )
            .where(analytics_events.c.success.is_(False))
            .order_by(analytics_events.c.created_at.desc())
            .limit(20)
        ).mappings().all()

    trend = []
    def row_local_date(value):
        if value.tzinfo is None:
            value = value.replace(tzinfo=timezone.utc)
        return value.astimezone().date()

    for offset in range(7):
        day = (seven_days_start + timedelta(days=offset)).date()
        day_rows = [row for row in recent_rows if row_local_date(row["created_at"]) == day]
        trend.append(
            {
                "date": day.isoformat(),
                "page_views": sum(1 for row in day_rows if row["event_name"] == "page_view"),
                "uploads": sum(1 for row in day_rows if row["event_name"] == "upload_start"),
                "upload_success": sum(1 for row in day_rows if row["event_name"] == "upload_success"),
                "card_success": sum(1 for row in day_rows if row["event_name"] == "card_generate_success"),
                "pdf_downloads": sum(1 for row in day_rows if row["event_name"] == "pdf_download"),
            }
        )

    upload_rate = f"{(today_upload_success / today_uploads * 100):.1f}%" if today_uploads else "0.0%"
    return render_template_string(
        """
<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <title>工卡工具使用统计</title>
  <style>
    body{font-family:Arial,'Microsoft YaHei',sans-serif;margin:32px;background:#f6f7f9;color:#222}
    .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:12px}
    .card{background:#fff;border:1px solid #e5e7eb;border-radius:8px;padding:16px}
    .num{font-size:28px;font-weight:700;margin-top:6px}
    table{width:100%;border-collapse:collapse;background:#fff}
    th,td{border-bottom:1px solid #eee;text-align:left;padding:10px;font-size:13px}
    h1,h2{margin:24px 0 12px}
  </style>
</head>
<body>
  <h1>工卡工具使用统计</h1>
  <div class="grid">
    <div class="card">今日独立访问会话数<div class="num">{{ today_sessions }}</div></div>
    <div class="card">今日 page_view 次数<div class="num">{{ today_page_views }}</div></div>
    <div class="card">今日上传次数<div class="num">{{ today_uploads }}</div></div>
    <div class="card">今日上传成功次数<div class="num">{{ today_upload_success }}</div></div>
    <div class="card">今日上传成功率<div class="num">{{ upload_rate }}</div></div>
    <div class="card">今日工卡生成成功次数<div class="num">{{ today_card_success }}</div></div>
    <div class="card">今日 PDF 下载次数<div class="num">{{ today_pdf_download }}</div></div>
    <div class="card">今日模板下载次数<div class="num">{{ today_template_download }}</div></div>
  </div>

  <h2>最近 7 天每日趋势</h2>
  <table>
    <thead><tr><th>日期</th><th>PV</th><th>上传</th><th>上传成功</th><th>生成成功</th><th>PDF 下载</th></tr></thead>
    <tbody>
      {% for row in trend %}
      <tr><td>{{ row.date }}</td><td>{{ row.page_views }}</td><td>{{ row.uploads }}</td><td>{{ row.upload_success }}</td><td>{{ row.card_success }}</td><td>{{ row.pdf_downloads }}</td></tr>
      {% endfor %}
    </tbody>
  </table>

  <h2>最近 20 条失败事件</h2>
  <table>
    <thead><tr><th>时间</th><th>事件</th><th>页面</th><th>错误</th><th>会话</th></tr></thead>
    <tbody>
      {% for row in failures %}
      <tr><td>{{ row.created_at }}</td><td>{{ row.event_name }}</td><td>{{ row.page_path }}</td><td>{{ row.error_message }}</td><td>{{ row.session_id }}</td></tr>
      {% endfor %}
    </tbody>
  </table>
</body>
</html>
        """,
        today_sessions=today_sessions,
        today_page_views=today_page_views,
        today_uploads=today_uploads,
        today_upload_success=today_upload_success,
        upload_rate=upload_rate,
        today_card_success=today_card_success,
        today_pdf_download=today_pdf_download,
        today_template_download=today_template_download,
        trend=trend,
        failures=failures,
    )


@app.route("/admin/stats", methods=["GET", "POST"])
def admin_stats():
    configured = os.environ.get("ADMIN_STATS_PASSWORD")
    if not configured:
        return Response("Not Found", status=404)
    if request.method == "POST":
        password = request.form.get("password", "")
        if hmac.compare_digest(password, configured):
            session["admin_stats_ok"] = True
            return redirect(url_for("admin_stats"))
        return _admin_stats_html("密码错误")
    return _admin_stats_html()


@app.post("/api/upload")
def upload_excel():
    prune_artifacts()
    artifact_tmpdir = None
    stage = "upload"
    session_id = get_session_id(request)
    track_event("upload_start", page_path=request.path, session_id=session_id)
    file = request.files.get("file")
    if not file or not file.filename:
        track_event("upload_failed", success=False, page_path=request.path, error_message="请选择 Excel 文件", session_id=session_id)
        return jsonify({"ok": False, "error": "请选择 Excel 文件"}), 400
    if not file.filename.lower().endswith(".xlsx"):
        track_event("upload_failed", success=False, page_path=request.path, error_message="请上传 .xlsx 格式的 Excel 文件", session_id=session_id)
        return jsonify({"ok": False, "error": "请上传 .xlsx 格式的 Excel 文件"}), 400

    try:
        if request.content_length and request.content_length > MAX_UPLOAD_BYTES:
            max_mb = MAX_UPLOAD_BYTES // (1024 * 1024)
            track_event("upload_failed", success=False, page_path=request.path, error_message=f"Excel 文件不能超过 {max_mb}MB", session_id=session_id)
            return jsonify({"ok": False, "error": f"Excel 文件不能超过 {max_mb}MB"}), 413

        file_bytes = file.read()
        if len(file_bytes) > MAX_UPLOAD_BYTES:
            max_mb = MAX_UPLOAD_BYTES // (1024 * 1024)
            track_event("upload_failed", success=False, page_path=request.path, error_message=f"Excel 文件不能超过 {max_mb}MB", session_id=session_id)
            return jsonify({"ok": False, "error": f"Excel 文件不能超过 {max_mb}MB"}), 413

        stage = "image_extract"
        staff_list, warnings = parse_excel(file_bytes)
        row_count = len(staff_list)
        track_event("image_extract_success", success=True, page_path=request.path, row_count=row_count, session_id=session_id)
        track_event("upload_success", success=True, page_path=request.path, row_count=row_count, session_id=session_id)
        stage = "card_generate"
        cards = build_browser_cards(staff_list)
        client_employees = build_client_employees(cards)
        del staff_list
        total_cards = len(cards)
        total_pages = max(1, math.ceil(total_cards / browser_export.CARDS_PER_PAGE))

        if total_cards > MAX_CARDS_PER_UPLOAD:
            track_event(
                "card_generate_failed",
                success=False,
                page_path=request.path,
                card_count=total_cards,
                error_message=f"单次最多生成 {MAX_CARDS_PER_UPLOAD} 张工牌",
                session_id=session_id,
            )
            return jsonify({"ok": False, "error": f"单次最多生成 {MAX_CARDS_PER_UPLOAD} 张工牌"}), 400

        html_content = browser_export.build_pdf_html(cards)
        preview_html = browser_export.build_preview_html(cards)
        del cards

        artifact_id = uuid.uuid4().hex
        artifact_tmpdir = tempfile.mkdtemp(prefix=f"store-card-{artifact_id}-")
        artifact_dir = Path(artifact_tmpdir)
        html_path = artifact_dir / "print.html"
        preview_path = artifact_dir / "preview.html"
        pdf_path = artifact_dir / "badges.pdf"
        png_zip_path = artifact_dir / "badges.zip"

        html_path.write_text(html_content, encoding="utf-8")
        preview_path.write_text(preview_html, encoding="utf-8")

        ARTIFACTS[artifact_id] = {
            "created_at": time.time(),
            "tmpdir": artifact_tmpdir,
            "total_cards": total_cards,
            "total_pages": total_pages,
            "warnings": warnings,
            "html_path": str(html_path),
            "preview_path": str(preview_path),
            "pdf_path": str(pdf_path),
            "pdf_ready": False,
            "png_zip_path": str(png_zip_path),
            "png_zip_ready": False,
        }

        TMP_PREVIEW_PATH.write_text(preview_html, encoding="utf-8")
        del file_bytes, html_content, preview_html
        track_event("card_generate_success", success=True, page_path=request.path, card_count=total_cards, row_count=row_count, session_id=session_id)

        return jsonify(
            {
                "ok": True,
                "artifactId": artifact_id,
                "totalCards": total_cards,
                "totalPages": total_pages,
                "warnings": warnings,
                "employees": client_employees,
                "previewUrl": url_for("preview_artifact", artifact_id=artifact_id),
                "printUrl": url_for("print_artifact", artifact_id=artifact_id),
                "downloadHtmlUrl": url_for("download_html", artifact_id=artifact_id),
                "downloadPdfUrl": url_for("download_pdf", artifact_id=artifact_id),
                "downloadPngUrl": url_for("download_png_zip", artifact_id=artifact_id),
            }
        )
    except ValueError as exc:
        if artifact_tmpdir:
            shutil.rmtree(artifact_tmpdir, ignore_errors=True)
        gc.collect()
        if stage == "image_extract":
            track_event("image_extract_failed", success=False, page_path=request.path, error_message=str(exc), session_id=session_id)
            track_event("upload_failed", success=False, page_path=request.path, error_message=str(exc), session_id=session_id)
        elif stage == "upload":
            track_event("upload_failed", success=False, page_path=request.path, error_message=str(exc), session_id=session_id)
        else:
            track_event("card_generate_failed", success=False, page_path=request.path, error_message=str(exc), session_id=session_id)
        return jsonify({"ok": False, "error": str(exc)}), 400
    except Exception as exc:
        if artifact_tmpdir:
            shutil.rmtree(artifact_tmpdir, ignore_errors=True)
        gc.collect()
        if stage == "image_extract":
            track_event("image_extract_failed", success=False, page_path=request.path, error_message=str(exc), session_id=session_id)
            track_event("upload_failed", success=False, page_path=request.path, error_message=str(exc), session_id=session_id)
        elif stage == "upload":
            track_event("upload_failed", success=False, page_path=request.path, error_message=str(exc), session_id=session_id)
        else:
            track_event("card_generate_failed", success=False, page_path=request.path, error_message=str(exc), session_id=session_id)
        return jsonify({"ok": False, "error": str(exc)}), 500
    finally:
        gc.collect()


@app.get("/preview/<artifact_id>")
def preview_artifact(artifact_id: str):
    try:
        artifact = get_artifact(artifact_id)
    except KeyError as exc:
        return Response(str(exc), status=404, mimetype="text/plain")
    return send_file(artifact["preview_path"], mimetype="text/html; charset=utf-8")


@app.get("/print/<artifact_id>")
def print_artifact(artifact_id: str):
    try:
        artifact = get_artifact(artifact_id)
    except KeyError as exc:
        return Response(str(exc), status=404, mimetype="text/plain")
    html = Path(artifact["html_path"]).read_text(encoding="utf-8").replace(
        "</body>",
        "<script>window.onload=function(){setTimeout(function(){window.print();},120);};</script></body>",
    )
    return Response(html, mimetype="text/html; charset=utf-8")


@app.get("/download/html/<artifact_id>")
def download_html(artifact_id: str):
    artifact = get_artifact(artifact_id)
    return send_file(
        artifact["preview_path"],
        as_attachment=True,
        download_name="store-card-latest-preview.html",
        mimetype="text/html",
    )


@app.get("/download/pdf/<artifact_id>")
def download_pdf(artifact_id: str):
    try:
        artifact = get_artifact(artifact_id)
        pdf_path = ensure_artifact_pdf(artifact)
        track_event(
            "pdf_download",
            success=True,
            page_path=request.path,
            card_count=artifact.get("total_cards"),
            meta_json={"total_pages": artifact.get("total_pages")},
        )
        return send_file(
            pdf_path,
            as_attachment=True,
            download_name="挂卡打印.pdf",
            mimetype="application/pdf",
        )
    except KeyError as exc:
        return jsonify({"ok": False, "error": str(exc)}), 404
    except Exception as exc:
        track_event(
            "pdf_download_failed",
            success=False,
            page_path=request.path,
            error_message=str(exc),
        )
        return jsonify({"ok": False, "error": f"PDF 下载失败：{exc}"}), 500


@app.get("/download/png/<artifact_id>")
def download_png_zip(artifact_id: str):
    artifact = get_artifact(artifact_id)
    pdf_path = ensure_artifact_pdf(artifact)
    if not artifact["png_zip_ready"]:
        build_png_zip(pdf_path, artifact["total_pages"], Path(artifact["png_zip_path"]))
        artifact["png_zip_ready"] = True
    return send_file(
        artifact["png_zip_path"],
        as_attachment=True,
        download_name="挂卡打印.zip",
        mimetype="application/zip",
    )


@app.get("/latest-preview")
def latest_preview():
    if TMP_PREVIEW_PATH.exists():
        return redirect(TMP_PREVIEW_PATH.as_uri())
    return Response("尚未生成预览，请先上传 Excel 并生成工牌。", mimetype="text/plain; charset=utf-8")


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8502, debug=False)
