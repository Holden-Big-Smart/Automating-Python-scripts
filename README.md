# 🛠️ Automating Python Scripts

[![Python](https://img.shields.io/badge/Python-3.x-blue.svg)](https://www.python.org/)
[![Status](https://img.shields.io/badge/Status-Active-success.svg)]()
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)]()

A comprehensive collection of Python automation scripts and utilities designed to streamline daily office administrative workflows, financial reporting, and marketing tasks.

This repository contains tools developed to handle repetitive tasks such as generating reports, managing course schedules, updating posters, and automating communication.

🔗 **Repository:** [Automating-Python-scripts](https://github.com/Holden-Big-Smart/Automating-Python-scripts.git)

---

## 📖 Table of Contents

- [Project Overview](#project-overview)
- [Key Features](#key-features)
  - [🏢 Office Administration](#-office-administration)
  - [💰 Finance & Accounting](#-finance--accounting)
  - [📢 Marketing & Communication](#-marketing--communication)
  - [🔧 Utilities & Extensions](#-utilities--extensions)
- [Directory Structure](#directory-structure)
- [Prerequisites](#prerequisites)
- [Disclaimer](#disclaimer)

---

## 🏗️ Project Overview

This project serves as a centralized hub for automation scripts used in an organizational environment (specifically tailored for course management and NGO operations). It leverages Python's powerful libraries (like `pandas`, `python-docx`, `python-pptx`, `pyautogui`) to interact with Excel, Word, PowerPoint, and web browsers.

## 🚀 Key Features

### 🏢 Office Administration

* **Attendance Sheet Generator (`/点名纸`)**
    * Automatically generates student attendance sheets in `.docx` format.
    * Supports both horizontal and vertical layouts based on input templates.
* **Course Management (`/自动生成下期课程`)**
    * Analyzes current course data (`.csv`) to generate schedules for the upcoming term.
    * Handles date updates and status transitions (Ongoing vs. Waiting).
* **Member Data Entry (`/会员录入`)**
    * A GUI-based tool (`fill_form_gui.py`) to assist in entering member information into systems.

### 💰 Finance & Accounting

* **Receipt Printing (`/打印收据单`)**
    * Automates the printing of receipts from Excel data (`待打印数据.xlsx`).
    * Includes logic for handling odd/even pages and A4 layout formatting.
* **Receipt Statistics (`/收据单统计`)**
    * Aggregates and merges receipt data from CSV exports.
    * Provides statistical analysis and data cleaning for financial reporting.
* **Requisition Forms & Checklists (`/行政清单领款单相关`)**
    * Complex automation for generating administrative checklists and payment requisition forms.
    * Exports data to Excel templates compatible with accounting systems.
* **Payment Reminders (`/缴费单提醒`)**
    * Identifies unpaid members from course lists and generates reminder alerts.

### 📢 Marketing & Communication

* **Poster Auto-Update (`/海报自动更新`)**
    * Dynamically updates PowerPoint (`.pptx`) posters with the latest course information (e.g., Badminton, Art classes).
    * Includes scripts for filtering data and merging multiple poster files.
* **WhatsApp Automation (`/Whatsapp發送腳本`)**
    * Automates sending WhatsApp messages to contact lists.
    * Uses image recognition (GUI automation) to locate interface elements and send text/images.
* **Contact Management (`/联系人爬虫测试`)**
    * Tools for splitting and merging contact lists (`.xlsx`), likely used for organizing bulk messaging data.

### 🔧 Utilities & Extensions

* **Browser Extensions:**
    * **Auto-Login (`/账号密码自动输入插件`):** A Chrome extension structure to auto-fill credentials.
    * **Contact Scraper (`/联系人爬虫测试`):** Extension components (`manifest.json`, `content.js`) for extracting data from web pages.
* **Video Merger (`/视频合并脚本`):** Simple script to combine multiple video files.
* **Scheduled Shutdown (`/定时关机`):** Batch script for timing system shutdowns.
* **GUI Tools (`/可视化窗口`):** Helper scripts for creating centered windows and file selection dialogs using Python (Tkinter/PyQt).

---

## 📂 Directory Structure

```text
Automating-Python-scripts/
├── 📂打印收据单 (Receipt Printing)
├── 📂点名纸 (Attendance Sheet Generator)
├── 📂定时关机 (Scheduled Shutdown)
├── 📂海报自动更新 (Poster Auto-Update)
├── 📂会员录入 (Member Entry GUI)
├── 📂缴费单提醒 (Payment Reminders)
├── 📂可视化窗口 (GUI Helpers)
├── 📂联系人爬虫测试 (Contact Scraper & Tools)
├── 📂视频合并脚本 (Video Merger)
├── 📂收据单统计 (Receipt Statistics)
├── 📂行政清单领款单相关 (Admin Checklists & Requisitions)
├── 📂账号密码自动输入插件 (Auto-login Browser Extension)
├── 📂自动生成下期课程 (Course Schedule Generator)
└── 📂Whatsapp發送腳本 (WhatsApp Automation)
```

🛠️ Prerequisites
To run these scripts, you will likely need Python installed along with the following common libraries (depending on the specific script):

Bash
```
pip install pandas openpyxl python-docx python-pptx pyautogui selenium
Note: Specific folders may have their own requirements or dependency logic.
```
⚠️ Disclaimer
All data in the project files has been anonymized and is for project demonstration purposes only.
These scripts were developed for a specific organizational workflow (Tuen Mun District Women's Association context). While the logic is reusable, file paths, template names (e.g., "屯門婦聯..."), and data structures (Excel columns) may need modification to fit other environments.

Author: Holden-Big-Smart
