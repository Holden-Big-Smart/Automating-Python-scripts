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


---

# 行政清单与领款单自动化生成工具使用说明

## 📖 项目简介

本项目旨在自动化处理和生成“课程行政清单”及“各类领款单（课程导师费、打印费、网费、Facebook宣传费等）”。通过读取 Excel/CSV 数据源和 PDF 账单文件，批量生成 Word 文档，自动注入会计编号，并最终转换为 PDF 格式进行归档。

## 📂 目录结构说明

```text
行政清单领款单相关/
├── 0_模板文件及初始化/          # 存放 Word/Excel 模板及初始化脚本
│   ├── config_paths.py        # 路径配置文件
│   ├── 0_初始化清理.py         # 用于清理旧的输出文件
│   └── [各类 .docx 模板文件]
├── 1A_课程行政清单_领款单/      # 处理课程相关的清单和领款单
│   ├── 生成行政清单-领款单.py   # [核心脚本] 生成课程领款单
│   ├── output/                # 1A 类生成的 Word 文档输出目录
│   └── [课程及收据源数据 .csv]
├── 1B_杂费领款单/              # 处理打印费、网费、FB宣传费
│   ├── 生成杂费领款单.py       # [核心脚本] 扫描 PDF 生成杂费领款单
│   ├── 此处放入...文件/        # 存放待处理的原始 PDF 账单
│   └── output/                # 1B 类生成的 Word 文档输出目录
├── 2_Excel滙入記錄模板-支出賬文件/
│   └── [Excel 记账模板].xlsx   # 脚本会自动追加记录到此文件
├── 3_會計及財務記賬系統 - 下載文件/
│   └── [支出賬源数据].csv      # 用于匹配和注入编号的数据源
├── 4_合并处理脚本/             # 后期处理脚本
│   ├── 1_注入编号1.2.py        # [核心脚本] 将编号注入到 Word 文档
│   └── 2_合并转换.py           # [核心脚本] Word 转 PDF 并分类/合并
└── 5_Word转PDF/               # [最终产物] 存放转换后的 PDF 文件
    ├── 行政清单/
    ├── 领款单/
    └── 最终汇总/               # 合并后的总 PDF 文件

```

## 🛠️ 环境依赖

在运行脚本前，请确保安装了 Python 3.x 以及以下依赖库：

```bash
pip install pandas openpyxl python-docx docxtpl PyPDF2 pymupdf docx2pdf

```

*注意：`docx2pdf` 依赖于 Microsoft Word，请确保运行环境为 Windows 且已安装 Word。*

## 🚀 使用流程 (Step-by-Step)

建议按照以下顺序执行脚本，以完成全套工作流：
<img width="4235" height="1190" alt="未命名" src="https://github.com/user-attachments/assets/50cb533b-1309-4cae-a6b1-2b5d572ddee6" />


### 第一步：初始化与清理 (可选)

运行 `0_模板文件及初始化/0_初始化清理.py`。

* **功能**：清空之前生成的 `output` 文件夹，避免旧文件干扰。

### 第二步(可选)：生成课程行政清单与领款单

1. 确保在 `1A_课程行政清单_领款单` 目录下放入最新的课程数据 CSV 文件。
2. 运行 `1A_课程行政清单_领款单/生成行政清单-领款单.py`。

* **产出**：在 `1A.../output` 文件夹中生成对应的 Word 文档。

### 第三步(可选)：生成杂费领款单 (打印费/网费/FaceBook)

1. 将原始 PDF 账单放入 `1B_杂费领款单` 下对应的文件夹中：
* `此处放入打印费文件`
* `此处放入上网费文件`
* `此处放入FaceBook宣传费文件`


2. 运行 `1B_杂费领款单/生成杂费领款单.py`。

* **产出**：
* 在 `1B.../output` 中生成 Word 领款单。
* 自动将数据追加到 `2_Excel滙入記錄模板...` 的 Excel 文件中。
* 原始 PDF 会被移动到 `已处理文件` 归档。



### 第四步：注入电脑编号

1. 确保 `3_會計及財務記賬系統 - 下載文件` 中有最新的 `支出賬.csv` 文件（包含编号信息）。
2. 运行 `4_合并处理脚本/1_注入编号1.2.py`。

* **功能**：脚本会扫描 `1A` 和 `1B` 的输出目录，根据文件名或科目类型，从 CSV 中匹配并填入“电脑编号”。

### 第五步：格式转换与归档

运行 `4_合并处理脚本/2_合并转换.py` (或最新版脚本)。

* **功能**：
1. 将所有 Word 文档批量转换为 PDF。
2. 根据文件类型自动分类存放到 `5_Word转PDF/行政清单` 或 `5_Word转PDF/领款单`。
3. (可选) 自动合并同类 PDF 为一个总文件。



## ⚙️ 关键配置说明

* **config_paths.py**：项目中大部分路径配置集中在此文件中，如果文件夹名称变更，请修改此文件。
* **Excel 追加逻辑**：`生成杂费领款单.py` 会从目标 Excel 的第 9 行开始寻找空行追加数据，请勿随意更改 Excel 模板的前 8 行结构。
* **日期逻辑**：
* 每月 **15号** 前运行：日期设为本月 15 日（第 2 期）。
* 每月 **15号** 后运行：日期设为下月 1 日（第 1 期）。



## ⚠️ 常见问题

1. **编号注入失败**：
* 检查 `3_.../支出賬.csv` 的表头是否包含空格（脚本 v1.4 已修复此问题）。
* 确认 CSV 中的“收款人”或“种类”名称与脚本中的匹配规则一致。


2. **Word 转 PDF 报错**：
* 确保运行脚本时不要打开生成的 Word 文件。
* 确保 Windows 系统中安装了 Microsoft Office Word。


3. **找不到文件**：
* 请严格遵守目录结构，不要随意重命名核心文件夹。
