# 网络基线分析工具

[![版本](https://img.shields.io/badge/version-v1.0.0-blue.svg)](https://github.com/haisi-ai/Network-Baseline-Analysis)
[![Python](https://img.shields.io/badge/python-3.10+-green.svg)](https://www.python.org/)
[![PySide6](https://img.shields.io/badge/GUI-PySide6-orange.svg)](https://doc.qt.io/qtforpython/)
[![License](https://img.shields.io/badge/license-MIT-yellow.svg)](LICENSE)

规则表驱动的离线网络设备基线数据采集工具，通过 TextFSM 模板批量解析交换机/路由器/防火墙日志，自动提取设备资产信息，校验后以"一台设备一行"格式输出 Excel。

---

## 功能特性

- **规则表驱动**：所有采集规则、厂商识别、命令校验、数据校验集中在一个 Excel 规则表中，无需改代码即可扩展厂商和模板
- **多厂商支持**：H3C v5/v7、UNIS、华为、Cisco、锐捷等主流网络设备
- **自动日志识别**：根据特征码自动识别设备厂商，支持手动指定厂商
- **命令完整性校验**：采集前检查日志中是否包含所需命令，缺失命令明确提示
- **细粒度模板**：每个模板只提取一类数据，模板名体现提取内容
- **数据校验**：支持正则表达式校验提取结果，异常数据红色标记
- **多值字段合并**：电源、风扇、序列号等多值数据合并到一个单元格
- **编码自动检测**：兼容 UTF-8 / GBK / GB2312 等常见日志编码
- **程序更新检查**：内置版本检测，一键检查 GitHub 最新版本

---

## 支持的采集项

| 采集项 | 说明 |
|--------|------|
| 设备型号 | 从 display version 或 dir 输出中提取 |
| 软件版本 | 设备运行的软件版本号 |
| 运行时间 | 设备连续运行时长 |
| 设备序列号 | 机箱、板卡、电源、风扇等部件序列号 |
| 电源状态 | 各电源模块状态（Normal/Absent/Fault） |
| 风扇状态 | 各风扇模块状态 |
| 内存使用率 | 内存使用百分比 |
| CPU使用率 | 5秒/1分钟/5分钟 CPU 使用率 |

---

## 快速开始

### 方式一：直接运行 exe（推荐）

1. 下载最新版本 `网络基线分析工具.exe`
2. 双击运行，程序会自动在同目录下生成 `templates/` 和 `RuleBook.xlsx`
3. 点击「选择日志文件夹」加载设备日志
4. 勾选需要采集的设备，点击「开始采集」
5. 采集完成后点击「导出Excel」保存结果

### 方式二：源码运行

```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python main.py
```

---

## 使用说明

### 操作流程

1. **选择日志文件夹**：程序自动识别厂商并校验命令完整性
2. **勾选设备**：在左侧文件列表勾选需要采集的设备（未识别厂商的文件默认不勾选）
3. **开始采集**：按规则表中启用的模板提取数据
4. **导出Excel**：一台设备一行，异常数据红色标记

### 菜单功能

| 菜单 | 功能 |
|------|------|
| 菜单 → 获取更多模板 | 从 ntc-templates 同步1000+各厂商模板（同名跳过不覆盖） |
| 菜单 → 重新识别日志 | 重新扫描日志文件夹，执行厂商识别和命令校验 |
| 菜单 → 恢复默认模板/规则表 | 误删后恢复内置默认资源 |
| 帮助 → 使用说明 | 查看程序使用说明 |
| 关于 → 检查更新 | 从 GitHub 检查最新版本 |
| 关于 → 关于本程序 | 查看程序版本和作者信息 |

### 快捷操作

- **双击日志文件**：用系统记事本打开，方便查找问题
- **右键日志文件**：手动设置厂商、用记事本打开
- **拖拽文件夹**：支持拖拽日志文件夹到窗口

---

## 规则表说明

规则表（RuleBook.xlsx）是程序的"大脑"，所有配置集中在一个 Sheet（规则配置）：

| 列 | 列名 | 说明 |
|----|------|------|
| A | 厂商 | 设备厂商名称（H3C v5、H3C v7、UNIS、华为、Cisco、锐捷） |
| B | 特征码 | 日志识别用的特征字符串，匹配到则标记为A列厂商 |
| C | 模板名称 | 对应 templates 目录中的 .textfsm 文件名（不含后缀） |
| D | 相关命令 | 该模板对应的日志命令（含简写，分号分隔），用于日志校验 |
| E | 数据描述 | 采集内容说明，同时作为输出 Excel 的列头 |
| F | 模板是否启用 | 是/否，控制该模板是否参与采集 |
| G | 数据校验是否启用 | 是/否，控制是否校验提取值 |
| H | 数据格式 | 校验用的正则表达式 |

### 规则表的三种用途

同一行规则同时驱动三个功能：
- **日志识别**：用 B 列特征码在日志中匹配
- **日志校验**：检查 D 列相关命令是否存在于日志中
- **数据采集**：用 C 列模板文件解析日志，G/H 列控制校验

---

## 目录结构

```
网络基线分析工具/
├── main.py                  # 程序入口
├── config.json              # 配置文件（自动生成）
├── RuleBook.xlsx            # 规则表（首次运行自动生成）
├── version.txt              # 版本号
├── requirements.txt         # Python 依赖
├── app_icon.ico             # 程序图标
├── templates/               # TextFSM 模板库
│   ├── hp_comware_*.textfsm      # H3C v7 模板
│   ├── hp_comware_v5_*.textfsm   # H3C v5 模板
│   ├── unis_comware_*.textfsm    # UNIS 模板
│   ├── huawei_vrp_*.textfsm      # 华为模板
│   ├── cisco_ios_*.textfsm       # 思科模板
│   └── ruijie_os_*.textfsm       # 锐捷模板
├── core/
│   ├── config.py            # 配置管理
│   ├── rule_book.py         # 规则表读取与生成
│   ├── template_engine.py   # TextFSM 模板引擎
│   ├── template_sync.py     # ntc-templates 同步
│   ├── log_analyzer.py      # 日志识别与校验
│   ├── data_collector.py    # 数据采集调度
│   ├── data_validator.py    # 数据校验
│   ├── excel_exporter.py    # Excel 导出
│   ├── resource_manager.py  # 资源管理（封装支持）
│   └── updater.py           # 程序更新检查
└── gui/
    ├── main_window.py       # 主窗口
    └── workers.py           # 后台工作线程
```

---

## 技术栈

- **GUI**：PySide6（Qt for Python）
- **模板解析**：TextFSM
- **Excel 读写**：openpyxl
- **编码检测**：charset-normalizer
- **多线程**：concurrent.futures
- **模板来源**：ntc-templates（可选，用于获取更多厂商模板）

---

## 更新日志

### v1.0.0 (2026-09-12)
- 初始版本发布
- 支持 H3C v5/v7、UNIS、华为、Cisco、锐捷多厂商设备
- 规则表单 Sheet 驱动，8列配置
- 自动日志识别与命令完整性校验
- 细粒度 TextFSM 模板，每模板提取一类数据
- 数据校验与异常标记
- 一台设备一行 Excel 输出
- 内置程序更新检查

---

## 作者

- **海斯** - [haisi.cc](http://haisi.cc)
- **豆包** - [doubao.com](https://doubao.com)

---

## 许可证

MIT License

---

## 致谢

- [ntc-templates](https://github.com/networktocode/ntc-templates) - 提供丰富的网络设备 TextFSM 模板
- [TextFSM](https://github.com/google/textfsm) - Google 开发的模板解析引擎
