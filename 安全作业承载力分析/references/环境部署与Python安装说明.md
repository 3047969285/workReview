# 环境部署与 Python 安装说明

本技能的计算引擎 `scripts/calcCapacity.py` 由 **Python** 编写,仅依赖 Python 标准库(要求 Python 3.8 及以上)。本说明面向 Windows 环境,从零完成 Python 安装、依赖配置与运行验证。

---

## 1. 环境要求

| 项目 | 要求 | 说明 |
|---|---|---|
| 操作系统 | Windows 10 / 11 | 64 位 |
| Python | **3.10.8**(推荐)或 3.8+ | 内网标准版安装包即 3.10.8 |
| pip | 随 Python 自带 | 用于安装依赖 |
| 内网 | 可访问内网镜像源 | 依赖均从内网镜像安装 |

---

## 2. 检查本机是否已安装 Python

打开命令行(PowerShell 或 CMD),执行:

```bash
python --version
# 或
py -3 --version
```

- 能输出版本号(如 `Python 3.10.8`)→ 已安装,直接跳到第 4 步。
- 提示 `python 不是内部或外部命令` → 未安装或未加入 PATH,继续第 3 步。

---

## 3. 安装 Python

### 3.1 内网安装包(推荐)

从内网下载官方安装包:

```
http://25.40.168.143/oss/iAssistant/python-3.10.8-amd64.exe
```

打开安装程序,安装向导中务必勾选:

- ✅ **Add python.exe to PATH**(加入环境变量,步骤 2 验证需要)
- 安装路径建议默认:`C:\Users\<用户名>\AppData\Local\Programs\Python\Python310\`

点击 **Install Now** 完成安装,重启命令行窗口后验证。

### 3.2 其它来源

- 若机器已有 Python 3.8+ 任意版本(含 Anaconda / Miniconda),可跳过安装,直接使用现有解释器。
- 确认 pip 可用:`python -m pip --version`。

---

## 4. 配置 pip 使用内网镜像源

内网隔离环境无法访问公网 PyPI,必须先配置内网镜像,否则安装依赖会失败。

```bash
# 一次性配置(写入用户级配置)
pip config set global.index-url http://25.40.168.143:50021/pip/simple
pip config set global.trusted-host 25.40.168.143
```

查看确认:

```bash
pip config list
```

应能看到 `global.index-url` 与 `global.trusted-host` 两项内网配置。

---

## 5. 安装技能依赖

按各脚本实际使用场景安装(内网镜像见第 4 步),全套装齐才能一键出报告:

```bash
# xlrd —— 读取 .xls/.xlsm 作业计划数据源（calcCapacity 依赖）  2.0.2 已含中文列兼容
pip install xlrd==2.0.2

# openpyxl —— 读取 .xlsx 数据源 / 输出结果表（calcCapacity、renderReportCharts）
pip install openpyxl

# matplotlib —— 渲染 PNG 图表（renderReportCharts：承载/管理/作业类型图）
pip install matplotlib

# Pillow —— 读取 PNG 原始宽高供等比缩放插图（applyCharts）
pip install Pillow

# pywin32 —— Word COM 自动化（fillReport 填模板 / applyCharts 插图 / renderNativeCharts 原生 Chart），含 pythoncom
pip install pywin32
```

安装验证:

```bash
python -c "import xlrd, openpyxl, matplotlib, PIL, win32com; print('all ok')"
```

---

## 6. 运行验证

在技能目录下(含 `scripts/` 子目录)执行:

```bash
# Windows 命令行
python scripts/calcCapacity.py --demo

# 或带路径:
python "……\安全作业承载力分析\scripts\calcCapacity.py" --demo
```

预期结果:

- stdout 输出一段 **JSON**(含 `persons_beta` / `days` / `periods`);
- 退出码 `0`,无报错 → **环境就绪**。

更多用法:

```bash
python scripts/calcCapacity.py 数据文件.json           # 从 JSON 文件计算
python scripts/calcCapacity.py 数据文件.json --out 结果.json   # 结果写盘
python scripts/calcCapacity.py --demo --out 结果.json  # 演示并写盘
```

---

## 7. 常见问题排查

| 现象 | 原因与处理 |
|---|---|
| `python 不是内部或外部命令` | 未加入 PATH:重装并勾选 *Add python.exe to PATH*,或手动把 `Python310` 与 `Python310\Scripts` 加入系统环境变量 |
| `pip install` 报连接超时 / Could not fetch | 内网镜像未配置,重新执行第 4 步;或临时加 `-i http://25.40.168.143:50021/pip/simple --trusted-host 25.40.168.143` |
| 运行脚本正常但中文输出乱码 | 命令行编码问题:先执行 `chcp 65001` 切换 UTF-8 再运行 |
| `xlrd` 已装却 ImportError | 多个 Python 并存导致调用错解释器;统一用 `python -m pip install …` 安装到当前解释器 |
| 退出码 3 / 4 并输出 `errorCode` JSON | 输入数据或文件问题,按返回的错误码(`INPUT_INVALID` / `FILE_NOT_FOUND` / `FILE_INVALID`)对照 `references/作业计划数据字段说明.md` 检查输入 |

---

## 8. 文件独立性说明

- 计算引擎 `scripts/calcCapacity.py` + `scripts/capacity/` 为**独立子模块**,不含报告/图表依赖(仅 Excel 输入时需 xlrd/openpyxl),可单独拷贝测算。
- 报告成链依赖 `matplotlib`(PNG 图)、`Pillow`(插图缩放)、`pywin32`(Word/WPS COM:填模板、插图、原生 Chart)、`xlrd`/`openpyxl`(读计划源),出报告需全装齐。
- `references/` 下各说明文档仅供阅读核对口径,运行计算不依赖这些文档。
