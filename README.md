# 乙方宝招标信息爬取工具

一款面向 Windows 的乙方宝/千里马企业站招标信息采集与 Excel 整理工具。当前版本针对山东济南、莱芜地区，按“监测、水土保持、测绘、测量”四组关键词检索，并提取项目名称、建设单位、资质要求、报名时间和投标截止时间等信息。

> 本项目为个人效率工具，与乙方宝、千里马及相关网站无官方关联。请遵守目标网站的服务条款、访问规则和适用法律，仅采集自己有权访问的信息。

## 主要功能

- 提供 Windows 桌面界面，也可通过 Python 命令行运行。
- 支持近 7 天、近 1 个月、近 3 个月及自定义时间范围。
- 同时执行“全文/标题检索 × 智能/精准检索”四种组合，合并并去重结果。
- 限定山东济南、莱芜地区，依次检索监测、水土保持、测绘、测量。
- 根据标题和公告正文划分“明确相关”与“疑似相关”。
- 自动提取建设单位、资格要求、报名时间和投标截止时间。
- 主表保留摘要，长文本写入“公告详情”工作表，并通过超链接互相跳转。
- 写入前按项目标题去重，保存前自动备份原 Excel。
- 支持演练模式，只显示结果而不修改 Excel。

## 项目结构

```text
.
├─ 01_直接使用软件/          # Windows 成品及分发压缩包
├─ 02_源码和说明/            # Python 源码、模板、说明和打包脚本
│  ├─ yfb_bid_spider.py      # 爬取、筛选、解析和 Excel 写入
│  ├─ yfb_spider_app.py      # Tkinter 桌面界面
│  ├─ requirements.txt       # 运行依赖
│  ├─ requirements-dev.txt   # 打包依赖
│  ├─ 打包成EXE.bat
│  └─ 乙方宝爬虫使用说明.md
├─ 03_归档备份/              # 本地旧版本和调试资料，不建议发布
├─ CHANGELOG.md
└─ PROJECT_LOG.txt
```

## 直接使用软件

1. 从 GitHub Releases 下载最新的 Windows 压缩包并完整解压。
2. 双击 `乙方宝招标信息爬取工具.exe`。
3. 登录乙方宝，在浏览器开发者工具中复制当前登录 Cookie。
4. 将 Cookie 粘贴到软件中，选择 Excel 文件和时间范围。
5. 点击“开始爬取”，完成后打开 Excel 检查结果。

Cookie 获取方法：

1. 打开并登录[乙方宝搜索页](https://qiye.qianlima.com/new_qd_yfbsite/#/infoCenter/search)。
2. 按 `F12`，在中文开发者工具中选择“网络”。
3. 在网页中执行一次搜索。
4. 在请求列表中找到名称包含 `queryZBInfo` 的请求。
5. 打开“标头”或“请求标头”，复制 `Cookie` 的完整值。

Cookie 相当于临时登录凭证。不要上传、提交到 GitHub、写入截图或发给无关人员；失效后需重新登录并获取。

## 从源码运行

### 运行环境

- 普通用户：使用 GitHub Release 中的 EXE，不需要安装 Python，也不需要创建 `.venv`。
- 源码运行或开发：Windows 10/11，Python 3.10 或更高版本，推荐 Python 3.12。
- 当前发布构建验证环境：Python 3.12.7、openpyxl 3.1.5、PyInstaller 6.21.0。

项目使用根目录的 `.venv` 隔离运行和构建依赖。该目录已被 `.gitignore` 排除，不会上传到 GitHub。首次配置：

```powershell
cd "项目根目录"
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r ".\02_源码和说明\requirements.txt"
.\.venv\Scripts\python.exe ".\02_源码和说明\yfb_spider_app.py"
```

也可以双击 `02_源码和说明\启动乙方宝爬虫工具.bat`，脚本会在缺少环境时自动创建 `.venv` 并安装依赖。

命令行运行爬虫：

```powershell
$env:YFB_COOKIE='在这里粘贴当前Cookie'
.\.venv\Scripts\python.exe ".\02_源码和说明\yfb_bid_spider.py" --days 7
```

指定 Excel 或仅演练：

```powershell
.\.venv\Scripts\python.exe ".\02_源码和说明\yfb_bid_spider.py" --days 31 --xlsx "D:\数据\招标信息.xlsx"
.\.venv\Scripts\python.exe ".\02_源码和说明\yfb_bid_spider.py" --days 31 --dry-run
```

## 当前检索规则

地区：

- 山东济南
- 山东莱芜

关键词：

- 监测
- 水土保持
- 测绘
- 测量

每个关键词执行以下四种检索组合：

| 检索范围 | 检索方式 |
| --- | --- |
| 全文检索 | 智能检索 |
| 全文检索 | 精准检索 |
| 标题检索 | 智能检索 |
| 标题检索 | 精准检索 |

脚本会再次按发布时间、地区、相关词和噪声词做本地过滤。为了尽可能减少漏项，正文命中但标题不明确的项目会标记为“疑似相关”，仍需人工复核。

## Excel 输出

主表包含以下字段：

1. 序号
2. 日期
3. 项目名称
4. 建设单位
5. 项目位置
6. 资质
7. 报名时间
8. 投标截止时间
9. 基本情况
10. 相关性
11. 备注

资质和公告正文等长文本保存在“公告详情”工作表。主表中的“公告资质”或备注链接可跳转到对应详情，详情页也提供返回主表链接。

## 打包 Windows 软件

进入 `02_源码和说明`，双击 `打包成EXE.bat`。脚本会使用根目录的 `.venv`，安装 `requirements-dev.txt` 中固定版本的构建依赖，并在 `dist\乙方宝招标信息爬取工具` 生成成品。

手动打包命令：

```powershell
cd "项目根目录"
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r ".\02_源码和说明\requirements-dev.txt"
cd ".\02_源码和说明"
..\.venv\Scripts\python.exe -m PyInstaller --noconfirm --clean --windowed --name "乙方宝招标信息爬取工具" --hidden-import yfb_bid_spider yfb_spider_app.py
```

正式发布时建议：

1. 使用不含真实业务数据的空白 Excel 模板。
2. 删除 Cookie、Token、备份表格、抓包文件和个人调试记录。
3. 将完整成品文件夹压缩为 ZIP，作为 GitHub Release 附件上传。
4. 源码仓库只保留必要源码、公开说明和脱敏模板。
5. 发布前在一台未配置 Python 的 Windows 电脑上试运行。

仓库中的 `.gitignore` 默认排除 `.venv`、本地成品、归档、构建缓存和 Excel 备份。发布 ZIP 时请在 GitHub Release 页面手动上传，不要为提交二进制文件而删除这些排除规则。

## 已知限制

- Cookie 会过期，需要用户重新获取。
- 网站接口、字段或访问策略变化后，脚本可能需要更新。
- 部分公告正文可能无法完整获取，软件会保留链接并标记人工核对。
- 建设单位和时间字段来自不同公告格式的规则提取，无法保证全部自动识别。
- 为兼顾召回率，疑似相关项目可能包含少量无关结果。
- 软件版运行中的内部任务不能强制中断，需要等待本轮结束。

## 常见问题

**提示认证失败**

重新登录乙方宝并获取新的 Cookie。

**Excel 保存失败**

先关闭正在由 Excel/WPS 打开的目标文件，再重新运行。

**没有新增数据**

可能是当前时间范围内没有符合条件的项目，也可能是项目已存在并被去重。可先使用演练模式查看日志。

**结果不完整或存在无关项目**

先检查“相关性”和“备注”列，再通过原公告链接人工确认。大范围采集后建议进行一次人工审核。

## 维护与贡献

主要配置位于 `02_源码和说明/yfb_bid_spider.py`：

- `SEARCH_KEYWORDS`：搜索关键词
- `AREA_IDS`：地区编号
- `PROTECTED_TITLE_TERMS`：优先保留词
- `NOISE_TITLE_TERMS`：常规排除词
- `STRICT_NOISE_TITLE_TERMS`：强排除词

提交问题时请提供软件版本、时间范围、项目标题和脱敏后的日志。请勿提交 Cookie、Admin-Token 或包含个人业务数据的 Excel。

版本变化见 [CHANGELOG.md](CHANGELOG.md)，开发过程与维护记录见 [PROJECT_LOG.txt](PROJECT_LOG.txt)。

## 许可证

项目目前尚未选择开源许可证。公开仓库前请根据预期的使用、修改和分发方式添加合适的 `LICENSE`；在未添加许可证时，不应默认他人已获得复制、修改或分发源码的授权。

