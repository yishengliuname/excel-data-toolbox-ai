# 表格快处 AI｜本地优先的 Excel 智能工作台

[![CI](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/ci.yml/badge.svg)](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/ci.yml)
[![CodeQL](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/codeql.yml/badge.svg)](https://github.com/yishengliuname/excel-data-toolbox-ai/actions/workflows/codeql.yml)
[English README](README_EN.md)

一句自然语言 → AI 规划 → 本地安全执行 → 可审计 Excel 交付。适合把真实客户的清洗、合并、对账、经营诊断和可视化需求，转换成可复用、可验收的工作流。

[5 分钟演示](docs/DEMO.md) · [开始贡献](CONTRIBUTING.md) · [贡献者路线图](docs/CONTRIBUTOR_ROADMAP.md)

这是一个开源、本地优先的 Excel 数据处理与经营分析工作台。上传数据后直接描述需求，系统会判断需要清洗、合并、对账、分析、可视化，还是生成 VBA、Power BI、数据库或人工核验交付物。数据计算和 Excel 图表在本机完成；启用 DeepSeek 时，默认只发送需求和表结构目录，不发送完整单元格原值。

> **重要：本仓库不包含任何 API Key、客户文件或真实业务数据。** `.env`、`user_data/`、`outputs/`、Excel/CSV、数据库和日志均默认禁止提交。公开 Fork 前请运行 `python scripts/check_secrets.py`。

## 快速开始

需要 Python 3.11 或更高版本。

```powershell
git clone <repository-url> excel_data_toolbox
cd excel_data_toolbox
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[automation]"
Copy-Item .env.example .env
python -m excel_data_toolbox.server
```

打开 `http://127.0.0.1:8501/`。不填写 DeepSeek API Key 也可使用本地确定性处理和报告能力。

## 为什么不是“模型直接改 Excel”

统一链路为：

`一句自然语言 → 意图/领域编译 → 白名单参数计划 → 本地安全校验 → 本机执行 → 自动验收 → 图表或交付包`

模型负责理解口语；本地代码负责字段存在性、数据类型、聚合语义、资源上限和文件安全。证据不足时报告缺口，不为了生成答案而猜测。

## 项目级 AI 配置

复制 `.env.example` 为本机 `.env`，再填写你在 DeepSeek 控制台新建的密钥：

```text
DEEPSEEK_API_KEY=
DEEPSEEK_MODEL=deepseek-v4-flash

# 可选：配置后 Power BI 任务会无人值守发布到 Fabric 工作区
POWER_BI_TENANT_ID=
POWER_BI_CLIENT_ID=
POWER_BI_CLIENT_SECRET=
POWER_BI_WORKSPACE_ID=
```

保存后重启程序即可。生产环境建议运行 `scripts/store_deepseek_key.py`，把密钥保存到 Windows DPAPI 本机加密保险箱；环境变量和 `.env` 仅作为兼容回退。不要使用曾经发到聊天、截图或公开仓库中的旧密钥。

## 核心能力

- AI 一句话完成：支持 `deepseek-v4-flash` 与 `deepseek-v4-pro`；模型夹带解释或 Markdown 时自动提取 JSON，解析失败会自动纠正重试一次
- DeepSeek 连接检测：单独判断网络/防火墙、密钥、余额、限流和模型是否可用；检测不发送表格数据
- 安全自动执行：AI 输出先经过兼容归一化、白名单和资源上限校验，再在任务副本中执行；失败不会覆盖原始文件
- 28 类本地白名单能力：专业财务、销售、库存、员工经营分析和通用自适应报告，可与清洗、验收、连接、模糊匹配、容差对账、趋势、透视、贡献、异常、RFM、脱敏等串成复杂工作流
- 通用分析编译器：根据需求、文件、工作表和字段共同识别领域、表角色、粒度、标准概念和证据缺口，再动态选择指标、排名、结构、趋势、风险与图表
- 参数兼容：兼容 `fill_missing`、`date_format`、`group_by`、`metrics`、图表 `series` 对象等常见模型变体；未知危险字段继续拦截
- 员工经营报告：员工主档、考勤、绩效、薪资调整和老板备注按字段结构自动识别，本机生成十表人效报告和四张原生图表；已知业务不依赖 DeepSeek JSON 或网络
- 高级工程订单：VBA 静态扫描代码包、SQLite/ODBC 只读查询、业务决策矩阵；Power BI 可生成自包含星型模型、Power Query、DAX 和 PBIP/PBIR 工程包，完整配置服务主体后自动发布并回读验证
- OCR 与大数据：PDF 结构化表格提取、图片中英文 OCR、DuckDB 查询 CSV/Parquet；缺少系统组件时健康检查会准确提示
- 一单一目录：任务表、图表历史、撤销/重做和交付物持久化隔离；支持恢复、清空和按保留期自动清理
- 自动交付验收：导出后重新打开文件，核对表名、行列数、内容指纹和数值合计，失败则阻止交付并生成 JSON 报告
- 保真导出：可在原工作簿副本上替换或追加结果表，尽可能保留公式、样式、隐藏表、图表、图片、命名区域和 `.xlsm` VBA
- 隐私分层：默认只发送需求、表名、字段名、类型及行列规模；密钥可由 Windows DPAPI 加密保存，不进入任务、配方、浏览器存储、导出文件或操作日志

- 可复用处理方案：把清洗、替换、筛选、去重、字段整理和分组汇总保存成安全 JSON 配方；支持先预演、再正式复跑
- 可配置质量验收：非空、唯一、范围、正则和允许值规则逐条检查，自动生成规则汇总、失败明细并送入人工核验中心
- 高级容差对账：多键匹配、Decimal 金额容差、日期窗口、重复键隔离、候选评分，以及保守的一对多/多对一待确认机制
- 人工核验中心：统一收纳相似名称、模糊匹配、异常值、质量失败和对账差异，支持接受、拒绝、备注与状态统计
- 专业交付包：处理摘要、验收清单、风险提示、统一表头/冻结窗格/筛选器/列宽及公式注入防护
- 更稳的任务操作：原表不覆盖，支持撤销和重做；上传 Excel 时预检公式、宏、图表、透视表、外链和隐藏工作表

## 从随口描述到自动执行

你可以直接说业务口语，不需要自己编写 JSON 或参数。处理链路是：

`随便一句口语 → DeepSeek 标准业务话术 → 白名单参数计划 → 本地安全校验 → 自动执行与验收`

例如你说：

> 把各门店上半年的乱订单合起来，重复的别算；客户名字八成八以上相似的和档案对一下，再跟银行回款核账。金额差五分钱以内、晚七天可以接受，但拆分回款和重复流水都让我确认。最后按区域月份看业绩，找异常订单和高价值客户，别改原表。

DeepSeek 会把它整理为类似下面的标准话术：

> 清理并按订单号去重后纵向合并所选订单表；客户名称以 88% 阈值模糊匹配客户主数据；订单与回款按订单号进行金额容差 0.05 元、日期容差 7 天的保守对账，重复键和拆分候选进入人工核验；生成区域月度汇总、IQR 异常明细及 RFM 客户分群，所有输出另存新表。

随后程序生成并校验参数计划。`88%`、`0.05 元`、`7 天` 等会改变结果的条件如果没有给出，程序会要求补充，不静默猜测。DeepSeek API Key 只在服务端请求时读取，不会返回网页或写入计划、日志和导出文件。

## 智能分析能力

- 一键数据体检：质量评分、问题清单、缺失/重复/混合类型/常量字段、描述统计与相关性
- 16 类专业可视化：柱状、横向条形、分组柱、堆叠柱、折线、面积、环形、雷达、漏斗、瀑布、矩形树、直方、散点回归、箱线、二维热力和甘特图
- 散点回归自动给出线性方程、R² 与相关系数，并明确提示相关/回归不代表因果
- Power BI 风格经营看板：最多固定 6 张图，单图导出 PNG，看板可打印或导出 PDF
- 高价值分析：IQR/Z-Score 异常检测、交叉透视、趋势聚合、分类贡献与 RFM 客户价值分群
- 分析交付包：原数据、质量概览、问题建议、描述统计、相关性和异常明细
- 销售演示数据：不上传客户文件也可以完整体验分析流程并制作宣传截图

## 专业财务能力

- 应收账龄：按截止日计算未结金额和逾期天数，输出未到期、1-30、31-60、61-90、90 天以上账龄，以及客户账龄矩阵和无效日期提示
- 预算差异：输出实际、预算、差异额、差异率，并按收入或成本口径区分有利/不利差异
- 现金流：自动识别流入/流出，生成月度流入、流出、净现金流、累计净现金流和分类现金流
- 财务比率：按现有字段计算毛利率、净利率、流动/速动比率、资产负债率、ROA、ROE、经营现金流比率及周转率；缺失字段不补造
- 凭证审计：按凭证号汇总借贷方，检测不平衡凭证、负数/无效金额和缺少科目的异常分录
- 专业导出：金额自动使用千分位，百分比、周转率和流动比率使用对应财务格式，负值红色括号显示

示例命令：

> 分析当前应收表，以 2026-08-31 为截止日，用“到期日、应收金额、已收金额、客户、发票号”计算未结金额，按未到期、1-30、31-60、61-90、90天以上生成账龄明细、账龄汇总和客户账龄表，原始数据不要覆盖。

> 检查当前凭证表，凭证号为“凭证号”，借贷字段为“借方金额、贷方金额”，允许 0.01 元尾差，输出不平衡凭证和异常分录。

财务计算由本地确定性规则完成，DeepSeek 只负责理解口语、识别字段和生成白名单计划。结果适合数据整理和辅助复核，不替代注册会计师、税务师或企业会计政策判断。

## 数据处理能力

- 导入多个 `.xlsx`、`.xlsm`、`.csv`，并支持 `.pdf`、图片、`.db/.sqlite` 与 `.parquet` 高级输入
- 数据质量概览：行列数、空值、重复行、字段类型和前 100 行预览
- 一键清洗：文本空格/换行、空行空列、按关键字段去重、空值填充、类型识别
- 纵向追加：多文件/多表按列名对齐，支持全部字段或共同字段
- 关键字段匹配：左连接、内连接、全连接，并在重复键可能造成数据膨胀时拦截
- 新旧数据比对：新增、删除、修改、未变化、重复键五类结果
- 分组汇总：计数、去重计数、求和、平均、最大、最小
- 批量拆分：按字段值拆成一个多工作表 Excel，或多个 Excel 的 ZIP 包
- 数据脱敏：手机号、邮箱、姓名、身份证号、银行卡号与通用部分隐藏
- 原文件不覆盖、处理结果独立生成、支持撤销与重做
- 导出 Excel/CSV ZIP，并附“处理摘要”；默认防护表格公式注入

## 一键启动

在本目录双击：

`启动表格快处.cmd`

启动窗口不要关闭。浏览器会自动打开 `http://127.0.0.1:8501`；关闭窗口即可停止应用。任务默认在本机隔离保留 30 天，也可在界面一键清空。

当前电脑会优先使用 Codex 随附的 Python。若以后把项目复制到另一台电脑，请安装 Python 3.12，然后在本目录运行：

```powershell
python -m pip install -r requirements.lock -r requirements-optional.txt
python server.py
```

## 推荐工作流

1. 建立任务号，导入客户的脱敏样例；先查看工作簿风险预检与数据体检。
2. 配置质量验收规则或载入复用方案，先“预演”确认影响行数。
3. 需要自动编排时，在“AI 一句话完成”里选择允许使用的表、填写 DeepSeek Key 与需求，先审查计划和预演结果。
4. 勾选人工确认后执行；差异、歧义和异常会进入“人工核验中心”。
5. 与客户确认待核验记录，保留接受/拒绝状态和处理记录。
6. 在“脱敏导出”中勾选“专业 V3 交付包”，生成结果、摘要与验收清单。
7. 客户验收后点击“清空本任务”，并按约定删除客户文件。

详细资料： [架构说明](docs/ARCHITECTURE.md)、[隐私说明](docs/PRIVACY.md)、[领域能力包指南](docs/ADDING_DOMAIN_PACK.md)、[路线图](docs/ROADMAP.md) 和 [贡献指南](CONTRIBUTING.md)。

## 当前边界

- VBA 会生成经过危险指令扫描的模块和测试清单，但不会绕过 Office 宏信任中心、自动启用宏或破解密码。
- Power BI 会生成可审计工程包并本地验收；云发布仍需要客户合法的 Microsoft Entra/Fabric 许可证、服务主体和工作区权限，AI 不能绕过登录、MFA 或付费授权。
- SQLite/ODBC 只允许 `SELECT/WITH` 单条只读查询，不写回数据库；连接权限由客户提供。
- 保真导出以原工作簿副本为底稿，但第三方插件、损坏文件、部分透视/切片器缓存或私有格式无法承诺 100% 无损。
- 不承接论文/考试数据造假、流水/票据/证明伪造；模糊匹配、拆分回款、主观业务审批等高风险结论必须保留复核痕迹。
- 默认单文件上限 50 MB、普通表上限 300,000 行；CSV/Parquet 汇总可切换 DuckDB，但结果集仍受安全上限控制。

## 隐私与安全

- 服务只监听 `127.0.0.1`，同一局域网其他设备无法直接访问。
- 不使用 AI 功能时不会连接 DeepSeek；使用时只发送需求和所选表的结构目录，默认不发送单元格值。
- API Key 优先保存在 Windows DPAPI 本机加密保险箱，也兼容 `.env`/进程环境变量；打包脚本不会包含 `.env`。
- 每个订单使用独立任务目录和 SQLite 索引；默认保留 30 天，可恢复、删除或一键清空。
- 配方只保存规则定义；任务数据只保存在对应任务目录，不跨单污染。
- 原始文件与结果目录分开，处理函数不会修改传入的原始数据表。
- 日志只记录处理规则和行数变化，不记录客户原始内容。
- 未经客户明确许可，不应把客户文件或截图作为案例。

## 项目结构

```text
excel_data_toolbox/
├─ server.py                  本地 HTTP 服务与应用会话
├─ core.py                    数据处理核心
├─ analytics.py               质量诊断、异常、趋势、透视与客户分群
├─ fuzzy.py                   相似值候选组与保守模糊匹配
├─ recipes.py                 安全、可复跑的声明式处理方案
├─ validation.py              可配置质量验收规则与失败明细
├─ reconciliation.py          金额/日期容差与可解释高级对账
├─ nl_agent.py               DeepSeek 规划、严格 JSON 校验与白名单执行器
├─ analysis_compiler.py      通用需求、领域、语义、证据与图表编译器
├─ domain_packs.json         可配置行业身份词和标准业务概念
├─ metric_semantics.py       金额、余额、比例和评分聚合语义
├─ delivery_qa.py            交付文件重开、指纹与合计自动验收
├─ workbook_fidelity.py      原工作簿保真副本导出
├─ task_store.py             一单一目录、恢复与留存策略
├─ large_data.py             DuckDB 大 CSV/Parquet 只读查询
├─ advanced_automation.py    VBA、数据库、PDF 与 OCR 安全适配
├─ io_utils.py                Excel/CSV 导入导出与安全处理
├─ models.py                  配置、报告与操作日志模型
├─ web/                       中文界面（HTML/CSS/JS）
├─ tests/                     核心回归测试
├─ .github/                   CI、CodeQL、Dependabot、Issue/PR 模板
├─ docs/                      架构、隐私、扩展和路线图
└─ 启动表格快处.cmd           Windows 双击启动器
```

## 参与贡献

欢迎修复 Bug、改善图表、补充测试、扩展领域包和增强安全能力。提交 PR 前运行：

如果你第一次参与，先看 [5 分钟演示](docs/DEMO.md) 和 [贡献者路线图](docs/CONTRIBUTOR_ROADMAP.md)，从标有 `good first issue` 或 `help wanted` 的任务开始。欢迎先在 [Discussions](https://github.com/yishengliuname/excel-data-toolbox-ai/discussions) 交流方案，再提交 PR。

```bash
python scripts/check_secrets.py
python -m ruff check . --select E9,F63,F7,F82
python -m pytest -q
python -m build
```

请使用 GitHub 提供的 Bug、Feature 或 Domain Pack 模板。安全漏洞按照 [SECURITY.md](SECURITY.md)私密报告，不要公开有效凭据或客户数据。

## 许可证

本项目使用 [Apache License 2.0](LICENSE)。你可以使用、修改和商业化，但须遵守许可证中的版权、通知和专利条款。软件按“现状”提供，不构成财务、税务、法律、审计、人事或投资建议。
