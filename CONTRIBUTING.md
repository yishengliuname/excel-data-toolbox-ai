# 贡献指南

感谢你帮助完善表格快处 AI。项目欢迎 Bug 修复、测试、文档、可视化改进、领域词典和新的安全数据能力。

## 开始之前

- 不要提交真实客户文件、截图、数据库、日志或任何可识别个人/企业的信息。
- 不要把 API Key、密码、令牌、私钥或连接字符串放进代码、Issue、PR、提交记录或测试夹具。
- 财务、人事、审批和风险结论必须保持“事实、推断、建议、人工核验”边界。
- 新功能必须使用本地白名单执行和参数校验，不能让模型直接执行任意代码、SQL 或文件操作。

## 本地开发

```bash
git clone <your-fork-url> excel_data_toolbox
cd excel_data_toolbox
python -m venv .venv
```

Windows PowerShell：

```powershell
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -e ".[automation,dev]"
Copy-Item .env.example .env
```

Linux/macOS：

```bash
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -e '.[automation,dev]'
cp .env.example .env
```

DeepSeek 是可选能力。没有 API Key 时，本地确定性清洗、分析、报告和测试仍应可运行。

## 提交前检查

```bash
python scripts/check_secrets.py
python -m ruff check .
python -m pytest -q
python -m build
```

所有测试数据必须由代码生成，或明确证明为无版权、无隐私的虚构数据。不要通过 `.gitignore` 的强制参数绕过秘密和客户数据保护。

## 贡献流程

1. 先搜索已有 Issue；较大功能先开 Feature Request 讨论边界。
2. 从 `main` 创建短分支，例如 `fix/chart-axis-title`。
3. 保持单个 PR 目标清晰，补充或更新自动测试。
4. 在 PR 中说明输入结构、预期结果、风险边界和验证命令。
5. 图表或 Excel 布局变更应附脱敏截图；不得包含客户数据。
6. 至少一名维护者评审通过、CI 全绿后合并。

## 添加新行业，不新增客户专用分支

优先扩展 `domain_packs.json`，再补充跨行业测试。只有当业务计算存在稳定、可审计的确定性规则时，才增加专用能力模块。详见 [领域能力包指南](docs/ADDING_DOMAIN_PACK.md)。

## PR 验收标准

- 原始输入不被覆盖。
- 历史输出不会重新进入新任务输入。
- 比率、余额、金额和评分使用正确聚合语义。
- 证据不足时输出缺口，不伪造结论。
- Excel 导出能重新打开，表名、行列、内容指纹和数值合计一致。
- 新依赖有明确用途，并同步更新依赖和安全说明。

提交贡献即表示你同意按本项目 Apache-2.0 许可证提供该贡献。
