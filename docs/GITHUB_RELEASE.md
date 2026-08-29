# GitHub 发布与协作指南

这份项目按“代码公开、密钥私有、客户数据不入库”的方式准备。发布前请先完成本页检查；不要把客户 Excel、`.env`、日志或导出报告提交到 GitHub。

## 首次发布

在 GitHub 新建一个空仓库（建议不要勾选自动生成 README、许可证或 `.gitignore`），然后在本地项目目录执行：

```powershell
git init -b main
git add .
git diff --cached --name-only
python scripts/check_secrets.py
git commit -m "chore: prepare open-source release"
git remote add origin https://github.com/<owner>/<repository>.git
git push -u origin main
```

`git diff --cached --name-only` 是人工复核点：列表中不应出现 `.env`、真实客户数据、`outputs/`、`user_data/`、日志或打包产物。密钥扫描只输出文件名和规则编号，不会回显密钥。

后续开发建议使用分支和 Pull Request：

```powershell
git switch -c feat/<short-description>
git push -u origin feat/<short-description>
```

## GitHub 仓库设置

创建仓库后，在 Settings 中完成：

1. 启用 Actions，并将默认工作流权限保持为“只读”；本项目的 CI 已声明最小权限。
2. 启用 Secret scanning 和 Push protection；如果密钥曾经进入历史，先撤销并轮换，再清理历史。
3. 启用 Dependabot alerts 和 security updates；仓库中的 `.github/dependabot.yml` 已配置每周检查 Python 与 Actions 依赖。
4. 启用 Private vulnerability reporting，让安全问题通过私下渠道提交。
5. 为 `main` 配置分支保护：要求 CI 通过、要求 PR、禁止直接推送和强制推送。

## 邀请协作者

在 Settings → Collaborators 中按 GitHub 用户名或团队邀请。权限按职责分配：

- `Triage`：只处理 Issue，适合业务测试人员。
- `Write`：提交分支和 Pull Request，适合日常开发者。
- `Maintain`：管理 Issue、Actions 和发布，适合核心维护者。
- `Admin`：仅给项目所有者或受信任的发布负责人。

不要为了“方便修复”给所有人 `Admin`。客户原始数据仍通过私下渠道提供，不通过公开 Issue 或 PR 附件传递。

## 发布前验收

本地建议按以下顺序执行：

```powershell
python -m pip install -e ".[automation,dev]"
python scripts/check_secrets.py
python -m ruff check . --select E9,F63,F7,F82
python -m pytest -q
python -m build
```

如果测试使用假密钥，必须明确带有 `unit-test` 或 `fake-` 标记；生产密钥绝不能出现在源码、测试、截图、Issue、Actions 日志或构建产物中。

## 客户数据交付边界

源码仓库只包含通用引擎、领域包、示例配置和测试。真实客户文件应保存在客户授权的本地或私有存储中。对外提交时只提交脱敏后的最小复现样例，并确认没有姓名、电话、地址、账号、订单号、财务凭证或 API key。

## 官方参考

- [GitHub 社区健康文件](https://docs.github.com/en/communities/setting-up-your-project-for-healthy-contributions/creating-a-default-community-health-file)
- [GitHub Secret scanning](https://docs.github.com/en/code-security/concepts/secret-security/secret-scanning)
- [GitHub Dependabot 配置参考](https://docs.github.com/en/code-security/reference/supply-chain-security/dependabot-options-reference)
