# 更新记录

## 9.9.1 - 2026-11-10

- 新增本机电商订单对账演示和可下载的合成工作簿。样例覆盖少结、未匹配及重复单号隔离。
- 修复销售演示的数据选择与 Windows 包中的领域配置路径；发布时实际运行两个演示并回读 Excel。
- 收紧本地静态资源路由，移除将保存密钥恢复到明文配置的隐藏命令，完善 Excel 布尔值与小数的验收。
- 更新双语 README、安装步骤、演示结果和社区参与入口。

- 新增共享业务内核与任务范围契约：明确缺失/零、事件粒度、主档冲突、退款分摊守恒及库存期间边界，保留现有界面与 API。
- 修复订单、库存、考勤、餐饮及电商反例；业务约束验收与文件验收分开，内核变更隔离旧缓存与授权。
- 修复任务存储和定时器连接释放；启动不自动删除过期历史资料，显式清理保护在用和后台任务。
- 增加内核属性测试、跨行业反例和实际 HTTP→后台任务→Excel 下载回读测试；迁移边界见 [业务内核说明](docs/BUSINESS_KERNEL.md)。
- 新增版本化业务规则准备层、确认面板与本机客户配置复用，拒绝字段漂移、未确认状态与不安全关联。
- 通用编译器支持确认的自定义指标、时间口径与维度属性关系；执行按契约计算，不因换字段名恢复猜测聚合。
- 规则路径新增原始记录留档、排除明细与完整规则审计；有需求证据缺口时不声明完整经营诊断。
- 新增合成规则验收及实际 localhost HTTP → 确认 → 执行 → Excel 下载回读测试。
- 保留现有行业路径。

## 9.9.0 - 2026-08-25

- Published the open-source repository with an English-first README, repeatable synthetic demo, community templates, and release documentation.
- Added a version-checked GitHub Release workflow and reproducible Windows PyInstaller packaging with executable smoke tests.
- 增加版本化数据合同、结构漂移检测和跨字段规则校验。
- 增加不保存客户原值的处理血缘与审计证据包。
- 增加持久化后台任务、进度、取消、失败重试和重启恢复。
- 增加任务级会话隔离与连续追问上下文。
- 增加 pandas、DuckDB、Polars 自动引擎路由与安全回退。
- 增加加密数据库连接中心和 AST/规则双层只读 SQL 校验。
- 增加工具注册表、AI 调用追踪、回归场景库与费用确认门。
- 增加专业质量配置、持续集成和真实客户逐单验收流程。
