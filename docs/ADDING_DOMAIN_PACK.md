# 添加领域能力包

领域能力包位于 `domain_packs.json`，用于把新行业术语映射到通用业务概念，不应包含客户名称、客户专用规则或真实数据。

## 最小结构

```json
{
  "id": "education",
  "label": "教育培训",
  "identity_anchors": ["教育培训", "学校", "课程"],
  "anchors": ["学员", "班级", "课时", "学费"],
  "concepts": {
    "date": ["上课日期", "月份"],
    "customer": ["学员", "学生"],
    "product": ["课程", "班型"],
    "revenue": ["学费收入", "实收学费"],
    "cost": ["教师成本", "场地成本"],
    "score": ["结课成绩", "满意度"]
  }
}
```

## 设计要求

- `identity_anchors` 只放能确认行业身份的词；“库存、成本、客户”等跨行业词放在 `anchors`。
- `concepts` 使用已有标准概念；确需新增概念时同时更新编译器测试和文档。
- 别名保持短、明确，避免一个别名同时映射到多个概念。
- 不在 JSON 中写公式、Python、SQL 或客户阈值。
- 新能力包至少提供一个与现有行业无关的单元测试。

## 验证

```bash
python -m pytest -q tests/test_analysis_compiler.py
python scripts/check_secrets.py
```

测试应验证领域识别、主表选择、指标、维度、趋势能力和证据缺口，而不是只断言报告能够生成。
