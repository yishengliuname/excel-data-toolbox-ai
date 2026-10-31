"""Compact decision view without discarding the underlying risk evidence."""
from collections import OrderedDict


def summarise_risks(rows: list[dict]) -> list[dict]:
    groups = OrderedDict()
    fields = ('优先级', '问题', '建议行动', '责任岗位', '建议期限', '验收指标', '性质')
    for row in rows:
        key = tuple(row.get(k) for k in fields)
        groups.setdefault(key, []).append(row)
    results = []
    for key, items in groups.items():
        objects = list(dict.fromkeys(str(r.get('对象', '')) for r in items))
        evidence = list(dict.fromkeys(str(r.get('数据证据', '')) for r in items))
        results.append({'优先级': key[0], '问题': key[1], '涉及对象数': len(objects),
            '对象示例': '、'.join(objects[:4]) + (f' 等{len(objects)}个' if len(objects) > 4 else ''),
            '数据证据': ('示例：' if len(evidence) > 2 else '') + '；'.join(evidence[:2])[:600], '建议行动': key[2], '责任岗位': key[3],
            '建议期限': key[4], '验收指标': key[5], '性质': key[6],
            '完整明细': '风险行动明细：按问题和对象筛选；本页未将不同口径金额相加'})
    return sorted(results, key=lambda r: (r['优先级'], -r['涉及对象数'], r['问题']))
