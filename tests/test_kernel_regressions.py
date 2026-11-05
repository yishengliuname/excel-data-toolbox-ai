"""Independent, small counterexamples for shared business contracts."""
from decimal import Decimal

import pandas as pd
import pytest
from openpyxl import Workbook
from openpyxl.worksheet.table import Table

from excel_data_toolbox.adaptive_report import build_adaptive_analysis_report
from excel_data_toolbox.customer_order_analysis import build_customer_order_analysis
from excel_data_toolbox.hr_report import build_hr_management_report
from excel_data_toolbox.inventory_report import build_inventory_management_report
from excel_data_toolbox.restaurant_report import build_restaurant_diagnosis_report
from excel_data_toolbox.workbook_fidelity import _write_frame
from excel_data_toolbox.ecommerce_report import build_ecommerce_diagnosis_report
from excel_data_toolbox.tests.test_inventory_report import _inventory_tables
from excel_data_toolbox.tests.test_hr_report import _hr_tables
from excel_data_toolbox.tests.test_restaurant_multifact_report import _multifact_fixture

INV = ['商品资料', '期初库存', '采购入库', '销售出库', '库存调整', '仓库说明']
HR = ['员工基础信息', '月度考勤记录', '绩效数据', '薪资调整记录', '老板需求说明']


def orders(*, cancel=True, refund=.01, state='退款成功'):
    sales = pd.DataFrame([
        ['A', '2026-08-01', '一店', 'P1', 1, 10, 0, '已完成'],
        ['A', '2026-08-01', '一店', 'P2', 1, 10, 0, '已完成'],
        ['A', '2026-08-01', '一店', 'P3', 1, 10, 0, '已完成'],
        ['B', '2026-08-02', '一店', 'P1', 1, 20, 0, '已完成'],
        *([['C', '2026-08-02', '一店', 'P1', 1, 10, 0, '已取消']] if cancel else []),
    ], columns=['订单号', '下单时间', '门店名称', 'SKU', '数量', '单价', '优惠金额', '订单状态'])
    master = pd.DataFrame({'商品编码': ['P1', 'P2', 'P3'], '标准商品名称': ['甲', '乙', '丙'],
                           '标准类目': ['食品'] * 3, '成本价': [2, 2, 2]})
    refunds = pd.DataFrame({'退款单号': ['R1'], '订单编号': ['A'], '退款金额': [refund], '退款状态': [state]})
    return [sales, master, refunds]


def order_report(fs):
    return build_customer_order_analysis(fs, source_names=['销售', '商品主数据', '退款'], user_request='分析订单及退款')


def test_clean_orders_have_typed_empty_exceptions():
    result = order_report(orders(cancel=False))
    assert result.outputs['异常数据'].empty
    assert '排除原因' in result.outputs['异常数据']


@pytest.mark.parametrize('amount', [.01, 1., 5.])
def test_refund_allocation_conserves_source_cents(amount):
    result = order_report(orders(refund=amount))
    assert result.report['refund_amount'] == pytest.approx(amount, abs=1e-9)
    assert sum(Decimal(str(x)) for x in result.outputs['清洗明细']['退款金额']) == Decimal(str(amount))


def test_pending_refund_is_not_realised():
    result = order_report(orders(refund=5, state='待审核'))
    assert result.report['refund_amount'] == 0
    assert not result.outputs['人工核验'].empty


def test_conflicting_refund_ids_are_not_silently_accumulated():
    fs = orders(refund=5)
    other = fs[2].copy()
    other['退款金额'] = 7.
    fs[2] = pd.concat([fs[2], other], ignore_index=True)
    with pytest.raises(ValueError, match='冲突'):
        order_report(fs)


def test_invalid_discount_is_not_zero():
    fs = orders()
    fs[0]['优惠金额'] = fs[0]['优惠金额'].astype(object)
    fs[0].loc[0, '优惠金额'] = '待核对'
    with pytest.raises(ValueError, match='优惠'):
        order_report(fs)


def test_percent_price_is_not_money():
    fs = orders()
    fs[0]['单价'] = fs[0]['单价'].astype(object)
    fs[0].loc[0, '单价'] = '10%'
    result = order_report(fs)
    assert 2 in set(result.outputs['异常数据']['源行号'])


def test_same_receipt_multiple_skus_are_not_duplicates():
    fs = _inventory_tables()
    fs[2].loc[1, ['商品编码', '入库数量', '采购价']] = ['P002', 5, 20]
    result = build_inventory_management_report(fs, source_names=INV)
    assert result.outputs['商品库存分析'].set_index('商品编码').loc['P002', '已入库数量'] == 5


def test_cancelled_inventory_events_never_win_on_success_substring():
    fs = _inventory_tables()
    fs[2].loc[0, '状态'] = fs[2].loc[1, '状态'] = '已取消（原已入库）'
    fs[3].loc[0, '状态'] = '已取消（原已完成）'
    fs[4].loc[0, '状态'] = '已确认后撤销'
    detail = build_inventory_management_report(fs, source_names=INV).outputs['商品库存分析'].set_index('商品编码')
    assert detail.loc['P001', '已入库数量'] == 0
    assert detail.loc['P001', '已出库数量'] == 10
    assert detail.loc['P001', '已确认调整'] == 0


def test_future_movement_does_not_change_as_of_stock():
    fs = _inventory_tables()
    new = fs[3].iloc[[0]].copy()
    new.loc[:, '出库日期'], new.loc[:, '出库单号'], new.loc[:, '数量'] = '2026/08/25', 'FUTURE', 100
    fs[3] = pd.concat([fs[3], new], ignore_index=True)
    detail = build_inventory_management_report(fs, source_names=INV).outputs['商品库存分析'].set_index('商品编码')
    assert detail.loc['P001', '当前账面库存'] == 4


@pytest.mark.parametrize('change', ['missing_opening', 'unknown_unit'])
def test_unknown_stock_evidence_cannot_generate_replenishment(change):
    fs = _inventory_tables()
    code = 'P002'
    if change == 'missing_opening':
        fs[1].loc[1, '期初库存'] = float('nan')
    else:
        code = 'P001'
        fs[0]['计量单位'] = '件'
        fs[2]['入库数量'] = fs[2]['入库数量'].astype(object)
        fs[2].loc[[0, 1], '入库数量'] = '1箱'
    result = build_inventory_management_report(fs, source_names=INV)
    row = result.outputs['商品库存分析'].set_index('商品编码').loc[code]
    assert pd.isna(row['当前账面库存'])
    assert code not in set(result.outputs['补货建议']['商品编码'])


def test_missing_attendance_is_not_absence_or_turnover_evidence():
    fs = _hr_tables()
    fs[1] = fs[1].loc[fs[1]['员工编号'].ne('E001')].copy()
    row = build_hr_management_report(fs, source_names=HR).outputs['员工综合分析'].set_index('员工编号').loc['E001']
    assert pd.isna(row['缺勤天数']) and pd.isna(row['考勤得分'])
    assert row['离职风险代理等级'] != '高'


def test_missing_dish_cost_invalidates_complete_contribution():
    fs, names = _multifact_fixture()
    fs[1].loc[0, '标准食材成本'] = float('nan')
    result = build_restaurant_diagnosis_report(fs, source_names=names)
    overview = result.outputs['管理层诊断总览'].set_index('指标')
    assert pd.isna(overview.loc['标准食材成本', '结果']) or overview.loc['标准食材成本', '结果'] == '未提供'
    dishes = result.outputs['菜品盈利分析']
    assert dishes.loc[dishes['菜品编码'].eq('D1'), '退款后贡献'].isna().all()


def test_explicit_scope_is_executed_not_merely_echoed():
    raw = pd.DataFrame({'订单号': ['A', 'B', 'C', 'D'], '日期': pd.to_datetime(['2026-08-01', '2026-08-01', '2026-09-01', '2026-09-01']),
                        '地区': ['华东', '华西', '华东', '华西'], '销售额': [100., 200., 300., 400.], '币种': ['CNY'] * 4})
    result = build_adaptive_analysis_report([raw], source_names=['订单'],
        user_request='只分析2026年8月华东地区的销售额、排名和月度趋势，生成Excel报表。')
    overview = result.outputs['管理层通用总览'].set_index('指标')
    assert overview.loc['核心指标：销售额', '结果'] == 100
    assert set(result.outputs['时间趋势']['月份']) == {'2026-08'}
    assert len(raw) == 4


def test_preserve_writer_retains_literal_text_and_updates_native_table():
    book = Workbook()
    try:
        ws = book.active
        ws.append(['编号', '备注'])
        ws.append(['A', '原记录'])
        ws.add_table(Table(displayName='ExistingData', ref='A1:B2'))
        _write_frame(ws, pd.DataFrame({'编号': ['A', 'B'], '备注': ['=1+1', '文字']}), replace_existing=True)
        assert ws['B2'].data_type == 's'
        assert ws.tables['ExistingData'].ref == 'A1:B3'
    finally:
        book.close()


def ecommerce_frames():
    return [
        pd.DataFrame({'SKU': ['P1'], '商品名称': ['甲'], '标准单位成本': [10.], '安全库存': [1.]}),
        pd.DataFrame({'下单日期': ['2026-08-01'], '订单号': ['O1'], '渠道': ['天猫'], '客户ID': ['C1'], 'SKU': ['P1'], '商品名称': ['甲'], '数量': [1.], '商品原价': [100.], '优惠金额': [0.], '买家实付': [100.], '订单状态': ['已完成']}),
        pd.DataFrame({'申请日期': ['2026-09-01'], '售后单号': ['R1'], '原订单号': ['O1'], 'SKU': ['P1'], '退款金额': [20.], '售后状态': ['退款成功']}),
        pd.DataFrame({'结算月份': ['2026-09'], '渠道': ['天猫'], '订单结算基数': [80.], '平台佣金': [0.], '技术服务费': [0.], '物流/服务扣款': [0.], '退款冲减': [0.], '实际到账': [80.]}),
        pd.DataFrame({'月份': ['2026-09'], '渠道': ['天猫'], '广告花费': [5.], '归因成交额': [50.]}),
        pd.DataFrame({'月份': ['2026-09'], 'SKU': ['P1'], '可售库存': [3.], '库存金额': [30.]}),
    ]


def ecommerce_report(fs):
    return build_ecommerce_diagnosis_report(fs, source_names=['商品主数据', '订单明细', '售后退款', '平台结算', '广告投放', '月末库存'])


def test_ecommerce_calendar_includes_month_without_sales():
    report = ecommerce_report(ecommerce_frames())
    occurrence = report.outputs['利润驱动分析'].query('时间口径 == "发生月视角"').set_index('月份')
    assert set(occurrence.index) == {'2026-08', '2026-09'}
    assert occurrence.loc['2026-09', '退款金额'] == 20
    assert occurrence.loc['2026-09', '实际到账'] == 80
    assert occurrence.loc['2026-09', '广告费'] == 5


def test_conflicting_product_cost_has_no_row_order_override():
    fs = ecommerce_frames()
    other = fs[0].copy()
    other['标准单位成本'] = 99.
    fs[0] = pd.concat([fs[0], other], ignore_index=True)
    with pytest.raises(ValueError, match='冲突'):
        ecommerce_report(fs)


def test_ecommerce_refund_document_is_idempotent():
    fs = ecommerce_frames()
    fs[2] = pd.concat([fs[2], fs[2]], ignore_index=True)
    assert ecommerce_report(fs).report['refund_amount'] == 20
