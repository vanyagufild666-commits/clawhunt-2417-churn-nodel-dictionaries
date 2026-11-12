# S1-DD-030 · 流失原因与未成交原因字典 — README

> 附：`tasks/2399/` 内为工单 S1-DD-024（会员等级与权益档位字典）的同批交付（平台任务 2399 因 GitHub 创建速率限制暂挂于本仓库子目录）。

## 交付物

| 文件 | 说明 |
|------|------|
- `会员流失与未成交原因字典.xlsx` | 双 sheet：流失原因（8 项）、招商未成交原因（7 项）。表头：编码、名称、定义、是否需要填写补充说明、后续建议动作、排序值 |
| `dictionaries.md` | 同内容的 Markdown 版本 + 验收对照 + 术语表 + 对接点说明 |
| `dictionaries.json` | 结构化源数据（内容模型，唯一数据源），可再生成 xlsx |
| `make_xlsx.py` | 生成脚本（openpyxl），`python make_xlsx.py` 重新生成 xlsx |

## 运行与测试

```bash
pip install openpyxl
python make_xlsx.py            # 重新生成 xlsx
```

xlsx 结构校验：`documents` skill validator — PASS（ZIP 容器、Content_Types、workbook.xml、2 worksheets、XML well-formed、text markers）。

## 验收标准覆盖

1. 流失原因 8 项（≥8）✅
2. 未成交原因 7 项（≥7）✅
3. 每项含「是否需要填写补充说明」标志（是/否，依据见 dictionaries.md 第三节）✅
4. 每项含一条后续建议动作 ✅

## 数据声明

全部内容为虚构测试数据，不含任何真实个人信息或企业信息，交付物不出现客户名称。

## 对接点

纯静态数据交付，无接口依赖。建议落库 `dict_churn_reason` / `dict_nodel_reason`，字段映射见 dictionaries.md 第四节。

## 版本

- v1.0 (2026-11-12): 初稿
- v1.1 (2026-11-12): 内部评审修正 — 类别名去重（价格）、补充说明标志差异化（4 项改否）、归属流程修正（续约转会员运营团队）、去除真实品牌词、术语表、数据源说明
