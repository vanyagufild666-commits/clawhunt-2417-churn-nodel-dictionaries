"""Generate 2947 行业分类标签字典: 3 levels, ~108 labels aligned to GB/T 4754-2017, JSON + xlsx."""
import json

# ---- Level 1: sectors (our grouping; GB code = dominant 门类 letter) ----
SECTORS = [
    ('IND_TEC', '科技', None, 'M', '以研发、软件信息、互联网与专业技术为主营的企业'),
    ('IND_MFG', '先进制造', None, 'C', '以设备/整机/器件实物生产制造为主营的企业'),
    ('IND_ENT', '企业服务', None, 'L', '面向企业提供专业服务（法律/咨询/人力/物流等）的企业'),
    ('IND_CUL', '文化创意', None, 'R', '以内容生产、文艺展演、文博为主营的企业'),
    ('IND_MED', '医疗健康', None, 'Q', '以医疗、医药流通与健康服务为主营的企业'),
    ('IND_CON', '消费', None, 'F', '以消费品制造与零售、文旅休闲为主营的企业'),
    ('IND_FIN', '金融服务', None, 'J', '以非银行金融、资本服务与保险为主营的企业'),
]

# ---- Level 2: groups (aligned to GB 大类) ----
GROUPS = [
    ('IND_TEC_SW', '软件与信息技术', 'IND_TEC', 'I65', '软件开发、集成、运维与数据处理类服务'),
    ('IND_TEC_NET', '互联网服务', 'IND_TEC', 'I64', '依托互联网提供信息服务、平台与安全服务'),
    ('IND_TEC_RD', '研发与设计', 'IND_TEC', 'M73/M74', '研究与试验发展及工程技术设计服务'),
    ('IND_TEC_QC', '检验检测', 'IND_TEC', 'M74', '质检技术与环境监测检测服务'),
    ('IND_MFG_ELE', '电子信息制造', 'IND_MFG', 'C39', '计算机、通信、电子器件与智能设备制造'),
    ('IND_MFG_EQP', '通用与专用设备', 'IND_MFG', 'C34/C35', '通用机械与专用设备整机制造'),
    ('IND_MFG_AUTO', '汽车与高端装备', 'IND_MFG', 'C36/C37', '汽车零部件、改装车、船舶与航空航天装备'),
    ('IND_MFG_PWR', '电气与新能源装备', 'IND_MFG', 'C38', '输配电、电池与照明器具制造'),
    ('IND_MFG_PHARMA', '医药制造', 'IND_MFG', 'C27', '化学药、中药与生物药品制造'),
    ('IND_ENT_LAW', '法律财税咨询', 'IND_ENT', 'L72', '法律服务与管理咨询（含会计审计税务）'),
    ('IND_ENT_HR', '人力资源与会展', 'IND_ENT', 'L72', '人力资源服务与会议展览服务'),
    ('IND_ENT_ADS', '广告营销', 'IND_ENT', 'L72', '广告设计、发布与代理服务'),
    ('IND_ENT_LOG', '物流仓储', 'IND_ENT', 'G54/G56/G59', '道路货运、运输代理、装卸搬运与仓储'),
    ('IND_CUL_AV', '影视内容出版', 'IND_CUL', 'R85/R86', '影视节目制作、录音制作与出版'),
    ('IND_CUL_ART', '文艺展演', 'IND_CUL', 'R87', '文艺创作、演出场馆与群众文体活动'),
    ('IND_MED_SVC', '医疗与健康服务', 'IND_MED', 'Q83', '医院、基层医疗与其他健康服务'),
    ('IND_CON_FNB', '食品饮料', 'IND_CON', 'C14/C15', '焙烤、乳品、调味品、饮料与精制茶'),
    ('IND_CON_APP', '服装日用', 'IND_CON', 'C18', '服装与服饰制造'),
    ('IND_CON_RETAIL', '零售与电商', 'IND_CON', 'F52', '综合零售、专业零售与无店铺零售'),
    ('IND_CON_TOUR', '文旅休闲', 'IND_CON', 'L72/R88', '旅行社服务与健身休闲活动'),
    ('IND_FIN_BANK', '多元金融服务', 'IND_FIN', 'J66/J69', '非货币银行服务与金融信息服务'),
    ('IND_FIN_CAP', '资本市场', 'IND_FIN', 'J67/J69', '证券市场服务与金融信托管理'),
    ('IND_FIN_INS', '保险', 'IND_FIN', 'J68', '人身保险与财产保险服务'),
]

# ---- Level 3: leaves (aligned to GB 中类) ----
LEAVES = [
    # 科技-软件与信息技术 I65
    ('IND_TEC_SW_DEV', '软件开发', 'IND_TEC_SW', '651', '主营定制或通用软件研发（应用/系统/嵌入式）'),
    ('IND_TEC_SW_INT', '信息系统集成与物联网技术服务', 'IND_TEC_SW', '653', '主营系统集成、智能建筑弱电与物联网方案落地'),
    ('IND_TEC_SW_OM', '运行维护服务', 'IND_TEC_SW', '654', '主营信息系统与基础设施运维托管'),
    ('IND_TEC_SW_CONSULT', '信息技术咨询服务', 'IND_TEC_SW', '656', '主营 IT 规划、架构与技术咨询服务'),
    ('IND_TEC_SW_DATA', '互联网数据服务', 'IND_TEC_SW', '645', '主营数据中心、云计算与大数据处理服务'),
    # 科技-互联网 I64
    ('IND_TEC_NET_INFO', '互联网信息服务', 'IND_TEC_NET', '642', '主营新闻/搜索/社交等内容信息服务'),
    ('IND_TEC_NET_PLAT', '互联网平台', 'IND_TEC_NET', '643', '主营电商、出行、生产协同等互联网交易平台'),
    ('IND_TEC_NET_SEC', '互联网安全服务', 'IND_TEC_NET', '644', '主营网络安全产品与安全运营服务'),
    # 科技-研发与设计 M73/M74
    ('IND_TEC_RD_NAT', '自然科学研究和试验发展', 'IND_TEC_RD', '731', '主营物理/化学/材料等自然科学方向研发'),
    ('IND_TEC_RD_ENG', '工程和技术研究和试验发展', 'IND_TEC_RD', '732', '主营工程技术方向研究与中试'),
    ('IND_TEC_RD_MED', '医学研究和试验发展', 'IND_TEC_RD', '734', '主营医药与临床方向研究'),
    ('IND_TEC_RD_EDS', '工程技术与设计服务', 'IND_TEC_RD', '748', '主营工程设计与工业设计服务'),
    # 科技-检验检测 M74
    ('IND_TEC_QC_INSPECT', '质检技术服务', 'IND_TEC_QC', '745', '主营产品检验、检测与认证服务'),
    ('IND_TEC_QC_ENV', '环境与生态监测检测服务', 'IND_TEC_QC', '746', '主营环境质量与污染源监测'),    # 先进制造-电子信息 C39
    ('IND_MFG_ELE_COMP', '计算机制造', 'IND_MFG_ELE', '391', '主营整机与服务器等计算设备生产'),
    ('IND_MFG_ELE_SMART', '智能消费设备制造', 'IND_MFG_ELE', '396', '主营智能穿戴、智能家居等智能终端'),
    ('IND_MFG_ELE_DEVICE', '电子器件制造', 'IND_MFG_ELE', '397', '主营半导体分立器件、集成电路与显示器件'),
    ('IND_MFG_ELE_PARTS', '电子元件及电子专用材料制造', 'IND_MFG_ELE', '398', '主营 PCB、连接器与电子材料'),
    # 先进制造-通用与专用设备 C34/C35
    ('IND_MFG_EQP_METAL', '金属加工机械制造', 'IND_MFG_EQP', '342', '主营机床与金属成形设备'),
    ('IND_MFG_EQP_PUMP', '泵、阀门、压缩机及类似机械制造', 'IND_MFG_EQP', '344', '主营流体机械与通用部件'),
    ('IND_MFG_EQP_MEDDEV', '医疗仪器设备及器械制造', 'IND_MFG_EQP', '358', '主营诊断、治疗与康复器械（归制造，医疗场景）'),
    ('IND_MFG_EQP_ECO', '环保、邮政、社会公共服务专用设备制造', 'IND_MFG_EQP', '359', '主营水/气/固废治理与市政专用设备'),
    # 先进制造-汽车与高端装备 C36/C37
    ('IND_MFG_AUTO_PARTS', '汽车零部件及配件制造', 'IND_MFG_AUTO', '367', '主营汽车零部件与配件'),
    ('IND_MFG_AUTO_REFIT', '改装汽车制造', 'IND_MFG_AUTO', '363', '主营专用车与改装车'),
    ('IND_MFG_AUTO_AERO', '航空航天器及设备制造', 'IND_MFG_AUTO', '374', '主营航空航天器与配套设备'),
    ('IND_MFG_AUTO_SHIP', '船舶及相关装置制造', 'IND_MFG_AUTO', '373', '主营船舶总装与船用装置'),
    # 先进制造-电气与新能源 C38
    ('IND_MFG_PWR_GRID', '输配电及控制设备制造', 'IND_MFG_PWR', '382', '主营变压器、开关与配电控制设备'),
    ('IND_MFG_PWR_CELL', '电池制造', 'IND_MFG_PWR', '384', '主营动力/储能电池及电池材料'),
    ('IND_MFG_PWR_LIGHT', '照明器具制造', 'IND_MFG_PWR', '387', '主营照明灯具与光源'),
    # 先进制造-医药 C27
    ('IND_MFG_PHARMA_API', '化学药品原料药制造', 'IND_MFG_PHARMA', '271', '主营原料药与中间体'),
    ('IND_MFG_PHARMA_FORM', '化学药品制剂制造', 'IND_MFG_PHARMA', '272', '主营化学制剂成品药'),
    ('IND_MFG_PHARMA_TCM', '中药饮片加工', 'IND_MFG_PHARMA', '273', '主营中药饮片炮制加工'),
    ('IND_MFG_PHARMA_BIO', '生物药品制品制造', 'IND_MFG_PHARMA', '276', '主营疫苗、抗体等生物药'),
    ('IND_MFG_PHARMA_MAT', '卫生材料及医药用品制造', 'IND_MFG_PHARMA', '277', '主营医用敷料与卫生材料'),
    # 企业服务-法律财税咨询 L72
    ('IND_ENT_LAW_LEGAL', '法律服务', 'IND_ENT_LAW', '722', '主营法律咨询与知识产权代理服务'),
    ('IND_ENT_LAW_CONSULT', '咨询与调查服务（含财税审计）', 'IND_ENT_LAW', '723', '主营管理咨询、市场调查与会计审计税务服务'),
    # 企业服务-人力资源与会展 L72
    ('IND_ENT_HR_HR', '人力资源服务', 'IND_ENT_HR', '725', '主营招聘、派遣、灵活用工与人评服务'),
    ('IND_ENT_HR_MICE', '会议、展览及相关服务', 'IND_ENT_HR', '728', '主营会议展览策划与执行'),
    # 企业服务-广告营销 L72
    ('IND_ENT_ADS_ADS', '广告业', 'IND_ENT_ADS', '724', '主营广告设计、投放与代理'),
    ('IND_ENT_ADS_KOL', '文化活动服务（营销活动执行）', 'IND_ENT_ADS', '729', '主营品牌活动与线下营销活动执行'),
    # 企业服务-物流仓储 G54/G56/G59
    ('IND_ENT_LOG_TRUCK', '道路货物运输', 'IND_ENT_LOG', '543', '主营道路货运与同城配送'),
    ('IND_ENT_LOG_AGENT', '运输代理业', 'IND_ENT_LOG', '562', '主营货代与多式联运代理'),
    ('IND_ENT_LOG_HANDLE', '装卸搬运', 'IND_ENT_LOG', '591', '主营装卸搬运作业服务'),
    ('IND_ENT_LOG_WH', '通用仓储', 'IND_ENT_LOG', '592', '主营普通货物仓储与仓配一体'),
    # 文化创意-影视内容出版 R85/R86
    ('IND_CUL_AV_FILM', '电影和广播电视节目制作', 'IND_CUL_AV', '861', '主营影视节目策划与制作'),
    ('IND_CUL_AV_AUDIO', '录音制作', 'IND_CUL_AV', '864', '主营音频内容与录音制作'),
    ('IND_CUL_AV_PUB', '出版业', 'IND_CUL_AV', '852', '主营图书、期刊与数字出版'),
    # 文化创意-文艺展演 R87
    ('IND_CUL_ART_CREATE', '文艺创作与表演', 'IND_CUL_ART', '871', '主营文艺创作、剧目与演出'),
    ('IND_CUL_ART_VENUE', '艺术表演场馆', 'IND_CUL_ART', '872', '主营剧场与演出场馆运营'),
    ('IND_CUL_ART_PUBLIC', '群众文体活动', 'IND_CUL_ART', '876', '主营社区/园区群众文体活动组织'),
    # 医疗健康-医疗与健康服务 Q83
    ('IND_MED_SVC_HOSP', '医院', 'IND_MED_SVC', '831', '主营综合或专科医院运营'),
    ('IND_MED_SVC_PRIM', '基层医疗卫生服务', 'IND_MED_SVC', '832', '主营社区诊所与基层医疗'),
    ('IND_MED_SVC_OTHER', '其他卫生服务（康养护理）', 'IND_MED_SVC', '839', '主营康养、护理与健康管理机构'),
    # 医疗健康-医药流通 F525 (leaf挂 2 级方向节点)
    ('IND_MED_DIST_RX', '医药及医疗器材专门零售', 'IND_MED', '525', '主营药店与医疗器械零售'),
    # 消费-食品饮料 C14/C15
    ('IND_CON_FNB_BAKE', '焙烤食品制造', 'IND_CON_FNB', '141', '主营烘焙食品与糕点生产'),
    ('IND_CON_FNB_DAIRY', '乳制品制造', 'IND_CON_FNB', '144', '主营乳制品加工'),
    ('IND_CON_FNB_COND', '调味品、发酵制品制造', 'IND_CON_FNB', '146', '主营调味品与发酵食品'),
    ('IND_CON_FNB_BEV', '饮料制造', 'IND_CON_FNB', '152', '主营瓶装水、果汁与功能饮料'),
    ('IND_CON_FNB_TEA', '精制茶加工', 'IND_CON_FNB', '153', '主营茶叶精制与茶产品'),
    # 消费-服装日用 C18
    ('IND_CON_APP_WOVEN', '机织服装制造', 'IND_CON_APP', '181', '主营梭织服装设计与生产'),
    ('IND_CON_APP_ACC', '服饰制造', 'IND_CON_APP', '183', '主营配饰、帽类与服饰周边'),
    # 消费-零售与电商 F52
    ('IND_CON_RETAIL_GEN', '综合零售', 'IND_CON_RETAIL', '521', '主营百货、超市综合零售'),
    ('IND_CON_RETAIL_ELEC', '家用电器及电子产品专门零售', 'IND_CON_RETAIL', '527', '主营家电与数码零售'),
    ('IND_CON_RETAIL_ECOM', '无店铺及其他零售（含电商）', 'IND_CON_RETAIL', '529', '主营线上零售与直播带货'),
    # 消费-文旅休闲 L72/R88
    ('IND_CON_TOUR_TRAVEL', '旅行社及相关服务', 'IND_CON_TOUR', '726', '主营旅行服务与定制游'),
    ('IND_CON_TOUR_FIT', '健身休闲活动', 'IND_CON_TOUR', '883', '主营健身、瑜伽与休闲运动场馆'),
    # 金融服务-多元金融 J66/J69
    ('IND_FIN_BANK_NBFI', '非货币银行服务（融资租赁/小额贷款）', 'IND_FIN_BANK', '663', '主营融资租赁、商业保理与小贷'),
    ('IND_FIN_BANK_INFO', '金融信息服务', 'IND_FIN_BANK', '694', '主营金融数据与信息服务'),
    # 金融服务-资本市场 J67/J69
    ('IND_FIN_CAP_SEC', '证券市场服务', 'IND_FIN_CAP', '671', '主营证券服务与投资顾问'),
    ('IND_FIN_CAP_TRUST', '金融信托与管理服务', 'IND_FIN_CAP', '691', '主营信托与资产管理服务'),
    # 金融服务-保险 J68
    ('IND_FIN_INS_LIFE', '人身保险', 'IND_FIN_INS', '681', '主营寿险与健康险服务'),
    ('IND_FIN_INS_PROP', '财产保险', 'IND_FIN_INS', '682', '主营财产险与责任险服务'),
]

# ---- fictional example name pools (deterministic, unique) ----
PRE = ['云汇', '博衍', '星桥', '拓远', '安泰', '明澈', '劲风', '青芜', '朗拾', '珩光',
       '恒岳', '映真', '合致', '知行', '沄舟', '曦和', '简芯', '朴石', '枢衡', '元叙',
       '启澄', '焕新', '澜图', '北辰', '鹿鸣', '覃石', '岸屿', '岑越']
SUF = ['科技', '智造', '数联', '服务', '文化', '健康', '商贸', '新材', '工程', '设计',
       '生物', '食品', '咨询', '传媒', '器械', '环保', '服饰', '电气']
_seed_state = {'n': 0}


def examples(seed):
    # injective mapping: n -> (PRE[n % 28], SUF[(n // 28) % 18]), unique for n < 504
    a_n = _seed_state['n'] * 2
    b_n = _seed_state['n'] * 2 + 1
    _seed_state['n'] += 1
    a = PRE[a_n % len(PRE)] + SUF[(a_n // len(PRE)) % len(SUF)]
    b = PRE[b_n % len(PRE)] + SUF[(b_n // len(PRE)) % len(SUF)]
    return a, b


rows = []
json_out = []
for code, name, parent, gb, cal in SECTORS:
    a, b = examples(0)
    rows.append([code, name, parent or '—', '1', gb, cal, f'{a}（虚构）', f'{b}（虚构）'])
    json_out.append({'编码': code, '名称': name, '上级编码': parent, '层级': 1, '国标代码': gb,
                     '判定口径': cal, '虚构企业示例': [f'{a}（虚构）', f'{b}（虚构）']})
for code, name, parent, gb, cal in GROUPS:
    a, b = examples(0)
    rows.append([code, name, parent, '2', gb, cal, f'{a}（虚构）', f'{b}（虚构）'])
    json_out.append({'编码': code, '名称': name, '上级编码': parent, '层级': 2, '国标代码': gb,
                     '判定口径': cal, '虚构企业示例': [f'{a}（虚构）', f'{b}（虚构）']})
for code, name, parent, gb, cal in LEAVES:
    a, b = examples(0)
    rows.append([code, name, parent, '3', gb, cal,
                 f'{a}（虚构）', f'{b}（虚构）'])
    json_out.append({'编码': code, '名称': name, '上级编码': parent, '层级': 3, '国标代码': gb,
                     '判定口径': cal, '虚构企业示例': [f'{a}（虚构）', f'{b}（虚构）']})

content = {'sheets': [
    {
        'title': '行业标签字典',
        'section': '一、行业分类标签字典（3 级层级：方向 → 大类组 → 中类标签，共 %d 个标签，对齐 GB/T 4754-2017）' % len(rows),
        'headers': ['编码', '名称', '上级编码', '层级', '对应的国标代码', '判定口径（一句话）', '虚构企业示例 1', '虚构企业示例 2'],
        'widths': [26, 34, 18, 8, 14, 44, 22, 22],
        'rows': rows,
    },
    {
        'title': '使用与无法归类规则',
        'section': '二、使用规则与「无法归类」处理规则',
        'headers': ['主题', '内容'],
        'widths': [22, 94],
        'rows': [
            ['层级与选择', '标签层级不超过 3 级；企业打标只允许挂第 3 级（叶子）标签，1–5 个；按叶子自动继承上级方向用于统计，不重复挂父级。'],
            ['对齐说明', '国标代码为 GB/T 4754-2017 的门类/大类/中类代码；个别标签按会员业务场景归组（如医疗器械制造归「先进制造」而场景服务于「医疗健康」），国标代码保持原位，归组仅影响平台展示。'],
            ['无法归类处理', '按全部叶子判定口径都无法归入的企业：置「待归类」标记（企业档案字段，不入本字典标签集），由运营每月集中复盘——能归则归入最接近叶子，确属新业态则申请新增叶子（走字典版本化发布，编码不复用）；不强制归入不匹配标签。'],
            ['虚构示例', '全部企业示例均为虚构名称（统一带「（虚构）」标记），仅演示判定口径，不对应任何真实企业。'],
            ['维护规则', '标签编码发布后不可复用；调整判定口径需运营与技术双侧评审；字典按版本发布（企业打标记录打标时点版本）。']
        ],
    },
]}

json.dump(content, open('content.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
json.dump(json_out, open('行业分类标签字典.json', 'w', encoding='utf-8'), ensure_ascii=False, indent=1)

# validations
codes = [r[0] for r in rows]
assert len(codes) == len(set(codes)), 'dup codes'
cs = set(codes)
dang = [r[0] for r in rows if r[2] not in cs and r[3] != '1']
assert not dang, dang
assert 80 <= len(rows) <= 120, len(rows)
lv = max(int(r[3]) for r in rows)
assert lv <= 3
# fictional names all unique
names = []
for r in rows:
    names += [r[6], r[7]]
assert len(names) == len(set(names)), 'dup example names'
# GB codes present on every row
assert all(r[4] and r[4] != '' and r[4] != '—' for r in rows)
per_lv = {l: sum(1 for r in rows if r[3] == l) for l in '123'}
print('labels:', len(rows), '| per level:', per_lv, '| JSON:', len(json_out), '| all checks PASS')
