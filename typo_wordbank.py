# -*- coding: utf-8 -*-
"""
typo_wordbank.py — 专利文档常见错别字/错误用词词库
每项格式: {"wrong": "错误写法", "suggestion": "正确写法"}
可持续手工扩充。
"""

WORDBANK = [
    # ── 附图/说明书 ──
    {"wrong": "附图说名",     "suggestion": "附图说明"},
    {"wrong": "说明书附名",   "suggestion": "说明书附图"},
    {"wrong": "说名书",       "suggestion": "说明书"},

    # ── 权利要求 ──
    {"wrong": "权力要求",     "suggestion": "权利要求"},
    {"wrong": "权力要求书",   "suggestion": "权利要求书"},
    {"wrong": "权利要球",     "suggestion": "权利要求"},

    # ── 实施例/实施方式 ──
    {"wrong": "实示例",       "suggestion": "实施例"},
    {"wrong": "实施列",       "suggestion": "实施例"},
    {"wrong": "实是例",       "suggestion": "实施例"},
    {"wrong": "具体实施列",   "suggestion": "具体实施例"},
    {"wrong": "具体实示方式", "suggestion": "具体实施方式"},
    {"wrong": "具体施实方式", "suggestion": "具体实施方式"},

    # ── 技术领域/背景 ──
    {"wrong": "技术领与",     "suggestion": "技术领域"},
    {"wrong": "北京技术",     "suggestion": "背景技术"},
    {"wrong": "背景几术",     "suggestion": "背景技术"},

    # ── 发明/实用新型 ──
    {"wrong": "发名内容",     "suggestion": "发明内容"},
    {"wrong": "发民内容",     "suggestion": "发明内容"},
    {"wrong": "实用型新",     "suggestion": "实用新型"},
    {"wrong": "实新用型",     "suggestion": "实用新型"},

    # ── 连接/固定类动词 ──
    {"wrong": "固顶",         "suggestion": "固定"},
    {"wrong": "联结",         "suggestion": "连接"},   # 注意：部分场景"联结"可接受，词库可按需删除
    {"wrong": "安转",         "suggestion": "安装"},
    {"wrong": "设直",         "suggestion": "设置"},
    {"wrong": "包扩",         "suggestion": "包括"},
    {"wrong": "包扣",         "suggestion": "包括"},

    # ── 常见形近字 ──
    {"wrong": "其中一端固顶", "suggestion": "其中一端固定"},
    {"wrong": "其特证在于",   "suggestion": "其特征在于"},
    {"wrong": "其特正在于",   "suggestion": "其特征在于"},
    {"wrong": "其特征子于",   "suggestion": "其特征在于"},
    {"wrong": "至少一个的",   "suggestion": "至少一个"},

    # ── 标点/符号类误用（文字层面）──
    {"wrong": "摘要说明",     "suggestion": "摘要"},

    # ── 数字/编号类 ──
    {"wrong": "第一实施列",   "suggestion": "第一实施例"},
    {"wrong": "第二实施列",   "suggestion": "第二实施例"},
    {"wrong": "第三实施列",   "suggestion": "第三实施例"},

    # ── 章节常见错字 ──
    {"wrong": "技术邻域",     "suggestion": "技术领域"},
    {"wrong": "技术领或",     "suggestion": "技术领域"},
    {"wrong": "背景计术",     "suggestion": "背景技术"},
    {"wrong": "发明容内",     "suggestion": "发明内容"},
    {"wrong": "发明内荣",     "suggestion": "发明内容"},
    {"wrong": "附图说眀",     "suggestion": "附图说明"},
    {"wrong": "附图说朋",     "suggestion": "附图说明"},
    {"wrong": "说眀书",       "suggestion": "说明书"},
    {"wrong": "说明书附眀",   "suggestion": "说明书附图"},
    {"wrong": "权力要求",     "suggestion": "权利要求"},
    {"wrong": "权利要术",     "suggestion": "权利要求"},
    {"wrong": "其待征在于",   "suggestion": "其特征在于"},
    {"wrong": "其特微在于",   "suggestion": "其特征在于"},

    # ── 实施 / 实例 类 ──
    {"wrong": "实拖例",       "suggestion": "实施例"},
    {"wrong": "实拖方式",     "suggestion": "实施方式"},
    {"wrong": "具体实拖方式", "suggestion": "具体实施方式"},
    {"wrong": "具体实施防式", "suggestion": "具体实施方式"},
    {"wrong": "实施例子",     "suggestion": "实施例"},

    # ── 结构 / 部件用词 ──
    {"wrong": "构件",         "suggestion": "构件"},   # 占位（同义保留）
    {"wrong": "件部",         "suggestion": "部件"},
    {"wrong": "组装件",       "suggestion": "组件"},
    {"wrong": "装配件",       "suggestion": "组件"},
    {"wrong": "联接件",       "suggestion": "连接件"},
    {"wrong": "固结",         "suggestion": "固定"},
    {"wrong": "固持",         "suggestion": "固定"},
    {"wrong": "枢接于",       "suggestion": "铰接于"},
    {"wrong": "鉸接",         "suggestion": "铰接"},
    {"wrong": "螺纹连结",     "suggestion": "螺纹连接"},
    {"wrong": "螺栓连结",     "suggestion": "螺栓连接"},
    {"wrong": "栓接",         "suggestion": "螺栓连接"},

    # ── 形近 / 同音错字 ──
    {"wrong": "园柱",         "suggestion": "圆柱"},
    {"wrong": "园周",         "suggestion": "圆周"},
    {"wrong": "园弧",         "suggestion": "圆弧"},
    {"wrong": "圆孤",         "suggestion": "圆弧"},
    {"wrong": "圆桩",         "suggestion": "圆柱"},
    {"wrong": "桩体",         "suggestion": "柱体"},
    {"wrong": "园锥",         "suggestion": "圆锥"},
    {"wrong": "椎形",         "suggestion": "锥形"},
    {"wrong": "锯齿装",       "suggestion": "锯齿状"},
    {"wrong": "形装",         "suggestion": "形状"},
    {"wrong": "园环",         "suggestion": "圆环"},
    {"wrong": "圆桶",         "suggestion": "圆筒"},

    # ── 描述 / 动词 ──
    {"wrong": "用与",         "suggestion": "用于"},
    {"wrong": "适合用与",     "suggestion": "适合用于"},
    {"wrong": "用以于",       "suggestion": "用于"},
    {"wrong": "通过过",       "suggestion": "通过"},
    {"wrong": "在通过",       "suggestion": "通过"},
    {"wrong": "配合于",       "suggestion": "配合"},
    {"wrong": "相互连结",     "suggestion": "相互连接"},
    {"wrong": "相互联结",     "suggestion": "相互连接"},
    {"wrong": "相联接",       "suggestion": "相连接"},
    {"wrong": "穿设过",       "suggestion": "穿过"},
    {"wrong": "穿装",         "suggestion": "穿设"},
    {"wrong": "贯设",         "suggestion": "贯穿"},
    {"wrong": "形成有有",     "suggestion": "形成有"},
    {"wrong": "设制",         "suggestion": "设置"},
    {"wrong": "设至",         "suggestion": "设置"},
    {"wrong": "设臵",         "suggestion": "设置"},
    {"wrong": "安排在",       "suggestion": "设置在"},
    {"wrong": "包刮",         "suggestion": "包括"},
    {"wrong": "包活",         "suggestion": "包括"},
    {"wrong": "包栝",         "suggestion": "包括"},
    {"wrong": "包含有",       "suggestion": "包括"},

    # ── 作用 / 效果 ──
    {"wrong": "起着到",       "suggestion": "起到"},
    {"wrong": "起到了",       "suggestion": "起到"},
    {"wrong": "实现了",       "suggestion": "实现"},
    {"wrong": "效果好",       "suggestion": "效果好"},  # 占位

    # ── 数量 / 程度 ──
    {"wrong": "至少有一个",   "suggestion": "至少一个"},
    {"wrong": "至少一以上",   "suggestion": "至少一个"},
    {"wrong": "若干个的",     "suggestion": "若干"},
    {"wrong": "多个个",       "suggestion": "多个"},
    {"wrong": "二个以上",     "suggestion": "两个以上"},
    {"wrong": "其它",         "suggestion": "其他"},

    # ── 标准 / 范围常见 ──
    {"wrong": "本实用新性",   "suggestion": "本实用新型"},
    {"wrong": "本发名",       "suggestion": "本发明"},
    {"wrong": "本实用型",     "suggestion": "本实用新型"},
    {"wrong": "本申清",       "suggestion": "本申请"},
    {"wrong": "本申晴",       "suggestion": "本申请"},
    {"wrong": "申清人",       "suggestion": "申请人"},

    # ── 物理 / 机械常用 ──
    {"wrong": "驱动机构",     "suggestion": "驱动机构"},  # 占位
    {"wrong": "驱动装制",     "suggestion": "驱动装置"},
    {"wrong": "驱动装至",     "suggestion": "驱动装置"},
    {"wrong": "传动装制",     "suggestion": "传动装置"},
    {"wrong": "联动",         "suggestion": "联动"},   # 占位
    {"wrong": "电机器",       "suggestion": "电机"},
    {"wrong": "马达机",       "suggestion": "电机"},
    {"wrong": "电源源",       "suggestion": "电源"},
    {"wrong": "传感气",       "suggestion": "传感器"},

    # ── 副词 / 连词错用 ──
    {"wrong": "另外的",       "suggestion": "另一个"},
    {"wrong": "可以为",       "suggestion": "可以是"},
    {"wrong": "因为而",       "suggestion": "因而"},
    {"wrong": "并且且",       "suggestion": "并且"},
    {"wrong": "或者者",       "suggestion": "或者"},
    {"wrong": "以及和",       "suggestion": "以及"},
]

# 移除占位（wrong == suggestion）的条目，避免无效规则
WORDBANK = [e for e in WORDBANK if e["wrong"] != e["suggestion"]]
