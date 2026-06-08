# -*- coding: utf-8 -*-
"""
Word 验证脚本：读出 .docx 全部内容，写到 scripts/word_verify.txt
并检查：
1) 大纲 v2 全部 18 小节标题都在
2) 3 套授课设计的核心内容（主题/教学目标/教学流程）都在
3) 没有红线词（必涨/稳赚/无风险/具体股票推荐）
"""
import os
import sys
import re
from docx import Document

ROOT = r'D:\Learning\AI\面试\AI散户投资课程'
DOCX = os.path.join(ROOT, 'AI散户投资课程_完整交付.docx')
OUT = os.path.join(ROOT, 'scripts', 'word_verify.txt')

# 大纲 v2 必含的 18 小节标题
OUTLINE_SECTIONS = [
    '0.1 工具开箱',
    '0.2 第一次问 AI',
    '1.1 散户亏钱都栽在这 5 大坑',
    '1.2 AI 投资能帮你什么、不能帮你什么',
    '1.3 怎么用 AI 才不被忽悠',
    '2.1 选股助手',
    '2.2 阅读助手',
    '2.3 盯盘机器人',
    '3.1 怎么用 AI 给一只票做"多维度梳理"',
    '3.2 怎么用 AI 选一只适合你的基金',
    '3.3 港美股怎么玩',
    '4.1 写你的第一份 AI 投资流程',
    '4.2 拿 AI 跑回测',
    '4.3 仓位怎么算',
    '5.1 止损纪律',
    '5.2 为什么我一买就跌',
    '5.3 极端行情',
    '6.1 AI 交易复盘官',
    '6.2 月度复盘',
    '6.3 长期迭代',
]

# 3 套授课设计必含的核心关键词
LESSON_KEYWORDS = {
    'A_决策沙盘': [
        '10 分钟让 AI 给你 5 维红绿灯',
        '决策推演沙盘',
        '5 维推演',
        '红绿灯',
        '3 种典型场景的"学员画像适配"',
        '风险厌恶型',
        '均衡型',
        '进取型',
        '决策陪练',
        'AI 信号 → 人工复核',
    ],
    'C_投资人格': [
        '5 分钟测出你的"投资人格"',
        '趋势型',
        '价值型',
        '短线型',
        '抄作业型',
        '强制冷却',
        '工具组合',
        '人格会变',
        '不是心理测评',
    ],
    'D_主力逆向': [
        '从一只票回溯 3 年',
        '"主力建仓痕迹"逆向追踪',
        '资金流',
        '龙虎榜',
        '大宗交易',
        '公告联动',
        '建仓时间窗',
        '逆向追踪的 3 个局限',
        '看见过去',
    ],
}

# 红线词检查
RED_FLAG_PATTERNS = [
    (r'必涨', '绝对收益承诺'),
    (r'稳赚不赔', '绝对收益承诺'),
    (r'稳赚(?!.*营销)', '绝对收益承诺（稳赚）'),  # 排除"任何"年化 30%+ 稳赚"都是营销"这类
    (r'100%', '绝对化'),
    (r'无风险', '无风险'),
    (r'保证收益', '保证收益'),
]


def extract_docx_text(path):
    """完整提取 docx 文本（段落 + 表格 + 页眉页脚）"""
    doc = Document(path)
    parts = []

    # 正文段落
    for p in doc.paragraphs:
        parts.append(p.text)

    # 表格
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                parts.append(cell.text)

    # 页眉页脚
    for section in doc.sections:
        for p in section.header.paragraphs:
            parts.append(p.text)
        for p in section.footer.paragraphs:
            parts.append(p.text)

    return '\n'.join(parts)


def main():
    print(f'Reading: {DOCX}')
    assert os.path.exists(DOCX), f'.docx not found: {DOCX}'

    text = extract_docx_text(DOCX)
    print(f'Total chars: {len(text):,}')

    # 写全文
    with open(OUT, 'w', encoding='utf-8') as f:
        f.write('=' * 80 + '\n')
        f.write('AI 散户投资课程_完整交付.docx · 全文提取\n')
        f.write(f'文件：{DOCX}\n')
        f.write(f'字符数：{len(text):,}\n')
        f.write('=' * 80 + '\n\n')
        f.write(text)
        f.write('\n\n' + '=' * 80 + '\n')
        f.write('验证结果\n')
        f.write('=' * 80 + '\n\n')

        # 1) 大纲 v2 18 小节
        f.write('## [1] 大纲 v2 全部 18 小节检查\n\n')
        missing = []
        for sec in OUTLINE_SECTIONS:
            ok = sec in text
            mark = '✓' if ok else '✗'
            f.write(f'  {mark}  {sec}\n')
            if not ok:
                missing.append(sec)
        f.write(f'\n小结：{len(OUTLINE_SECTIONS) - len(missing)}/{len(OUTLINE_SECTIONS)} 通过\n\n')
        if missing:
            f.write(f'MISSING: {missing}\n\n')

        # 2) 3 套授课设计核心关键词
        f.write('## [2] 3 套授课设计核心内容检查\n\n')
        for lesson, kws in LESSON_KEYWORDS.items():
            f.write(f'### {lesson}\n')
            miss_kw = []
            for kw in kws:
                ok = kw in text
                mark = '✓' if ok else '✗'
                f.write(f'  {mark}  {kw}\n')
                if not ok:
                    miss_kw.append(kw)
            f.write(f'  小结：{len(kws) - len(miss_kw)}/{len(kws)} 通过\n\n')
            if miss_kw:
                f.write(f'  MISSING: {miss_kw}\n\n')

        # 3) 红线词
        f.write('## [3] 红线词检查\n\n')
        red_found = []
        for pat, desc in RED_FLAG_PATTERNS:
            matches = re.findall(pat, text)
            if matches:
                red_found.append((pat, desc, len(matches)))
                f.write(f'  ⚠️  {pat}（{desc}）：出现 {len(matches)} 次\n')
        if not red_found:
            f.write('  ✓  无红线词\n')

        # 4) 关键章节
        f.write('\n## [4] 关键章节锚点检查\n\n')
        anchors = [
            'AI 散户投资课程',
            '完整交付包',
            '作者：方逸之',
            '2026-06-05',
            '目  录',
            '一、课程定位与卖课话术',
            '二、课程大纲 v2',
            '三、面试授课设计 · 3 套候选',
            '四、附：交付物清单',
            'A 套｜决策沙盘',
            'C 套｜投资人格',
            'D 套｜主力逆向',
            'AI 散户投资课程 · 课程大纲 + 面试授课设计',  # 页眉
        ]
        for a in anchors:
            ok = a in text
            mark = '✓' if ok else '✗'
            f.write(f'  {mark}  {a}\n')

        # 5) 总结
        f.write('\n## [5] 总结\n\n')
        pass_count = sum(1 for sec in OUTLINE_SECTIONS if sec in text)
        lesson_total = sum(len(kws) for kws in LESSON_KEYWORDS.values())
        lesson_pass = sum(1 for kws in LESSON_KEYWORDS.values() for kw in kws if kw in text)
        f.write(f'大纲 v2 18 小节：{pass_count}/{len(OUTLINE_SECTIONS)}\n')
        f.write(f'3 套授课设计关键词：{lesson_pass}/{lesson_total}\n')
        f.write(f'红线词命中：{len(red_found)}（{red_found if red_found else "无"}）\n')

    print(f'Verify written: {OUT}')

    # 打印关键验证结果
    pass_count = sum(1 for sec in OUTLINE_SECTIONS if sec in text)
    lesson_total = sum(len(kws) for kws in LESSON_KEYWORDS.values())
    lesson_pass = sum(1 for kws in LESSON_KEYWORDS.values() for kw in kws if kw in text)
    print(f'大纲 v2 18 小节：{pass_count}/{len(OUTLINE_SECTIONS)} 通过')
    print(f'3 套授课设计关键词：{lesson_pass}/{lesson_total} 通过')
    if red_found:
        print(f'⚠️  红线词：{red_found}')
    else:
        print('✓  无红线词')


if __name__ == '__main__':
    main()
