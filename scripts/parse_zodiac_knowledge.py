#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
解析 生肖命理 DOCX，生成 SQL INSERT 腳本
策略：先提取全文，再用生肖標記切割大塊，再拆子分類
輸出：/home/adamtsai/projects/Ecanapi/sql/20260510_ZodiacKnowledge_data.sql
"""
import zipfile, re, os

# ─── 常數 ──────────────────────────────────────────────────────────────
BRANCH_TO_ZODIAC = {
    '子':'鼠','丑':'牛','寅':'虎','卯':'兔',
    '辰':'龍','巳':'蛇','午':'馬','未':'羊',
    '申':'猴','酉':'雞','戌':'狗','亥':'豬'
}
ZODIAC_TO_BRANCH = {v:k for k,v in BRANCH_TO_ZODIAC.items()}

MONTH_TO_BRANCH = {
    '正':'寅','一':'寅','二':'卯','三':'辰','四':'巳',
    '五':'午','六':'未','七':'申','八':'酉','九':'戌',
    '十':'亥','十一':'子','十二':'丑'
}

POS_KW = [
    '大吉大利','財運亨通','旺上加旺','喜上加喜','一帆風順','財源廣進',
    '貴人扶持','大展鴻圖','步步高升','滿堂吉慶','諸事順遂',
    '大吉','大利','貴人','吉星','順遂','吉祥','順利','幸福','昌盛','興旺',
    '喜慶','財旺','喜事','喜氣','得財','升遷','事業有成',
]
NEG_KW = [
    '喪門','白虎','弔客','天耗','歲破','劫煞','陰煞','病符','死符','大耗',
    '大凶','官非','破財','災難','損失','刑剋','凶星','多災','官司','牢獄',
    '破耗','損財','口舌是非','小人作梗','大破財','血光之災',
]

def fortune_level(text):
    p = sum(text.count(k) for k in POS_KW)
    n = sum(text.count(k) for k in NEG_KW)
    t = p + n
    if t == 0: return None
    r = p / t
    if r >= 0.70: return '大吉'
    elif r >= 0.55: return '吉'
    elif r >= 0.40: return '平'
    elif r >= 0.25: return '凶'
    else: return '大凶'

def sql_val(v):
    if v is None: return 'NULL'
    return "'" + str(v).replace("'","''") + "'"

def make_row(bb, bz, cat, sub, tb, tl, content, fl):
    content = re.sub(r'\s+', ' ', content).strip()
    if len(content) < 15: return None
    cols = 'birth_branch,birth_zodiac,category,subcategory,target_branch,target_label,content,fortune_level,uid'
    vals = ','.join([sql_val(x) for x in [bb, bz, cat, sub, tb, tl, content, fl]] + ['0'])
    return f'INSERT INTO "生肖命理庫" ({cols}) VALUES ({vals});'

# ─── DOCX 抽取（合併為整段文字，盡量保持段落）──────────────────────────
def extract_text(path):
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8')
    # 段落標記 <w:p> 之間換行，其他標記換空格
    text = re.sub(r'</w:p>', '\n', xml)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = re.sub(r'[ \t]+', ' ', text)
    # 清理行
    lines = []
    for l in text.split('\n'):
        l = l.strip()
        if l: lines.append(l)
    return '\n'.join(lines)

# ─── 切出 12 個生肖大塊 ──────────────────────────────────────────────────
# 原文異體字對照（文件實際使用字元 → 標準地支/生肖）
DOC_ZODIAC_NORM = {
    '難': ('酉', '雞'),  # 酉雞 異體字1
    '雖': ('酉', '雞'),  # 酉雞 異體字2
    '猪': ('亥', '豬'),  # 亥豬 簡體
    '鼠': ('子', '鼠'), '牛': ('丑', '牛'), '虎': ('寅', '虎'),
    '兔': ('卯', '兔'), '龍': ('辰', '龍'), '蛇': ('巳', '蛇'),
    '馬': ('午', '馬'), '羊': ('未', '羊'), '猴': ('申', '猴'),
    '雞': ('酉', '雞'), '狗': ('戌', '狗'), '豬': ('亥', '豬'),
}

def split_zodiac_blocks(text):
    """
    找每個生肖的起點：「代表肖X，」（X 含異體字）
    回傳 [(branch, zodiac, block_text), ...]
    """
    all_chars = '|'.join(DOC_ZODIAC_NORM.keys())
    pattern = re.compile(r'代表肖(' + all_chars + r')[，,]')
    matches = list(pattern.finditer(text))
    blocks = []
    for i, m in enumerate(matches):
        raw_char = m.group(1)
        branch, zodiac = DOC_ZODIAC_NORM.get(raw_char, ('', raw_char))
        start = m.start()
        end = matches[i+1].start() if i+1 < len(matches) else len(text)
        block = text[start:end]
        blocks.append((branch, zodiac, block))
        print(f'  [{branch}/{zodiac}(原:{raw_char})] block={len(block)}字')
    return blocks

# ─── 解析子分類 ─────────────────────────────────────────────────────────
# 用 { } 包住的標記來切段
SUBCAT_TAGS = [
    (r'名人堂',                 '名人堂'),
    (r'風水魚',                 '風水魚'),
    (r'健康運勢',               '健康'),
    (r'財運氣勢',               '財運'),
    (r'事業發展',               '事業'),
    (r'職業運勢',               '職業'),
    (r'工作最佳拍檔',           '工作拍檔'),
    (r'最佳合作夥伴|合作最佳夥伴', '合作夥伴'),
    (r'愛情生活',               '愛情'),
    (r'婚姻家庭',               '婚姻'),
    (r'生年運勢',               '生年'),
    (r'生月運勢|生月運势',       '生月'),
    (r'生時特性',               '生時特性'),
    (r'生時辰運勢|生時辰運势',   '生時辰'),
    (r'[鼠牛虎兔龍蛇馬羊猴雞狗豬]人取名', '取名'),
]

def find_sections(block):
    """
    在大塊文字中找各子分類的起止位置
    回傳 [(subcat_name, start, end), ...]
    """
    events = []
    # 守護神 / 本命錦囊 本身（起始）
    m = re.search(r'代表肖[鼠牛虎兔龍蛇馬羊猴雞狗豬]', block)
    if m:
        events.append(('守護神', m.start()))

    for pattern, name in SUBCAT_TAGS:
        for m in re.finditer(r'\{[^}]*(?:' + pattern + r')[^}]*\}|(?:' + pattern + r')', block):
            events.append((name, m.start()))

    # 大分類切換標記
    for m in re.finditer(r'精批[榮樂藥禰祿禰禄]|精批\S{1,2}禄', block):
        events.append(('__精批', m.start()))
    for m in re.finditer(r'詳批大[運筵_]', block):
        events.append(('__詳批', m.start()))
    for m in re.finditer(r'改名增運', block):
        events.append(('__改名', m.start()))
    for m in re.finditer(r'流\s*年\s*運\s*程', block):
        events.append(('__流年', m.start()))
    for m in re.finditer(r'流\s*月\s*運\s*程', block):
        events.append(('__流月', m.start()))

    # 排序
    events.sort(key=lambda x: x[1])

    # 找出第一個 __流年 或 __流月 的位置
    first_stream_pos = None
    for name, pos in events:
        if name in ('__流年', '__流月'):
            first_stream_pos = pos
            break

    # 去重，且排除在流年/流月之後出現的子分類標記（避免流年內文誤切）
    seen = set()
    unique = []
    for name, pos in events:
        if pos in seen:
            continue
        seen.add(pos)
        # 流年/流月 之後只保留 __流年 / __流月 / __改名 等大分類事件
        if first_stream_pos and pos > first_stream_pos:
            if name not in ('__流年', '__流月', '__改名', '__精批', '__詳批'):
                continue
        unique.append((name, pos))

    # 建立 (name, start, end) 列表
    result = []
    for i, (name, pos) in enumerate(unique):
        end = unique[i+1][1] if i+1 < len(unique) else len(block)
        result.append((name, pos, end))
    return result

def parse_block(branch, zodiac, block, rows):
    sections = find_sections(block)
    category = '本命特性'

    for sec_name, start, end in sections:
        content_text = block[start:end].strip()

        # 大分類切換
        if sec_name == '__精批':
            category = '精批榮祿'
            continue
        if sec_name == '__詳批':
            category = '詳批大運'
            continue
        if sec_name == '__改名':
            category = '改名增運'
            continue

        # 流年：細切「X逢Y年」條目（含異體字 難=雞/酉, 猪=豬/亥）
        if sec_name == '__流年':
            ZO_ALL = r'[鼠牛虎兔龍蛇馬羊猴雞難雖狗豬猪]'  # 含異體字難/雖/猪
            # 允許字間有空格（部分文件 "猪 逢 鼠 年" 格式），用捕獲群組取目標生肖
            pattern = re.compile(ZO_ALL + r'\s*逢\s*(' + ZO_ALL + r')\s*年')
            entries = list(pattern.finditer(content_text))
            for i, m in enumerate(entries):
                raw_target = m.group(1)  # 捕獲群組 = 逢後的目標生肖字
                norm = DOC_ZODIAC_NORM.get(raw_target, (ZODIAC_TO_BRANCH.get(raw_target,''), raw_target))
                tb = norm[0]
                tl = f'逢{norm[1]}年'
                entry_start = m.start()
                entry_end = entries[i+1].start() if i+1 < len(entries) else len(content_text)
                entry_text = content_text[entry_start:entry_end]
                fl = fortune_level(entry_text)
                r = make_row(branch, zodiac, '流年運程', None, tb, tl, entry_text, fl)
                if r: rows.append(r)
            continue

        # 流月：細切「逢X月」條目
        if sec_name == '__流月':
            pattern = re.compile(r'逢(正|十二|十一|十|一|二|三|四|五|六|七|八|九)月[：:]?')
            entries = list(pattern.finditer(content_text))
            for i, m in enumerate(entries):
                month_str = m.group(1)
                tb = MONTH_TO_BRANCH.get(month_str, '')
                tl = f'{month_str}月'
                entry_start = m.start()
                entry_end = entries[i+1].start() if i+1 < len(entries) else len(content_text)
                entry_text = content_text[entry_start:entry_end]
                fl = fortune_level(entry_text)
                r = make_row(branch, zodiac, '流月運程', None, tb, tl, entry_text, fl)
                if r: rows.append(r)
            continue

        # 普通子分類
        subcat = sec_name
        # category 根據 subcat 判斷
        if subcat in ('生年','生月','生時特性','生時辰'):
            cat = '精批榮祿'
        elif subcat == '取名':
            cat = '改名增運'
        elif subcat in ('守護神','名人堂','風水魚','健康','財運','事業','職業',
                        '工作拍檔','合作夥伴','愛情','婚姻'):
            cat = '本命特性'
        else:
            cat = category

        fl = fortune_level(content_text) if subcat in ('健康','財運','事業') else None
        r = make_row(branch, zodiac, cat, subcat, None, None, content_text, fl)
        if r: rows.append(r)

# ─── 主程式 ─────────────────────────────────────────────────────────────
FILES = [
    '/mnt/d/命理知識庫/生肖命理/子丑辰酉.docx',
    '/mnt/d/命理知識庫/生肖命理/卯巳戌申.docx',
    '/mnt/d/命理知識庫/生肖命理/寅午未亥.docx',
]
OUT = '/home/adamtsai/projects/Ecanapi/sql/20260510_ZodiacKnowledge_data.sql'

rows = []
for fpath in FILES:
    fname = os.path.basename(fpath)
    print(f'\n=== {fname} ===')
    text = extract_text(fpath)
    blocks = split_zodiac_blocks(text)
    for branch, zodiac, block in blocks:
        parse_block(branch, zodiac, block, rows)

print(f'\n=== 總計 {len(rows)} rows ===')

# 統計
from collections import Counter
cats = Counter()
subs = Counter()
zodiacs = Counter()
levels = Counter()
for r in rows:
    m = re.search(r"VALUES \('([^']+)','([^']+)','([^']+)',([^,]+),", r)
    if m:
        zodiacs[m.group(2)] += 1
        cats[m.group(3)] += 1
        sub_raw = m.group(4).strip("'NULL")
        subs[sub_raw] += 1
    m2 = re.search(r",'([大吉平凶]+)',0\);$", r)
    if m2: levels[m2.group(1)] += 1
    elif ',NULL,0);' in r: levels['NULL'] += 1

print('\n[生肖]:', dict(sorted(zodiacs.items())))
print('[分類]:', dict(sorted(cats.items())))
print('[子分類]:', dict(sorted(subs.items())))
print('[吉凶]:', dict(sorted(levels.items())))

# 寫 SQL
with open(OUT, 'w', encoding='utf-8') as f:
    f.write('-- 生肖命理庫資料\n-- 來源：D:\\命理知識庫\\生肖命理\\*.docx\n-- 生成：2026-05-10\n\n')
    f.write('BEGIN;\n\nDELETE FROM "生肖命理庫";\n\n')
    for r in rows:
        f.write(r + '\n')
    f.write('\nCOMMIT;\n')

print(f'\n輸出：{OUT}')
