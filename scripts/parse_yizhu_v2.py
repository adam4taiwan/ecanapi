#!/usr/bin/env python3
"""
Parse 一柱論命六十甲子.docx + 一柱論命點竅.docx -> SQL INSERT data
"""
import zipfile, re, json
from xml.etree import ElementTree as ET

STEMS = list("甲乙丙丁戊己庚辛壬癸")
BRANS = list("子丑寅卯辰巳午未申酉戌亥")
JIAZI = [STEMS[i % 10] + BRANS[i % 12] for i in range(60)]
JIAZI_SET = set(JIAZI)
BRAN_SET = set(BRANS)

# OCR correction map for 點竅.docx
OCR_FIX = {
    '甲貝': '甲寅',   # 寅 → 貝
    '丁已': '丁巳',   # 巳 → 已
    '辛己': '辛巳',   # 巳 → 己
    '壬貞': '壬寅',   # 寅 → 貞
}

def extract_text(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml = z.read("word/document.xml")
    root = ET.fromstring(xml)
    paras = []
    for p in root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
        texts = [r.text for r in p.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t") if r.text]
        line = "".join(texts).strip()
        if line:
            paras.append(line)
    return paras

DAY_RE = re.compile(r'([甲乙丙丁戊己庚辛壬癸][子丑寅卯辰巳午未申酉戌亥])日')

def is_xun_kongwang(line):
    return '旬' in line and '空亡' in line

def extract_monthly_dict(lines):
    """Extract XX月 entries from lines into a {branch: text} dict."""
    monthly = {}
    segment_re = re.compile(
        r'([子丑寅卯辰巳午未申酉戌亥]{1,4})月'
        r'[，。、：\'\u2018\u2019\s]*'
        r'((?:(?![子丑寅卯辰巳午未申酉戌亥]{1,4}月).)*)'
    )
    for line in lines:
        if is_xun_kongwang(line):
            continue
        for m in segment_re.finditer(line):
            branches_str = m.group(1)
            text = m.group(2).strip().rstrip('。，、').strip()
            if not text:
                continue
            for b in branches_str:
                if b in BRAN_SET and b not in monthly:
                    monthly[b] = text
    return monthly

def parse_jiazi_docx(docx_path):
    """Parse 六十甲子.docx using XX日 markers."""
    paras = extract_text(docx_path)
    print(f"六十甲子 paragraphs: {len(paras)}")

    sections = {}
    current_dp = None
    for line in paras:
        m = DAY_RE.search(line)
        if m and m.group(1) in JIAZI_SET:
            dp = m.group(1)
            current_dp = dp
            sections.setdefault(dp, [])
        elif current_dp is not None:
            sections[current_dp].append(line)

    print(f"Sections found: {len(sections)}")
    missing = [j for j in JIAZI if j not in sections]
    if missing:
        print(f"Missing: {missing}")

    result = {}
    for dp, lines in sections.items():
        personality_lines = [l for l in lines if not is_xun_kongwang(l)]
        result[dp] = {
            "personality": "\n".join(personality_lines).strip(),
            "monthly": extract_monthly_dict(lines)
        }
    return result

def normalize_diqiao_header(line):
    """Normalize a potential day-pillar header from 點竅.docx."""
    # Apply OCR fixes
    for ocr, fix in OCR_FIX.items():
        if ocr in line:
            line = line.replace(ocr, fix)

    # Strip trailing punctuation
    clean = line.rstrip('：:。，、.；;\'"\u2018\u2019\u201c\u201d\s')

    # Check if the cleaned version ends with a JIAZI
    if len(clean) >= 2:
        tail = clean[-2:]
        if tail in JIAZI_SET:
            return tail
    return None

def parse_diqiao_docx(docx_path):
    """Parse 點竅.docx using 2-char 干支 headers (with OCR correction)."""
    paras = extract_text(docx_path)
    print(f"點竅 paragraphs: {len(paras)}")

    STEM_HEADER_RE = re.compile(r'^[甲乙丙丁戊己庚辛壬癸][：:\s]')

    sections = {}
    current_dp = None

    for line in paras:
        # Skip stem-level headers like 甲：
        if STEM_HEADER_RE.match(line) and len(line) <= 5:
            current_dp = None
            continue

        # Try to detect a day-pillar header at the START of this line
        dp_from_header = normalize_diqiao_header(line[:6])  # check first 6 chars

        # Or at the END of a line (jiazi embedded at line end from previous entry)
        dp_at_end = None
        if len(line) > 2:
            tail_candidate = line[-2:]
            if normalize_diqiao_header(tail_candidate) == tail_candidate and tail_candidate in JIAZI_SET:
                dp_at_end = tail_candidate

        if dp_from_header and dp_from_header in JIAZI_SET and line.strip()[:2] in JIAZI_SET or \
           (dp_from_header and len(line.strip()) <= 4 and dp_from_header in JIAZI_SET):
            # Standalone header or short line
            current_dp = dp_from_header
            sections.setdefault(current_dp, [])
        elif dp_at_end and dp_at_end in JIAZI_SET:
            # The line ends with a jiazi header; preceding content is for current section
            prefix = line[:-2].strip().rstrip('。，、 ')
            if current_dp and prefix:
                sections.setdefault(current_dp, []).append(prefix)
            current_dp = dp_at_end
            sections.setdefault(current_dp, [])
        elif current_dp is not None:
            sections[current_dp].append(line)

    # Second pass: look for missed headers (OCR errors treated as 2-char lines)
    # Re-scan with broader detection
    sections2 = {}
    current_dp = None
    for line in paras:
        if STEM_HEADER_RE.match(line) and len(line) <= 5:
            current_dp = None
            continue
        stripped = line.strip().rstrip('：:。，、.；;\'"\u2018\u2019\u201c\u201d\s ')
        # Apply OCR fix to the whole line if short
        if len(stripped) <= 4:
            for ocr, fix in OCR_FIX.items():
                stripped = stripped.replace(ocr, fix)
            if stripped in JIAZI_SET:
                current_dp = stripped
                sections2.setdefault(current_dp, [])
                continue
        # Check end-of-line jiazi
        m_end = re.search(r'([甲乙丙丁戊己庚辛壬癸][子丑寅卯辰巳午未申酉戌亥])[.；;\s]*$', line)
        if m_end and m_end.group(1) in JIAZI_SET:
            prefix = line[:m_end.start()].strip().rstrip('。，、 ')
            if current_dp and prefix:
                sections2.setdefault(current_dp, []).append(prefix)
            current_dp = m_end.group(1)
            sections2.setdefault(current_dp, [])
        elif current_dp:
            sections2[current_dp].append(line)

    print(f"點竅 sections: {len(sections2)}")
    missing = [j for j in JIAZI if j not in sections2]
    print(f"Missing ({len(missing)}): {sorted(missing)}")

    return {dp: "\n".join(lines).strip() for dp, lines in sections2.items() if lines}

def generate_sql(jiazi_data, diqiao_data):
    lines = []
    lines.append("-- 一柱論命·六十甲子日柱定數 資料插入")
    lines.append("-- 自動產生，請在生產環境 NeonDB 手動執行")
    lines.append("")

    def esc(s):
        if not s:
            return "NULL"
        return "'" + s.replace("'", "''") + "'"

    count = 0
    for dp in JIAZI:
        jd = jiazi_data.get(dp, {})
        dq = diqiao_data.get(dp, "")
        personality = jd.get("personality", "")
        monthly = jd.get("monthly", {})
        monthly_json = json.dumps(monthly, ensure_ascii=False) if monthly else ""

        if not (personality or monthly or dq):
            continue

        count += 1
        lines.append(f"INSERT INTO \"YiZhuLunMings\" (\"DayPillar\", \"Personality\", \"Poem\", \"MonthlyAnalysis\", \"VoidAnalysis\")")
        lines.append(f"VALUES (")
        lines.append(f"  {esc(dp)},")
        lines.append(f"  {esc(personality)},")
        lines.append(f"  NULL,")
        lines.append(f"  {esc(monthly_json)},")
        lines.append(f"  {esc(dq)}")
        lines.append(f") ON CONFLICT (\"DayPillar\") DO UPDATE SET")
        lines.append(f"  \"Personality\" = EXCLUDED.\"Personality\",")
        lines.append(f"  \"Poem\" = EXCLUDED.\"Poem\",")
        lines.append(f"  \"MonthlyAnalysis\" = EXCLUDED.\"MonthlyAnalysis\",")
        lines.append(f"  \"VoidAnalysis\" = EXCLUDED.\"VoidAnalysis\";")
        lines.append("")

    print(f"\nSQL entries: {count}")
    return "\n".join(lines)

if __name__ == "__main__":
    jiazi_path = "/mnt/d/命理知識庫/千里/一柱論命六十甲子.docx"
    diqiao_path = "/mnt/d/命理知識庫/千里/一柱論命點竅.docx"
    out_sql = "/home/adamtsai/projects/Ecanapi/sql/20260616_YiZhuLunMing_Data.sql"

    print("=== Parsing 六十甲子.docx ===")
    jiazi_data = parse_jiazi_docx(jiazi_path)

    print("\n=== Parsing 點竅.docx ===")
    diqiao_data = parse_diqiao_docx(diqiao_path)

    print("\n=== Generating SQL ===")
    sql = generate_sql(jiazi_data, diqiao_data)

    with open(out_sql, "w", encoding="utf-8") as f:
        f.write(sql)
    print(f"SQL written to {out_sql}")

    # Coverage stats
    total_monthly = sum(len(d.get('monthly', {})) for d in jiazi_data.values())
    no_monthly = [dp for dp in JIAZI if not jiazi_data.get(dp, {}).get('monthly')]
    dq_count = sum(1 for dp in JIAZI if diqiao_data.get(dp))
    print(f"\nStats:")
    print(f"  六十甲子: 60/60 sections")
    print(f"  Monthly avg entries: {total_monthly/60:.1f}")
    print(f"  Pillars without monthly: {no_monthly}")
    print(f"  點竅 coverage: {dq_count}/60")
