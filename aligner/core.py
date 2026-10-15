"""自动列识别、单向包含匹配和稳定排序；不执行外部网络操作。"""
from dataclasses import dataclass, field
from pathlib import Path
from collections import Counter
import csv
import io
import re
import unicodedata
import zipfile
import posixpath
import xml.etree.ElementTree as ET

NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
HEADERS = ('图档编号', '图纸编号', '零件编号', '零件件号', '待排序项', 'component number', 'part number', '编号', '短文本', '物料描述')

# 核心层的数据契约
# 行号与列号始终从 1 开始，与 Excel 显示坐标一致。
# 缺失单元格与显式空字符串在匹配时都视为空编号。
# rows 存放原始读取值，规范化不会覆盖这些原值。
# 因此匹配可忽略大小写，导出仍能保留来源拼写。
# 一个 Table 只对应一个工作表，而不是整个工作簿。
# 工作簿中的其他工作表由导出层原样保留。
# col 为零表示尚未识别，header 为零表示无表头。
# 候选评分是启发式排序分数，不解释为概率或置信度。
# 参考工作表的选择依赖两张表，不能只比较单表分数。
# 核心不缓存全局文件对象，连续处理不同文件不会交叉污染。
# 所有函数均为本地计算，不承担网络调用或身份认证。
# 输入样本是数据，单元格中的自然语言不改变算法规则。
# 多个参考候选属于业务歧义，不能用首个结果掩盖其存在。
# 算法保留所有候选位置，同时给出确定且可复现的排序键。
# 对业务行的保留优先于生成看似完整的匹配覆盖率。

def norm(value):
    """只做可解释的字宽、大小写和首尾空白归一，不删除内部连字符。"""
    # 数值型旧 XLS 编号可能由 xlrd 读为 3411.0。
    # 仅去掉真正整数的 .0，不改变非整数数值。
    # 文本型前导零必须保留，例如 0861 与 861 不强制等同。
    if value is None:
        return ''
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    return unicodedata.normalize('NFKC', str(value)).strip().upper()

@dataclass
class Table:
    name: str
    sheet: str
    rows: dict
    path: str = ''
    part: str = ''
    col: int = 0
    header: int = 0
    score: float = 0
    candidates: list = field(default_factory=list)

    def value(self, row, col):
        return self.rows.get(row, {}).get(col, '')

    def records(self):
        """标题区不参与重排；非空数据行全部保留，包括空编号行。"""
        # 标题之前的行留在原位，避免移动合并的封面区。
        # 识别签批页脚时检查整行，兼容页脚标签在不同列。
        # 空编号但其他列有内容的业务行依然需要保留。
        # 完全空白行属于版式，不能算成未匹配业务项。
        # 页脚之后的格式行不参与数据数量和排序统计。
        result = []
        for r in sorted(self.rows):
            if r <= self.header:
                continue
            values = self.rows[r]
            if any(re.match(r'^(客户码|制表|审核|申请人|批准|采购部门)[:：]', str(v).strip()) for v in values.values()):
                break
            if any(norm(v) for v in values.values()):
                result.append(r)
        return result

def letters_to_col(s):
    result = 0
    for c in s:
        result = result * 26 + ord(c) - 64
    return result

def load_tables(path):
    """稀疏读取实际单元格，避免格式扩展到 XFD 导致巨量空列扫描。"""
    # 读取入口只接受题目要求的两类扩展名。
    # 不自动执行宏，不用 Excel 外部链接获取任何值。
    # XLSX 公式读取缓存值，导出层另行检查公式安全性。
    # 异常消息保留文件名，便于用户定位导入问题。
    path = Path(path)
    try:
        if path.suffix.lower() == '.xls':
            import xlrd
            book = xlrd.open_workbook(path, on_demand=True)
            tables = []
            for sheet in book.sheets():
                rows = {r + 1: {c + 1: sheet.cell_value(r, c) for c in range(sheet.ncols)
                                if sheet.cell_value(r, c) != ''} for r in range(sheet.nrows)}
                tables.append(Table(path.name, sheet.name, rows, str(path)))
            book.release_resources()
            return tables
        if path.suffix.lower() != '.xlsx':
            raise ValueError('请选择 .xls 或 .xlsx 文件')
        with zipfile.ZipFile(path) as z:
            # 先检查解压总量，避免损坏文件和异常压缩包耗尽内存。
            if sum(i.file_size for i in z.infolist()) > 256 * 1024 * 1024:
                raise ValueError('工作簿解压后超过 256MB，请拆分后处理')
            # sharedStrings 中的富文本由各段可见文字拼接。
            # 不能假定所有字符串都以普通 v 元素保存。
            # inlineStr 是另一种合法的 Excel 文本表示。
            # 关系文件给出的实际工作表路径优于猜测 sheet1.xml。
            strings = []
            if 'xl/sharedStrings.xml' in z.namelist():
                strings = [''.join(e.itertext()) for e in ET.fromstring(z.read('xl/sharedStrings.xml'))]
            rels = {e.attrib['Id']: e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))}
            tables = []
            for sheet in ET.fromstring(z.read('xl/workbook.xml')).find('s:sheets', NS):
                rid = sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
                target = rels[rid]
                part = target.lstrip('/') if target.startswith('/') else posixpath.normpath('xl/' + target)
                rows = {}
                for row in ET.fromstring(z.read(part)).findall('s:sheetData/s:row', NS):
                    values = {}
                    for cell in row:
                        addr = cell.attrib.get('r', '')
                        if not addr:
                            continue
                        c = letters_to_col(re.match(r'[A-Z]+', addr)[0])
                        v = cell.find('s:v', NS)
                        value = v.text if v is not None else ''
                        if cell.attrib.get('t') == 's':
                            value = strings[int(value)] if value else ''
                        elif cell.attrib.get('t') == 'inlineStr':
                            value = ''.join(cell.find('s:is', NS).itertext())
                        if value not in ('', None):
                            values[c] = value
                    rows[int(row.attrib['r'])] = values
                tables.append(Table(path.name, sheet.attrib['name'], rows, str(path), part))
            return tables
    except (OSError, zipfile.BadZipFile, ET.ParseError, KeyError, IndexError) as e:
        raise ValueError('无法读取工作簿，请确认文件完整且未加密：' + path.name) from e

def from_text(text, name='粘贴表格'):
    """支持 Excel 制表符、CSV 和单列文本，不执行公式。"""
    # Excel 区域通常使用制表符，优先于 CSV 逗号。
    # csv.reader 保留引号包围的逗号和换行单元格。
    # 文本不做 eval，等号开头的输入仍然是普通字符串。
    # 粘贴来源没有工作簿版式，导出层会建立普通 XLSX。
    if not text.strip():
        raise ValueError('请先粘贴表格内容')
    delim = '\t' if '\t' in text else ','
    rows = {r: {c: v for c, v in enumerate(row, 1)}
            for r, row in enumerate(csv.reader(io.StringIO(text), delimiter=delim), 1)}
    return Table(name, '粘贴数据', rows)

def detect(table):
    """候选评分综合表头、非空率、长度、唯一性和编号字符形态。"""
    # 只对实际出现的列评分，不依赖工作表声明的最大列数。
    # 表头搜索限制在前 40 个实际行，避免把数据当表头。
    # 字母和数字共存有助于区分编号与姓名、材料描述。
    # 唯一性不会单独决定结果，否则序号列会总是胜出。
    # 长度项抑制过长说明列，密度项抑制稀疏备注列。
    # 数字编号使用独立特征，兼容纯四位零件后缀。
    # 相同分数时使用确定的列号次序，避免每轮结果变化。
    cols = sorted({c for row in table.rows.values() for c in row})
    candidates = []
    for col in cols:
        headers = []
        for r in sorted(table.rows)[:40]:
            v = str(table.value(r, col)).strip().lower()
            # 短表头或多语言表头；排除包含编号词的长业务描述。
            if len(v) < 100 and any(h in v for h in HEADERS):
                headers.append(r)
        header = headers[0] if headers else 0
        vals = [norm(table.value(r, col)) for r in sorted(table.rows) if r > header]
        nonempty = [v for v in vals if v]
        if not nonempty:
            continue
        n = len(nonempty)
        unique = len(set(nonempty)) / n
        density = n / max(len(vals), 1)
        length = sum(3 <= len(v) <= 80 for v in nonempty) / n
        shape = sum(bool(re.search(r'[A-Z]', v) and re.search(r'\d', v)) for v in nonempty) / n
        numeric_id = sum(bool(re.fullmatch(r'\d{4,}', v)) for v in nonempty) / n
        score = 2 * unique + density + length + 2 * shape + numeric_id + (3 if header else 0)
        if density < .08:
            score -= 4
        candidates.append((score, col, header))
    if not candidates:
        raise ValueError('表格没有可识别的编号数据')
    table.candidates = sorted(candidates, reverse=True)
    table.score, table.col, table.header = table.candidates[0]
    return table

def choose(tables):
    """优先选择有明确编号表头且数据丰富的工作表。"""
    # 空页或只有模板边框的页不会阻止其他工作表参与选择。
    # 这里用于来源表的单表识别；参考表还需跨表核对。
    # 数据行数只作为分数相同情况下的有限次级依据。
    valid = []
    for t in tables:
        try:
            detect(t)
            valid.append(t)
        except ValueError:
            pass
    if not valid:
        raise ValueError('没有包含编号的有效工作表')
    return max(valid, key=lambda t: (t.score, min(len(t.records()), 100)))

def choose_reference(tables, source):
    """用跨表包含覆盖率消除多工作表的微小评分差异，表头评分作次级依据。"""
    # 先去重来源键，避免重复补料行放大工作表覆盖权重。
    # 覆盖率用于工作表选择，不作为官方匹配准确率。
    # 分数相近时，包含更多实际来源编号的页应当优先。
    # 仍保留表头分数，防止随意拼接的备注页覆盖主清单。
    # 全部候选都低覆盖时仍可处理，未匹配会明确显示。
    valid = []
    keys = {norm(source.value(r, source.col)) for r in source.records()} - {''}
    for t in tables:
        try:
            detect(t)
            values = [norm(t.value(r,t.col)) for r in t.records()]
            coverage = sum(any(k in v for v in values) for k in keys) / max(len(keys),1)
            valid.append((coverage * 10 + t.score,t))
        except ValueError:
            pass
    if not valid:
        raise ValueError('参考文件没有有效编号数据')
    return max(valid,key=lambda item:item[0])[1]

def align(ref, source, policy='tail'):
    """完全匹配优先，否则 source 编号必须是 reference 编号的子串。"""
    # policy 只允许两种已定义行为，拼写错误不能悄悄降级。
    # 完全匹配必须先于包含匹配，防止短编号命中额外修订项。
    # 只有来源编号是参考编号子串时才满足题目的方向要求。
    # 重复来源键共享候选列表，但每条业务行保留独立身份。
    # 空编号不允许按空子串规则命中全部参考行。
    # 输出审计的 rank 是参考表原始行号，便于人工检查。
    # 固定策略改变可排序槽位，不改变未匹配业务行位置。
    # 两种策略都不新增业务行，输出行数应与输入完全一致。
    if policy not in ('tail', 'fixed'):
        raise ValueError('未知未匹配策略')
    if not ref.col:
        detect(ref)
    if not source.col:
        detect(source)
    refs = [(r, norm(ref.value(r, ref.col))) for r in ref.records()]
    exact = {}
    for r, key in refs:
        if key:
            exact.setdefault(key, []).append(r)
    cache = {}
    details = []
    for row in source.records():
        key = norm(source.value(row, source.col))
        if key not in cache:
            hits = exact.get(key, []) or ([r for r, value in refs if key in value] if key else [])
            cache[key] = hits
        hits = cache[key]
        details.append({'original_row': row, 'key': key, 'reference_rows': hits,
                        'rank': hits[0] if hits else None,
                        'status': ('多候选' if len(hits) > 1 else '已匹配') if hits else '未匹配'})
    # 稳定排序：同一参考项保留来源顺序；一对多不复制业务行。
    matched = sorted((d for d in details if d['rank'] is not None), key=lambda d: d['rank'])
    unmatched = [d for d in details if d['rank'] is None]
    order = matched + unmatched
    if policy == 'fixed':
        it = iter(matched)
        order = [next(it) if d['rank'] is not None else d for d in details]
    for slot, d in zip(source.records(), order):
        d['output_row'] = slot
    return {'version': '1.0.0', 'policy': policy, 'reference_sheet': ref.sheet,
            'source_sheet': source.sheet, 'reference_column': ref.col, 'source_column': source.col,
            'reference_header': ref.header, 'source_header': source.header,
            'total': len(details), 'matched': len(matched), 'unmatched': len(unmatched),
            'ambiguous': sum(d['status'] == '多候选' for d in details), 'rows': order}
