"""原子导出和本地审计。XLSX 只改选定工作表中的行片段。"""
import copy
import csv
import html
import json
import os
from pathlib import Path
import re
import tempfile
import zipfile
from datetime import datetime, timezone, timedelta
from .core import NS

# 导出层与匹配层的职责边界
# 排序层只决定行映射，导出层决定如何安全保存文件。
# 输入文件始终以只读方式打开，所有写入使用独立路径。
# 输出的主工作表不插入审计列，保持原有业务列结构。
# 审计 CSV 便于人工筛选，JSON 保留完整候选数组。
# 原始行号与输出行号都使用 Excel 一基坐标。
# 匹配数量不写成准确率，审计只记录实际处理事实。
# XLSX 的样式对象存放于独立部件，移动行时保留其引用。
# 原包其他部件直接复制，避免整个工作簿重建造成样式漂移。
# 字节保留不能保证任意复杂 Excel 对象的引用语义。
# 因此明确限制移动区域公式与跨行合并，避免静默损坏。
# 临时文件与目标放在同一目录，便于使用原子替换。
# 审计文件在工作簿成功保存后生成，不伪造导出成功状态。
# 日志只记录操作元信息，不记录整张表或完整图纸内容。
# 日志不会联网发送，也没有自动删除计时器。

def log_dir():
    """日志按月归档，不自动清理，满足至少 180 天保留。"""
    # 用户本地数据目录可写，不依赖程序目录写权限。
    # 非 Windows 测试环境回退到用户目录，便于核心测试。
    # 目录每次存在即可，不为每个文件创建独立日志目录。
    p = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'Xuheng' / 'logs'
    p.mkdir(parents=True, exist_ok=True)
    return p

def audit(event):
    # 使用明确的东八区时戳，避免主机默认时区造成歧义。
    # JSONL 每条事件占一行，可流式检索和按月归档。
    # 月文件只追加，满足不少于 180 天的保留策略。
    now = datetime.now(timezone(timedelta(hours=8)))
    event = {'time': now.isoformat(), **event}
    with (log_dir() / (now.strftime('%Y-%m') + '.jsonl')).open('a', encoding='utf-8') as f:
        f.write(json.dumps(event, ensure_ascii=False) + '\n')

def xlsx_reorder(table, mapping, output):
    """保持 ZIP 中其他部件逐字节不变，含原样式、列宽、图片和其他工作表。"""
    # 不通过加载保存整个 openpyxl 工作簿进行重建。
    # 原始命名空间和扩展元素尽可能留在原 XML 中。
    # 每行移动时将行级属性一起移动，包含自定义行高。
    # 单元格 r 属性同时更新，避免 XML 行号与坐标冲突。
    # 标题或页脚没有出现在映射中，因此保持原始位置。
    with zipfile.ZipFile(table.path) as z:
        xml = z.read(table.part).decode('utf-8')
        # 暂不移动跨数据行合并块，防止拆散业务结构；标题合并不受影响。
        import xml.etree.ElementTree as ET
        root = ET.fromstring(xml)
        for merge in root.findall('s:mergeCells/s:mergeCell', NS):
            a, b = merge.attrib['ref'].split(':') if ':' in merge.attrib['ref'] else (merge.attrib['ref'],) * 2
            ra, rb = int(re.search(r'\d+', a)[0]), int(re.search(r'\d+', b)[0])
            if ra != rb and any(r in mapping and mapping[r] != r for r in range(ra, rb + 1)):
                raise ValueError('数据区含跨行合并单元格，请先拆分合并；原文件未修改')
        fragments = list(re.finditer(r'<row\b[^>]*(?:/>|>.*?</row>)', xml, re.S))
        # 先生成完整的新行字典，再一次性替换 sheetData 片段。
        # 排序循环不在原 XML 中原地交换，避免行互相覆盖。
        # 行号最终升序序列化，符合 Excel 的常规读取预期。
        moved = {}
        for m in fragments:
            raw = m[0]
            old = int(re.search(r'\br="(\d+)"', raw)[1])
            new = mapping.get(old, old)
            if old != new:
                # 带坐标对象和共享公式需依赖完整 Excel 引擎；明确阻止不安全改写。
                if re.search(r'<f\b', raw):
                    raise ValueError('移动区域含公式，请将公式转为值后重试；原文件未修改')
                raw = re.sub(r'\br="(?:([A-Z]+))?\d+"', lambda x: 'r="' + (x[1] or '') + str(new) + '"', raw)
            moved[new] = raw
        if fragments:
            xml = xml[:fragments[0].start()] + ''.join(moved[r] for r in sorted(moved)) + xml[fragments[-1].end():]
        # 同行合并的地址随行移动；保留跨行标题合并。
        def merge_move(m):
            a, r1, b, r2 = m.group(1, 2, 3, 4)
            if r1 == r2:
                nr = mapping.get(int(r1), int(r1))
                return f'ref="{a}{nr}:{b}{nr}"'
            return m[0]
        xml = re.sub(r'ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', merge_move, xml)
        with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as dest:
            for info in z.infolist():
                dest.writestr(info, xml.encode('utf-8') if info.filename == table.part else z.read(info))

def xls_reorder(table, mapping, output):
    """保留 BIFF 单元格格式和列宽；拒绝公式与跨行合并的风险区域。"""
    # XLS 是 BIFF 容器，不能使用 XLSX 的 XML 重排方法。
    # xlrd formatting_info 保留样式索引和列宽等属性。
    # 旧格式公式表达式并非 xlrd 的可编辑公开接口。
    # 检查 BIFF 的公式记录后主动阻止不安全的值化输出。
    # 使用 xlutils 完整复制工作簿后，仅替换选定表的数据行。
    import xlrd
    from xlutils.copy import copy as copy_book
    book = xlrd.open_workbook(table.path, formatting_info=True)
    # xlrd 不暴露公式表达式，旧格式导出仅支持值型表格。
    from xlrd.compdoc import CompDoc
    cd = CompDoc(Path(table.path).read_bytes())
    stream = cd.get_named_stream('Workbook') or cd.get_named_stream('Book')
    import struct
    pos = 0
    while pos + 4 <= len(stream):
        opcode, size = struct.unpack_from('<HH', stream, pos)
        if opcode == 6:
            raise ValueError('XLS 含公式，需在 Excel 中另存为值型表格后再排序')
        pos += 4 + size
    idx = book.sheet_names().index(table.sheet)
    source = book.sheet_by_index(idx)
    for lo, hi, c1, c2 in source.merged_cells:
        if hi - lo > 1 and any(r in mapping and mapping[r] != r for r in range(lo + 1, hi + 1)):
            raise ValueError('XLS 数据区含跨行合并单元格，请先拆分')
    wb = copy_book(book)
    target = wb.get_sheet(idx)
    # xlutils 复制生成的样式序号与 cell_xf_index 不保证相同，读取复制行对象样式。
    snapshots = {r: copy.deepcopy(target.row(r - 1)._Row__cells) for r in mapping}
    for old, new in mapping.items():
        target.row(new - 1)._Row__cells = copy.deepcopy(snapshots[old])
        for c, cell in target.row(new - 1)._Row__cells.items():
            cell.rowx = new - 1
        target.row(new - 1).height = target.row(old - 1).height
    wb.save(str(output))

def export(table, result, destination, reference_path=None):
    """临时文件落盘成功后再替换目标；禁止覆盖输入。"""
    # 扩展名必须与实际格式一致，不能把 XLS 伪装成 XLSX。
    # 解析绝对路径比较，阻止相对路径绕过来源覆盖保护。
    # 目标父目录不存在时创建，适配 GUI 直接输入路径。
    # 临时文件句柄立即关闭，避免 Windows 文件锁阻止保存。
    dest = Path(destination)
    if table.path and dest.resolve() == Path(table.path).resolve():
        raise ValueError('输出位置不能覆盖原文件')
    if reference_path and dest.resolve() == Path(reference_path).resolve():
        raise ValueError('输出位置不能覆盖参考文件')
    expected = Path(table.path).suffix.lower() if table.path else '.xlsx'
    if dest.suffix.lower() != expected:
        raise ValueError('输出扩展名应为 ' + expected)
    dest.parent.mkdir(parents=True, exist_ok=True)
    fd, tmp = tempfile.mkstemp(suffix=expected, dir=dest.parent)
    os.close(fd)
    mapping = {d['original_row']: d['output_row'] for d in result['rows']}
    try:
        if table.path:
            (xlsx_reorder if expected == '.xlsx' else xls_reorder)(table, mapping, tmp)
        else:
            # 粘贴内容没有原始 Excel 样式，创建一个普通工作表。
            # 强制文本数据类型，阻止等号开头内容变成公式。
            # 匹配用规范化值，但保存仍使用来源表原始文本。
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = '排序结果'
            for r, values in table.rows.items():
                for c, value in values.items():
                    cell = ws.cell(mapping.get(r, r), c, str(value))
                    cell.data_type = 's'  # 防止粘贴内容被当作可执行公式。
            wb.save(tmp)
        os.replace(tmp, dest)
    finally:
        if os.path.exists(tmp):
            os.unlink(tmp)
    # 先输出机器可读审计，再输出可用 Excel 查看的人读审计。
    # UTF-8 BOM 使 Windows Excel 默认打开 CSV 时显示中文。
    # 候选数组转换为参考行列表，避免丢掉多候选细节。
    report = dest.with_suffix('.audit.json')
    report.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding='utf-8')
    csvpath = dest.with_suffix('.audit.csv')
    with csvpath.open('w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['原始行', '输出行', '编号', '状态', '全部参考行'])
        for d in result['rows']:
            # Excel 公式注入防护适用于审计 CSV。
            key = d['key']
            if key.startswith(('=', '+', '-', '@')):
                key = "'" + key
            writer.writerow([d['original_row'], d['output_row'], key, d['status'], ','.join(map(str, d['reference_rows']))])
    audit({'action': 'export', 'file': dest.name, **{k: result[k] for k in ('total', 'matched', 'unmatched', 'ambiguous')}})
    return dest
