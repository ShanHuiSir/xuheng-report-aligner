"""附件只在本机读取，报告只存汇总。无官方答案时不计算匹配准确率。"""
from pathlib import Path
import time,json,platform,hashlib,zipfile
from aligner.core import *
from aligner.export import export

def main():
    base=Path('.private/报表智能排序对齐算法/数据')
    output=Path('.private/attachment_outputs');output.mkdir(exist_ok=True)
    reports=[]
    for sourcepath in sorted((base/'待排序表格').glob('*.xlsx')):
        prefix=sourcepath.stem.removeprefix('待排序-')
        refpath=next(p for p in (base/'参考表格').iterdir() if p.name.lower().startswith(prefix.lower()))
        before=hashlib.sha256(sourcepath.read_bytes()).hexdigest()
        start=time.perf_counter();src=choose(load_tables(sourcepath));ref=choose_reference(load_tables(refpath),src)
        parsed=time.perf_counter();result=align(ref,src);aligned=time.perf_counter()
        dest=output/sourcepath.name;export(src,result,dest);end=time.perf_counter()
        # 独立暴力基准：不使用生产匹配缓存与排序逻辑。
        rr=ref.records();expected=[]
        for sr in src.records():
            key=norm(src.value(sr,src.col))
            exact=[r for r in rr if key and norm(ref.value(r,ref.col))==key]
            hits=exact or [r for r in rr if key and key in norm(ref.value(r,ref.col))]
            expected.append((hits[0] if hits else float('inf'),sr))
        expected=[r for _,r in sorted(expected)]
        assert [d['original_row'] for d in result['rows']]==expected
        loaded=load_tables(dest)[0]
        assert len(result['rows'])==len(src.records())
        for d in result['rows']:
            assert loaded.rows[d['output_row']]==src.rows[d['original_row']]
        assert hashlib.sha256(sourcepath.read_bytes()).hexdigest()==before
        with zipfile.ZipFile(sourcepath) as a,zipfile.ZipFile(dest) as b:
            assert all(a.read(n)==b.read(n) for n in a.namelist() if n!=src.part)
        reports.append({'case':len(reports)+1,'reference_column':ref.col,'reference_sheet':ref.sheet,
                        'source_column':src.col,'reference_rows':len(rr),
                        **{k:result[k] for k in ['total','matched','unmatched','ambiguous']},
                        'import_s':round(parsed-start,6),'align_s':round(aligned-parsed,6),
                        'end_to_end_s':round(end-start,6),'rule_oracle_order_pass':True,
                        'row_values_and_original_hash_pass':True,'other_xlsx_parts_unchanged':True})
    report={'environment':{'platform':platform.platform(),'python':platform.python_version()},
            'scope':'10 supplied input pairs; no independent official expected output; match counts are not accuracy',
            'cases':reports}
    Path('validation').mkdir(exist_ok=True)
    Path('validation/attachments_summary.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
