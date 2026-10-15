"""固定随机种子的千行独立顺序验证，同时记录完整导入导出耗时。"""
from pathlib import Path
import json,random,time,tempfile,platform,statistics
from openpyxl import Workbook
from aligner.core import load_tables,choose,choose_reference,align
from aligner.export import export

def main():
    samples=[]
    with tempfile.TemporaryDirectory() as td:
        td=Path(td);ref=td/'ref.xlsx';src=td/'src.xlsx'
        keys=[f'PART-{i:05d}' for i in range(1000)]
        shuffled=keys[:];random.Random(20261015).shuffle(shuffled)
        for p,values in [(ref,[k+'-REV2' for k in keys]),(src,shuffled)]:
            w=Workbook();s=w.active;s.append(['零件编号','数量'])
            for k in values:s.append([k,1])
            w.save(p)
        for _ in range(5):
            start=time.perf_counter();s=choose(load_tables(src));r=choose_reference(load_tables(ref),s)
            result=align(r,s);export(s,result,td/'out.xlsx');elapsed=time.perf_counter()-start
            # 独立预期：按构造时已知的编号列表，而非生产算法计算。
            assert [d['key'] for d in result['rows']]==keys
            assert result['matched']==1000
            samples.append(round(elapsed,6))
    report={'rows':1000,'repeats':5,'end_to_end_seconds':samples,'median_seconds':statistics.median(samples),
            'known_expected_order_pass':True,'platform':platform.platform(),'python':platform.python_version(),
            'scope':'synthetic fixed-seed 1000 rows; includes import, recognition, match, sort, export and audit'}
    Path('validation/benchmark.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
    print(report)

if __name__=='__main__':main()
