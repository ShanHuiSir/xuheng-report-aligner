"""可独立运行 72 小时耐久验证；短时实测不作为 72 小时证据。"""
import argparse,json,time,gc,tracemalloc
from pathlib import Path
from aligner.core import *
from aligner.export import export
import tempfile

def main():
    p=argparse.ArgumentParser();p.add_argument('--seconds',type=int,default=259200);args=p.parse_args()
    tracemalloc.start();start=time.monotonic();iterations=0;peak=0;checkpoints=[]
    with tempfile.TemporaryDirectory() as td:
        while time.monotonic()-start<args.seconds:
            source=choose(load_tables('samples/模拟待排序.xlsx'))
            ref=choose_reference(load_tables('samples/模拟参考.xlsx'),source)
            result=align(ref,source);export(source,result,Path(td)/'out.xlsx')
            assert (result['total'],result['matched'],result['unmatched'])==(8,7,1)
            iterations+=1
            if iterations%100==0:
                gc.collect();current,peak=tracemalloc.get_traced_memory()
                checkpoints.append({'iteration':iterations,'elapsed_s':round(time.monotonic()-start,2),'current_bytes':current,'peak_bytes':peak})
    report={'requested_seconds':args.seconds,'actual_seconds':round(time.monotonic()-start,2),'iterations':iterations,
            'errors':0,'checkpoints':checkpoints,'scope':'sample import-match-export loop; no 72h claim unless actual_seconds>=259200'}
    Path('validation/soak.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
    print({k:v for k,v in report.items() if k!='checkpoints'})

if __name__=='__main__':main()
