"""公开交付安全扫描：递归检查源码与 Office/ZIP 容器，不打印命中内容。"""
import argparse,json,re,zipfile,hashlib
from pathlib import Path

PATTERNS=[('API credential',re.compile(rb'(?:sk-[A-Za-z0-9_-]{16,}|gh[opusr]_[A-Za-z0-9]{20,})')),
          ('mainland phone',re.compile(rb'(?<!\d)1[3-9]\d{9}(?!\d)'))]
TRUSTED_VENDOR={}
REVIEWED=[]

def scan_bytes(data,label,issues):
    for kind,pattern in PATTERNS:
        if pattern.search(data):
            digest=hashlib.sha256(data).hexdigest()
            # Only an exact, separately verified vendor hash may suppress a phone
            # pattern. Credentials are always findings, including in dependencies.
            if kind=='mainland phone' and digest in TRUSTED_VENDOR:
                REVIEWED.append({'file':label,'sha256':digest,'reason':TRUSTED_VENDOR[digest]})
            else:issues.append({'file':label,'kind':kind})
    if data[:2]==b'PK':
        import io
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for name in z.namelist():
                    if not name.endswith('/'):
                        scan_bytes(z.read(name),label+'!'+name,issues)
        except zipfile.BadZipFile:pass

def main():
    p=argparse.ArgumentParser();p.add_argument('targets',nargs='+');p.add_argument('--report',default='validation/public_scan.json');args=p.parse_args()
    vendor_manifest=Path(__file__).with_name('vendor_scan_allowlist.json')
    if vendor_manifest.exists():
        for entry in json.loads(vendor_manifest.read_text(encoding='utf-8')):
            TRUSTED_VENDOR[entry['sha256']]=entry['reason']
    issues=[];files=[]
    for target in args.targets:
        path=Path(target)
        for f in ([path] if path.is_file() else sorted(path.rglob('*'))):
            if not f.is_file():continue
            scan_bytes(f.read_bytes(),str(f),issues);files.append(str(f))
    report={'checked_files':len(files),'issues':issues,'reviewed_vendor_patterns':REVIEWED,'scope':'API credential and phone patterns, recursive ZIP/OOXML; only exact verified vendor hashes suppress phone false positives; private identity checked separately before publishing'}
    Path(args.report).write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(report,ensure_ascii=False))
    if issues:raise SystemExit(1)

if __name__=='__main__':main()
