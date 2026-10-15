"""只从 Git 索引白名单打包，私有附件和本地申报书不进入公开 ZIP。"""
from pathlib import Path
import subprocess,zipfile,json,hashlib,shutil

def main():
    root=Path(__file__).resolve().parents[1]
    files=subprocess.check_output(['git','ls-files','-z'],cwd=root).decode('utf-8').split('\0')
    files=[f for f in files if f and (root/f).is_file()]
    forbidden=('.private/','delivery/','.env')
    assert not any(f.startswith(forbidden) for f in files)
    out=root/'delivery/序衡完整源码-v1.0.0.zip'
    with zipfile.ZipFile(out,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
        for name in files:z.write(root/name,'Xuheng/'+name)
        for p in sorted((root/'dist/Xuheng').rglob('*')):
            relative=p.relative_to(root/'dist/Xuheng')
            if p.is_file() and 'output' not in relative.parts:
                z.write(p,'Xuheng/Windows/'+relative.as_posix())
    shutil.copy2(root/'docs/项目申报书公开版.docx',root/'delivery/项目申报书公开版.docx')
    targets=[out,root/'delivery/序衡演示.mp4',root/'delivery/项目申报书公开版.docx']
    (root/'delivery/SHA256SUMS.txt').write_text('\n'.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name for p in targets)+'\n',encoding='utf-8')
    print(out,'bytes',out.stat().st_size)

if __name__=='__main__':main()
