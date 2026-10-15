"""录制真实 Windows 操作。此工具仅采集画面，不生成应用结果。"""
import imageio_ffmpeg,subprocess,time,json
from pathlib import Path

def main():
    base=Path('.private/video');base.mkdir(parents=True,exist_ok=True)
    stop=base/'stop'
    if stop.exists():stop.unlink()
    log=(base/'capture.log').open('w',encoding='utf-8')
    # 标题录制限定应用窗口，避免录入其他桌面文件或私人聊天。
    p=subprocess.Popen([imageio_ffmpeg.get_ffmpeg_exe(),'-y','-f','gdigrab','-framerate','15',
                        '-draw_mouse','1','-i','title=序衡 | 本地报表智能排序对齐',
                        '-c:v','libx264','-preset','ultrafast','-crf','18','-pix_fmt','yuv420p',str(base/'raw.mp4')],
                        stdin=subprocess.PIPE,stdout=log,stderr=log)
    (base/'start.json').write_text(json.dumps({'epoch':time.time()}),encoding='utf-8')
    while not stop.exists() and p.poll() is None:time.sleep(.2)
    if p.poll() is None:p.communicate(b'q',timeout=30)
    print('capture exit',p.returncode)

if __name__=='__main__':main()
