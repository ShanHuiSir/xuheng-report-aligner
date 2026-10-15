"""以视频像素为基准排字幕，全部字幕放入底部留白，不遮挡结果。"""
from pathlib import Path
import re,subprocess,json
import imageio_ffmpeg

def seconds(s):
    h,m,rest=s.split(':');return int(h)*3600+int(m)*60+float(rest.replace(',','.'))

def ass_stamp(t):
    cs=round(t*100);h,cs=divmod(cs,360000);m,cs=divmod(cs,6000);s,cs=divmod(cs,100)
    return f'{h}:{m:02}:{s:02}.{cs:02}'

def main():
    root=Path(__file__).resolve().parents[1];raw=root/'.private/video'
    text=(root/'demo/演示字幕.srt').read_text(encoding='utf-8')
    entries=[]
    for block in text.strip().split('\n\n'):
        lines=block.splitlines();a,b=lines[1].split(' --> ')
        entries.append([seconds(a),seconds(b),'\\N'.join(lines[2:])])
    for i in range(len(entries)-1):entries[i][1]=min(entries[i][1],entries[i+1][0]-.01)
    header='''[Script Info]
ScriptType: v4.00+
PlayResX: 1280
PlayResY: 980
WrapStyle: 2
[V4+ Styles]
Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding
Style: Default,Microsoft YaHei,24,&H00FFFFFF,&H00FFFFFF,&H00372F10,&H00372F10,0,0,0,0,100,100,0,0,1,1,0,2,25,25,22,1
[Events]
Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text
'''
    ass=header+'\n'.join(f'Dialogue: 0,{ass_stamp(a)},{ass_stamp(b)},Default,,0,0,0,,{t}' for a,b,t in entries)
    (raw/'captions.ass').write_text(ass,encoding='utf-8');(root/'demo/演示字幕.ass').write_text(ass,encoding='utf-8')
    r=subprocess.run([imageio_ffmpeg.get_ffmpeg_exe(),'-hide_banner','-y','-i','joined.mp4','-vf','ass=captions.ass','-c:v','libx264','-preset','fast','-crf','19','-c:a','copy','-movflags','+faststart',str(root/'delivery/序衡演示.mp4')],cwd=raw,capture_output=True)
    if r.returncode:raise RuntimeError(r.stderr.decode('utf-8','replace')[-2500:])
    print('Burned pixel-sized subtitles; no application text covered')

if __name__=='__main__':main()
