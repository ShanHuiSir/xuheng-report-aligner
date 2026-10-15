"""真实窗口录屏剪辑、中文神经配音和同步字幕。仅旁白文本发送到语音服务。"""
import asyncio,json,re,subprocess
from pathlib import Path
import edge_tts,imageio_ffmpeg

FF=imageio_ffmpeg.get_ffmpeg_exe()
ROOT=Path(__file__).resolve().parents[1]
RAW=ROOT/'.private/video'
DEMO=ROOT/'demo'
SECTIONS=[
 ('业务问题','日常拿到两份报表，编号列和行顺序经常对不上。这里是序衡，一个在本机处理的桌面工具。我们用明确标注的模拟数据，实际跑一遍。'),
 ('输入资料','先填入两份模拟表格，再点导入并识别。参考表用第一列，待排序表用第二列，程序已经自己找到了。下面看到的，是文件里原来的顺序。'),
 ('参考顺序','切到参考顺序。注意，排在最前面的是零零三，然后才是零零一。我们要遵循这张表的顺序，不是把编号简单从小到大排列。'),
 ('运行处理','现在点开始对齐。结果出来了，一共八行，七行匹配成功，一行没有匹配。零零三已经排到了前面。两条零零二仍然保持它们原来的先后关系。'),
 ('核对结果','往下看，黄色这一行对应两个参考位置，程序把候选都列出来，按最早的位置排序。橙色的零九九没有参考项，放在最后。原始行号也保留着，方便回去核对。'),
 ('导出成果','最后输入保存位置，点击导出成果。这里实际生成了新的表格，还有同名的审计文件。每一行从哪里来、排到哪里、匹配了谁，都能查到。原始文件保持不变。')]

def run(args):
    r=subprocess.run([FF,'-hide_banner','-y',*args],capture_output=True)
    if r.returncode:raise RuntimeError(r.stderr.decode('utf-8','replace')[-2500:])
    return r

def duration(p):
    r=subprocess.run([FF,'-i',str(p)],capture_output=True)
    m=re.search(rb'Duration: (\d+):(\d+):(\d+\.\d+)',r.stderr)
    return int(m[1])*3600+int(m[2])*60+float(m[3])

def stamp(s):
    ms=round(s*1000);h,ms=divmod(ms,3600000);m,ms=divmod(ms,60000);s,ms=divmod(ms,1000)
    return f'{h:02}:{m:02}:{s:02},{ms:03}'

async def main():
    events=json.loads((RAW/'events.json').read_text())
    bounds=[e['t'] for e in events]
    timeline=[];offset=0;srt=[];caption_index=0;transcript=[]
    for i,(title,text) in enumerate(SECTIONS):
        audio=RAW/f'voice{i}.mp3';entries=[]
        voice=edge_tts.Communicate(text,'zh-CN-XiaoxiaoNeural',rate=['-2%','+1%','-1%','+2%','-2%','-1%'][i],boundary='SentenceBoundary')
        with audio.open('wb') as f:
            async for chunk in voice.stream():
                if chunk['type']=='audio':f.write(chunk['data'])
                elif chunk['type'] in ('SentenceBoundary','WordBoundary'):entries.append(chunk)
        dur=duration(audio)+1.2
        start,end=bounds[i],bounds[i+1]
        # 只调整真实操作片段的速度；长于旁白的空等候可加速，结果停留可放慢。
        factor=dur/(end-start)
        chapter=RAW/f'chapter{i}.txt';chapter.write_text(f'序衡  /  {i+1:02d} {title}    ·    模拟数据  真实程序运行',encoding='utf-8')
        # 字体路径使用正斜线和转义冒号以适配 FFmpeg 滤镜。
        filt=f"setpts={factor:.8f}*(PTS-STARTPTS),fps=15,scale=1280:820,pad=1280:980:0:60:color=0x102f37,drawtext=fontfile='C\\:/Windows/Fonts/msyh.ttc':textfile=chapter{i}.txt:fontsize=24:fontcolor=white:x=24:y=18"
        r=subprocess.run([FF,'-hide_banner','-y','-ss',str(start),'-t',str(end-start),'-i','raw.mp4','-i',f'voice{i}.mp3',
                          '-vf',filt,'-af','apad','-t',str(dur),'-c:v','libx264','-preset','fast','-crf','20','-c:a','aac','-b:a','160k',f'clip{i}.mp4'],cwd=RAW,capture_output=True)
        if r.returncode:raise RuntimeError(r.stderr.decode('utf-8','replace')[-3000:])
        for entry in entries:
            a=offset+entry['offset']/10_000_000;b=a+entry['duration']/10_000_000
            caption_index+=1
            # 一句自然口语按可读行长分行，不把字幕堆成一段说明书。
            sentence=entry['text'];lines=[]
            while len(sentence)>30:
                cut=max(sentence.rfind(c,0,30)+1 for c in '，。；、')
                cut=cut if cut>9 else 26
                lines.append(sentence[:cut]);sentence=sentence[cut:]
            lines.append(sentence)
            srt.append(f'{caption_index}\n{stamp(a)} --> {stamp(b)}\n'+ '\n'.join(lines)+'\n')
        transcript.append(f'{i+1:02d} {title}\n{text}\n')
        timeline.append({'chapter':title,'source_start':start,'source_end':end,'output_start':round(offset,3),'output_duration':round(dur,3),'speed_factor':round(factor,4),'tts_voice':'zh-CN-XiaoxiaoNeural'})
        offset+=dur
    (RAW/'concat.txt').write_text('\n'.join(f"file 'clip{i}.mp4'" for i in range(6)),encoding='utf-8')
    run(['-f','concat','-safe','0','-i',str(RAW/'concat.txt'),'-c','copy',str(RAW/'joined.mp4')])
    (DEMO/'演示字幕.srt').write_text('\n'.join(srt),encoding='utf-8')
    (DEMO/'演示解说词.txt').write_text('\n'.join(transcript),encoding='utf-8')
    # 硬字幕便于任意播放器直接观看，同时保留可编辑的 SRT。
    sub=DEMO/'演示字幕.srt'
    import shutil
    shutil.copy2(sub,RAW/'captions.srt')
    r=subprocess.run([FF,'-hide_banner','-y','-i','joined.mp4','-vf',"subtitles=captions.srt:force_style='FontName=Microsoft YaHei,FontSize=19,PrimaryColour=&HFFFFFF,OutlineColour=&H372F10,BorderStyle=1,Outline=1,Shadow=0,MarginV=23'",'-c:v','libx264','-preset','fast','-crf','19','-c:a','copy','-movflags','+faststart',str(ROOT/'delivery/序衡演示.mp4')],cwd=RAW,capture_output=True)
    if r.returncode:raise RuntimeError(r.stderr.decode('utf-8','replace')[-3000:])
    report={'version':'1.0.0','source':'real Windows application screen capture, 15 fps; no fabricated UI or result frames',
            'duration_s':duration(ROOT/'delivery/序衡演示.mp4'),'chapters':timeline,
            'actual_result':{k:json.loads((DEMO/'本轮对齐成果.audit.json').read_text(encoding='utf-8'))[k] for k in ['total','matched','unmatched','ambiguous']},
            'audio':'Chinese neural voice; only authored narration sent to TTS; no drawing files or private data transmitted',
            'subtitles':'burned in and editable SRT; sentence timestamps from TTS metadata'}
    (DEMO/'视频实测记录.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf-8')
    # 最终使用指定像素坐标的 ASS 排版，避免播放器默认字幕缩放遮住表格。
    from burn_subtitles import main as burn
    burn()
    print(json.dumps(report,ensure_ascii=False,indent=2))

if __name__=='__main__':asyncio.run(main())
