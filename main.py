"""桌面与命令行共同调用同一套核心，不依赖浏览器或网络。"""
import argparse
import json
import time
from aligner.core import load_tables, choose, choose_reference, align
from aligner.export import export

# 两个启动方式只在交互层不同，算法与导出路径必须共享。
# 完整提供三个文件参数时进入命令行，便于自动化复现。
# 没有文件参数时才导入 GUI，核心测试不依赖图形桌面。
# 时间从文件读取前开始，覆盖识别、匹配、排序、导出和审计。
# 不把核心函数的微小耗时当成整个业务流程的处理时间。
# 命令行输出只报告本次真实运行的统计，没有样例预期数字。
# 参考工作表必须在来源列识别后选取，避免备件页误选。
# 终端输出不包含 API Key，因为应用本身没有模型调用路径。
# 桌面打包仍使用这个入口，避免两个产品版本实现不一致。
# 入口保持薄层，业务规则应在可测试的核心模块里修改。

def main():
    parser = argparse.ArgumentParser(description='序衡 本地报表智能排序对齐')
    parser.add_argument('--reference')
    parser.add_argument('--source')
    parser.add_argument('--output')
    parser.add_argument('--policy', choices=['tail', 'fixed'], default='tail')
    args = parser.parse_args()
    if args.reference and args.source and args.output:
        start = time.perf_counter()
        source = choose(load_tables(args.source))
        ref = choose_reference(load_tables(args.reference),source)
        result = align(ref, source, args.policy)
        export(source, result, args.output, ref.path)
        print(json.dumps({k: v for k, v in result.items() if k != 'rows'}, ensure_ascii=False))
        print(f'导入到导出耗时 {time.perf_counter() - start:.4f} 秒')
    else:
        from aligner.gui import run
        run()

if __name__ == '__main__':
    main()
