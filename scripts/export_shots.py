# -*- coding: utf-8 -*-
"""把指定 pptx 每页导出 PNG 供排版自查（压线/溢出/挡页脚）。
用法：python export_shots.py <pptx绝对路径> <输出目录>
两个参数都必须给。优先用桌面 PowerPoint；没有则自动尝试 WPS 演示（KWPP）。
COM 调用可能被拒绝（RPC_E_CALL_REJECTED，Office 忙碌时常态），内置指数重试。
都没有可用 COM → 输出 NO_COM，跳过截图自查，改用坐标人工核对（见 SKILL.md 降级阶梯）。
"""
import os
import sys
import time

if len(sys.argv) < 3:
    print('用法: python export_shots.py <pptx绝对路径> <输出目录>')
    sys.exit(1)
SRC = os.path.abspath(sys.argv[1])
DST = os.path.abspath(sys.argv[2])
os.makedirs(DST, exist_ok=True)

try:
    import pythoncom
    import win32com.client
except ImportError:
    print('NO_PYWIN32')
    sys.exit(1)

# 竞争场景（多 agent 并行批量备课共用一台 Office）可用环境变量加大等待窗口：
#   set EXPORT_SHOTS_WAIT=300 && python export_shots.py ...
RETRY_SECONDS = float(os.environ.get('EXPORT_SHOTS_WAIT', '60'))  # 单次 COM 调用累计重试上限
RETRY_DELAY = 1.5

# 可重试的 COM 瞬态错误：调用被拒绝 / 调用失败 / 服务器不可用 / 进程启动失败
BUSY_HRESULTS = (-2147418111, -2147417846, -2147023174, -2146959355)


def _busy(e):
    """判断是否为 COM 忙碌/瞬态故障（可重试）。"""
    if getattr(e, 'hresult', None) in BUSY_HRESULTS:
        return True
    msg = str(e)
    return ('-2147418111' in msg or '拒绝' in msg or 'rejected' in msg.lower()
            or 'call failed' in msg.lower())


def com_retry(fn, *args, **kw):
    """Office COM 忙碌时会拒绝调用（RPC_E_CALL_REJECTED），退避重试直到成功或超时。"""
    deadline = time.time() + RETRY_SECONDS
    delay = RETRY_DELAY
    while True:
        try:
            return fn(*args, **kw)
        except Exception as e:
            if not _busy(e) or time.time() > deadline:
                raise
            time.sleep(delay)
            delay = min(delay * 1.6, 8)


done = False
for progid, label in (('PowerPoint.Application', 'PowerPoint'), ('KWPP.Application', 'WPS')):
    try:
        def _open_app():
            # 静态派发优先：dynamic Dispatch 偶发把 Open 返回对象解析坏
            # （AttributeError: Open.SaveAs），EnsureDispatch 生成静态包装可避开
            try:
                return win32com.client.gencache.EnsureDispatch(progid)
            except Exception:
                return win32com.client.Dispatch(progid)
        app = com_retry(_open_app)
        deck = com_retry(lambda: app.Presentations.Open(SRC, ReadOnly=True, WithWindow=False))
        print('USING', label)
        try:
            com_retry(lambda: deck.SaveAs(DST, 18))  # ppSaveAsPNG（WPS 兼容此枚举值）
        except Exception as e:
            # 二级降级：SaveAs 被持续拒绝（如桌面端挂模态框）→ 逐页 Slide.Export
            print('SaveAs failed:', str(e)[:80], '-> 改用逐页 Slide.Export')
            os.makedirs(DST, exist_ok=True)
            n = com_retry(lambda: deck.Slides.Count)
            for i in range(1, n + 1):
                com_retry(lambda i=i: deck.Slides(i).Export(
                    os.path.join(DST, '幻灯片%d.PNG' % i), 'PNG', 1280, 720))
        com_retry(lambda: deck.Close())
        try:
            app.Quit()
        except Exception:
            pass
        done = True
        break
    except Exception as e:
        print(label, 'unavailable:', str(e)[:100])
        deck = None

if not done:
    print('NO_COM')
    sys.exit(2)
print('OK ->', DST)
for f in sorted(os.listdir(DST)):
    print(' ', f)
