import win32gui
import win32process
import win32con
import win32com.client
import time
import subprocess
import json
# 获取窗口句柄
def getWindowHandle(pid):
    def callback(hwnd, winList):
        winList.append(hwnd)
    winList = []
    win32gui.EnumWindows(callback, winList)
    for hwnd in winList:
        # 获取PID
        winTid, winPid = win32process.GetWindowThreadProcessId(hwnd)
        # 窗口强制置顶
        if pid == winPid:
            return hwnd
    return False
try:
    try:
        # 读取配置文件
        with open('./fxwl-config.json', 'r', encoding='utf-8') as f:
            config = json.load(f)
    except:
        # 创建配置
        config = {}
        print('请输入软件地址：')
        config['app'] = input()
        config['interval'] = 3
        with open('./fxwl-config.json', 'w') as f:
            f.write(json.dumps(config))
    # 创建句柄
    WshShell = win32com.client.Dispatch('WScript.Shell')
    # 打开软件
    mainExe = subprocess.Popen(config['app'])
    print('打开软件：' + config['app'])
    # 主进程
    while True:
        try:
            time.sleep(config['interval'])
        except:
            time.sleep(3)
        # 获取当前时间
        nowTime = time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(time.time()))
        # 获取顶部窗口
        foreExe = win32gui.GetForegroundWindow()
        # 获取PID
        foreTid, forePid = win32process.GetWindowThreadProcessId(foreExe)
        # 窗口句柄校准
        if mainExe.pid != forePid:
            # 获取主窗口句柄
            mainHandle = getWindowHandle(mainExe.pid)
        else:
            mainHandle = foreExe
        # 获取主窗口坐标
        left, top, right, bottom = win32gui.GetWindowRect(mainHandle)
        # 窗口强制置顶
        if mainExe.pid != forePid:
            print('窗口未置顶：' + nowTime)
            # 激活窗口
            win32gui.SetForegroundWindow(mainHandle)
            # 最小化窗口
            win32gui.ShowWindow(mainHandle, win32con.SW_SHOWMINIMIZED)
            # 还原窗口
            win32gui.ShowWindow(mainHandle, win32con.SW_RESTORE)
        elif left != 0 and top != 0:
            print('窗口未全屏：' + nowTime)
            WshShell.SendKeys('{esc}')
            time.sleep(0.05)
            WshShell.SendKeys('{f11}')
except Exception as e:
    print(e)
    print('程序将在5秒后自动关闭')
    time.sleep(5)