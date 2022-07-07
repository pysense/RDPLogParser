#SingleInstance Off ; 允许运行多个实例
#NoEnv              ; 避免检查空变量是否为环境变量
#KeyHistory 0       ; 禁用鼠标及键盘击键历史

; 更新环境变量
binPATH = %A_ScriptDir%;%A_ScriptDir%\bin;%A_ScriptDir%\..\NirSoftSuite\FullEventLogView_x64_sps;%A_ScriptDir%\..\..\Console\Log Parser 2.2
EnvGet, PATH, PATH
If PATH Not Contains %binPATH%
    EnvSet, PATH, %binPATH%;%PATH%

; 运行前检查：1）脚本需要管理员权限，如果不是则自动提升权限
full_command_line := DllCall("GetCommandLine", "str")
If Not (A_IsAdmin or RegExMatch(full_command_line, " /restart(?!\S)"))
{
    try
    {
        if A_IsCompiled
            Run *RunAs "%A_ScriptFullPath%" /restart
        else
            Run *RunAs "%A_AhkPath%" /restart "%A_ScriptFullPath%"
    }
    ExitApp
}

; 运行前检查：2）依赖文件是否存在
DepCommand = LogParser.exe|FullEventLogView.exe
Loop, Parse, DepCommand, |
    If Not CheckCommand(A_LoopField)
    {
        MsgBox, 缺少依赖文件 %A_LoopField%，程序退出。
        ExitApp
    }

; 程序设置
Author = 1057
AppName = RDPLogParser
Title_SettingTSPortNumber = 更新远程桌面端口
Title_TermServiceSettings = 查看远程桌面配置
Title_RegeditLocate = 注册表跳转
Title_EventLogAnalyzer = 事件日志查询

About =
(
> Powered by AutoHotkey.

最开始是为了分析远程登录日志，写了一个界面方便 LogParser 查询系统日志，所以暂时命名为 RDPLogParser 吧。后面慢慢加了一些用处不大的功能，随性更新。

—— by %Author%
)

; 定义热键
Gosub, SetHotkey

; 主界面
Gui, Add, Tab2, -Wrap w600, 远程桌面|日志分析|辅助工具|关于

Gui, Tab, 远程桌面
Gui, Add, Button, gTermServiceSettings vTermServiceSettings Section xm+10 ym+35, %Title_TermServiceSettings%
Gui, Add, Button, gSettingTSPortNumber vSettingTSPortNumber ys, %Title_SettingTSPortNumber%
Gui, Add, Button, gEnableTermService ys, 启用远程桌面
Gui, Add, Button, gRDPProtect vRDPProtect xs, 启动远程登录守护

Gui, Tab, 日志分析
Gui, Add, Button, gEventLogAnalyzer Section xm+10 ym+35, %Title_EventLogAnalyzer%
Gui, Add, Button, gDetectBruteForce vDetectBruteForce Section xs, 远程登录暴力破解检测

Gui, Tab, 辅助工具
Gui, Add, Button, gRegeditLocate Section xm+10 ym+35, 注册表跳转
Gui, Add, Button, gRunFullEventLogView xs, FullEventLogView

Gui, Tab, 关于
Gui, Add, Edit, Multi ReadOnly r13 w575, %About%

;~ Gui, +AlwaysOnTop +Resize
FileGetTime, ScriptModifyTime, %A_ScriptFullPath%
ToolsVersion := "Build" . SubStr(ScriptModifyTime, 1, 8)
Gui, Show,, %AppName% by %Author% %ToolsVersion%
Return

GuiEscape:
Return

GuiClose:
If DetectBruteForceStarted
{
    MsgBox, 4,, 远程登录守护仍在运行，是否确认退出？
    IfMsgBox, No
        Return
}
ExitApp
Return

; ------------------------------------------------------------------------------
;
; TermServiceSettings BEGIN
;
; ------------------------------------------------------------------------------
TermServiceSettings:
IfWinExist, %Title_TermServiceSettings%
{
    WinActivate
    Return
}
Gui, TermService:Add, ListView, vTermServiceListView Grid NoSort -Hdr -Multi r20 w700, 检查项|值
Gui, TermService:Default
GuiControl, TermService:-Redraw, TermServiceListView ; 在加载时禁用重绘来提升性能

LV_Add("", "远程（RDP）连接要求使用指定的安全层")
LV_Add("", "远程（RDP）要求使用 NLA 对远程连接的用户进行身份验证")
LV_Add("", "TermService 服务状态")
LV_Add("", "TermService 进程 PID")
LV_Add("", "TermService 监听端口")
LV_Add("", "TermService 配置端口")
LV_Add("", "TermService 允许连接")
LV_Add("", "日志文件 Microsoft-Windows-RemoteDesktopServices-RdpCoreTS/Operational 最大大小")
LV_Add("", "CredSSP 加密数据库修正策略（设置针对加密 Oracle 漏洞的防护级别）")

LV_ModifyCol(1, 550)
LV_ModifyCol(2, 140)
LV_ModifyCol(2, "Right")

Gosub, TermServiceRefresh
GuiControl, TermService:+Redraw, TermServiceListView ; 重新启用重绘
Gui, TermService:Show,, %Title_TermServiceSettings%
Return

; 刷新列表
TermServiceRefresh:
Gui, TermService:Default
Gui, +Disabled ; 更新数据时禁用窗口点击
LV_Modify(1, "Col2", getSecurityLayer())
LV_Modify(2, "Col2", getUserAuthentication())
LV_Modify(3, "Col2", queryTermServiceStatus())
LV_Modify(4, "Col2", queryTermServicePID())
LV_Modify(5, "Col2", queryTermServiceListeningPort())
LV_Modify(6, "Col2", getTermServicePortNumber())
LV_Modify(7, "Col2", getfDenyTSConnections())
LV_Modify(8, "Col2", getRdCoreTSLogMaxSize())
LV_Modify(9, "Col2", getAllowEncryptionOracle())
Gui, -Disabled
Gui, 1:Default
Return

; 右键菜单
TermServiceGuiContextMenu:
Try, Menu, TermServiceListViewMenu, Delete
Menu, TermServiceListViewMenu, Add, 刷新列表, TermServiceRefresh
Menu, TermServiceListViewMenu, Add
Menu, TermServiceListViewMenu, Add, %Title_SettingTSPortNumber%, SettingTSPortNumber
Menu, TermServiceListViewMenu, Add, 启用远程桌面, EnableTermService
Menu, TermServiceListViewMenu, Show
Return

TermServiceGuiClose:
Gui, TermService:Destroy
Gui, 1:Default
Return
; ------------------------------------------------------------------------------
;
; TermServiceSettings END
;
; ------------------------------------------------------------------------------
; ------------------------------------------------------------------------------
;
; SettingTSPortNumber BEGIN
;
; ------------------------------------------------------------------------------
SettingTSPortNumber:
IfWinExist, %Title_SettingTSPortNumber%
{
    WinActivate
    Return
}
GuiControl, Disable, SettingTSPortNumber

Gui, SettingTSPortNumber:Font, s11 c000080, Verdana
Gui, SettingTSPortNumber:Add, Text, Section w100, 注册表值
Gui, SettingTSPortNumber:Font
Gui, SettingTSPortNumber:Add, Edit, ys Number Limit5 w50 vPortNumber, % getTermServicePortNumber()
Gui, SettingTSPortNumber:Add, Button, ys gUpdatePortNumber, 更新
Gui, SettingTSPortNumber:Font, s11 c000080, Verdana
Gui, SettingTSPortNumber:Add, Text, Section xs w100, 重启服务
Gui, SettingTSPortNumber:Add, Checkbox, ys Checked1 visAddFWRule, 添加防火墙放行策略
Gui, SettingTSPortNumber:Add, Checkbox, ys Checked1 visFallbackTSPortNumber, 1 分钟内未确认则回退端口
Gui, SettingTSPortNumber:Font
Gui, SettingTSPortNumber:Add, Button, ys gRestartTermService vRestartTermService, 执行重启
Gui, SettingTSPortNumber:Show,, %Title_SettingTSPortNumber%

GuiControl, Enable, SettingTSPortNumber
Return

; 重建窗口
SettingTSPortNumberReload:
Gui, SettingTSPortNumber:Destroy
GoSub, SettingTSPortNumber
Return

SettingTSPortNumberGuiClose:
SettingTSPortNumberGuiEscape:
Gui, SettingTSPortNumber:Destroy
Return

UpdatePortNumber:
GuiControlGet, NewPortNumber,, PortNumber
; 判断端口是否正确
If NewPortNumber Not Between 1 And 65535
{
    MsgBox, 端口必须在 1 到 65535 之间
    Return
}
; 判断新端口是否与系统中其他监听端口冲突
bat =
(LTrim
@echo off
pushd `%~dp0
for /f "delims=: tokens=2" `%`%i in ('netstat -an ^| findstr /c:":%NewPortNumber% "') do ^
echo `%`%i | findstr "%NewPortNumber% " && del `%0
del `%0
)
If RunWaitBatOutput(bat)
{
    MsgBox, 4,, 当前系统已存在监听 %NewPortNumber% 的进程，是否继续更改为 %NewPortNumber% 端口？
    IfMsgBox, No
        Return
}
; 更新注册表值
RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp, PortNumber, % NewPortNumber
MsgBox, 已更新。
Return

; 重启远程桌面服务
RestartTermService:
RegRead, PortNumber, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp, PortNumber
ListeningPort := queryTermServiceListeningPort()
IfEqual, ListeningPort,, SetEnv, ListeningPort, 3389 ; 未获取到监听端口情况下，默认回退到 3389 端口
MsgBox, 305,, 即将重启 TermService 服务，请提前检查端口【%PortNumber%】是否放行以免丢失远程连接！`r`n如果选中【1 分钟内未确认则回退端口】，请在服务重启后重新登录远程桌面并确认，否则将回退到当前服务监听端口【%ListeningPort%】。
IfMsgBox, Cancel
    Return
FallbackFlag = ; 每次执行重启时都要重置下
GuiControl, Disable, RestartTermService
GuiControlGet, isAddFWRule,, isAddFWRule
GuiControlGet, isFallbackTSPortNumber,, isFallbackTSPortNumber
Gosub, AddFWAllowRDPPort
RestartTermService1:
ServicetoStop = UmRdpService|TermService ; 停止相关服务
Loop, Parse, ServicetoStop, |
{
    RunWait, sc stop %A_LoopField%,, Hide
    Loop
    {
        If A_Index = 10 ; 等待 10 秒后强制结束服务进程
        {
            RunWait, taskkill /FI "SERVICES eq %A_LoopField%" /F,, Hide
            Break
        }
        Sleep, 1000
        BatOutput := RunWaitBatOutput("@sc query " A_LoopField)
        If RegExMatch(BatOutput, "i)STATE.*STOPPED")
            Break
    }
}
SetTimer, FallbackTSPortNumber, -2000 ; 端口回退机制，计时器只运行一次
RunWait, net start TermService,, Hide
Gosub, SettingTSPortNumberReload
Return

FallbackTSPortNumber:
IfEqual, isFallbackTSPortNumber, 1
{
    IfEqual, FallbackFlag, 1 ; 回退操作仅执行一次，防止无法登录远程的情况下无限重启服务
        Return
    MsgBox, 262144,, 请点击【确定】按钮以确认正常登录远程桌面，否则远程端口将在 1 分钟后变更为：【%ListeningPort%】。, 60
    IfMsgBox, Timeout ; 超时未确认
    {
        RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp, PortNumber, % ListeningPort
        ; TODO 消除固定文本
        RunWait, cmd /d/c netsh.exe advfirewall firewall delete rule name="放行远程桌面端口：[%PortNumber%] %A_YYYY%-%A_MM%-%A_DD% -- by RDPLogParser",, Hide
        RunWait, cmd /d/c netsh.exe advfirewall firewall add rule name="放行远程桌面端口：[%ListeningPort%] %A_YYYY%-%A_MM%-%A_DD% -- by RDPLogParser" dir=in interface=any action=allow protocol=tcp localport=%ListeningPort% description="%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%",, Hide
        RunWait, cmd /d/c netsh.exe advfirewall firewall add rule name="放行远程桌面端口：[%ListeningPort%] %A_YYYY%-%A_MM%-%A_DD% -- by RDPLogParser" dir=in interface=any action=allow protocol=udp localport=%ListeningPort% description="%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%",, Hide
        FallbackFlag = 1
        Gosub, RestartTermService1
    }
}
Return

; 添加防火墙放行策略
AddFWAllowRDPPort:
GuiControlGet, PortNumber,, PortNumber
GuiControlGet, isAddFWRule,, isAddFWRule
IfEqual, isAddFWRule, 1
{
    ; 删除已有规则
    Loop, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Services\SharedAccess\Parameters\FirewallPolicy\FirewallRules
    {
        If A_LoopRegType = KEY
            value =
        Else
        {
            RegRead, value
            If ErrorLevel
                value = *error*
        }
        IfInString, value, 放行远程桌面端口：
        {
            ;~ RegDelete ; 似乎手动删除注册表不会生效
            FWRuleName := RegExReplace(value, "i).*\|Name=(放行.*RDPLogParser)\|.*", "$1") ; TODO 消除固定文本
            RunWait, cmd /d/c netsh.exe advfirewall firewall delete rule name="%FWRuleName%",, Hide
        }
    }
    ; 添加当前窗口获取的注册表对应的端口值
    RunWait, cmd /d/c netsh.exe advfirewall firewall add rule name="放行远程桌面端口：[%PortNumber%] %A_YYYY%-%A_MM%-%A_DD% -- by RDPLogParser" dir=in interface=any action=allow protocol=tcp localport=%PortNumber% description="%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%",, Hide
    RunWait, cmd /d/c netsh.exe advfirewall firewall add rule name="放行远程桌面端口：[%PortNumber%] %A_YYYY%-%A_MM%-%A_DD% -- by RDPLogParser" dir=in interface=any action=allow protocol=udp localport=%PortNumber% description="%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec%",, Hide
}
Return
; ------------------------------------------------------------------------------
;
; SettingTSPortNumber END
;
; ------------------------------------------------------------------------------
; ------------------------------------------------------------------------------
;
; EnableTermService BEGIN
;
; ------------------------------------------------------------------------------
; 启用远程桌面
EnableTermService:
MsgBox, 4, 启用远程桌面, 是否执行以下操作：`r`n`r`n1. 允许连接远程桌面服务`r`n2. 远程（RDP）要求使用 NLA 对远程连接的用户进行身份验证`r`n3. 启动远程桌面服务
IfMsgBox, Yes
{
    ; 允许连接远程桌面服务（开始监听端口）
    RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server, fDenyTSConnections, 0
    ; 远程（RDP）要求使用 NLA 对远程连接的用户进行身份验证
    RegWrite, REG_DWORD, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp, UserAuthentication, 1
    ; 启动远程桌面服务
    RunWait, net start TermService,, Hide
    MsgBox,, 启用远程桌面, 操作完成，请检查防火墙是否放行相应端口。
}
Return
; ------------------------------------------------------------------------------
;
; EnableTermService END
;
; ------------------------------------------------------------------------------
; ------------------------------------------------------------------------------
;
; RDPProtect BEGIN
;
; ------------------------------------------------------------------------------
RDPProtect:
GuiControl, Disable, RDPProtect

If DetectBruteForceStarted
{
    SetTimer, DetectBruteForceTimer, Off
    DetectBruteForceStarted =
    MsgBox, 已停止远程登录守护。
    GuiControl, Enable, RDPProtect
    GuiControl, Text, RDPProtect, 启动远程登录守护
    Return
}
DetectBruteForceCycle = 60 ; 检测周期 60s
SetTimer, DetectBruteForceTimer, % DetectBruteForceCycle * 1000
Gosub, DetectBruteForceTimer
MsgBox, 已启动远程登录守护。`r`n`r`n统计周期 %DetectBruteForceCycle% 秒内身份验证失败次数超过 5 次将自动添加防火墙拦截策略。

GuiControl, Enable, RDPProtect
GuiControl, Text, RDPProtect, 停止远程登录守护
Return

DetectBruteForceTimer:
DetectBruteForceStarted = 1
IfNotExist, RDPCheckpoint.lpc
{
    m = 统计周期【首次启动】
    n = 50 ; 失败记录超过此数值则定义为暴力破解 IP
}
Else
{
    m = 统计周期【%DetectBruteForceCycle%秒】
    n = 5
}
cmd =
(Join`s LTrim
LogParser.exe -q -stats:OFF -i:EVT -o:CSV -headers:OFF -iCheckpoint:RDPCheckpoint.lpc "
    SELECT
        EXTRACT_TOKEN(Strings, 19, '|') AS IpAddress
    USING
        COUNT(*) AS Total,
        EXTRACT_TOKEN(Strings, 10, '|') AS LogonType
    INTO %A_Temp%\LogParserDetectBruteForce.out
    FROM Security
    WHERE
        EventID = 4625 AND
        LogonType LIKE '3' AND
        IpAddress NOT LIKE '-'
    GROUP BY IpAddress
    HAVING
        Total > %n%
    "
)
RunWait, %cmd%,, Hide UseErrorLevel ; 未更改 WorkingDir，将在当前目录下生成 RDPCheckpoint.lpc 文件
IfNotEqual, ErrorLevel, 0, Return
Loop, Read, %A_Temp%\LogParserDetectBruteForce.out
{
    RunWait, cmd /d/c netsh.exe advfirewall firewall delete rule name="拦截远程桌面暴力破解 IP：[%A_LoopReadLine%] -- by RDPLogParser",, Hide
    RunWait, cmd /d/c netsh.exe advfirewall firewall add rule name="拦截远程桌面暴力破解 IP：[%A_LoopReadLine%] -- by RDPLogParser" dir=in interface=any action=block remoteip=%A_LoopReadLine% description="%A_YYYY%-%A_MM%-%A_DD% %A_Hour%:%A_Min%:%A_Sec% %m%，爆破次数大于【%n%】次",, Hide
}
FileDelete, %A_Temp%\LogParserDetectBruteForce.out
Return
; ------------------------------------------------------------------------------
;
; RDPProtect END
;
; ------------------------------------------------------------------------------
; ------------------------------------------------------------------------------
;
; EventLogAnalyzer BEGIN
;
; ------------------------------------------------------------------------------
EventLogAnalyzer:
IfWinExist, %Title_EventLogAnalyzer%
{
    WinActivate
    Return
}
; 运行前检查：3）日志服务是否启动
cmd = tasklist /FI "SERVICES eq EventLog" /NH
CmdOutput := RunWaitBatOutput("@"cmd)
If CmdOutput Not Contains svchost.exe
    MsgBox, 日志服务未启动，当前系统日志查询将无法正常执行。
Gui, EventLogAnalyzer:Font, s9 c000080
Gui, EventLogAnalyzer:Add, ListBox, vEventLogAnalyzer_LB_ConfigFiles gEventLogAnalyzer_LoadConfig w260 h520 HScroll400
Gui, EventLogAnalyzer:Font
Gui, EventLogAnalyzer:Add, Edit, Section ys vLogDir w373, %A_WinDir%\System32\winevt\Logs
Gui, EventLogAnalyzer:Add, Button, ys gEventLogAnalyzer_Select_System, 选择当前系统
Gui, EventLogAnalyzer:Add, Button, ys gEventLogAnalyzer_Select_Dir, 选择目录
Gui, EventLogAnalyzer:Add, Edit, xs vCurrentConfig Multi w540 h295,
Gui, EventLogAnalyzer:Add, Edit, vCmdOutput Multi Readonly w540 h190,
Gui, EventLogAnalyzer:Add, GroupBox, Section ys h0 w0, 输出
Gui, EventLogAnalyzer:Add, Button, Section xs gEventLogAnalyzer_OpenFolder, 打开配置文件目录
Gui, EventLogAnalyzer:Add, Button, Section xs gShortCutInfo, 快捷键说明
Gui, EventLogAnalyzer:Add, Button, Section xs gRunEventvwr, 事件查看器
Gui, EventLogAnalyzer:Add, Button, Section xs gOpenCMD, 打开 CMD
Gui, EventLogAnalyzer:Add, Button, ys gOpenPowerShell, 打开 PowerShell
Gui, EventLogAnalyzer:Add, Button, Section xs gEventLogAnalyzer_CopyCmd, 复制命令
Gui, EventLogAnalyzer:Add, Button, ys gEventLogAnalyzer_ExecCmd, 执行命令
Gui, EventLogAnalyzer:Add, Button, Section xs gEventLogAnalyzer_ListConfigFiles, 刷新列表
;~ Gui, EventLogAnalyzer:+Resize ; TODO
Gui, EventLogAnalyzer:Show,, %Title_EventLogAnalyzer%
Gosub, EventLogAnalyzer_ListConfigFiles
Return

EventLogAnalyzerGuiEscape:
EventLogAnalyzerGuiClose:
Gui, EventLogAnalyzer:Destroy
Return

; 右键菜单
EventLogAnalyzerGuiContextMenu:
IfEqual, A_GuiControl, EventLogAnalyzer_LB_ConfigFiles
{
    Try, Menu, EventLogAnalyzer_LB_Menu, Delete
    Menu, EventLogAnalyzer_LB_Menu, Add, 刷新列表, EventLogAnalyzer_ListConfigFiles
    Menu, EventLogAnalyzer_LB_Menu, Add
    Menu, EventLogAnalyzer_LB_Menu, Add, 复制命令, EventLogAnalyzer_CopyCmd
    Menu, EventLogAnalyzer_LB_Menu, Add, 打开目录, EventLogAnalyzer_OpenFolder
    Menu, EventLogAnalyzer_LB_Menu, Add, 编辑文件, EventLogAnalyzer_EditConfigFile
    Menu, EventLogAnalyzer_LB_Menu, Add
    Menu, EventLogAnalyzer_LB_Menu, Add, 执行命令, EventLogAnalyzer_ExecCmd
    Menu, EventLogAnalyzer_LB_Menu, Show
}
Return

; 更新 ListView（左侧配置文件列表）
EventLogAnalyzer_ListConfigFiles:
EventLogAnalyzer_LB_ConfigFiles =
FileCreateDir, %A_ScriptDir%\Modules\EventLogs
Loop, %A_ScriptDir%\Modules\EventLogs\*
{
    StringTrimRight, OutputVar, A_LoopFileName, % StrLen(A_LoopFileExt) + 1
    EventLogAnalyzer_LB_ConfigFiles := EventLogAnalyzer_LB_ConfigFiles "|" OutputVar
}
GuiControl, EventLogAnalyzer:, EventLogAnalyzer_LB_ConfigFiles, % EventLogAnalyzer_LB_ConfigFiles
GuiControl, EventLogAnalyzer:Choose, EventLogAnalyzer_LB_ConfigFiles, 1
Gosub, EventLogAnalyzer_LoadConfig
Return

; 加载配置文件内容到编辑框控件
EventLogAnalyzer_LoadConfig:
IfEqual, A_GuiEvent, DoubleClick
    Gosub, EventLogAnalyzer_ExecCmd
GuiControlGet, CurrentConfigFileName, EventLogAnalyzer:, EventLogAnalyzer_LB_ConfigFiles
IfNotExist, %A_ScriptDir%\Modules\EventLogs\%CurrentConfigFileName%.txt
{
    MsgBox, %CurrentConfigFileName% 需要以 txt 为后缀。
    GuiControl,, CurrentConfig
    Return
}
FileRead, OutputVar, *P65001 %A_ScriptDir%\Modules\EventLogs\%CurrentConfigFileName%.txt ; *P65001 以 UTF-8 编码读取文本
GuiControl, EventLogAnalyzer:, CurrentConfig, % OutputVar
Return

; 设置分析当前系统日志
EventLogAnalyzer_Select_System:
GuiControl, EventLogAnalyzer:, LogDir, %A_WinDir%\System32\winevt\Logs
Return

; 设置分析指定目录日志
EventLogAnalyzer_Select_Dir:
GuiControlGet, LogDir, EventLogAnalyzer:, LogDir
FileSelectFolder, _LogDir, *%LogDir%, 7, 请选择存放离线日志的目录
IfNotEqual, ErrorLevel, 1
    GuiControl, EventLogAnalyzer:, LogDir, %_LogDir%
Return

EventLogAnalyzer_OpenFolder:
Run, %A_ScriptDir%\Modules\EventLogs
Return

ShortCutInfo:
ShortCut =
(
# 当前窗口快捷键（焦点不处于编辑框）
F1              快捷键说明
F5              刷新配置文件列表
Ctrl+Up         选择上一个配置文件
Ctrl+Down       选择下一个配置文件
Ctrl+E          编辑配置文件
Ctrl+Enter      执行命令
Alt+C           打开 CMD
Alt+P           打开 PowerShell

# 列表控件热键（焦点处于配置文件列表）
Ctrl+C          复制命令
Enter           执行命令
)
MsgBox, % ShortCut
Return

RunEventvwr:
Run, eventvwr.exe
Return

OpenCMD:
GuiControlGet, LogDir, EventLogAnalyzer:, LogDir
EnvSet, LogDir, %LogDir%
Run, cmd /d/k cls && for /f "delims=" `%i in ('ver') do @title CMD - `%i && prompt $S$S[$T]$S$P$_$G$S
Return

OpenPowerShell:
Run, powershell -noe -nol -nop -ex unrestricted
Return

EventLogAnalyzer_CopyCmd:
GuiControlGet, cmd, EventLogAnalyzer:, CurrentConfig
GuiControlGet, LogDir, EventLogAnalyzer:, LogDir ; 在函数 CmdCompress 中引用
Clipboard := CmdCompress(cmd) ; 复制命令时处理下 LogParser 的换行
Return

EventLogAnalyzer_EditConfigFile:
Run, %A_ScriptDir%\Modules\EventLogs\%CurrentConfigFileName%.txt
Return

EventLogAnalyzer_SelectUp:
GuiControl, EventLogAnalyzer:Focus, EventLogAnalyzer_LB_ConfigFiles
Send, {Up}
Return

EventLogAnalyzer_SelectDown:
GuiControl, EventLogAnalyzer:Focus, EventLogAnalyzer_LB_ConfigFiles
Send, {Down}
Return

; 执行 LogParser 需要处理分析源的两种情况
EventLogAnalyzer_ExecCmd:
GuiControlGet, cmd, EventLogAnalyzer:, CurrentConfig
GuiControlGet, LogDir, EventLogAnalyzer:, LogDir
EnvSet, LogDir, %LogDir%
;~ CmdOutput := RunWaitCmdOutput(CmdCompress(cmd))
CmdOutput := RunWaitBatOutput("@echo off`r`n"CmdCompress(cmd))
GuiControl, EventLogAnalyzer:, CmdOutput, % "> 选择的日志目录：" LogDir "`n> 执行结果：`n" CmdOutput
Return
; ------------------------------------------------------------------------------
;
; EventLogAnalyzer END
;
; ------------------------------------------------------------------------------
; ------------------------------------------------------------------------------
;
; DetectBruteForce BEGIN
;
; ------------------------------------------------------------------------------
; 检查是否存在暴力破解登陆成功行为
DetectBruteForce:
GuiControl, Disable, DetectBruteForce

; 获取远程登录成功记录，输出：TargetUserName,IpAddress
cmd =
(Join`s LTrim
LogParser.exe -q -stats:OFF -i:EVT -o:CSV -headers:OFF "
    SELECT
        EXTRACT_TOKEN(Strings, 5, '|') AS TargetUserName,
        EXTRACT_TOKEN(Strings, 18, '|') AS IpAddress
    USING
        EXTRACT_TOKEN(Strings, 8, '|') AS LogonType
    INTO %A_Temp%\LogParserDetectBruteForce.out
    FROM Security
    WHERE
        EventID = 4624 AND
        LogonType IN ('10';'11')
    GROUP BY IpAddress,TargetUserName -- 对登陆成功记录的 IP 去重
    ORDER BY IpAddress
    "
)
RunWait, %cmd%,, Hide UseErrorLevel
IfNotEqual, ErrorLevel, 0
{
    MsgBox, 执行出错，退出代码：%ErrorLevel%。
    Return
}
IfNotExist, %A_Temp%\LogParserDetectBruteForce.out ; 查询结果为空时不会生成文件，不需要处理
{
    MsgBox, 未查询到远程登录成功记录。
    GuiControl, Enable, DetectBruteForce
    Return
}

Loop, Read, %A_Temp%\LogParserDetectBruteForce.out
{
    StringSplit, OutputArray, A_LoopReadLine, `,
    TargetUserName := OutputArray1
    IpAddress := OutputArray2

    ; 指定 IP 查询身份验证成功记录
    cmd = ; 首次登陆成功时间，格式：2021-01-12 09:35:44
    (Join`s LTrim
    LogParser.exe -stats:OFF -i:EVT -o:CSV -headers:OFF "
        SELECT
        TOP 1
            TimeGenerated
        USING
            EXTRACT_TOKEN(Strings, 8, '|') AS LogonType,
            EXTRACT_TOKEN(Strings, 18, '|') AS IpAddress
        INTO %A_Temp%\LogParserDetectBruteForce.out1
        FROM Security
        WHERE
            EventID = 4624 AND
            LogonType IN ('10';'11') AND
            IpAddress LIKE '%IpAddress%'
        "
    )
    RunWait, %cmd%,, Hide UseErrorLevel
    IfNotEqual, ErrorLevel, 0, Return
    FileReadLine, first_success_logon, %A_Temp%\LogParserDetectBruteForce.out1, 1

    ; 指定 IP 查询身份验证失败记录
    cmd = ; 首次登陆失败时间，格式：2021-01-12 09:35:44
    (Join`s LTrim
    LogParser.exe -stats:OFF -i:EVT -o:CSV -headers:OFF "
        SELECT
        TOP 1
            TimeGenerated
        USING
            EXTRACT_TOKEN(Strings, 10, '|') AS LogonType,
            EXTRACT_TOKEN(Strings, 19, '|') AS IpAddress
        INTO %A_Temp%\LogParserDetectBruteForce.out1
        FROM Security
        WHERE
            EventID = 4625 AND
            LogonType LIKE '3' AND -- TODO LogonType 3 的匿名登录记录大多数都没有记录 IP，得考虑用其他日志文件查找
            IpAddress LIKE '%IpAddress%'
        "
    )
    RunWait, %cmd%,, Hide UseErrorLevel
    IfNotEqual, ErrorLevel, 0
        first_failed_logon = 日志中未找到登录失败记录
    Else
        FileReadLine, first_failed_logon, %A_Temp%\LogParserDetectBruteForce.out1, 1

    ; 指定 IP 在登录成功之前有多少次失败记录
    cmd =
    (Join`s LTrim
    LogParser.exe -stats:OFF -i:EVT -o:CSV -headers:OFF "
        SELECT
            COUNT(*)
        USING
            EXTRACT_TOKEN(Strings, 10, '|') AS LogonType,
            EXTRACT_TOKEN(Strings, 19, '|') AS IpAddress
        INTO %A_Temp%\LogParserDetectBruteForce.out2
        FROM Security
        WHERE
            EventID = 4625 AND
            LogonType LIKE '3' AND
            IpAddress LIKE '%IpAddress%' AND
            TimeGenerated < TIMESTAMP('%first_success_logon%', 'yyyy-MM-dd hh:mm:ss')
        "
    )
    RunWait, %cmd%,, Hide UseErrorLevel
    IfNotEqual, ErrorLevel, 0
        failed_logon_times = 日志中未找到登录失败记录
    Else
        FileReadLine, failed_logon_times, %A_Temp%\LogParserDetectBruteForce.out2, 1

    ; 指定 IP 总计登陆失败次数
    cmd =
    (Join`s LTrim
    LogParser.exe -stats:OFF -i:EVT -o:CSV -headers:OFF "
        SELECT
            COUNT(*)
        USING
            EXTRACT_TOKEN(Strings, 10, '|') AS LogonType,
            EXTRACT_TOKEN(Strings, 19, '|') AS IpAddress
        INTO %A_Temp%\LogParserDetectBruteForce.out3
        FROM Security
        WHERE
            EventID = 4625 AND
            LogonType LIKE '3' AND
            IpAddress LIKE '%IpAddress%'
        "
    )
    RunWait, %cmd%,, Hide UseErrorLevel
    IfNotEqual, ErrorLevel, 0
        total_failed_logon_times = 日志中未找到登录失败记录
    Else
        FileReadLine, total_failed_logon_times, %A_Temp%\LogParserDetectBruteForce.out3, 1

    ; TODO 实现：在某个 IP 首次成功登陆前失败次数达到 M 次为可疑登陆，为 N 次为暴力破解成功，并输出该 IP 最早登陆成功时间，最早登陆失败时间，登陆失败次数
    ; 同个账号下，没有登陆成功记录，但登陆失败次数多于指定次数，可确定为爆破行为
    ; 同个账号下，如果登陆失败时间早于登陆成功时间，并且登陆失败次数多于指定次数，可确定为爆破成功
    ; 同个账号下，如果登陆失败时间晚于登陆成功时间，并且登陆失败次数多于指定次数，为可疑爆破行为

    a = 5   ; 在登录成功前允许的最大失败次数
    b = 20
    c = 60
    ; 判断是否爆破成功
    IfLess, failed_logon_times, %a%                 ; 首次成功登陆前失败次数小于 a 次为可允许的失败次数
        status = 否
    If failed_logon_times Between %a% and %b%       ; 首次成功登陆前失败次数大于 a 次小于 b 次为可疑
        status = 可疑
    IfGreaterOrEqual, failed_logon_times, %c%       ; 首次成功登陆前失败次数大于 c 次认为是暴力破解成功
        status = 是
    IfGreater, failed_logon_times, %b%, IfLess, total_failed_logon_times, %c% ; 首次成功登陆前失败次数大于 b 次，但总计失败次数小于 c 次为可疑
        status = 可疑
    IfGreaterOrEqual, total_failed_logon_times, %c% ; 登陆失败次数大于 c 次认为是暴力破解成功
        status = 是
    ; msgbox, % IpAddress "," TargetUserName "," first_success_logon "," first_failed_logon "," failed_logon_times "," total_failed_logon_times "," status
    ; CSV：登陆成功 IP, 登陆成功账号, 首次登陆成功时间, 首次登陆失败时间, 登陆成功前尝试次数, 总计登陆失败次数, 是否爆破成功（可疑/是/否）
    FileAppend, % IpAddress "," TargetUserName "," first_success_logon "," first_failed_logon "," failed_logon_times "," total_failed_logon_times "," status "`r`n", %A_Temp%\LogParserDetectBruteForce.out4
}
cmd =
(Join`s LTrim
LogParser.exe -stats:OFF -i:CSV -o:DATAGRID -headerRow:OFF "
    SELECT
        Field1 AS 登陆成功IP,
        Field2 AS 登陆成功账号,
        Field3 AS 首次登陆成功时间,
        Field4 AS 首次登陆失败时间,
        Field5 AS 登陆成功前尝试次数,
        Field6 AS 总计登陆失败次数,
        Field7 AS 是否爆破成功（可疑/是/否）
    FROM %A_Temp%\LogParserDetectBruteForce.out4
    "
)
RunLogParser(cmd)
FileDelete, %A_Temp%\LogParserDetectBruteForce.out
FileDelete, %A_Temp%\LogParserDetectBruteForce.out1
FileDelete, %A_Temp%\LogParserDetectBruteForce.out2
FileDelete, %A_Temp%\LogParserDetectBruteForce.out3
FileDelete, %A_Temp%\LogParserDetectBruteForce.out4

GuiControl, Enable, DetectBruteForce
Return
; ------------------------------------------------------------------------------
;
; DetectBruteForce END
;
; ------------------------------------------------------------------------------
; ------------------------------------------------------------------------------
;
; RegeditLocate BEGIN
;
; ------------------------------------------------------------------------------
RegeditLocate:
IfWinExist, %Title_RegeditLocate%
{
    WinActivate
    Return
}
GuiControl, Disable, RegeditLocate

RegeditPathFile = %A_ScriptDir%\Modules\Registry\RegistryPath.txt
IfNotExist, %RegeditPathFile%
{
    FileCreateDir, %A_ScriptDir%\Modules\Registry
    FileAppend,, %RegeditPathFile%
}
Gosub, RefreshRegeditPath

Gui, RegeditLocate:Add, ComboBox, vRegeditPath w600, %RegeditPathText%
Gui, RegeditLocate:Add, Button, gJumpRegeditPath ys Default, Go
Gui, RegeditLocate:Add, Button, gEditRegeditPath Section xs, 编辑列表
Gui, RegeditLocate:Add, Button, gRefreshRegeditPath ys, 更新列表
Gui, RegeditLocate:Show,, %Title_RegeditLocate%

GuiControl, Enable, RegeditLocate
Return

RegeditLocateGuiEscape:
RegeditLocateGuiClose:
Gui, RegeditLocate:Destroy
Return

JumpRegeditPath:
GuiControlGet, RegeditPath,, RegeditPath

; 替换字符
StringReplace, RegeditPath, RegeditPath, ＼, \, 1

; 早期操作系统的 Regedit 不支持类似 HKLM 的路径简写，需要替换为完整路径
RegeditPath := RegExReplace(RegeditPath, "^HKLM", "HKEY_LOCAL_MACHINE")
RegeditPath := RegExReplace(RegeditPath, "^HKCU", "HKEY_CURRENT_USER")
RegeditPath := RegExReplace(RegeditPath, "^HKU", "HKEY_USERS")
RegeditPath := RegExReplace(RegeditPath, "^HKCR", "HKEY_CLASSES_ROOT")
RegeditPath := RegExReplace(RegeditPath, "^HKCC", "HKEY_CURRENT_CONFIG")

RegWrite, REG_SZ, HKCU, SOFTWARE\Microsoft\Windows\CurrentVersion\Applets\Regedit, LastKey, %RegeditPath%
Run, regedit -m ; -m 允许运行多个注册表实例
Return

EditRegeditPath:
Run, %RegeditPathFile%
Return

; 更新 ComboBox 控件预设值
RefreshRegeditPath:
FileRead, RegeditPathText, %RegeditPathFile%
StringReplace, RegeditPathText, RegeditPathText, `r`n, |, 1
StringReplace, RegeditPathText, RegeditPathText, `n, |, 1
GuiControl,, RegeditPath, % "|"RegeditPathText
Return
; ------------------------------------------------------------------------------
;
; RegeditLocate END
;
; ------------------------------------------------------------------------------

; ------------------------------------------------------------------------------
;
; RunFullEventLogView BEGIN
;
; ------------------------------------------------------------------------------
RunFullEventLogView:
Run, bin\FullEventLogView.exe
Return
; ------------------------------------------------------------------------------
;
; RunFullEventLogView END
;
; ------------------------------------------------------------------------------

; 隐藏 LogParser 控制台窗口
HideCmdWindow:
DetectHiddenWindows, On
Loop, 100
{
    IfWinExist, ahk_pid %RunWaitCmdPID%
    {
        WinHide, ahk_pid %RunWaitCmdPID% ahk_class ConsoleWindowClass ; 隐藏控制台窗口
        Return
    }
    Sleep, 50
}
Return

ResizeLogParserWindow:
; TODO
Return

Reload:
Gosub, clean
Reload
Return

clean:
; TODO 是否需要结束 LogParser 进程
; LogParser 在执行查询未完成时文件为占用状态无法删除，统一在重载/退出时执行删除
FileDelete, Microsoft-Windows-TerminalServices-LocalSessionManager`%4Operational.evtx
FileDelete, Microsoft-Windows-RemoteDesktopServices-RdpCoreTS`%4Operational.evtx
Return

SetHotkey:
; 全局热键
Hotkey, ^F12, reload

; 窗口热键
Hotkey, IfWinActive, %Title_EventLogAnalyzer%
Hotkey, F1, ShortCutInfo
Hotkey, F5, EventLogAnalyzer_ListConfigFiles
Hotkey, ^Up, EventLogAnalyzer_SelectUp
Hotkey, ^Down, EventLogAnalyzer_SelectDown
Hotkey, ^e, EventLogAnalyzer_EditConfigFile
Hotkey, ^Enter, EventLogAnalyzer_ExecCmd
Hotkey, !c, OpenCMD
Hotkey, !p, OpenPowerShell
Hotkey, IfWinActive

; 控件热键
ActiveControlIsClass(Class) {
    global Title_EventLogAnalyzer
    ControlGetFocus, OutputVar, %Title_EventLogAnalyzer%
    IfEqual, OutputVar, % Class
        Return 1
    Else
        Return 0
}

#If, WinActive(Title_EventLogAnalyzer) and ActiveControlIsClass("ListBox1")
^c::Gosub, EventLogAnalyzer_CopyCmd
Enter::Gosub, EventLogAnalyzer_ExecCmd
#If
Return

; ------------------------------------------------------------------------------
;
; 通用函数
;
; ------------------------------------------------------------------------------
; 检查依赖命令
CheckCommand(str) {
    global binPATH
    Loop, Parse, binPATH, `;
    {
        IfExist, %A_LoopField%\%str%
            Return, 1
    }
}

; 执行 CMD 命令，支持执行多行命令，所有命令执行结束后返回执行结果（有弹出 CMD 窗口）
RunWaitCmdOutput(cmd) {
    global RunWaitCmdPID =
    shell := ComObjCreate("WScript.Shell")
    exec := shell.Exec("cmd /d/q/k echo off")
    SetTimer, HideCmdWindow, -100, 5
    RunWaitCmdPID := exec.ProcessID
    exec.StdIn.WriteLine(cmd "`nexit")
    Loop ; 等待命令执行完毕
    {
        If exec.Status = 1 ; https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/windows-scripting/443b45a5(v=vs.84)
            Break
        Sleep, 100
    }
    If exec.ExitCode = 0 ; 退出状态码（成功）
        Return exec.StdOut.ReadAll() exec.StdErr.ReadAll() 
    Else
        Return exec.StdOut.ReadAll() exec.StdErr.ReadAll() 
}

; 执行 Bat 脚本，需要生成临时文件，脚本执行完毕后返回执行结果（无弹出 CMD 窗口）
RunWaitBatOutput(bat) {
    BatExecFile = ~%A_Now%.bat
    FileDelete, %A_Temp%\%BatExecFile%
    FileAppend, %bat%, %A_Temp%\%BatExecFile%
    RunWait, cmd /d/c %BatExecFile% > %BatExecFile%.out, %A_Temp%, Hide
    FileRead, BatOutput, %A_Temp%\%BatExecFile%.out
    FileDelete, %A_Temp%\%BatExecFile%
    FileDelete, %A_Temp%\%BatExecFile%.out
    Return % BatOutput
}

; 执行 LogParser 命令
; HideConsoleWindow 是否隐藏控制台窗口
; AutoResize        是否在加载完成后发送按键点击 Auto Resize（针对 DATAGRID 视图）
RunLogParser(cmd, HideConsoleWindow = 1, AutoResize = 1) {
    Run, %cmd%,, UseErrorLevel, OutputVarPID
    IfNotEqual, ErrorLevel, 0
        Return
    IfEqual, HideConsoleWindow, 1
    {
        WinWait, ahk_pid %OutputVarPID%
        WinHide, ahk_pid %OutputVarPID% ahk_class ConsoleWindowClass ; 隐藏控制台窗口
    }
    IfEqual, AutoResize, 1
    {
        WinActivate, ahk_pid %OutputVarPID%
        WinWaitActive, ahk_pid %OutputVarPID%,, 3
        ControlSend, Button3, &r, ahk_pid %OutputVarPID% ; Auto Resize
    }
}
; 执行 LogParser 命令并返回输出
RunWaitLogParserOutput(cmd) {
    RunWait, %cmd%, %A_Temp%, Hide
    FileRead, LogParserOutput, %A_Temp%\LogParser.out
    FileDelete, %A_Temp%\LogParser.out
    Return % LogParserOutput
}

; ------------------------------------------------------------------------------
;
; 功能函数
;
; ------------------------------------------------------------------------------
; 获取远程桌面服务配置端口
getTermServicePortNumber() {
    RegRead, PortNumber, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp, PortNumber
    Return PortNumber
}
getUserAuthentication() {
    RegRead, UserAuthentication, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp, UserAuthentication
    Return % UserAuthentication ? "是" : "否"
}
getSecurityLayer() {
    RegRead, SecurityLayer, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp, SecurityLayer
    ; https://docs.microsoft.com/en-us/windows-hardware/customize/desktop/unattend/microsoft-windows-terminalservices-rdp-winstationextensions-securitylayer
    IfEqual, SecurityLayer, 0, Return "使用 RDP 进行身份验证"
    IfEqual, SecurityLayer, 1, Return "协商身份验证方法"
    IfEqual, SecurityLayer, 2, Return "使用 TLS 进行身份验证"
}
; 获取 CredSSP 加密数据库修正策略
getAllowEncryptionOracle() {
    RegRead, AllowEncryptionOracle, HKEY_LOCAL_MACHINE, SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\CredSSP\Parameters, AllowEncryptionOracle
    ; https://go.microsoft.com/fwlink/?linkid=866660
    IfEqual, AllowEncryptionOracle, 0, Return "强制更新的客户端"
    IfEqual, AllowEncryptionOracle, 1, Return "缓解"
    IfEqual, AllowEncryptionOracle, 2, Return "易受攻击"
    Return "未配置"
}
; 是否允许连接远程桌面服务
getfDenyTSConnections() {
    RegRead, fDenyTSConnections, HKEY_LOCAL_MACHINE, SYSTEM\CurrentControlSet\Control\Terminal Server, fDenyTSConnections
    IfEqual, fDenyTSConnections, 0, Return "是"
    IfEqual, fDenyTSConnections, 1, Return "否"
}
; 获取 TermService 服务状态
queryTermServiceStatus() {
    ; Win7 有个奇怪的 bug，for 命令输出第一个字符有可能是个 \xFF 字符，以空格为分隔符时就可能取错字段
    bat =
    (% LTrim
    @echo off
    pushd %~dp0
    for /f "tokens=2 delims=:" %%i in ('sc query TermService ^| findstr STATE') do ^
    for /f "tokens=2" %%j in ("%%i") do set STATE=%%j
    set/p=%STATE%<nul
    del %0
    )
    BatOutput := RunWaitBatOutput(bat)
    If RegExMatch(BatOutput, "[A-Z_]+")
        Return BatOutput
}
; 获取 TermService 服务的 PID
queryTermServicePID() {
    ; Win7 有个奇怪的 bug，for 命令输出第一个字符有可能是个 \xFF 字符，以空格为分隔符时就可能取错字段
    bat =
    (% LTrim
    @echo off
    pushd %~dp0
    for /f "tokens=2" %%i in ('tasklist /FI "SERVICES eq TermService" /NH ^| find "svchost.exe"') do set PID=%%i
    set/p=%PID%<nul
    del %0
    )
    BatOutput := RunWaitBatOutput(bat)
    If RegExMatch(BatOutput, "\d+")
        Return BatOutput
}
; 获取 TermService 服务的监听端口
queryTermServiceListeningPort() {
    ; Win7 有个奇怪的 bug，for 命令输出第一个字符有可能是个 \xFF 字符，以空格为分隔符时就可能取错字段
    bat =
    (% LTrim
    @echo off
    pushd %~dp0
    for /f "tokens=2" %%i in ('tasklist /FI "SERVICES eq TermService" /NH ^| find "svchost.exe"') do set PID=%%i
    for /f "tokens=2 delims=:" %%j in ('netstat -ano ^| findstr "LISTENING.*%PID%" ^| findstr /v "::"') do ^
    for /f "tokens=1" %%k in ("%%j") do set PORT=%%k
    set/p=%PORT%<nul
    del %0
    )
    BatOutput := RunWaitBatOutput(bat)
    If RegExMatch(BatOutput, "\d+")
        Return BatOutput
}
getRdCoreTSLogMaxSize() {
    bat =
    (% LTrim
    @echo off
    pushd %~dp0
    for /f "delims=: tokens=2" %%i in ('wevtutil.exe gl Microsoft-Windows-RemoteDesktopServices-RdpCoreTS/Operational ^| findstr maxSize') do set maxSize=%%i
    set/p=%maxSize%<nul
    del %0
    )
    BatOutput := RunWaitBatOutput(bat)
    If RegExMatch(BatOutput, "\d+")
    {
        BatOutput /= 1024
        Return BatOutput " KB"
    }
}
; 构造 LogParser 执行命令（分析三大系统日志、其他系统日志及离线日志）
; 需要处理两种场景：分析当前系统日志以及分析指定目录下的日志文件
;
; 场景1：日志为 Security,System,Application
;
; - 当分析当前系统日志时保留不带 evtx 后缀的日志名即可
; - 当分析指定目录日志文件时，更新 FROM 语句为完整路径名
;
; 场景2：日志非 Security,System,Application
;
; - 需要传递 LogDir 变量给子进程（通过 EnvSet）以正确执行 CMD 命令
;     - 情况1：为 OpenCMD【打开 CMD】配置（用于传递给 CMD 子进程）
;     - 情况2：为 EventLogAnalyzer_ExecCmd【执行命令】配置
; - 不需要更新 FROM 语句，但配置文件需按格式编写
;     - 将日志文件拷贝到临时目录：copy /y %LogDir%\%logname% %tmp%
;     - 进临时目录执行查询：pushd %tmp%
CmdCompress(cmd) {
    global LogDir
    Loop, Parse, cmd, `n, `r
    {
        If SubStr(A_LoopField, 1, 2) = "::" ; 以 :: 开头的为自定义的 CMD 命令
            _cmd := _cmd "`r`n" SubStr(A_LoopField, 3) "`r`n"
        Else ; 否则为 LogParser 命令及参数
        {
            StringReplace, _LogParseCmd, A_LoopField, `t, %A_Space%, 1
            If RegExMatch(_LogParseCmd, "i)FROM\s+") ; 根据 FROM 语句指定的日志名做额外处理（是否为日志名添加 evtx 后缀及完整路径）
            {
                NewStr := RegExReplace(_LogParseCmd, "i)\s+FROM\s+(Security|System|Application)$", "$1")
                If NewStr in Security,System,Application
                {
                    IfNotEqual, LogDir, %A_WinDir%\System32\winevt\Logs
                        _LogParseCmd = FROM "%LogDir%\%NewStr%.evtx"
                }
                ; 非 Security,System,Application 日志需涉及 SetWorkingDir、手动复制命令执行、删除临时文件，目前在配置文件通过 CMD 命令实现
            }
            Loop ; 压缩空白字符
            {
                StringReplace, _LogParseCmd, _LogParseCmd, %A_Space%%A_Space%, %A_Space%, UseErrorLevel
                If ErrorLevel = 0
                    Break
            }
            _cmd := _cmd " " _LogParseCmd
        }
    }
    Return _cmd
}
