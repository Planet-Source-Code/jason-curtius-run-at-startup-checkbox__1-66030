VERSION 5.00
Begin VB.UserControl Run_At_StartUp 
   BackStyle       =   0  'Transparent
   ClientHeight    =   210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1410
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   210
   ScaleWidth      =   1410
   ToolboxBitmap   =   "run_at_startup.ctx":0000
   Begin VB.CheckBox RunAtStart 
      Caption         =   "Run at Startup"
      Height          =   195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "Run_At_StartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'Run at Startup user control by Jason Curtius
'Tested on windows XP only
'Credit to KPD-Team 2000, URL: http://www.allapi.net/
'for some of their registry functions and api tutorials.
'Credit to all those who give feedback on peoples code,
'I learn as much from reading peoples comments
'as I do from reading code.
'
'Thankyou to Roger Gilchrist for his valuable help
'on setting up the propertybag.
'
'Enjoy.
Option Explicit
Public Event Change() ' The Change event
Private MyFont As Font
Private Cmd_Line As String ' Command Line Arguements
'Registry read, write, delete, Query
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal Hkey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

' function prototypes, constants, and type definitions for Windows 32-bit Registry API
 Const HKEY_CLASSES_ROOT = &H80000000
 Const HKEY_CURRENT_USER = &H80000001
 Const HKEY_LOCAL_MACHINE = &H80000002
 Const HKEY_USERS = &H80000003
 Const HKEY_PERFORMANCE_DATA = &H80000004
 Const ERROR_SUCCESS = 0&
' Registry API prototypes
 Const REG_SZ = 1 ' Unicode nul terminated string
 Const REG_DWORD = 4 ' 32-bit number
'Create the properties.
Public Property Get Caption() As Variant 'Caption get, this reads from the caption property
Attribute Caption.VB_Description = "Change the label infront of the text box"
Attribute Caption.VB_ProcData.VB_Invoke_Property = ";Appearance"
Caption = RunAtStart.Caption ' set the RunAtStart.Caption to the propoperty caption
End Property
Public Property Let Caption(ByVal vNewValue As Variant) ' Set the Caption property
RunAtStart.Caption = vNewValue ' Set's RunAtStart.Caption to new value
PropertyChanged ' Triggers the WriteProperties event
UserControl_Resize 'resize caption
End Property
Public Property Get Cmd_Line_arg() As Variant
Attribute Cmd_Line_arg.VB_Description = "Change the text in the text box"
Attribute Cmd_Line_arg.VB_ProcData.VB_Invoke_Property = ";Misc"
Cmd_Line_arg = Cmd_Line
End Property
Public Property Let Cmd_Line_arg(ByVal vNewValue As Variant)
Cmd_Line = vNewValue
PropertyChanged
End Property
Public Property Get Alignment() As AlignmentConstants
Alignment = RunAtStart.Alignment
End Property
Public Property Let Alignment(ByVal vNewValue As AlignmentConstants)
If vNewValue = 2 Then vNewValue = 0
RunAtStart.Alignment = vNewValue
PropertyChanged
End Property
Public Property Get Run_At_StartUp() As Boolean
Run_At_StartUp = RunAtStart.Value
End Property
Public Property Let Run_At_StartUp(ByVal vNewValue As Boolean)
RunAtStart.Value = cBoolInt(vNewValue)
PropertyChanged
End Property
Public Property Get BackColor() As OLE_COLOR
BackColor = RunAtStart.BackColor
End Property
Public Property Let BackColor(ByVal vNewValue As OLE_COLOR)
RunAtStart.BackColor = vNewValue
PropertyChanged
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = RunAtStart.ForeColor
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
RunAtStart.ForeColor = vNewValue
PropertyChanged
End Property
Public Property Get Font() As Font
Set Font = MyFont
End Property
Public Property Set Font(ByVal vData As Font)
Set MyFont = vData
Set UserControl.Font = vData
Set RunAtStart.Font = MyFont
PropertyChanged "Font"
UserControl_Resize
End Property

Private Sub UserControl_Initialize()
qurey
End Sub

Private Sub UserControl_InitProperties()
Set Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'The propertybag

    With PropBag
        Set Font = .ReadProperty("Font", Ambient.Font)        'NOTE because strings are likely to trigger code that needs a font you should set it early
        Cmd_Line = .ReadProperty("prbCmd_Line_arg", "")
        BackColor = .ReadProperty("prbBackColor", vbButtonFace)
        ForeColor = .ReadProperty("prbForeColor", vbButtonText)
        Alignment = .ReadProperty("prbAlignment", 0)
        Caption = .ReadProperty("prbCaption", "Run At StartUp")
    End With '

End Sub
Private Sub UserControl_Resize()
'Make the caption text the same size as the UserControl
RunAtStart.Width = UserControl.TextWidth(RunAtStart.Caption) + UserControl.TextWidth(Mid$(RunAtStart.Caption, 1, 1)) + 210
RunAtStart.Height = UserControl.TextHeight(RunAtStart.Caption)
UserControl.Width = RunAtStart.Width
UserControl.Height = RunAtStart.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

' Write to your propertybag

    With PropBag
        .WriteProperty "Font", MyFont, Ambient.Font            ' Write  font property
        .WriteProperty "prbCaption", Caption, "Run At Startup" ' Write  caption property
        .WriteProperty "prbCmd_Line_arg", Cmd_Line, ""         ' Write  command line arguments property
        .WriteProperty "prbBackColor", BackColor, vbButtonFace ' Write Background color property
        .WriteProperty "prbForeColor", ForeColor, vbButtonText ' Write Text color
        .WriteProperty "prbAlignment", Alignment, 0
    End With
End Sub


Private Sub RunAtStart_Click()
If RunAtStart.Value = 0 Then del Else sve
RaiseEvent Change ' Triggers the Change event
End Sub

Private Function getstring(Hkey As Long, strPath As String, strValue As String)
'read string from registry
Dim keyhand As Long
Dim lValueType As Long
Dim r As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(Hkey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            getstring = Left$(strBuf, intZeroPos - 1)
        Else
            getstring = strBuf
        End If
    End If
End If
End Function

Private Sub savestring(Hkey As Long, strPath As String, strValue As String, strdata As String)
'save string to the rgistry
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(Hkey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub

Private Function DeleteValue(ByVal Hkey As Long, ByVal strPath As String, ByVal strValue As String)
'Delete from registry
Dim keyhand As Long
Dim r As Long
r = RegOpenKey(Hkey, strPath, keyhand)
r = RegDeleteValue(keyhand, strValue)
r = RegCloseKey(keyhand)
End Function
Private Sub sve()
Dim strString As String
'Save app path and name to registry makes Run_At_StartUp = True
If Len(App.Path) > 3 Then
Call savestring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\", App.EXEName & ".exe", Chr$(34) & App.Path + "\" + App.EXEName & ".exe" & Chr$(34) & " " & Cmd_Line)
Else
Call savestring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\", App.EXEName & ".exe", Chr$(34) & App.Path + App.EXEName & ".exe" & Chr$(34) & " " & Cmd_Line)
End If
End Sub
Public Function qurey() As Boolean
Dim strString As String
'Get a String out the Registry
strString = getstring(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\", App.EXEName & ".exe")
RunAtStart.Value = cBoolInt(CBool(Len(strString))) 'if reg entry exists RunAtStart.value = "1" else "0"
qurey = CBool(RunAtStart.Value) '(qurey = True/False)
End Function
Private Sub del()
'remove app path and name from registry  makes Run_At_StartUp = False
Call DeleteValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run\", App.EXEName & ".exe")
End Sub
Private Function cBoolInt(ByVal bool As Boolean) As Long
cBoolInt = IIf(bool = True, 1, 0) 'Returns true or false from a "0" or "1"
End Function
