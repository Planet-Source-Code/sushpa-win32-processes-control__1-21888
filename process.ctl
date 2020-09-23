VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl Process 
   Alignable       =   -1  'True
   ClientHeight    =   4230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7395
   PropertyPages   =   "process.ctx":0000
   ScaleHeight     =   4230
   ScaleWidth      =   7395
   ToolboxBitmap   =   "process.ctx":002B
   Begin VB.Timer Timer1 
      Left            =   1530
      Top             =   270
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   630
      Top             =   2970
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "process.ctx":033D
            Key             =   "file"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "process.ctx":04A1
            Key             =   "id"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "process.ctx":05FD
            Key             =   "threads"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "process.ctx":0919
            Key             =   "path"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   4020
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7350
      _ExtentX        =   12965
      _ExtentY        =   7091
      SortKey         =   1
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Module"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Process ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Parent ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "#"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
   End
End
Attribute VB_Name = "Process"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit
'-----------------------------------------------
Const MAX_PATH = 260
Const TH32CS_SNAPPROCESS = 2&
'-----------------------------------------------
Private Type PROCESSENTRY32
    lSize            As Long
    lUsage           As Long
    lProcessId       As Long
    lDefaultHeapId   As Long
    lModuleId        As Long
    lThreads         As Long
    lParentProcessId As Long
    lPriClassBase    As Long
    lFlags           As Long
    sExeFile         As String * MAX_PATH
End Type
'------------------------------------------------------
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'------------------------------------------------------
Private bCanResize As Boolean
'------------------------------------------------------
Public Enum EnumViewType
Icon = 2
Full = 3
End Enum

Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "ppMain"
Enabled = lvMain.Enabled
End Property

Property Let Enabled(vData As Boolean)
lvMain.Enabled = vData
PropertyChanged "Enabled"
End Property

Property Get Resizable() As Boolean
Attribute Resizable.VB_ProcData.VB_Invoke_Property = "ppMain"
Resizable = bCanResize
End Property

Property Let Resizable(vData As Boolean)
bCanResize = vData
PropertyChanged "Resizable"
End Property

Property Get ViewMode() As EnumViewType
ViewMode = lvMain.View
End Property

Property Let ViewMode(vData As EnumViewType)
lvMain.View = vData
PropertyChanged "ViewMode"
End Property

Property Get ListFont() As String
Font = lvMain.Font.Name
End Property

Property Let ListFont(vData As String)
lvMain.Font = vData
PropertyChanged "ListFont"
End Property

Public Property Get FontSize() As Integer
FontSize = lvMain.Font.Size
End Property

Public Property Let FontSize(vData As Integer)
lvMain.Font.Size = vData
PropertyChanged "FontSize"
End Property

Property Let ForeColor(vData As OLE_COLOR)
lvMain.ForeColor = vData
PropertyChanged "ForeColor"
End Property

Property Get ForeColor() As OLE_COLOR
ForeColor = lvMain.ForeColor
End Property

Property Let BackColor(vData As OLE_COLOR)
lvMain.BackColor = vData
PropertyChanged "BackColor"
End Property

Property Get BackColor() As OLE_COLOR
BackColor = lvMain.BackColor
End Property

Private Sub Load()
On Error Resume Next
'Declarations
Dim sExeName   As String
Dim sPid       As String
Dim sParentPid As String
Dim lSnapShot  As Long
Dim r          As Long
Dim uProcess   As PROCESSENTRY32
'Get the items
lSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
If lSnapShot <> 0 Then
    With lvMain
    'Arrange coloumns
    .ListItems.Clear
    .ColumnHeaders(1).Width = 1500
    .ColumnHeaders(2).Width = 1000
    .ColumnHeaders(3).Width = 1000
    .ColumnHeaders(4).Width = 375
    .ColumnHeaders(5).Width = lvMain.Width - 3940
    uProcess.lSize = Len(uProcess)
    r = ProcessFirst(lSnapShot, uProcess)
    Do While r
        'uProcess properties are to be filled in lvmain
        sExeName = (Left(uProcess.sExeFile, InStr(1, uProcess.sExeFile, vbNullChar) - 1))
        sPid = Hex$(uProcess.lProcessId)
        sParentPid = Hex$(uProcess.lParentProcessId)
        .ListItems.Add 1, , GetFile(sExeName)
        .ListItems.Item(1).ListSubItems.Add 1, , sPid
        .ListItems(1).ListSubItems.Add 2, , sParentPid
        .ListItems(1).ListSubItems.Add 3, , CStr(uProcess.lThreads)
        .ListItems.Item(1).ListSubItems.Add 4, , GetPath(sExeName)
        r = ProcessNext(lSnapShot, uProcess)
    Loop
    CloseHandle (lSnapShot) 'Done
    End With
End If
End Sub

Private Sub UserControl_Initialize()
Load 'Load things
bCanResize = False 'by default
End Sub

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
MsgBox "Win32 Process Heap Viewer by Sushant Pandurangi." & vbNewLine & "Please send in your comments at sushant@phreaker.net.", vbInformation, "About"
End Sub

Private Sub UserControl_Resize()
'Arrange
If bCanResize = True Then
lvMain.Width = ScaleWidth
lvMain.Height = ScaleHeight
Refresh
Else
Width = lvMain.Width
Height = lvMain.Height
End If
End Sub

Private Function GetFile(sPath As String) As String
'Returns only file name
GetFile = Right(sPath, InStr(1, StrReverse(sPath), "\") - 1)
End Function

Private Function GetPath(sPath As String) As String
'Returns only path name
GetPath = Left(sPath, Len(sPath) - InStr(1, StrReverse(sPath), "\"))
End Function

Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
Load 'Refresh module list
End Sub

