VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WMML 0.1.11"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   7995
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   120
      TabIndex        =   11
      Text            =   "player"
      Top             =   3600
      Width           =   3135
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   270
      Left            =   120
      TabIndex        =   8
      Text            =   "1024"
      Top             =   2760
      Width           =   3135
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Text            =   "自定义"
      Top             =   360
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Text            =   "java "
      Top             =   720
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   3615
      Left            =   3360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   0
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "启动！"
      Height          =   615
      Left            =   5880
      TabIndex        =   0
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "用户名"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3360
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "内存大小(MB)"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "要启动的版本"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Java 路径"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 声明 Windows API 禁用文件系统重定向
Private Declare Function Wow64DisableWow64FsRedirection Lib "kernel32" (ByRef OldValue As Long) As Boolean
Private Declare Function Wow64RevertWow64FsRedirection Lib "kernel32" (ByVal OldValue As Long) As Boolean

' 在模块或窗体顶部定义
Private Type JavaInstallInfo
    FolderName As String    ' 文件夹名称（如 "jdk1.8.0_301"）
    JavaExePath As String  ' java.exe 的完整路径（如 "X:\Program Files\Java\jdk1.8.0_301\bin\java.exe"）
End Type

' 声明动态数组存储所有 Java 安装信息
Private JavaInstallations() As JavaInstallInfo
Private JavaInstallCount As Long  ' 记录已存储的 Java 安装数量

Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
    
Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" _
    (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
    
Private Declare Function FindClose Lib "kernel32" _
    (ByVal hFindFile As Long) As Long
    
Public Sub ScanJavaInstallations()
    Combo1.Clear
    Dim sJavaPath As String
    Dim sJavaExePath As String
    Dim OldRedirection As Long
    Dim hFind As Long
    Dim wfd As WIN32_FIND_DATA
    Dim sDirName As String
    
    JavaInstallCount = 1
    ReDim JavaInstallations(0)
    ReDim Preserve JavaInstallations(1)
    
    Wow64DisableWow64FsRedirection OldRedirection
    
    ' 检查多个可能的安装位置
    Dim searchPaths(3) As String
    searchPaths(0) = Environ("SystemDrive") & "\Program Files\Java\"
    searchPaths(1) = Environ("SystemDrive") & "\Program Files (x86)\Java\"
    searchPaths(2) = Environ("ProgramFiles") & "\Java\"
    searchPaths(3) = Environ("ProgramFiles(x86)") & "\Java\"
    
    Combo1.AddItem "自定义"
    JavaInstallations(1).FolderName = "自定义"
    JavaInstallations(1).JavaExePath = "java"
    
    Dim i As Integer
    For i = 0 To UBound(searchPaths)
        sJavaPath = searchPaths(i)
        If Right(sJavaPath, 1) <> "\" Then sJavaPath = sJavaPath & "\"
        
        ' 使用 API 查找目录
        hFind = FindFirstFile(sJavaPath & "*", wfd)
        If hFind <> INVALID_HANDLE_VALUE Then
            Do
                sDirName = Left$(wfd.cFileName, InStr(wfd.cFileName, vbNullChar) - 1)
                If (wfd.dwFileAttributes And vbDirectory) = vbDirectory Then
                    If sDirName <> "." And sDirName <> ".." Then
                        sJavaExePath = sJavaPath & sDirName & "\bin\java.exe"
                        If Dir(sJavaExePath) <> "" Then
                            ' 检查是否已经添加过这个版本
                            Dim bExists As Boolean
                            bExists = False
                            For j = 1 To JavaInstallCount
                                If JavaInstallations(j).FolderName = sDirName Then
                                    bExists = True
                                    Exit For
                                End If
                            Next j
                            
                            If Not bExists Then
                                JavaInstallCount = JavaInstallCount + 1
                                ReDim Preserve JavaInstallations(JavaInstallCount)
                                
                                Combo1.AddItem sDirName
                                
                                JavaInstallations(JavaInstallCount).FolderName = sDirName
                                JavaInstallations(JavaInstallCount).JavaExePath = sJavaExePath
                            End If
                        End If
                    End If
                End If
            Loop While FindNextFile(hFind, wfd)
            FindClose hFind
        End If
    Next i
    
    Wow64RevertWow64FsRedirection OldRedirection
    
    If Combo1.ListCount > 0 Then
        Combo1.ListIndex = 0
    End If
End Sub
' 函数：根据文件夹名称查找 java.exe 路径
Public Function GetJavaExePath(ByVal FolderName As String) As String
    Dim i As Long
    
    For i = 1 To JavaInstallCount
        If StrComp(JavaInstallations(i).FolderName, FolderName, vbTextCompare) = 0 Then
            GetJavaExePath = JavaInstallations(i).JavaExePath
            Exit Function
        End If
    Next i
    
    ' 没找到则返回空字符串
    GetJavaExePath = ""
End Function

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Text3.Enabled = False
Else
    Text3.Enabled = True
End If
End Sub

Private Sub Combo1_Change()
Text2.Text = GetJavaExePath(Combo1.Text)
End Sub

Private Sub Combo1_Click()
Text2.Text = GetJavaExePath(Combo1.Text)
End Sub

Private Sub Command1_Click()
LaunchMinecraft ".minecraft", Combo2.Text, Text4.Text
End Sub

Private Sub LoadMinecraftVersionsToCombo(spath As String)
    Dim sDir As String
    If Right(spath, 1) <> "\" Then spath = spath & "\"
    Combo2.Clear
    sDir = Dir(spath, vbDirectory)
    Do While sDir <> ""
        If sDir <> "." And sDir <> ".." Then
            If (GetAttr(spath & sDir) And vbDirectory) = vbDirectory Then
                Combo2.AddItem sDir
            End If
        End If
        sDir = Dir()
    Loop
    If Combo2.ListCount > 0 Then
        Combo2.ListIndex = 0
    End If
End Sub

Private Sub Form_Load()
Form1.Caption = "WMML " & App.Major & "." & App.Minor & "." & App.Revision
ScanJavaInstallations
LoadMinecraftVersionsToCombo App.Path & "\.minecraft\versions\"
End Sub
