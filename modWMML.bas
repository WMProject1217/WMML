Attribute VB_Name = "modWMML"
Option Explicit

' 该函数生成命令并启动 Minecraft
' (string)mcPath .minecraft文件夹路径
' (string)versionName 版本名称
' (string)playerName 玩家名称
Public Sub LaunchMinecraft(mcPath As String, versionName As String, playerName As String)
    On Error GoTo ErrorHandler
    
    ' 标准化路径
    If Right(mcPath, 1) <> "\" Then mcPath = mcPath & "\"
    
    ' 读取版本json文件
    Dim versionJsonPath As String
    Dim jsonContent As String
    versionJsonPath = mcPath & "versions\" & versionName & "\" & versionName & ".json"
    jsonContent = ReadTextFile(versionJsonPath)
    
    ' 解析JSON
    Dim versionJson As Variant
    ParseJSONString2 jsonContent, versionJson
    
    ' 获取主类
    Dim mainClass As String
    mainClass = versionJson("mainClass")
    
    ' 构建库路径
    Dim libraries As String
    libraries = BuildLibrariesPath(mcPath, versionJson)
    
    ' 构建游戏参数
    Dim gameArgs As String
    gameArgs = BuildGameArguments(mcPath, versionName, playerName, versionJson)
    
    ' 构建Java命令
    Dim javaCommand As String
    Dim conststr As String
    conststr = "-Dfile.encoding=GB18030 -Dsun.stdout.encoding=GB18030 -Dsun.stderr.encoding=GB18030 -Djava.rmi.server.useCodebaseOnly=true -Dcom.sun.jndi.rmi.object.trustURLCodebase=false -Dcom.sun.jndi.cosnaming.object.trustURLCodebase=false -Dlog4j2.formatMsgNoLookups=true -Dlog4j.configurationFile=.minecraft\versions\" & versionName & "\log4j2.xml "
    conststr = conststr & "-Dminecraft.client.jar=.minecraft\versions\" & versionName & "\" & versionName & ".jar -XX:+UnlockExperimentalVMOptions -XX:+UseG1GC -XX:G1NewSizePercent=20 -XX:G1ReservePercent=20 -XX:MaxGCPauseMillis=50 -XX:G1HeapRegionSize=32m -XX:-UseAdaptiveSizePolicy -XX:-OmitStackTraceInFastThrow -XX:-DontCompileHugeMethods -Dfml.ignoreInvalidMinecraftCertificates=true "
    conststr = conststr & "-Dfml.ignorePatchDiscrepancies=true -XX:HeapDumpPath=MojangTricksIntelDriversForPerformance_javaw.exe_minecraft.exe.heapdump -Djava.library.path=.minecraft\versions\" & versionName & "\natives-windows-x86_64 "
    conststr = conststr & "-Djna.tmpdir=.minecraft\versions\" & versionName & "\natives-windows-x86_64 -Dorg.lwjgl.system.SharedLibraryExtractPath=.minecraft\versions\" & versionName & "\natives-windows-x86_64 -Dio.netty.native.workdir=.minecraft\versions\" & versionName & "\natives-windows-x86_64 -Dminecraft.launcher.brand=WMML -Dminecraft.launcher.version=" & App.Major & "." & App.Minor & "." & App.Revision
    javaCommand = "cmd /K java " & conststr & " -cp """ & libraries & """ " & mainClass & " " & gameArgs
    
    ' 执行命令
    Form1.Text1.Text = javaCommand
    Shell javaCommand, vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "启动Minecraft时出错: " & Err.Description, vbCritical, "错误"
End Sub

' 构建库路径
Private Function BuildLibrariesPath(mcPath As String, versionJson As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim libs As Variant
    Dim lib As Variant
    Dim libPath As String
    Dim result As String
    
    ' 首先添加版本jar文件
    result = mcPath & "versions\" & versionJson("id") & "\" & versionJson("id") & ".jar"
    
    ' 添加所有库文件
    libs = versionJson("libraries")
    For Each lib In libs
        ' 检查规则(跳过不适用于当前系统的库)
        If Not CheckLibraryRules(lib) Then GoTo NextLib
        
        ' 获取库路径
        libPath = GetLibraryPath(mcPath, lib)
        
        ' 添加到结果
        If libPath <> "" Then
            If result <> "" Then result = result & ";"
            result = result & libPath
        End If
        
NextLib:
    Next
    
    BuildLibrariesPath = result
    Exit Function
    
ErrorHandler:
    BuildLibrariesPath = ""
End Function

' 检查库规则
Private Function CheckLibraryRules(lib As Variant) As Boolean
    On Error Resume Next
    
    ' 如果没有rules部分，则总是包含
    If IsEmpty(lib("rules")) Then
        CheckLibraryRules = True
        Exit Function
    End If
    
    Dim rule As Variant
    Dim osName As String
    Dim osArch As String
    
    ' 获取当前系统信息
    #If Win64 Then
        osArch = "x86_64"
    #Else
        osArch = "x86"
    #End If
    
    osName = "windows"
    
    ' 检查所有规则
    For Each rule In lib("rules")
        ' 检查action
        If rule("action") = "allow" Then
            ' 如果没有os部分，则允许
            If IsEmpty(rule("os")) Then
                CheckLibraryRules = True
                Exit Function
            End If
            
            ' 检查os条件
            If rule("os")("name") = osName Then
                ' 如果有arch条件，检查arch
                If Not IsEmpty(rule("os")("arch")) Then
                    If rule("os")("arch") = osArch Then
                        CheckLibraryRules = True
                        Exit Function
                    Else
                        CheckLibraryRules = False
                        Exit Function
                    End If
                Else
                    CheckLibraryRules = True
                    Exit Function
                End If
            Else
                CheckLibraryRules = False
                Exit Function
            End If
        ElseIf rule("action") = "disallow" Then
            ' 如果没有os部分，则不允许
            If IsEmpty(rule("os")) Then
                CheckLibraryRules = False
                Exit Function
            End If
            
            ' 检查os条件
            If rule("os")("name") = osName Then
                CheckLibraryRules = False
                Exit Function
            End If
        End If
    Next
    
    ' 默认允许
    CheckLibraryRules = True
End Function

' 获取库路径
Private Function GetLibraryPath(mcPath As String, lib As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim parts As Variant
    Dim artifactPath As String
    Dim nativePath As String
    Dim i As Integer
    
    ' 解析库名称
    parts = Split(lib("name"), ":")
    
    ' 构建基本路径
    artifactPath = mcPath & "libraries\" & Replace(parts(0), ".", "\") & "\" & parts(1) & "\" & parts(2) & "\" & parts(1) & "-" & parts(2)
    
    ' 检查是否有natives
    If Not IsEmpty(lib("natives")) Then
        ' 获取windows平台的native分类器
        If Not IsEmpty(lib("natives")("windows")) Then
            Dim classifier As String
            classifier = lib("natives")("windows")
            classifier = Replace(classifier, "${arch}", IIf(EnvironmentIs64Bit(), "64", "32"))
            
            ' 构建native路径
            nativePath = artifactPath & "-" & classifier & ".jar"
            
            ' 检查文件是否存在
            If FileExists(nativePath) Then
                GetLibraryPath = nativePath
                Exit Function
            End If
        End If
    End If
    
    ' 如果没有natives或找不到native文件，使用普通jar
    artifactPath = artifactPath & ".jar"
    If FileExists(artifactPath) Then
        GetLibraryPath = artifactPath
        Exit Function
    End If
    
ErrorHandler:
    GetLibraryPath = ""
End Function

' 构建游戏参数
Private Function BuildGameArguments(mcPath As String, versionName As String, playerName As String, versionJson As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim args As String
    Dim assetsPath As String
    Dim versionType As String
    
    ' 设置默认值
    assetsPath = mcPath & "assets"
    versionType = "VB6 Launcher"
    
    ' 尝试从json获取assets和versionType
    If Not IsEmpty(versionJson("assets")) Then
        assetsPath = mcPath & "assets\"
    End If
    
    If Not IsEmpty(versionJson("type")) Then
        versionType = versionJson("type")
    End If
    
    ' 构建参数
    'args = "--username " & playerName & " " & _
    '       "--version " & versionName & " " & _
    '       "--gameDir " & mcPath & " " & _
    '       "--assetsDir " & assetsPath & " " & _
    '       "--assetIndex " & versionJson("assets") & " " & _
    '       "--uuid 00000000-0000-0000-0000-000000000000 " & _
    '       "--accessToken 00000000000000000000000000000000 " & _
    '       "--userType legacy " & _
    '       "--versionType " & versionType
    args = ""
    
    ' 检查是否有minecraftArguments(旧版本)
    If Not IsEmpty(versionJson("minecraftArguments")) Then
        args = versionJson("minecraftArguments") & " " & args
    End If
    
    ' 检查是否有arguments(新版本)
    If Not IsEmpty(versionJson("arguments")) Then
        Dim arg As Variant
        For Each arg In versionJson("arguments")("game")
            If VarType(arg) = vbString Then
                args = args & " " & arg
            End If
        Next
    End If
    
    args = Replace(args, "${auth_player_name}", playerName)
    args = Replace(args, "${version_name}", versionName)
    args = Replace(args, "${game_directory}", mcPath)
    args = Replace(args, "${assets_root}", assetsPath)
    args = Replace(args, "${assets_index_name}", versionJson("assets"))
    args = Replace(args, "${auth_uuid}", "00000000-0000-0000-0000-000000000000")
    args = Replace(args, "${auth_access_token}", "00000000000000000000000000000000")
    'args = Replace(args, "${clientid}", "")
    'args = Replace(args, "${auth_xuid}", "")
    args = Replace(args, "${user_type}", "legacy")
    args = Replace(args, "${version_type}", Chr(34) & "WMML " & App.Major & "." & App.Minor & "." & App.Revision & Chr(34))
    
    BuildGameArguments = args
    Exit Function
    
ErrorHandler:
    BuildGameArguments = ""
End Function

' 辅助函数: 读取文本文件
Private Function ReadTextFile(filePath As String) As String
    On Error GoTo ErrorHandler
    
    Dim fileNum As Integer
    Dim content As String
    
    fileNum = FreeFile
    Open filePath For Input As #fileNum
    content = Input$(LOF(fileNum), fileNum)
    Close #fileNum
    
    ReadTextFile = content
    Exit Function
    
ErrorHandler:
    ReadTextFile = ""
End Function

' 辅助函数: 检查文件是否存在
Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
End Function

' 辅助函数: 检查是否是64位环境
Private Function EnvironmentIs64Bit() As Boolean
    #If Win64 Then
        EnvironmentIs64Bit = True
    #Else
        EnvironmentIs64Bit = False
    #End If
End Function

