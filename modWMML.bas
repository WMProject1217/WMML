Attribute VB_Name = "modWMML"
Option Explicit

' �ú�������������� Minecraft
' (string)mcPath .minecraft�ļ���·��
' (string)versionName �汾����
' (string)playerName �������
Public Sub LaunchMinecraft(mcPath As String, versionName As String, playerName As String)
    On Error GoTo ErrorHandler
    
    ' ��׼��·��
    If Right(mcPath, 1) <> "\" Then mcPath = mcPath & "\"
    
    ' ��ȡ�汾json�ļ�
    Dim versionJsonPath As String
    Dim jsonContent As String
    versionJsonPath = mcPath & "versions\" & versionName & "\" & versionName & ".json"
    jsonContent = ReadTextFile(versionJsonPath)
    
    ' ����JSON
    Dim versionJson As Variant
    ParseJSONString2 jsonContent, versionJson
    
    ' ��ȡ����
    Dim mainClass As String
    mainClass = versionJson("mainClass")
    
    ' ������·��
    Dim libraries As String
    libraries = BuildLibrariesPath(mcPath, versionJson)
    
    ' ������Ϸ����
    Dim gameArgs As String
    gameArgs = BuildGameArguments(mcPath, versionName, playerName, versionJson)
    
    ' ����Java����
    Dim javaCommand As String
    Dim conststr As String
    conststr = "-Dfile.encoding=GB18030 -Dsun.stdout.encoding=GB18030 -Dsun.stderr.encoding=GB18030 -Djava.rmi.server.useCodebaseOnly=true -Dcom.sun.jndi.rmi.object.trustURLCodebase=false -Dcom.sun.jndi.cosnaming.object.trustURLCodebase=false -Dlog4j2.formatMsgNoLookups=true -Dlog4j.configurationFile=.minecraft\versions\" & versionName & "\log4j2.xml "
    conststr = conststr & "-Dminecraft.client.jar=.minecraft\versions\" & versionName & "\" & versionName & ".jar -XX:+UnlockExperimentalVMOptions -XX:+UseG1GC -XX:G1NewSizePercent=20 -XX:G1ReservePercent=20 -XX:MaxGCPauseMillis=50 -XX:G1HeapRegionSize=32m -XX:-UseAdaptiveSizePolicy -XX:-OmitStackTraceInFastThrow -XX:-DontCompileHugeMethods -Dfml.ignoreInvalidMinecraftCertificates=true "
    conststr = conststr & "-Dfml.ignorePatchDiscrepancies=true -XX:HeapDumpPath=MojangTricksIntelDriversForPerformance_javaw.exe_minecraft.exe.heapdump -Djava.library.path=.minecraft\versions\" & versionName & "\natives-windows-x86_64 "
    conststr = conststr & "-Djna.tmpdir=.minecraft\versions\" & versionName & "\natives-windows-x86_64 -Dorg.lwjgl.system.SharedLibraryExtractPath=.minecraft\versions\" & versionName & "\natives-windows-x86_64 -Dio.netty.native.workdir=.minecraft\versions\" & versionName & "\natives-windows-x86_64 -Dminecraft.launcher.brand=WMML -Dminecraft.launcher.version=" & App.Major & "." & App.Minor & "." & App.Revision
    javaCommand = "cmd /K java " & conststr & " -cp """ & libraries & """ " & mainClass & " " & gameArgs
    
    ' ִ������
    Form1.Text1.Text = javaCommand
    Shell javaCommand, vbNormalFocus
    
    Exit Sub
    
ErrorHandler:
    MsgBox "����Minecraftʱ����: " & Err.Description, vbCritical, "����"
End Sub

' ������·��
Private Function BuildLibrariesPath(mcPath As String, versionJson As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim libs As Variant
    Dim lib As Variant
    Dim libPath As String
    Dim result As String
    
    ' ������Ӱ汾jar�ļ�
    result = mcPath & "versions\" & versionJson("id") & "\" & versionJson("id") & ".jar"
    
    ' ������п��ļ�
    libs = versionJson("libraries")
    For Each lib In libs
        ' ������(�����������ڵ�ǰϵͳ�Ŀ�)
        If Not CheckLibraryRules(lib) Then GoTo NextLib
        
        ' ��ȡ��·��
        libPath = GetLibraryPath(mcPath, lib)
        
        ' ��ӵ����
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

' �������
Private Function CheckLibraryRules(lib As Variant) As Boolean
    On Error Resume Next
    
    ' ���û��rules���֣������ǰ���
    If IsEmpty(lib("rules")) Then
        CheckLibraryRules = True
        Exit Function
    End If
    
    Dim rule As Variant
    Dim osName As String
    Dim osArch As String
    
    ' ��ȡ��ǰϵͳ��Ϣ
    #If Win64 Then
        osArch = "x86_64"
    #Else
        osArch = "x86"
    #End If
    
    osName = "windows"
    
    ' ������й���
    For Each rule In lib("rules")
        ' ���action
        If rule("action") = "allow" Then
            ' ���û��os���֣�������
            If IsEmpty(rule("os")) Then
                CheckLibraryRules = True
                Exit Function
            End If
            
            ' ���os����
            If rule("os")("name") = osName Then
                ' �����arch���������arch
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
            ' ���û��os���֣�������
            If IsEmpty(rule("os")) Then
                CheckLibraryRules = False
                Exit Function
            End If
            
            ' ���os����
            If rule("os")("name") = osName Then
                CheckLibraryRules = False
                Exit Function
            End If
        End If
    Next
    
    ' Ĭ������
    CheckLibraryRules = True
End Function

' ��ȡ��·��
Private Function GetLibraryPath(mcPath As String, lib As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim parts As Variant
    Dim artifactPath As String
    Dim nativePath As String
    Dim i As Integer
    
    ' ����������
    parts = Split(lib("name"), ":")
    
    ' ��������·��
    artifactPath = mcPath & "libraries\" & Replace(parts(0), ".", "\") & "\" & parts(1) & "\" & parts(2) & "\" & parts(1) & "-" & parts(2)
    
    ' ����Ƿ���natives
    If Not IsEmpty(lib("natives")) Then
        ' ��ȡwindowsƽ̨��native������
        If Not IsEmpty(lib("natives")("windows")) Then
            Dim classifier As String
            classifier = lib("natives")("windows")
            classifier = Replace(classifier, "${arch}", IIf(EnvironmentIs64Bit(), "64", "32"))
            
            ' ����native·��
            nativePath = artifactPath & "-" & classifier & ".jar"
            
            ' ����ļ��Ƿ����
            If FileExists(nativePath) Then
                GetLibraryPath = nativePath
                Exit Function
            End If
        End If
    End If
    
    ' ���û��natives���Ҳ���native�ļ���ʹ����ͨjar
    artifactPath = artifactPath & ".jar"
    If FileExists(artifactPath) Then
        GetLibraryPath = artifactPath
        Exit Function
    End If
    
ErrorHandler:
    GetLibraryPath = ""
End Function

' ������Ϸ����
Private Function BuildGameArguments(mcPath As String, versionName As String, playerName As String, versionJson As Variant) As String
    On Error GoTo ErrorHandler
    
    Dim args As String
    Dim assetsPath As String
    Dim versionType As String
    
    ' ����Ĭ��ֵ
    assetsPath = mcPath & "assets"
    versionType = "VB6 Launcher"
    
    ' ���Դ�json��ȡassets��versionType
    If Not IsEmpty(versionJson("assets")) Then
        assetsPath = mcPath & "assets\"
    End If
    
    If Not IsEmpty(versionJson("type")) Then
        versionType = versionJson("type")
    End If
    
    ' ��������
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
    
    ' ����Ƿ���minecraftArguments(�ɰ汾)
    If Not IsEmpty(versionJson("minecraftArguments")) Then
        args = versionJson("minecraftArguments") & " " & args
    End If
    
    ' ����Ƿ���arguments(�°汾)
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

' ��������: ��ȡ�ı��ļ�
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

' ��������: ����ļ��Ƿ����
Private Function FileExists(filePath As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(filePath) <> "")
End Function

' ��������: ����Ƿ���64λ����
Private Function EnvironmentIs64Bit() As Boolean
    #If Win64 Then
        EnvironmentIs64Bit = True
    #Else
        EnvironmentIs64Bit = False
    #End If
End Function

