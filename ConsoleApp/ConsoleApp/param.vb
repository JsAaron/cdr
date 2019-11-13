Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json

'定义参数
Module Param

    Public cmdCommand As String = "get:text"
    Public cmdPath As String
    Public cmdStylePath As String
    Public cmdExternalData
    '活动页面
    Public cmdActivePagte
    '字体名字
    Public cmdFontName As String

    '获取参数是有值
    Function hasValue(key)
        '为空
        If TypeName(cmdExternalData) = "Nothing" Then
            Return False
        End If

        '如果是JObject对象在去判断
        If TypeName(cmdExternalData(key)) = "JObject" Then
            If Len(cmdExternalData(key)("value")) > 0 Then
                Return True
            End If
        End If
    End Function


    '获取外部参数的值
    Function getExternalValue(key)
        '为空
        If TypeName(cmdExternalData) = "Nothing" Then
            Return ""
        End If

        Dim valueType = TypeName(cmdExternalData(key))
        '数据为空是删除
        If valueType = "Nothing" Then
            Return ""
        End If

        'JObject对象
        Return cmdExternalData(key)("value")

    End Function


    Sub decodeURI(key)
        If Param.hasValue(key) Then
            cmdExternalData(key)("value") = decodePath(cmdExternalData(key)("value"))
        End If
    End Sub


    Function decodePath(value)
        Dim e = CreateObject("MSScriptControl.ScriptControl")
        e.Language = "javascript"
        Return e.Eval("decodeURIComponent('" & value & "')")
    End Function

    '参数解析
    Public Sub parseCommand(command)
        Dim args() = Split(command, " ")
        Dim count = args.Count

        cmdCommand = args(0)
        If cmdCommand = "open" Then
            If count = 2 Then
                cmdPath = decodePath(args(1))
            End If
        ElseIf cmdCommand = "get:pageSize" Then
            If count = 2 Then
                cmdPath = decodePath(args(1))
            End If
        ElseIf cmdCommand = "get:fontJson" Then
            If count = 2 Then
                cmdPath = decodePath(args(1))
            End If
        ElseIf cmdCommand = "get:text" Then
            '如果是2个参数
            If count = 2 Then
                '如果是单页设置：(get:pageSize,page,path)
                '默认页面数不会多余100个，都算page的参数
                If Len(args(1)) < 3 Then
                    cmdActivePagte = args(1)
                Else
                    cmdPath = decodePath(args(1))
                End If
            End If

            '如果是3个参数
            If count = 3 Then
                '(get:pageSize,page,path)
                cmdActivePagte = args(1)
                cmdPath = decodePath(args(2))
            End If

        ElseIf cmdCommand = "set:text" Then
            If count = 1 Then
                globalData.errorlog = "没有传递设置参数"
            ElseIf count = 2 Then
                cmdExternalData = JsonConvert.DeserializeObject(args(1))
                decodeURI("logo")
                decodeURI("logo2")
                decodeURI("qrcode")
            ElseIf count = 3 Then
                '如果第3个参数，是页码
                If Len(args(2)) < 3 Then
                    cmdExternalData = JsonConvert.DeserializeObject(args(1))
                    decodeURI("logo")
                    decodeURI("logo2")
                    decodeURI("qrcode")
                    cmdActivePagte = args(2)
                Else
                    cmdExternalData = JsonConvert.DeserializeObject(args(1))
                    decodeURI("logo")
                    decodeURI("logo2")
                    decodeURI("qrcode")
                    cmdPath = decodePath(args(2))
                End If

            ElseIf count = 4 Then
                cmdExternalData = JsonConvert.DeserializeObject(args(1))
                decodeURI("logo")
                decodeURI("logo2")
                decodeURI("qrcode")
                cmdActivePagte = args(2)
                cmdPath = decodePath(args(3))
            End If
        ElseIf cmdCommand = "set:style" Then
            '参数不够
            If count = 1 Then
                globalData.errorlog = "必须传递样式路径参数"
            End If

            If count = 2 Then
                cmdStylePath = decodePath(args(1))
            End If

            '设置样式
            If count = 3 Then
                cmdPath = decodePath(args(1))
            End If

        ElseIf cmdCommand = "set:font" Then
            If count = 1 Then
                globalData.errorlog = "必须传递字体名"
            End If
            cmdFontName = args(1)
        End If

        ' Console.WriteLine(cmdExternalData)
        ' globalData.steps = "解析参数完成"
    End Sub


End Module
