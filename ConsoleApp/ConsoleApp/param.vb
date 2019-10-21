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

    '判断参数是否有key
    Function hasKey(key)
        '为空
        If TypeName(cmdExternalData) = "Nothing" Then
            Return False
        End If
    End Function


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
            Return False
        End If
        Return cmdExternalData(key)("value")
    End Function


    Sub decodeURI(key)
        If Param.hasValue(key) Then
            Dim e = CreateObject("MSScriptControl.ScriptControl")
            e.Language = "javascript"
            cmdExternalData(key)("value") = e.Eval("decodeURI('" & Param.getExternalValue(key) & "')")
        End If
    End Sub


    '参数解析
    Public Sub parseCommand(command)
        Dim args() = Split(command, " ")
        Dim count = args.Count
        cmdCommand = args(0)
        If cmdCommand = "open" Then
            If count = 2 Then
                cmdPath = args(1)
            End If
        ElseIf cmdCommand = "get:pageSize" Then
            If count = 2 Then
                cmdPath = args(1)
            End If
        ElseIf cmdCommand = "get:fontJson" Then
            If count = 2 Then
                cmdPath = args(1)
            End If
        ElseIf cmdCommand = "get:text" Then
            If count = 2 Then
                cmdPath = args(1)
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
                cmdExternalData = JsonConvert.DeserializeObject(args(1))
                decodeURI("logo")
                decodeURI("logo2")
                decodeURI("qrcode")
                cmdPath = args(2)
            End If
        ElseIf cmdCommand = "set:style" Then
            '参数不够
            If count = 1 Then
                globalData.errorlog = "必须传递样式路径参数"
            End If

            If count = 2 Then
                cmdStylePath = args(1)
                Dim e = CreateObject("MSScriptControl.ScriptControl")
                e.Language = "javascript"
                cmdStylePath = e.Eval("decodeURI('" & cmdStylePath & "')")
            End If

            '设置样式
            If count = 3 Then
                cmdPath = args(1)
            End If
        End If

        ' Console.WriteLine(cmdExternalData)
        ' globalData.steps = "解析参数完成"
    End Sub


End Module
