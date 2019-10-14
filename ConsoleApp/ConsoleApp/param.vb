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


    '获取参数是有值
    Function hasValue(key)
        '如果是JObject对象在去判断
        If TypeName(cmdExternalData(key)) = "JObject" Then
            If Len(cmdExternalData(key)("value")) > 0 Then
                Return True
            End If
        End If
    End Function


    '获取外部参数的值
    Function getExternalValue(key)
        Return cmdExternalData(key)("value")
    End Function



    Sub decodeURI(cmdExternalData, key)
        If cmdExternalData(key) <> "" Then
            Dim e = CreateObject("MSScriptControl.ScriptControl")
            e.Language = "javascript"
            cmdExternalData(key) = e.Eval("decodeURI('" & cmdExternalData(key) & "')")
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
            ElseIf count = 2 Then
                cmdExternalData = JsonConvert.DeserializeObject(args(1))
                decodeURI(cmdExternalData, "logo")
                decodeURI(cmdExternalData, "qrcode")
            ElseIf count = 3 Then
                cmdExternalData = JsonConvert.DeserializeObject(args(1))
                decodeURI(cmdExternalData, "logo")
                decodeURI(cmdExternalData, "qrcode")
                cmdPath = args(2)
            End If
        ElseIf cmdCommand = "set:style" Then
            '参数不够
            If count = 1 Then
                ' globalData.errorlog = "必须传递样式路径参数"
            End If

            If count = 2 Then
                cmdStylePath = args(1)
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
