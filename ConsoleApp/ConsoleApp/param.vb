Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json

'定义参数
Module Param
    Public cmdCommand As String
    Public cmdPath As String
    Public cmdExternalData

    '设置数据
    Public cmdPrintSettings


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

        '打印
        If cmdCommand = "print" Then
            cmdPrintSettings = JsonConvert.DeserializeObject(decodePath(args(1)))
        ElseIf cmdCommand = "import" Then
            '导出
            cmdExternalData = JsonConvert.DeserializeObject(args(1))
            cmdExternalData("path") = decodePath(cmdExternalData("path"))
        ElseIf cmdCommand = "save" Then
            '保存
            cmdExternalData = JsonConvert.DeserializeObject(args(1))
            cmdExternalData("path") = decodePath(cmdExternalData("path"))
        End If

        ' Console.WriteLine(cmdExternalData)
        ' globalData.steps = "解析参数完成"
    End Sub


End Module
