Imports System.IO
Imports System.Text
Imports Corel.Interop.VGCore
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

'定义参数
Module Param
    Public cmdCommand As String
    Public cmdPath As String
    Public cmdExternalData

    Private tempData

    Public cmdSavePath As String

    '打印数据
    Public cmdPrintSettings

    '导出图片设置
    Public cmdExportSettings


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
            '参数
            cmdPrintSettings = JsonConvert.DeserializeObject(decodePath(args(1)))
            '有路径的方法
            If count = 3 Then
                cmdExternalData = JsonConvert.DeserializeObject(args(2))
                cmdPrintSettings("Save") = decodePath(cmdExternalData("Save"))
                cmdPrintSettings("Load") = decodePath(cmdExternalData("Load"))
            End If
        ElseIf cmdCommand = "import" Then
            cmdExternalData = JsonConvert.DeserializeObject(args(1))
            cmdExternalData("path") = decodePath(cmdExternalData("path"))
        ElseIf cmdCommand = "save" Then
            cmdExternalData = JsonConvert.DeserializeObject(args(1))
            cmdExternalData("path") = decodePath(cmdExternalData("path"))
        ElseIf cmdCommand = "export-image" Then
            '导出图片
            cmdExportSettings = JsonConvert.DeserializeObject(decodePath(args(1)))
            '有路径的方法
            If count = 3 Then
                tempData = JsonConvert.DeserializeObject(args(2))
                cmdExportSettings("FileName") = decodePath(tempData("FileName"))
            End If
        ElseIf cmdCommand = "insert-image" Then
            '插入图片
            cmdExportSettings = JsonConvert.DeserializeObject(decodePath(args(1)))

            '有路径的方法
            If count = 3 Then
                tempData = JsonConvert.DeserializeObject(args(2))
                cmdExportSettings("FileName") = decodePath(tempData("FileName"))
            End If
        ElseIf cmdCommand = "SaveAsCopy" Then
            '保存文档
            tempData = JsonConvert.DeserializeObject(args(1))
            cmdSavePath = decodePath(tempData("FileName"))
        End If

        ' Console.WriteLine(cmdExternalData)
        ' globalData.steps = "解析参数完成"
    End Sub


End Module
