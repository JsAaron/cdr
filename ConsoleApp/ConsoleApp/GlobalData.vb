Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Corel.Interop.VGCore


'全局数据对象
Module globalData

    Public state = False '状态'
    Public pagesize = ""
    Public errorlog As String '错误日志
    Public steps As String  '步骤


    Dim recordlog As ArrayList = New ArrayList() '//记录一些有用数据
    Dim inputFiled = New JObject()
    Dim textData = New JObject()


    '增加日志记录
    Public Function addRecord(key)
        If Not recordlog.Contains(key) Then
            recordlog.Add(key)
        End If
    End Function


    '保存input数组显示的字段
    Public Function saveInputFiled(key)
        Dim type = TypeName(inputFiled(key))

        If type = "Nothing" Then

        Else
            '有值的情况下，判断
            If type = "JValue" Or type = "JObject" Then
                'Console.WriteLine("重复key:" & key)
                Return True
            End If
        End If

        '相关的情况处理
        Select Case key
            Case "url"
                inputFiled.Add("bjnews", "url")
            Case "bjnews"
                inputFiled.Add("url", "bjnews")
            Case "mobile"
                inputFiled.Add("phone", "mobile")
            Case "phone"
                inputFiled.Add("mobile", "phone")
            Case "email"
                inputFiled.Add("qq", "email")
            Case "qq"
                inputFiled.Add("email", "qq")
        End Select

        inputFiled.Add(key, True)


    End Function



    Function saveData(pageIndex, key, value)
        Dim json = New JObject()
        json.Add("pageIndex", pageIndex.ToString())
        json.Add("value", value.ToString())
        textData.Add(key, json)
    End Function


    '保存获取的值
    '1 可能有分组组合的情况，所以需要找到字段合计，然后找到分组的数组
    Public Function saveValue(pageIndex As String, key As String, tempShape As Shape, determine As Determine)
        Dim type = TypeName(textData(key))
        If type = "Nothing" Then

        Else
            '有值的情况下，判断
            If type = "JValue" Or type = "JObject" Then
                'Console.WriteLine("重复key:" & key)
                Return True
            End If
        End If

        '是否存在需要分解的数据
        Dim hasRange = determine.getRangeScope(key)
        If hasRange = True Then
            '一个字段有上下2行,可能是被改变过，需要分解
            If tempShape.Text.Story.Paragraphs.Count = 2 Then
                Dim p = tempShape.Text.Story.Paragraphs
                For i = 1 To p.Count
                    Console.WriteLine(p.Item(i).Text)
                Next
            Else
                '一行的情况下，直接保存
                saveData(pageIndex, key, tempShape.Text.Story.Text)
            End If
        Else
            '直接保存
            saveData(pageIndex, key, tempShape.Text.Story.Text)
        End If


    End Function


    Function retrunData()
        Dim json = New JObject()
        json.Add("state", state.ToString())

        If Param.cmdCommand = "get:text" Then
            json.Add("fileds", inputFiled)
            json.Add("text", textData)
        End If

        If Param.cmdCommand = "get:pageSize" Then
            json.Add("pagesize", pagesize.ToString())
        End If

        json.Add("recordlog", JsonConvert.SerializeObject(recordlog))
        json.Add("errorlog", errorlog)
        json.Add("steps", steps)
        Return JsonConvert.SerializeObject(json)
    End Function

End Module
