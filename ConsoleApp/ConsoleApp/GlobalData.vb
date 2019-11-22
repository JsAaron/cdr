Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Corel.Interop.VGCore


'全局数据对象
Module globalData

    Public state = False '状态'
    Public pagesize = ""
    Public errorlog As String '错误日志
    Public steps As String  '步骤
    Public totalPages = 0
    Public fnreturn = New JObject() '设置返回数据

    Dim recordlog As ArrayList = New ArrayList() '//记录一些有用数据
    Dim inputFiled = New JObject()
    Dim inputData = New JObject()



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
                inputFiled.Add("bjnews", "url+bjnews")
            Case "bjnews"
                inputFiled.Add("url", "url+bjnews")
            Case "mobile"
                inputFiled.Add("phone", "mobile+phone")
            Case "phone"
                inputFiled.Add("mobile", "mobile+phone")
            Case "email"
                inputFiled.Add("qq", "email+qq")
            Case "qq"
                inputFiled.Add("email", "email+qq")
        End Select

        inputFiled.Add(key, True)


    End Function



    '增加函数功能返回数据
    Public Function addFnReturn(key, value)
        Dim json = New JObject()
        fnreturn.Add(key, value.ToString())
    End Function



    Function saveData(pageIndex, key, value, overflow)
        Dim json = New JObject()

        '溢出了
        If overflow = True Then
            json.Add("overflow", True)
        End If

        json.Add("pageIndex", pageIndex.ToString())
        json.Add("value", value.ToString())
        inputData.Add(key, json)
    End Function



    '通过段落去匹配出key来
    Function valueTokey(pageIndex, name, tempShape)

        Dim p = tempShape.Text.Story.Paragraphs
        Dim v1 = p.Item(1).Text
        Dim v2 = p.Item(2).Text

        '电话手机一组
        If name = "mobile" Or name = "phone" Then
            saveData(pageIndex, "mobile", v1, False)
            saveData(pageIndex, "phone", v2, False)
        End If

        If name = "email" Or name = "qq" Then
            saveData(pageIndex, "email", v1, False)
            saveData(pageIndex, "qq", v2, False)
        End If

        If name = "url" Or name = "bjnews" Then
            saveData(pageIndex, "url", v1, False)
            saveData(pageIndex, "bjnews", v2, False)
        End If
    End Function

    '填充默认值给外部
    Function fillDefault(pageIndex, key, tempShape)

        '保存当前值
        saveData(pageIndex, key, tempShape.Text.Story.Text, False)

        '填充默认值
        Select Case key
            Case "url"
                saveData(pageIndex, "bjnews", "", False)
            Case "bjnews"
                saveData(pageIndex, "url", "", False)
            Case "mobile"
                saveData(pageIndex, "phone", "", False)
            Case "phone"
                saveData(pageIndex, "mobile", "", False)
            Case "email"
                saveData(pageIndex, "qq", "", False)
            Case "qq"
                saveData(pageIndex, "email", "", False)
        End Select

    End Function


    '保存获取的值
    '1 可能有分组组合的情况，所以需要找到字段合计，然后找到分组的数组
    Public Function saveValue(pageIndex As String, key As String, tempShape As Shape, determine As Determine, onlyFill As Boolean)

        Dim type = TypeName(inputData(key))
        If type = "Nothing" Then

        Else
            '有值的情况下，判断
            If type = "JValue" Or type = "JObject" Then
                'Console.WriteLine("重复key:" & key)
                Return True
            End If
        End If

        '如果只是填充默认值
        '仅针对图片的读
        If onlyFill Then
            saveData(pageIndex, key, "", False)
            Return True
        End If


        '是否存在需要分解的数据
        Dim hasRange = determine.getRangeScope(key)
        If hasRange = True Then
            '一个字段有上下2行,可能是被改变过，需要分解
            If tempShape.Text.Story.Paragraphs.Count = 2 Then
                valueTokey(pageIndex, key, tempShape)
            Else
                '填充默认值
                fillDefault(pageIndex, key, tempShape)
            End If
        Else
            '直接保存
            saveData(pageIndex, key, tempShape.Text.Story.Text, tempShape.Text.Overflow)
        End If


    End Function


    Function retrunData()
        Dim json = New JObject()
        json.Add("state", state.ToString())
        json.Add("totalpages", totalPages.ToString())

        If Param.cmdCommand = "get:text" Then
            json.Add("fileds", inputFiled)
            json.Add("text", inputData)
        End If

        If Param.cmdCommand = "get:pageSize" Then
            json.Add("pagesize", pagesize.ToString())
        End If

        json.Add("fnreturn", fnReturn)
        json.Add("recordlog", JsonConvert.SerializeObject(recordlog))
        json.Add("errorlog", errorlog)
        json.Add("steps", steps)
        Return JsonConvert.SerializeObject(json)
    End Function

End Module
