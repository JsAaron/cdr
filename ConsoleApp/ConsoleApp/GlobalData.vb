Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

'全局数据对象
Module globalData

    Public state = False '状态'
    Public pagesize = ""
    Public errorlog As String '错误日志
    Public steps As String  '步骤

    Dim inputFiled = New JObject()
    Dim textData = New JObject()

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
                inputFiled.Add("bjnews", True)
            Case "bjnews"
                inputFiled.Add("url", True)
            Case "mobile"
                inputFiled.Add("phone", True)
            Case "phone"
                inputFiled.Add("mobile", True)
            Case "email"
                inputFiled.Add("qq", True)
            Case "qq"
                inputFiled.Add("email", True)
        End Select

        inputFiled.Add(key, True)


    End Function


    Public Function setValue(pageIndex, key, value)
        Dim type = TypeName(textData(key))
        If type = "Nothing" Then

        Else
            '有值的情况下，判断
            If type = "JValue" Or type = "JObject" Then
                'Console.WriteLine("重复key:" & key)
                Return True
            End If
        End If

        Dim json = New JObject()
        json.Add("pageIndex", pageIndex.ToString())
        json.Add("value", value.ToString())
        textData.Add(key, json)
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


        json.Add("errorlog", errorlog)
        json.Add("steps", steps)
        Return JsonConvert.SerializeObject(json)
    End Function

End Module
