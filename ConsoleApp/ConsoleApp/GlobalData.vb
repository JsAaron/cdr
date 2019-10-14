Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

'全局数据对象
Module globalData

    Public state = False '状态'
    Public pagesize = ""
    Public errorlog As String '错误日志
    Public steps As String  '步骤

    Dim list As New ArrayList
    Dim textData = New JObject()

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
        json.Add("pagesize", pagesize.ToString())
        json.Add("text", textData)
        json.Add("errorlog", errorlog)
        json.Add("steps", steps)
        Return JsonConvert.SerializeObject(json)
    End Function

End Module
