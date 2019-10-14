Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

'全局数据对象
Module globalData

    Public state = False '状态'
    Public pagesize = ""
    Public errorlog As String '错误日志
    Public steps As String  '步骤

    '文本数据
    Private textData As JObject = New JObject()



    Public Function setValue(pageIndex, key, value)
        Dim json = New JObject()
        ' Console.WriteLine(key + "  " + value)
        json.Add("page", pageIndex.ToString())
        json.Add("value", value.ToString())

        ' Console.WriteLine(key & " " & value)

        textData.Add(key, json)

        Console.WriteLine(textData)
    End Function

    Function retrunData()





        Dim json = New JObject()
        json.Add("state", state.ToString())
        json.Add("pagesize", pagesize.ToString())
        json.Add("text", textData)
        json.Add("errorlog", errorlog)
        json.Add("steps", steps)
        ' Return JsonConvert.SerializeObject(json)
    End Function

End Module
