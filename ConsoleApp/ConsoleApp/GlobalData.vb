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
