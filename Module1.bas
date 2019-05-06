Sub Header()
    Dim objOutlook As Object
    Dim objMail As Object
    
    Dim email As String
    email = ""
    
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)
    
    email = "Mensagem=<br/>"
    email = email + "<table width='100%' border='0'>"
        email = email + "<tr>"
            email = email + "<td width='33%'>" + Environ$("username") + "</td>"
            email = email + "<td width='33%' style='text-align:center;'>" + Right(String(4, "0") & CStr(serial), 4) + "/" + CStr(Year(Now)) + "</td>"
            email = email + "<td width='16%' style='text-align:right;'>" + Format(Now(), "ddmmyy") + "</td>"
            email = email + "<td width='16%' style='text-align:right;'>" + Format(Now(), "HH:MM") + "</td>"
        email = email + "</tr>"
    email = email + "</table>"
    email = email + "<br/><br/><br/>"


    With objMail
        .HTMLBody = email
        .Display
    End With
End Sub


Function serial() As Integer
    Dim iArq As Long
    Dim index As Long
    Dim file As String
    Dim fileExists As Boolean
    Dim strTextLine
    
    index = 1
    iArq = FreeFile
    file = "C:/Macro/headerIndex.txt"
    fileExists = False
    
    If checkFileExists(file) Then
        Open file For Input As iArq
        
        Do While Not EOF(iArq)
            Line Input #iArq, strTextLine
        Loop
        
        serial = strTextLine
    Else
        Open file For Output As iArq
        Print #iArq, 1
        serial = 1
    End If

    Close #iArq
End Function

Function checkFileExists(ByVal file As String) As Boolean

    checkFileExists = False
    
    'Identifica se o arquivo existe
    If Not Dir(file) = vbNullString Then
        checkFileExists = True
    End If
    
End Function

Function usuarioRede() As String
    Dim GetUserN
    Dim ObjNetwork
    Set ObjNetwork = CreateObject("WScript.Network")
    GetUserN = ObjNetwork.UserName
    usuarioRede = GetUserN
End Function



