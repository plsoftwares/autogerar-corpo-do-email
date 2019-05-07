Sub Header()
    Dim objOutlook As Object
    Dim objMail As Object
    
    Dim email As String
    email = ""
    
    verifyAno
    
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

Function verifyAno() As Integer
    Dim iArq As Long
    Dim file As String
    Dim fileExists As Boolean
    Dim strTextLine
    Dim lineArq
    
    index = 1
    iArq = FreeFile
    file = "C:/Macro/headerIndex.txt"
    fileExists = False
    lineArq = 1
    verifyAno = 0
    
    If checkFileExists(file) Then
        Open file For Input As iArq
        
        Do While Not EOF(iArq)
            Line Input #iArq, strTextLine
            If lineArq = 1 Then
                verifyAno = strTextLine
            End If
            lineArq = lineArq + 1
        Loop
    End If
    

    Close #iArq
    If Not checkFileExists(file) Or verifyAno <> Year(Now) Then
        Open file For Output As iArq
        Print #iArq, Year(Now)
        Print #iArq, 1
        verifyAno = Year(Now)

        Close #iArq
    End If
End Function

Function serial() As Integer
    Dim iArq As Long
    Dim index As Long
    Dim file As String
    Dim fileExists As Boolean
    Dim strTextLine
    Dim line
    
    iArq = FreeFile
    file = "C:/Macro/headerIndex.txt"
    fileExists = False
    line = 1
    
    Open file For Input As iArq
    
    Do While Not EOF(iArq)
            Line Input #iArq, strTextLine
        If line = 2 Then
            serial = strTextLine
        End If
        line = line + 1
    Loop

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
