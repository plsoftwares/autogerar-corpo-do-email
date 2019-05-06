Attribute VB_Name = "Module1"
Sub Header()
    Dim objOutlook As Object
    Dim objMail As Object
    
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objMail = objOutlook.CreateItem(0)

    With objMail
        .HTMLBody = "Nome: " + usuarioRede + " | NR: " + Right(String(10, "0") & CStr(serial), 10) + "/" + CStr(Year(Now)) + "<br/>" + Format(Now(), "Long Date") + "<br/> <br/>"
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



