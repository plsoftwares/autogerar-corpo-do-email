Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
Dim xPrompt As String
Dim xOkOrCancel As Integer
On Error Resume Next
    incrementSerial
End Sub

Function incrementSerial() As Integer
    Dim iArq As Long
    Dim file As String
    Dim fileExists As Boolean
    Dim strTextLine
    Dim lineArq
    Dim serialValue
    
    index = 1
    iArq = FreeFile
    file = "C:/Macro/headerIndex.txt"
    fileExists = False
    lineArq = 1
    yearIndex = 0
    serialValue = serial
    
    If checkFileExists(file) Then
        Open file For Input As iArq
        
        Do While Not EOF(iArq)
            Line Input #iArq, strTextLine
            If lineArq = 1 Then
                yearIndex = strTextLine
            End If
            lineArq = lineArq + 1
        Loop
    End If
    Close #iArq
    
    If CInt(yearIndex) <> CInt(year(Now)) Then
        serialValue = 1
    Else
        serialValue = serialValue + 1
    End If
    
    Open file For Output As iArq
    Print #iArq, year(Now)
    Print #iArq, serialValue

    Close #iArq
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
