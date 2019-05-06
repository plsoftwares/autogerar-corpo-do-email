VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)
Dim xPrompt As String
Dim xOkOrCancel As Integer
On Error Resume Next
        Dim iArq As Long
        Dim file As String
        Dim strTextLine
        Dim Count As Integer
        
        
        iArq = FreeFile
        file = "C:/Macro/headerIndex.txt"
        fileExists = False

        Open file For Input As iArq
            
        Do While Not EOF(iArq)
        Line Input #iArq, strTextLine
        Loop
        
        Count = (CInt(strTextLine) + 1)
        Close #iArq
        
        Open file For Output As iArq
        Print #iArq, Count
        
        Close #iArq
End Sub
