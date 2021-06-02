Attribute VB_Name = "Module1"
Option Explicit

Sub ColorFormulae()

Dim regex As Object, rangeobj As Object, selectedcell, hasformula As Boolean
Dim indcol As New Collection
Dim operators As Object, operator
Dim i As Integer
Dim midstr As String, listtxt as string

Set regex = CreateObject("VBScript.RegExp")
regex.Pattern = "[^a-zA-Z0-9à-ú!$@:\[\]_]"
regex.Global = True

For Each selectedcell In Selection
    Set operators = regex.Execute(selectedcell.Formula)
    Set indcol = New Collection
    hasformula = False
    
    For Each operator In operators
        indcol.Add operator.firstindex
    Next
    
    If indcol.Count = 0 Then indcol.Add 0 'in case there is nothing except the first "=", assume it to be 0
    
    If indcol(indcol.Count) < Len(selectedcell.Formula) Then indcol.Add Len(selectedcell.Formula)
    
    For i = 1 To indcol.Count - 1
        midstr = Mid(selectedcell.Formula, indcol(i) + 2, indcol(i + 1) - indcol(i) - 1)
        
        If InStr(1, midstr, "[@") > 0 Then 'if reference is part of list
            listtxt = Replace(midstr, "@", "[#Headers],[") & "]" 'get address of the table header above the referenced cell
            midstr = Range(listtxt).Offset(selectedcell.Row() - Range(listtxt).Row(), 0).Address 'get the address relative to the analyzed cell        
        End If
                    
        On Error Resume Next
            Set rangeobj = Range(midstr) 'in case of error, will return Nothing or previous value (which is set to nothing in code below)
            
            If Not (rangeobj Is Nothing) Then
                hasformula = True
            End If
            
            Set rangeobj = Nothing 'required to clear existing, otherwise will not return to nothing in the set above
        Err.Clear
    
    Next i

    Select Case hasformula
    
        Case True
            selectedcell.Font.Color = RGB(0, 0, 255)
        Case False
            selectedcell.Font.Color = RGB(255, 0, 0)
    End Select
    
    Set operators = Nothing
Next selectedcell

Set rangeobj = Nothing
Set regex = Nothing
    
End Sub
