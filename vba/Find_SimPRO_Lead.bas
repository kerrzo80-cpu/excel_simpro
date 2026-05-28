Attribute VB_Name = "Find_SimPRO_Lead"
Option Explicit

Public Sub Find_SimPRO_Lead()
    Dim cmd As String
    Dim repoDir As String
    Dim scriptPath As String

    repoDir = Environ$("HOME") & "/Downloads/excel_simpro"
    scriptPath = repoDir & "/find_simpro_leads.py"

    cmd = "cd " & Chr(34) & repoDir & Chr(34) & " && /usr/bin/python3 " & Chr(34) & scriptPath & Chr(34)

    On Error GoTo ErrHandler

    MsgBox "Searching simPRO Leads. Type the customer name or address in Final Quote cell B3 first.", vbInformation, "Find simPRO Lead"

    MacScript "do shell script " & Chr(34) & Replace(cmd, Chr(34), "\" & Chr(34)) & Chr(34)

    MsgBox "Lead search complete. If a match was found, B3:B5 has been updated.", vbInformation, "Find simPRO Lead"
    Exit Sub

ErrHandler:
    MsgBox "Find simPRO Lead failed:" & vbCrLf & Err.Description, vbCritical, "Find simPRO Lead"
End Sub
