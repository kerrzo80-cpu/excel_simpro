Attribute VB_Name = "AI_Build_Scope"
Option Explicit

Public Sub AI_Build_Scope()
    Dim worksDescription As String
    Dim commandText As String
    Dim repoDir As String
    Dim logFile As String

    worksDescription = InputBox("Describe the works:" & vbCrLf & vbCrLf & _
                                "Example: toilet not filling" & vbCrLf & _
                                "Example: move radiator 3 metres", _
                                "AI Build Scope")

    If Trim(worksDescription) = "" Then Exit Sub

    repoDir = Environ("HOME") & "/Downloads/excel_simpro"
    logFile = Environ("HOME") & "/Desktop/ai_scope_last_run.txt"

    commandText = "/bin/zsh -lc " & Chr(34) & _
                  "cd " & ShellQuote(repoDir) & " && " & _
                  "python3 ai_scope_insert_v7.py --insert --yes " & ShellQuote(worksDescription) & _
                  " > " & ShellQuote(logFile) & " 2>&1" & _
                  Chr(34)

    Shell commandText, vbNormalFocus

    MsgBox "AI Build Scope has started." & vbCrLf & vbCrLf & _
           "Check the workbook shortly, then check:" & vbCrLf & _
           logFile, vbInformation, "AI Build Scope"
End Sub

Private Function ShellQuote(ByVal value As String) As String
    ShellQuote = "'" & Replace(value, "'", "'\''") & "'"
End Function
