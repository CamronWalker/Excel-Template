Attribute VB_Name = "Template"
' Written by Camron unless otherwise noted
'
' REQUIRES: (Goto Tools > References)
' Microsoft Visual Basic For Applications Extensibility 5.3
' Microsoft Scripting Runtime
'
' CHECK THIS BOX IN SETTINGS:
' File > Options > Trust Center > Trust Center Settings > Macro Settings > Check "Trust access to the VBA project object model"


Public Sub ExportModulesForGit()
    ' https://www.rondebruin.nl/win/s9/win002.htm
    Dim bExport As Boolean
    Dim wkbSource As Excel.Workbook
    Dim szSourceWorkbook As String
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    AddLog ("VBA Export Start")

    ''' The code modules will be exported in a folder named.
    ''' VBAProjectFiles in the Documents folder.
    ''' The code below create this folder if it not exist
    ''' or DELETE ALL FILES in the folder if it exist.
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    
    On Error Resume Next
        Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0

    ''' NOTE: This workbook must be open in Excel.
    szSourceWorkbook = ActiveWorkbook.Name
    Set wkbSource = Application.Workbooks(szSourceWorkbook)
    
    If wkbSource.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to export the code"
    Exit Sub
    End If
    
    szExportPath = FolderWithVBAProjectFiles & "\"
    
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        
        bExport = True
        szFileName = cmpComponent.Name

        ''' Concatenate the correct filename for export.
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                ''' This is a worksheet or workbook object.
                ''' Don't try to export.
                bExport = False
        End Select
        
        If bExport Then
            ''' Export the component to a text file.
            cmpComponent.Export szExportPath & szFileName
            
        ''' remove it from the project if you want
        '''wkbSource.VBProject.VBComponents.Remove cmpComponent
        
        End If
   
    Next cmpComponent

    AddLog ("VBA Export Completed With No Errors")
End Sub


Function FolderWithVBAProjectFiles() As String
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String

    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")

    SpecialPath = Application.ActiveWorkbook.Path

    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    
    If fso.FolderExists(SpecialPath & "includes") = False Then
        On Error Resume Next
        MkDir SpecialPath & "includes"
        On Error GoTo 0
    End If
    
    If fso.FolderExists(SpecialPath & "includes\src") = False Then
        On Error Resume Next
        MkDir SpecialPath & "includes\src"
        On Error GoTo 0
    End If
    
    If fso.FolderExists(SpecialPath & "includes\src") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "includes\src"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
    
End Function

Sub AddLog(LogEntry As String)
    Dim filesys, filetxt
    Dim logFileName As String: logFileName = Left(ActiveWorkbook.Name, Len(ActiveWorkbook.Name) - 5)
    If Range("Logging").Value = True Then
        Set filesys = CreateObject("scripting.filesystemobject")
        If Not filesys.FileExists(Application.ActiveWorkbook.Path & "\" & logFileName & " Log.txt") Then
            Set filetxt = filesys.OpenTextFile(Application.ActiveWorkbook.Path & "\" & logFileName & " Log.txt", ForWriting, True)
                filetxt.WriteLine ("End Log")
                filetxt.Close
        Else
            If FileLen(Application.ActiveWorkbook.Path & "\" & logFileName & " Log.txt") > 25000 Then
                With filesys
                    .CreateTextFile(Application.ActiveWorkbook.Path & "\" & logFileName & " Log.txt").Write Left(.OpenTextFile(Application.ActiveWorkbook.Path & "\" & logFileName & " Log.txt").ReadAll, 5000) & vbNewLine & "Log Trimmed " & Date & "_" & Time
                End With
            End If
        End If
        
        With filesys
            .CreateTextFile(Application.ActiveWorkbook.Path & "\" & logFileName & " Log.txt").Write Date & "_" & Time & ":  " & LogEntry & vbNewLine & .OpenTextFile(Application.ActiveWorkbook.Path & "\" & logFileName & " Log.txt").ReadAll
        End With
    End If
End Sub
' Procedure : TurnOffFunctionality
' Source    : www.ExcelMacroMastery.com
' Author    : Paul Kelly
' Purpose   : Turn off automatic calculations, events and screen updating
' https://excelmacromastery.com/
Public Sub TurnOffFunctionality()
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

' Procedure : TurnOnFunctionality
' Source    : www.ExcelMacroMastery.com
' Author    : Paul Kelly
' Purpose   : turn on automatic calculations, events and screen updating
' https://excelmacromastery.com/
Public Sub TurnOnFunctionality()
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
