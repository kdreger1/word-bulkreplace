Attribute VB_Name = "PMAmodule_renumber_v1"
Option Explicit

' declare functions
#If VBA7 Then
    'For 64 Bit Systems
    Declare PtrSafe Function PMA_renumber_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#Else
    'For 32 Bit Systems
    Declare Function PMA_renumber_GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
#End If

' constants for mode selection. The strings are displayed by the pop-up menu
Const PMA_RENUMBER_MODE_HyTn As String = "Highlight existing numbers, but do not make text changes"
Const PMA_RENUMBER_MODE_HyTy As Variant = "Highlight existing numbers, and make text changes"
Const PMA_RENUMBER_MODE_HnTy As Variant = "Do not highlight, and make text changes"
Public PMA_RENUMBER_MODE As Variant


Sub PMA_renumber()
Attribute PMA_renumber.VB_Description = "Macro designed to implement a bulk find-and-replace based on the patterns found in an Excel file that the user can select."
Attribute PMA_renumber.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.PMA_renumber"
On Error GoTo errme
'
' PMA_renumber Macro
' Macro designed to implement a bulk find-and-replace based on the patterns found in an Excel file that the user can select.
'
' Options are:
'   1. highlight what will be changed, but do not change it
'   2. make the changes and hightlight them
'   3. make the changes and do not hightlight them

    Dim inFileName As String    ' Excel file that has the numbers to change
    Dim outFileName As String   ' where to put resulting Word file
    Dim defaultDir As String    ' PMA uses Dropbox, so start search here (change to Teams eventually)
    'Dim renumberMode As Variant ' one of 3 modes
    Dim changeText As Boolean
    Dim popName As String       ' so we can destroy it after use
    
    Dim excelSrc As Workbook            ' object for the Excel input file
    Dim changeFrom(1 To 1000) As String ' read-in colum 1
    Dim changeTo(1 To 1000) As String   ' read-in column 2
    Dim maxDim As Integer               ' largest array size
    Dim totalRows As Integer            ' total number of rows read into arrays
    Dim activeSheetName As String       ' first sheet in the Excel file (usually called Sheet1, but just in case...)
    
    Dim prefixStr As String     ' to avoid cascading changes
    'Dim Word As Range           ' to loop through the document
    
    Dim i As Integer
    Dim j As Integer
    
    ' initialize
    maxDim = 1000
    prefixStr = "Z0Y1X2W3V"
    PMA_RENUMBER_MODE = Null
    changeText = False
    
    ' set defaultDir to be Dropbox root
    defaultDir = "C:\Users\" & PMA_renumber_getUser & "\Dropbox (Gates Institute)"
    
    ' double check directory, and try a general match if not found
    If Dir(defaultDir, vbDirectory) = "" Then
        defaultDir = Dir("C:\Users\" & PMA_renumber_getUser & "\Dropbox*", vbDirectory)  ' returns just the end directory
        If defaultDir <> "" Then
            defaultDir = "C:\Users\" & PMA_renumber_getUser & "\" & defaultDir
        End If
    End If
    
    ' get path and filename of Excel input file
    inFileName = PMA_renumber_getXLSXname("Please navigate to, and select, the pattern-match Excel file...", defaultDir)
    If inFileName = "" Then
        ' user canceled
        GoTo cleanup
    End If
    
    ' open up the Excel file in read-only hidden mode
    Application.ScreenUpdating = False
    
    Set excelSrc = Workbooks.Open(inFileName, False, True)
    activeSheetName = ActiveSheet.Name
    totalRows = excelSrc.Worksheets(activeSheetName).Range("A1:A" & Cells(Rows.Count, "A").End(xlUp).Row).Rows.Count
    If totalRows > maxDim Then
        MsgBox "There are too many rows (" & totalRows & ") in the Excel file." & Chr(13) & "Maximum number of rows allowed is " & maxDim & ".", vbCritical + vbOKOnly, "ERROR"
        excelSrc.Close False
        GoTo cleanup
    End If
    
    ' copy to arrays
    For i = 1 To totalRows
        changeFrom(i) = excelSrc.Worksheets(activeSheetName).Range("A" & i).Value
        changeTo(i) = excelSrc.Worksheets(activeSheetName).Range("B" & i).Value
    Next i
    
    ' remove blank rows (sometimes Excel thinks an empty row is real data)
    For i = 1 To totalRows
        If changeFrom(i) = "" Or changeTo(i) = "" Then
            For j = i To totalRows - 1
                changeFrom(j) = changeFrom(j + 1)
                changeTo(j) = changeTo(j + 1)
            Next j
            
            totalRows = totalRows - 1
        End If
    Next i
    
    ' explicitly close the source file
    excelSrc.Close False
    
    ' create a copy of the active document to make all the changes to
    ' this copy becomes the active document
    Application.Documents.Add ActiveDocument.FullName
    
    ' clear any existing parameters for search and replace
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    Options.DefaultHighlightColorIndex = wdTurquoise
    
    ' modes for this script
    'renumberMode = PMA_renumber_getMode
    Select Case PMA_renumber_Nz(PMA_renumber_getMode, "NULL")
        Case PMA_RENUMBER_MODE_HyTn
            Selection.Find.Replacement.Highlight = True
            changeText = False
        
        Case PMA_RENUMBER_MODE_HyTy
            Selection.Find.Replacement.Highlight = True
            changeText = True
            
        Case PMA_RENUMBER_MODE_HnTy
            Selection.Find.Replacement.Highlight = False
            changeText = True
        
        Case Else
            MsgBox "User did not select a mode for this script." & Chr(13) & "Exiting without making changes to copy.", vbCritical + vbOKOnly, "ERROR"
            GoTo cleanup
    End Select
    
    ' loop through changes to apply
    For i = 1 To totalRows
        With Selection.Find
            .Text = changeFrom(i)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = True
            .MatchCase = True
            .MatchWholeWord = True
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
            
            If changeText Then
                .Replacement.Text = prefixStr & changeTo(i)
            Else
                .Replacement.Text = ""  ' this does not change it to a blank, but rather indicates no text change
            End If
            
        End With

        Selection.Find.Execute Replace:=wdReplaceAll
    Next i
    
    ' remove prefix
    If changeText Then
        With ActiveDocument.Content.Find
            .Text = prefixStr
            .Replacement.Text = ""
            .Wrap = wdFindContinue
            .Execute Replace:=wdReplaceAll
        End With
    End If

cleanup:
'    If Not excelSrc Is Nothing Then
'        excelSrc.Close False
'    End If
    Set excelSrc = Nothing
    
    Application.ScreenUpdating = True

Exit Sub
errme:
    MsgBox Err.Description
    Resume cleanup

End Sub

Function PMA_renumber_getMode() As Variant
On Error GoTo errme
    ' function to get from the user the mode that PMA_renumber should operate in
    '   1. highlight_nochange
    '   2. highlight_change
    '   3. nohighlight_change
    
    Dim popName As String       ' so we can destroy it after use
    
    popName = "PMA_Popup_GetRenumberMode_34512303"
    
    Call PMA_renumber_CreatePopUp(popName)
    Call PMA_renumber_DisplayCustomPopUp(popName)
    Call PMA_renumber_DeleteCustomPopUp(popName)

    PMA_renumber_getMode = PMA_RENUMBER_MODE
        
    ' debug
    'PMA_renumber_getMode = "highlight_nochange"
    'PMA_renumber_getMode = "highlight_change"
    'PMA_renumber_getMode = "nohighlight_change"

    Exit Function
errme:
    MsgBox Err.Description

End Function

Public Function PMA_renumber_Nz(ByVal Value, Optional ByVal ValueIfNull = "")

    PMA_renumber_Nz = IIf(IsNull(Value), ValueIfNull, Value)

End Function

Public Function PMA_renumber_getUser() As String
On Error GoTo errme
    ' Display the name of the user currently logged on.
    Dim username As String  ' receives name of the user
    Dim slength As Long  ' length of the string
    Dim retval As Long  ' return value
    
    ' Create room in the buffer to receive the returned string.
    username = Space(255)  ' room for 255 characters
    slength = 255  ' initialize the size of the string
    
    ' Get the user's name and display it.
    retval = PMA_renumber_GetUserName(username, slength)  ' slength is now the length of the returned string
    username = Left(username, slength - 1)  ' extract the returned info from the buffer
                                            ' (We subtracted one because we don't want the null character in the trimmed string.)
    PMA_renumber_getUser = username
    
    'Debug.Print "User name is '" & Trim(USERNAME) & "' "
    
Exit Function
errme:
    PMA_renumber_getUser = "CANNOT_FIND_USER"
End Function

Function PMA_renumber_getXLSXname(dialogTitleStr As String, Optional startingDir As String = "") As String
On Error GoTo errme
    ' Open up a file picker that is filtered for Excel files (xlsx only)

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    Dim FileChosen As Integer

    fd.Title = dialogTitleStr
    'fd.InitialFileName = "c:\ActiveStudyDatabases\"    ' inital starting place
    'fd.InitialView = msoFileDialogViewSmallIcons       ' list, icons, detail, etc.
    
    ' Filters
    fd.Filters.Clear
    fd.Filters.Add "XLSX Files (*.xlsx)", "*.xlsx"
    fd.FilterIndex = 1  ' if there's more than one filter, you can control which one is selected by default
    
    ' starting directory
    If Dir(startingDir, vbDirectory) <> "" Then fd.InitialFileName = startingDir
    
    FileChosen = fd.Show
    
    If FileChosen <> -1 Then
        ' user canceled
        PMA_renumber_getXLSXname = ""
        GoTo cleanup
    Else
        PMA_renumber_getXLSXname = fd.SelectedItems(1)
    End If
    
cleanup:
    Set fd = Nothing
    
    Exit Function
errme:
    MsgBox Err.Description
End Function


Function PMA_renumber_DeleteCustomPopUp(popName As String)
On Error Resume Next
    ' deletes the indicated popup menu
    ' if there was none, moves on without error message
    
    CommandBars(popName).Delete
    
End Function

Function PMA_renumber_CreatePopUp(popName As String)
On Error GoTo errme

    ' creates the custom popup menu
    Dim cb As CommandBar
    
    ' clear any old one
    Call PMA_renumber_DeleteCustomPopUp(popName)
    
    Set cb = CommandBars.Add(popName, msoBarPopup, False, True)
    With cb
        
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "PMA_renumber_setMode"
            .FaceId = 71
            .Caption = PMA_RENUMBER_MODE_HyTn
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "PMA_renumber_setMode"
            .FaceId = 72
            .Caption = PMA_RENUMBER_MODE_HyTy
        End With
        With .Controls.Add(Type:=msoControlButton)
            .OnAction = "PMA_renumber_setMode"
            .FaceId = 73
            .Caption = PMA_RENUMBER_MODE_HnTy
        End With
    End With
    
cleanup:
    Set cb = Nothing
    Exit Function

errme:
    MsgBox Err.Description
    Resume cleanup
    
End Function

Function PMA_renumber_DisplayCustomPopUp(popName As String)
On Error GoTo errme

    ' displays the indicated custom popup menu
    Application.CommandBars(popName).ShowPopup
    
    Exit Function
errme:
    MsgBox Err.Description
End Function

Public Function PMA_renumber_setMode()
On Error GoTo errme

    PMA_RENUMBER_MODE = CommandBars.ActionControl.Caption
    
    Exit Function
errme:
    MsgBox Err.Description
    
End Function

