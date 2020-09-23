Attribute VB_Name = "Module1"
Type SHITEMID
    cb As Long
    abID As Byte
End Type
Type ITEMIDLIST
    mkid As SHITEMID
End Type
Type bkw
    nazad(4000) As String
    naprijed(4000) As String
    oloc As String
End Type
Public act_dir As String
Public mmm As bkw
Type stzu
    animacija As String
    caption As String
    datoteke As String
    folderi As String
    lokacija As String
    operacija As Integer
End Type
Public Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
Public Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long '  only used if FOF_SIMPLEPROGRESS
End Type
Type stl1
    copy As Boolean
    move As Boolean
End Type
Type clipboard_
    opr As stl1
    fls2 As String
    fldrs2 As String
End Type
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function ExtractAssociatedIcon Lib "shell32.dll" Alias "ExtractAssociatedIconA" (ByVal hInst As Long, ByVal lpIconPath As String, lpiIcon As Long) As Long
Public Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal rGetIcon As Long) As Long
Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long 'ITEMIDLIST
Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
   (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public old_lst As String
Public mv As stzu
Public frrec As String
Public clp As clipboard_
Public fls11, fldrs11 As String


Public Const SW_NORMAL = 1
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1
Public Const BIF_RETURNONLYFSDIRS = &H1
Public Const BIF_DONTGOBELOWDOMAIN = &H2
Public Const BIF_STATUSTEXT = &H4
Public Const BIF_RETURNFSANCESTORS = &H8
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const BIF_BROWSEFORPRINTER = &H2000
Public Const FO_DELETE = &H3
Public Const FOF_ALLOWUNDO = &H40

Type SHELLEXECUTEINFO
       cbSize As Long
       fMask As Long
       hWnd As Long
       lpVerb As String
       lpFile As String
       lpParameters As String
       lpDirectory As String
       nShow As Long
       hInstApp As Long
       lpIDList As Long
       lpClass As String
       hkeyClass As Long
       dwHotKey As Long
       hIcon As Long
       hProcess As Long
End Type
Const SEE_MASK_INVOKEIDLIST = &HC
Const SEE_MASK_NOCLOSEPROCESS = &H40
Const SEE_MASK_FLAG_NO_UI = &H400
Declare Function ShellExecuteEX Lib "shell32.dll" Alias "ShellExecuteEx" (SEI As SHELLEXECUTEINFO) As Long



Public Function OpenFile1(filename As String, OwnerhWnd As Long) As Long
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = 64 'Or 12 'Or 1024
        .hWnd = OwnerhWnd
        .lpVerb = "open"
        .lpFile = filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 1
        .hInstApp = 0
        .lpIDList = 0
    End With
    ShellExecuteEX SEI
    'ShowFileProperties = SEI.hInstApp
End Function



Function get_loc()
      Dim bi As BROWSEINFO
      Dim idl As ITEMIDLIST
      Dim rtn&, pidl&, path$, pos%
      bi.hOwner = frmExplore.hWnd
      bi.lpszTitle = "Select location"
      bi.ulFlags = BIF_RETURNONLYFSDIRS
      pidl = SHBrowseForFolder(bi)
      path$ = Space$(512)
      rtn& = SHGetPathFromIDList(ByVal pidl&, ByVal path$)
      If rtn& Then
       
       pos% = InStr(path$, Chr$(0))
      get_loc = Left(path$, pos - 1)
      End If
End Function

Function folderi()
On Error Resume Next
 folderi = ""
 For i = 1 To frmExplore.lv.ListItems.Count
    If frmExplore.lv.ListItems.Item(i).Selected = True Then
        If Right(frmExplore.lv.ListItems.Item(i).SubItems(2), 2) = "  " Then
                If folderi = "" Then
                    folderi = frmExplore.tv.SelectedItem.Key & "\" & frmExplore.lv.ListItems(i)
                Else
                    folderi = folderi & Chr(1) & frmExplore.tv.SelectedItem.Key & "\" & frmExplore.lv.ListItems(i)
                End If
        End If
    End If
 Next i
End Function

Function datoteke()
On Error Resume Next
 datoteke = ""
 For i = 1 To frmExplore.lv.ListItems.Count
    If frmExplore.lv.ListItems.Item(i).Selected = True Then
            If Right(frmExplore.lv.ListItems.Item(i).SubItems(2), 2) <> "  " Then
                If datoteke = "" Then
                    datoteke = frmExplore.tv.SelectedItem.Key & "\" & frmExplore.lv.ListItems(i)
                Else
                    datoteke = datoteke & Chr(1) & frmExplore.tv.SelectedItem.Key & "\" & frmExplore.lv.ListItems(i)
                End If
            End If
    End If
 Next i
End Function

Sub izvrši_1()
Screen.MousePointer = vbHourglass
On Error Resume Next
duse = True
    Dim fso As New FileSystemObject
    If Right(mv.lokacija, 1) <> "\" Then
        mv.lokacija = mv.lokacija & "\"
    End If
    If mv.operacija = "0" Then
        folderi1 = Split(mv.folderi, Chr(1))
        For i = 0 To UBound(folderi1)
            dio = folderi1(i)
            fso.MoveFolder dio, mv.lokacija
            DoEvents
        Next i
        datoteke1 = Split(mv.datoteke, Chr(1))
        For i = 0 To UBound(datoteke1)
            dio = datoteke1(i)
            fso.MoveFile dio, mv.lokacija
            DoEvents
        Next i
    ElseIf mv.operacija = "1" Then
        folderi1 = Split(mv.folderi, Chr(1))
        For i = 0 To UBound(folderi1)
            dio = folderi1(i)
            fso.CopyFolder dio, mv.lokacija
            DoEvents
        Next i
        datoteke1 = Split(mv.datoteke, Chr(1))
        For i = 0 To UBound(datoteke1)
            dio = datoteke1(i)
            fso.CopyFile dio, mv.lokacija
            DoEvents
        Next i
    ElseIf mv.operacija = "2" Then
        folderi1 = Split(mv.folderi, Chr(1))
        max_f = UBound(folderi1) + 1
        For i = 0 To UBound(folderi1)
            dio = folderi1(i)
            fso.DeleteFolder dio
            frmExplore.sbStatusBar.Panels(1).Text = "Deleting folder(s): " & i + 1 & "/" & max_f
            DoEvents
        Next i
        datoteke1 = Split(mv.datoteke, Chr(1))
        max_d = UBound(datoteke1) + 1
        For i = 0 To UBound(datoteke1)
            dio = datoteke1(i)
            fso.DeleteFile dio
            frmExplore.sbStatusBar.Panels(1).Text = "Deleting file(s): " & i + 1 & "/" & max_d
            DoEvents
        Next i
    End If
frmExplore.osvjezi_11
duse = False
Screen.MousePointer = vbArrow
End Sub
Public Function ShowPropWindow(filename As String, OwnerhWnd As Long) As Long
Dim SEI As SHELLEXECUTEINFO
Dim r As Long
If filename = "" Then
    ShowPropWindow = 0
    Exit Function
End If
With SEI
    .cbSize = Len(SEI)
    .fMask = SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
    .hWnd = OwnerhWnd
    .lpVerb = "properties"
    .lpFile = filename
    .lpParameters = vbNullChar
    .lpDirectory = vbNullChar
    .nShow = 0
    .hInstApp = 0
    .lpIDList = 0
End With
r = ShellExecuteEX(SEI)
ShowPropWindow = SEI.hInstApp
End Function
Public Function FileExist(filename As String) As Boolean
    On Error GoTo FileDoesNotExist
    Call FileLen(filename)
    FileExist = True
    Exit Function
FileDoesNotExist:
    FileExist = False
End Function

Public Function get_pth()
fls11 = ""
fldrs11 = ""
    If frmExplore.ActiveControl.Name = "lv" Then
        For i = 1 To frmExplore.lv.ListItems.Count
            If frmExplore.lv.ListItems(i).Selected = True Then
                If Right(frmExplore.lv.ListItems(i).SubItems(2), 2) = "  " Then
                    If fldrs11 = "" Then
                        fldrs11 = frmExplore.tv.SelectedItem & "\" & frmExplore.lv.ListItems(i)
                    Else
                        fldrs11 = fldrs11 & Chr(1) & frmExplore.tv.SelectedItem.FullPath & "\" & frmExplore.lv.ListItems(i)
                    End If
                Else
                    If fls11 = "" Then
                        fls11 = frmExplore.tv.SelectedItem.FullPath & "\" & frmExplore.lv.ListItems(i)
                    Else
                        fls11 = fls11 & Chr(1) & frmExplore.tv.SelectedItem.FullPath & "\" & frmExplore.lv.ListItems(i)
                    End If
                End If
            End If
        Next i
    ElseIf frmExplore.ActiveControl.Name = "tv" Then
        fldrs11 = frmExplore.tv.SelectedItem.FullPath
    End If
End Function
Public Function lfc1()
    If frmExplore.ActiveControl.Name = "lv" Then
        If Right(frmExplore.lv.SelectedItem.SubItems(2), 2) = "  " Then
            lfc1 = frmExplore.tv.SelectedItem.FullPath & "\" & frmExplore.lv.SelectedItem
        Else
            lfc1 = frmExplore.tv.SelectedItem.FullPath
        End If
    ElseIf frmExplore.ActiveControl.Name = "tv" Then
        lfc1 = frmExplore.tv.SelectedItem.FullPath
    End If
End Function

Public Function lfc()
If frmExplore.ActiveControl.Name = "tv" Then
    lfc = frmExplore.tv.SelectedItem.FullPath
ElseIf frmExplore.ActiveControl.Name = "lv" Then
k = 0
For i = 1 To frmExplore.lv.ListItems.Count
    If frmExplore.lv.ListItems(i).Selected = True Then
        k = k + 1
    End If
Next i
If k > 0 Then
    If Right(frmExplore.lv.SelectedItem.ListSubItems(2), 2) = "  " Then
        a = MsgBox("Nalijepi pored foldera ''" & frmExplore.lv.SelectedItem.Text & "'' ?" & vbCrLf & "Pri odabiru opcije ''No'' sadržaj ce biti nalijepljen u folder ''" & frmExplore.lv.SelectedItem.Text & "''", vbYesNoCancel + vbExclamation, "Nalijepi")
        If a = vbYes Then
            lfc = Left(frmExplore.lv.SelectedItem.Key, InStrRev(frmExplore.lv.SelectedItem.Key, "\"))
        ElseIf a = vbNo Then
            lfc = frmExplore.lv.SelectedItem.Key
        Else
            lfc = ""
        End If
    Else
        lfc = Left(frmExplore.lv.SelectedItem.Key, InStrRev(frmExplore.lv.SelectedItem.Key, "\"))
    End If
Else
    lfc = frmExplore.tv.SelectedItem.FullPath
End If

End If
End Function

