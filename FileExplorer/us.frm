VERSION 5.00
Begin VB.Form us 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   600
      Top             =   1680
   End
   Begin VB.Menu pregled 
      Caption         =   "pregled"
      Begin VB.Menu ikone 
         Caption         =   "Icons"
      End
      Begin VB.Menu male_ikone 
         Caption         =   "Small Icons"
      End
      Begin VB.Menu lista 
         Caption         =   "List"
      End
      Begin VB.Menu izvjestaj 
         Caption         =   "Details"
      End
   End
   Begin VB.Menu brisanje 
      Caption         =   "brisanje"
      Begin VB.Menu obr 
         Caption         =   "Delete"
      End
      Begin VB.Menu purb 
         Caption         =   "Send to Recycle Bin"
      End
   End
   Begin VB.Menu izbor 
      Caption         =   "izbor"
      Begin VB.Menu otvori 
         Caption         =   "Open"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu olue 
         Caption         =   "Open in Windows Explorer"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu izrezi 
         Caption         =   "Cut"
      End
      Begin VB.Menu kopiraj 
         Caption         =   "Copy"
      End
      Begin VB.Menu nalijepi 
         Caption         =   "Paste"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu obriši 
         Caption         =   "Delete"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu p_ime 
         Caption         =   "Rename"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu o_sve 
         Caption         =   "Select All"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu svojstva 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "us"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private txt_mem

Private Sub ikone_Click()
frmExplore.lv.View = lvwIcon
End Sub

Private Sub izrezi_Click()
frmExplore.cut_d
End Sub

Private Sub izvjestaj_Click()
frmExplore.lv.View = lvwReport
End Sub

Private Sub kopiraj_Click()
frmExplore.copy_d
End Sub

Private Sub lista_Click()
frmExplore.lv.View = lvwList
End Sub

Private Sub male_ikone_Click()
frmExplore.lv.View = lvwSmallIcon
End Sub

Private Sub nalijepi_Click()
frmExplore.paste_d
End Sub

Private Sub o_sve_Click()
For i = 1 To frmExplore.lv.ListItems.Count
    frmExplore.lv.ListItems(i).Selected = True
Next i
End Sub

Private Sub obr_Click()
If MsgBox("Potvrdi brisanje?", vbYesNo, "Brisanje") = vbYes Then
    If frmExplore.ActiveControl.Name = "lv" Then
        mv.caption = "Brisanje..."
        mv.folderi = folderi
        mv.datoteke = datoteke
        mv.operacija = 2
        mv.animacija = "Brisanje"
        izvrši_1
    ElseIf frmExplore.ActiveControl.Name = "tv" Then
        mv.caption = "Brisanje..."
        mv.folderi = frmExplore.tv.SelectedItem.FullPath
        mv.operacija = 2
        mv.animacija = "Brisanje"
        d = frmExplore.tv.SelectedItem.FullPath
        root_ = Left(d, InStrRev(d, "\") - 1)
        frmExplore.tv.Nodes(root_).Selected = True
        izvrši_1
    End If
End If
End Sub

Private Sub obriši_Click()
obr_Click
End Sub

Private Sub olue_Click()
Dim fn As String
If frmExplore.ActiveControl.Name = "tv" Then
    fn = frmExplore.tv.SelectedItem.FullPath
    
ElseIf frmExplore.ActiveControl.Name = "lv" Then
    If Right(frmExplore.lv.SelectedItem.ListSubItems(2), 2) = "  " Then
        fn = frmExplore.lv.SelectedItem.Key
    Else
        fn = Left(frmExplore.lv.SelectedItem.Key, InStrRev(frmExplore.lv.SelectedItem.Key, "\"))
        k = 1
    End If
End If
OpenFile1 fn, Me.hWnd
If k = 1 Then
    s = frmExplore.lv.SelectedItem.Key
    txt_mem = Right(s, Len(s) - InStrRev(s, "\"))
    Sleep 1000
    SendKeys txt_mem
End If
End Sub

Private Sub otvori_Click()
If frmExplore.ActiveControl.Name = "lv" Then
If Right(frmExplore.lv.SelectedItem.ListSubItems(2), 2) = "  " Then
    frmExplore.lvdblclck
Else
    OpenFile1 frmExplore.lv.SelectedItem.Key, Me.hWnd
End If
ElseIf frmExplore.ActiveControl.Name = "tv" Then
    If frmExplore.tv.SelectedItem.Expanded = False Then
        frmExplore.tv.SelectedItem.Expanded = True
    End If
End If
End Sub

Private Sub p_ime_Click()
If frmExplore.ActiveControl.Name = "tv" Then
    frmExplore.tv.StartLabelEdit
ElseIf frmExplore.ActiveControl.Name = "lv" Then
    frmExplore.lv.StartLabelEdit
End If
End Sub

Private Sub purb_Click()
Dim op As SHFILEOPSTRUCT
    If frmExplore.ActiveControl.Name = "lv" Then
        pth = frmExplore.tv.SelectedItem.FullPath & "\" & frmExplore.lv.SelectedItem
    ElseIf frmExplore.ActiveControl.Name = "tv" Then
        pth = frmExplore.tv.SelectedItem.FullPath
    End If
    With op
        .wFunc = FO_DELETE
        .pFrom = pth
        .fFlags = FOF_ALLOWUNDO
    End With
    SHFileOperation op
End Sub

Private Sub svojstva_Click()
frmExplore.svojstva_
End Sub

Private Sub Timer1_Timer()
If frmExplore.ActiveControl.Name = "tv" Then
    o_sve.Enabled = False
Else
    o_sve.Enabled = True
End If
End Sub

