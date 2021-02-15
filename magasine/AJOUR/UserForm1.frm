VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6855
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10560
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

End Sub

Private Sub CommandButton2_Click()
UserForm2.Show
End Sub

Private Sub CommandButton3_Click()
Unload Me
End Sub
'derniere ligne d une colone
Private Sub CommandButton4_Click()
Worksheets("medina").Select
Range("A65536").End(xlUp).Offset(8, 0).Select
ActiveCell.Offset(1, 0).Select 'Une case vers le bas

ActiveCell.Offset(-1, 0).Select 'Une case vers le haut
End Sub
'renvoi le numero de laligne
Private Sub CommandButton5_Click()
Dim rg As Range
 
Set rg = Range("b1:b100").Find("IMAGERUNNER", Range("b1"))
 
MsgBox rg.Address

End Sub

Private Sub CommandButton6_Click()
'Dim localise As Integer
localise = Cells.Find("IMAGERUNNER", , xlValues).Address
MsgBox localise
End Sub

Private Sub CommandButton7_Click()
'if not iserror(application.match(
End Sub
