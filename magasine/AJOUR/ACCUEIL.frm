VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ACCUEIL 
   Caption         =   "GESTION MAGASIN DU CONSEIL PREFECTORAL DE FES "
   ClientHeight    =   7455
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5160
   OleObjectBlob   =   "ACCUEIL.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ACCUEIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton2_Click()
UserForm2.Show
End Sub

Private Sub MouseMove()

End Sub

Private Sub BT1_Click()
Unload Me
AJOUTER.Show

End Sub

Private Sub BT10_Click()
Unload Me
SMEDINA.Show

End Sub

Private Sub BT11_Click()
Unload Me
SSIEGE.Show

End Sub

Private Sub BT3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
BT10.Visible = True
BT11.Visible = True

End Sub

Private Sub BT2_Click()
Unload Me
NOUVEAU.Show
End Sub
Private Sub BT5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
BT10.Visible = False
BT11.Visible = False
End Sub

Private Sub CommandButton10_Click()
Worksheets("SIEGE").Select
End Sub

Private Sub CommandButton11_Click()
Worksheets("DAPC").Select
End Sub

Private Sub CommandButton12_Click()
Worksheets("SAFM").Select
End Sub

Private Sub CommandButton13_Click()
Worksheets("SDE").Select
End Sub

Private Sub CommandButton14_Click()
Worksheets("SGRH").Select
End Sub

Private Sub CommandButton15_Click()
Worksheets("CAI").Select
End Sub

Private Sub CommandButton16_Click()
Worksheets("MRPRESIDENT").Select
End Sub

Private Sub CommandButton17_Click()
Worksheets("SMGP").Select
End Sub

Private Sub CommandButton18_Click()
Worksheets("DGS").Select
End Sub

Private Sub CommandButton19_Click()
Worksheets("LISTES").Select
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

End Sub

Private Sub CommandButton7_Click()

localise = Cells.Find("IMAGERUNNER", , xlValues).Address
MsgBox localise
End Sub

Private Sub CommandButton8_Click()
Dim Lig As Long
Worksheets("siege").Select
Lig = 1 'première ligne à vérifier
Do While Not IsEmpty(Range("C" & Lig))
    Lig = Lig + 1
Loop
MsgBox "La première ligne vide colonne C est la ligne : " & Lig
End Sub


Private Sub CommandButton9_Click()
Worksheets("medina").Select
End Sub

Private Sub BT4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
BT10.Visible = False
BT11.Visible = False
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub Frame1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
BT10.Visible = False
BT11.Visible = False
End Sub

Private Sub Frame2_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
BT10.Visible = True
BT11.Visible = True
End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()
BT10.Visible = False
BT11.Visible = False
End Sub
Private Sub vider_Click()
With Worksheets("medina")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("siege")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("sde")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("dapc")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("safm")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("sgrh")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("cai")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("dgs")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("mrpresident")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("smgp")
    .Activate
    .Range("a4:d23000").Select
    .Range("a4:d23000").Clear
End With
With Worksheets("listes")
    .Activate
    .Range("e4:e23000").Select
    .Range("e4:e23000").Clear
End With
End Sub
