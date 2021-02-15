VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AJOUTER 
   Caption         =   "AJOUT"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11925
   OleObjectBlob   =   "AJOUTER.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AJOUTER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub CB2_Change()
If CB2.Value <> "" And CB2.Value <> "Selectionner" Then
    cm4.Enabled = True
    cm1.Enabled = True
Else
    cm1.Enabled = False
End If
End Sub


Private Sub CB1_Change()
If CB1.Value <> "" And CB1.Value <> "Selectionner" Then
    cm4.Enabled = True
    cm1.Enabled = True
Else
    cm1.Enabled = False
End If
End Sub
Private Sub CB2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub CM1_Click()

Sheets("listes").Select
Range("e2").Select
Do Until ActiveCell = Me.TB2 Or ActiveCell = ""
    ActiveCell.Offset(1, 0).Select
Loop
If ActiveCell = Me.TB2 Then
MsgBox "ARTICLE DEJA EXISTANT!", vbCritical, "GMCPF"
Else




If TB2 <> "" And TB2 <> "Selectionner" And CB1 <> "" And CB1 <> "Selectionner" And CB2 <> "" And CB2 <> "Selectionner" And TB1 <> "" Then
    cm1.Enabled = True

    If MsgBox("VOULEZ VOUS VRAIMENT AJOUTER CET ARTICLE ?", vbYesNo, "GMCPF") = vbYes Then

    Dim ligvid As Integer
    ligvid = 0
    If CB2.Value = "SIEGE" Then
        With Worksheets("siege").Select
            ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            If Cells(ligvid - 1, "a") = "" Then
            MsgBox "LE NUMERO D'RTICLE DANS LA FEUILLE 'SIEGE' EST INCORRECTE VEUILLEZ-LE VERIFIER ! ", vbCritical, "GMCPF"
            Exit Sub
            Else
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
            End If
            
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "d") = Round(TB1.Value, 2)
            Cells(ligvid, "C") = CB1.Value
            Cells(ligvid, "e") = TB3.Value
            
        End With
        ligvid = 0
        With Worksheets("listes").Select
         ligvid = Columns("e").Find("", Range("e1"), xlValues).Row
            Cells(ligvid, "e") = TB2.Text
        End With
        ligvid = 0
        With Worksheets("medina").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
            Cells(ligvid, "e") = TB3.Value
        End With
        ligvid = 0
        'With Worksheets("services").Select
            'Cells(ligvid, "b") = TB2.Text
            'Cells(ligvid, "c") = CB1.Value
            'Cells(ligvid, "d") = 0
            'Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        'End With
        
        '------------------------------
        With Worksheets("SDE").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("DAPC").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("SAFM").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("SGRH").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("CAI").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("DGS").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("MRPRESIDENT").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("SMGP").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
   '---------------------------------------------
       
    End If
      ligvid = 0
    If CB2.Value = "MEDINA" Then
        Worksheets("medina").Select
            ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            If Cells(ligvid - 1, "a") = "" Then
            MsgBox "LE NUMERO D'RTICLE DANS LA FEUILLE 'MEDINA' EST INCORRECTE VEUILLEZ-LE VERIFIER ! ", vbCritical, "GMCPF"
            Exit Sub
            Else
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
            End If
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "d") = Round(TB1.Value, 2)
            Cells(ligvid, "C") = CB1.Value
            Cells(ligvid, "e") = TB3.Value
            
            
            
        ligvid = 0
            
        Worksheets("listes").Select
        ligvid = Columns("e").Find("", Range("e1"), xlValues).Row
            Cells(ligvid, "e") = TB2.Text
          ligvid = 0
        With Worksheets("siege").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
            Cells(ligvid, "e") = TB3.Value
        End With
        ligvid = 0
        
        'With Worksheets("services").Select
            'Cells(ligvid, "b") = TB2.Text
            'Cells(ligvid, "c") = CB1.Value
            'Cells(ligvid, "d") = 0
            'Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        'End With
        
         '------------------------------
        With Worksheets("SDE").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("DAPC").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("SAFM").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("SGRH").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("CAI").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("DGS").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("MRPRESIDENT").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
        With Worksheets("SMGP").Select
        ligvid = Columns("b").Find("", Range("b1"), xlValues).Row
            Cells(ligvid, "b") = TB2.Text
            Cells(ligvid, "c") = CB1.Value
            Cells(ligvid, "d") = 0
            Cells(ligvid, "a") = Cells(ligvid - 1, "a") + 1
        End With
        ligvid = 0
   '---------------------------------------------
        
        
        
        
        
        
        
        
        
    End If
    End If
    Else
                
                MsgBox "SVP REMPLIR TOUS LES CHAMPS !", vbCritical, "GMCPF"
End If
End If

CB1 = "Selectionner"
CB2 = "Selectionner"
TB1 = ""
TB2 = ""
cm1.Enabled = False
End Sub
Private Sub CM4_Click()
CB1 = "Selectionner"
CB2 = "Selectionner"
TB1 = ""
TB2 = ""
cm4.Enabled = False
End Sub
Private Sub CM5_Click()
Unload Me
ACCUEIL.Show
End Sub
'fonction chainepassOK
Private Function ChainePasOK(strpass As String) As Boolean
  If strpass = "" Then Exit Function
   If Len(Replace(strpass, ".", "")) <> Len(strpass) Then ChainePasOK = True: Exit Function
   strpass = Replace(strpass, ",", ".")
   If Len(CStr(Val(strpass))) <> Len(strpass) Then ChainePasOK = True
End Function
Private Sub CM2_Click()
Unload Me
SMEDINA.Show
End Sub
Private Sub CM3_Click()
Unload Me
SSIEGE.Show
End Sub


Private Sub Label7_Click()

End Sub

Private Sub TB1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Vérifie si le tx horaire est bien un nombre, avec des décimales, et opère aux modification le cas contraire
 Select Case KeyAscii
 
        'seulement des chiffres
        Case 48 To 57
 
        'virgule et point, séparateur décimal
        Case 44, 46
 
            'seulement la virgule
            If KeyAscii = 46 Then KeyAscii = 44
 
            'seulement une fois
            If InStr(TB1.Text, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
 
            'interdit le séparateur décimal en premier
            If TB1.Text = "" Then If KeyAscii = 44 Then KeyAscii = 0
            
 
    Case Else
            
            TB1.Value = Null
            MsgBox "SEULEMENT NUMERIQUE ! DEUX DECIMALES MAXIMUM", vbCritical, "GMCPF"
            cm4.Enabled = False
            cm1.Enabled = False
 
        End Select
        
End Sub

Private Sub TB2_Change()
If TB2.Value <> "" Then
    cm4.Enabled = True
    cm1.Enabled = True
Else
    cm4.Enabled = False
    cm1.Enabled = False
End If
End Sub
Private Sub CommandButton9_Click()
Dim bOk As Boolean
bOk = True
If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
   bOk = False
    If KeyAscii = Asc(".") Or KeyAscii = Asc(",") Then
        If KeyAscii = Asc(".") Then KeyAscii = Asc(",")
                bOk = True
    End If
End If
If Not bOk Then
    KeyAscii = 0
    Beep
End If
Dim strpass As String
   strpass = txtPrixUF1.Value
    
        If ChainePasOK(strpass) = True Then Cancel = True: txtPrixUF1.Value = "": Beep: MsgBox "Saisie invalide. Remplir par un chiffre,chiffre ex: 1256,52"
End Sub

Private Sub TB3_Change()

End Sub

Private Sub TB3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Vérifie si le tx horaire est bien un nombre, avec des décimales, et opère aux modification le cas contraire
 Select Case KeyAscii
 
        'seulement des chiffres
        Case 48 To 57
 
        'virgule et point, séparateur décimal
        Case 44, 46
 
            'seulement la virgule
            If KeyAscii = 46 Then KeyAscii = 44
 
            'seulement une fois
            If InStr(TB3.Text, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
 
            'interdit le séparateur décimal en premier
            If TB3.Text = "" Then If KeyAscii = 44 Then KeyAscii = 0
            
 
    Case Else
            
            TB3.Value = Null
            MsgBox "SEULEMENT NUMERIQUE ! DEUX DECIMALES MAXIMUM", vbCritical, "GMCPF"
            cm4.Enabled = False
            cm1.Enabled = False
 
        End Select
End Sub

Private Sub UserForm_Initialize()

cm1.Enabled = False
CB2.Value = "Selectionner"
CB1 = "Selectionner"
TB1.Text = ""
End Sub
