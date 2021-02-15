VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SMEDINA 
   Caption         =   "SORTIE MEDINA"
   ClientHeight    =   5145
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11085
   OleObjectBlob   =   "SMEDINA.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SMEDINA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CB1_Change()
If CB1.Value <> "" And CB1.Value <> "Selectionner" Then
    cm2.Enabled = True
    cm1.Enabled = True
Else
    cm2.Enabled = False
    cm1.Enabled = False
End If

Dim ligne As Integer

                If Not IsError(Application.Match(CB1, Sheets("MEDINA").Range("b:b"), 0)) Then
                    ligne = Application.Match(CB1, Sheets("MEDINA").Range("b:b"), 0)
                    If Sheets("MEDINA").Range("d" & ligne).Value > 5 Then
                    LB1.ForeColor = vbGreen
                    LB1 = Sheets("MEDINA").Range("d" & ligne).Value & " " & Sheets("MEDINA").Range("c" & ligne).Value
                    Else
                    LB1.ForeColor = vbRed
                    LB1 = Sheets("MEDINA").Range("d" & ligne).Value & " " & Sheets("MEDINA").Range("c" & ligne).Value
                    End If
                End If
                

End Sub

Private Sub CB1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub CB2_Change()
If CB2.Value <> "" And CB2.Value <> "Selectionner" Then
    cm2.Enabled = True
    cm1.Enabled = True
        If CB2.Value = "SERVICES" Then
            Cb3.Visible = True
            Label6.Visible = True
        Else
            Cb3.Visible = False
            Label6.Visible = False
        End If
Else
    cm2.Enabled = False
    cm1.Enabled = False
End If

Dim ligne As Integer

                If Not IsError(Application.Match(CB1, Sheets("MEDINA").Range("b:b"), 0)) Then
                    ligne = Application.Match(CB1, Sheets("MEDINA").Range("b:b"), 0)
                    If Sheets("MEDINA").Range("d" & ligne).Value > 5 Then
                    LB1.ForeColor = vbGreen
                    LB1 = Sheets("MEDINA").Range("d" & ligne).Value & " " & Sheets("MEDINA").Range("c" & ligne).Value
                    Else
                    LB1.ForeColor = vbRed
                    LB1 = Sheets("MEDINA").Range("d" & ligne).Value & " " & Sheets("MEDINA").Range("c" & ligne).Value
                    End If
                End If
                




End Sub

Private Sub CB2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub
Private Sub Cb3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub CM1_Click()

If CB1 <> "" And CB1 <> "Selectionner" And CB2 <> "" And CB2 <> "Selectionner" And TB1 <> "" Then    '1
    cm1.Enabled = True
    
 If MsgBox("VOULEZ VOUS VRAIMENT AJOUTER CETT SORTIE ?", vbYesNo, "GMCPF") = vbYes Then  '22
   
    
    Dim qte As Single
    qte = Round(TB1.Text, 2)
    If Not IsError(Application.Match(CB1, Sheets("MEDINA").Range("b:b"), 0)) Then '2
            ligne = Application.Match(CB1, Sheets("MEDINA").Range("b:b"), 0)
    End If  '2
    If qte > Sheets("MEDINA").Range("d" & ligne).Value Then  '3
    MsgBox "ATTENTION, QUANTITE INDISPONIBLE!!", vbCritical, "GMCPF"
    
    Else  '3
    
    Sheets("MEDINA").Range("d" & ligne).Value = Sheets("MEDINA").Range("d" & ligne).Value - Round(TB1.Text, 3)
    
    
    If CB2.Value = "SIEGE" Then  '4
        ligne = 0
        If Not IsError(Application.Match(CB1, Sheets("siege").Range("b:b"), 0)) Then  '5
            ligne = Application.Match(CB1, Sheets("siege").Range("b:b"), 0)
            Sheets("siege").Range("d" & ligne).Value = Sheets("siege").Range("d" & ligne).Value + Round(TB1.Text, 3)
        End If  '5
    End If  '4
    '----------------------------------------------
    If CB2.Value = "SERVICES" Then  '6
        ligne = 0
        If Cb3 = "SDE" Then  '7
        
            If Not IsError(Application.Match(CB1, Sheets("SDE").Range("b:b"), 0)) Then  '8
                ligne = Application.Match(CB1, Sheets("SDE").Range("b:b"), 0)
                Sheets("SDE").Range("d" & ligne).Value = Sheets("SDE").Range("d" & ligne).Value + Round(TB1.Text, 3)
            End If   '8
        End If     '7
        
        ligne = 0
        If Cb3 = "DAPC" Then   '9
        
            If Not IsError(Application.Match(CB1, Sheets("DAPC").Range("b:b"), 0)) Then  '10
                ligne = Application.Match(CB1, Sheets("DAPC").Range("b:b"), 0)
                Sheets("DAPC").Range("d" & ligne).Value = Sheets("DAPC").Range("d" & ligne).Value + Round(TB1.Text, 3)
            End If  '10
        End If  '9
        ligne = 0
        If Cb3 = "SAFM" Then  '11
        
            If Not IsError(Application.Match(CB1, Sheets("SAFM").Range("b:b"), 0)) Then  '12
                ligne = Application.Match(CB1, Sheets("SAFM").Range("b:b"), 0)
                Sheets("SAFM").Range("d" & ligne).Value = Sheets("SAFM").Range("d" & ligne).Value + Round(TB1.Text, 3)
            End If  '12
        End If  '11
        ligne = 0
        If Cb3 = "SGRH" Then  '13
        
            If Not IsError(Application.Match(CB1, Sheets("SGRH").Range("b:b"), 0)) Then  '14
                ligne = Application.Match(CB1, Sheets("SGRH").Range("b:b"), 0)
                Sheets("SGRH").Range("d" & ligne).Value = Sheets("SGRH").Range("d" & ligne).Value + Round(TB1.Text, 3)
            End If  '14
        End If  '13
        ligne = 0
        If Cb3 = "CAI" Then  '15
        
            If Not IsError(Application.Match(CB1, Sheets("CAI").Range("b:b"), 0)) Then  '16
                ligne = Application.Match(CB1, Sheets("CAI").Range("b:b"), 0)
                Sheets("CAI").Range("d" & ligne).Value = Sheets("CAI").Range("d" & ligne).Value + Round(TB1.Text, 3)
            End If  '16
        End If  '15
        ligne = 0
        If Cb3 = "DGS" Then  '17
        
            If Not IsError(Application.Match(CB1, Sheets("DGS").Range("b:b"), 0)) Then
                ligne = Application.Match(CB1, Sheets("DGS").Range("b:b"), 0)
                Sheets("DGS").Range("d" & ligne).Value = Sheets("DGS").Range("d" & ligne).Value + Round(TB1.Text, 3)
            End If
        End If  '17
        ligne = 0
        If Cb3 = "MR LE PRESIDENT" Then  '18
        
            If Not IsError(Application.Match(CB1, Sheets("MRPRESIDENT").Range("b:b"), 0)) Then  '19
                ligne = Application.Match(CB1, Sheets("MRPRESIDENT").Range("b:b"), 0)
                Sheets("MRPRESIDENT").Range("d" & ligne).Value = Sheets("MRPRESIDENT").Range("d" & ligne).Value + Round(TB1.Text, 3)
            End If  '19
        End If  '18
        ligne = 0
        If Cb3 = "SMGP" Then  '20
        
            If Not IsError(Application.Match(CB1, Sheets("SMGP").Range("b:b"), 0)) Then  '21
                ligne = Application.Match(CB1, Sheets("SMGP").Range("b:b"), 0)
                Sheets("SMGP").Range("d" & ligne).Value = Sheets("SMGP").Range("d" & ligne).Value + Round(TB1.Text, 3)
            End If  '21
        End If  '20
  ' --------------------------------------------
    End If  '2
    End If  '3
    End If '22

Else   '1
    MsgBox "SVP REMPLIR TOUS LES CHAMPS !", vbCritical, "GMCPF"
End If  '1

CB1 = "Selectionner"
CB2 = "Selectionner"
Cb3 = "Selectionner"
TB1 = ""
cm1.Enabled = False
cm2.Enabled = False
Cb3.Visible = False
Label6.Visible = False
LB1 = ""
End Sub

Private Sub CM2_Click()
CB1 = "Selectionner"
CB2 = "Selectionner"
Cb3 = "Selectionner"
TB1 = ""
LB1 = ""
cm2.Enabled = False
End Sub

Private Sub CM3_Click()
Unload Me
AJOUTER.Show
End Sub

Private Sub CM4_Click()
Unload Me
ACCUEIL.Show
End Sub




Private Sub LB1_Click()

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
            MsgBox "Seulement numérique ! Deux décimales maximum."
            cm2.Enabled = False
            cm1.Enabled = False
 
        End Select
End Sub

Private Sub UserForm_Activate()
cm2.Enabled = False
cm1.Enabled = False
Cb3.Visible = False
Label6.Visible = False
LB1 = ""
End Sub

Private Sub UserForm_Initialize()

cm1.Enabled = False
CB2.Value = "Selectionner"
TB1.Text = ""
Dim i As Integer

For i = 2 To Sheets("LISTES").Range("e65536").End(xlUp).Row
  CB1 = Sheets("LISTES").Range("e" & i)
  If CB1.ListIndex = -1 Then
  CB1.AddItem Sheets("LISTES").Range("e" & i)
  End If
Next i
CB1.Value = "Selectionner"
End Sub

