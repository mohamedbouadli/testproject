VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NOUVEAU 
   Caption         =   "NOUVELLE ENTREE"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11535
   OleObjectBlob   =   "NOUVEAU.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NOUVEAU"
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
If CB2.Value = "MEDINA" Then
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
                
End If
If CB2.Value = "SIEGE" Then
                If Not IsError(Application.Match(CB1, Sheets("SIEGE").Range("b:b"), 0)) Then
                    ligne = Application.Match(CB1, Sheets("SIEGE").Range("b:b"), 0)
                    If Sheets("SIEGE").Range("d" & ligne).Value > 5 Then
                    LB1.ForeColor = vbGreen
                    LB1 = Sheets("SIEGE").Range("d" & ligne).Value & " " & Sheets("SIEGE").Range("c" & ligne).Value
                    Else
                    LB1.ForeColor = vbRed
                    LB1 = Sheets("SIEGE").Range("d" & ligne).Value & " " & Sheets("SIEGE").Range("c" & ligne).Value
                    End If
                    
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
Else
    cm2.Enabled = False
    cm1.Enabled = False
End If



Dim ligne As Integer
If CB2.Value = "MEDINA" Then
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
                
End If
If CB2.Value = "SIEGE" Then
                If Not IsError(Application.Match(CB1, Sheets("SIEGE").Range("b:b"), 0)) Then
                    ligne = Application.Match(CB1, Sheets("SIEGE").Range("b:b"), 0)
                    If Sheets("SIEGE").Range("d" & ligne).Value > 5 Then
                    LB1.ForeColor = vbGreen
                    LB1 = Sheets("SIEGE").Range("d" & ligne).Value & " " & Sheets("SIEGE").Range("c" & ligne).Value
                    Else
                    LB1.ForeColor = vbRed
                    LB1 = Sheets("SIEGE").Range("d" & ligne).Value & " " & Sheets("SIEGE").Range("c" & ligne).Value
                    End If
                    
                End If
                
End If

End Sub

Private Sub CB2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
KeyAscii = 0
End Sub

Private Sub CM1_Click()
If CB1 <> "" And CB1 <> "Selectionner" And CB2 <> "" And CB2 <> "Selectionner" And TB1 <> "" Then
    cm1.Enabled = True
   


    If MsgBox("VOULEZ VOUS VRAIMENT FAIRE CETTE ENTREE ?", vbYesNo, "GMCPF") = vbYes Then

    
            If CB2.Value = "SIEGE" Then
                If Not IsError(Application.Match(CB1, Sheets("siege").Range("b:b"), 0)) Then
                    ligne = Application.Match(CB1, Sheets("siege").Range("b:b"), 0)
                    Sheets("siege").Range("d" & ligne).Value = Sheets("siege").Range("d" & ligne).Value + Round(TB1.Value, 3)
                End If
              
            End If
            
            If CB2.Value = "MEDINA" Then
                If Not IsError(Application.Match(CB1, Sheets("MEDINA").Range("b:b"), 0)) Then
                    ligne = Application.Match(CB1, Sheets("MEDINA").Range("b:b"), 0)
                    Sheets("MEDINA").Range("d" & ligne).Value = Sheets("MEDINA").Range("d" & ligne).Value + Round(TB1.Value, 3)
                End If
                
            End If
    
    End If
    
Else
    
    MsgBox "SVP REMPLIR TOUS LES CHAMPS !", vbOK, "GMCPF"
End If

CB1 = "Selectionner"
CB2 = "Selectionner"
TB1 = ""
LB1 = ""
cm1.Enabled = False
cm2.Enabled = False
End Sub



Private Sub CM2_Click()
CB1 = "Selectionner"
CB2 = "Selectionner"
TB1 = ""
LB1 = ""
cm2.Enabled = False
cm1.Enabled = False
End Sub

Private Sub CM3_Click()
Unload Me
AJOUTER.Show
End Sub

Private Sub CM4_Click()
Unload Me
ACCUEIL.Show
End Sub





Private Sub Frame2_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub TB1_Change()

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
            TB1.Text = ""
            cm2.Enabled = False
            cm1.Enabled = False
 
        End Select
End Sub

Private Sub TextBox1_Change()

End Sub



Private Sub UserForm_Activate()
cm2.Enabled = False
cm1.Enabled = False
LB1 = ""
End Sub
Private Sub UserForm_Initialize()
cm1.Enabled = False
CB2.Value = "Selectionner"
TB1.Text = ""
Dim i As Integer

    For i = 2 To Sheets("LISTES").Range("e65536").End(xlUp).Row
      CB1 = Sheets("LISTES").Range("b" & i)
      If CB1.ListIndex = -1 Then
      CB1.AddItem Sheets("LISTES").Range("e" & i)
      End If
    Next i
CB1.Value = "Selectionner"
End Sub
