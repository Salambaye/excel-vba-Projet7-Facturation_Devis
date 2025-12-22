VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntete 
   Caption         =   "Données de l'entête"
   ClientHeight    =   11745
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15855
   OleObjectBlob   =   "frmEntete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEntete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label9_Click()

End Sub

Private Sub lblRefPresta_Click()

End Sub

Private Sub LabelLibelle_Click()

End Sub

Private Sub Label1_Click()

End Sub

Private Sub lblTelGestionnaire_Click()

End Sub

Private Sub lblTitre_Click()

End Sub

'Private Sub UserForm_Click()
'
'End Sub
Private Sub UserForm_Initialize()
    ' Initialiser le formulaire
'    Me.Annule = False
'    Me.ControlerPTC = True
'
'    ' Configuration du UserForm
'    With Me
'        .Caption = "Point de comptage"
'        .BackColor = RGB(245, 248, 250)          ' Bleu-gris très clair
'        .Width = 480
'        .Height = 280
'    End With

    ' Label de titre
    With lblTitre
        .Caption = "Veuillez renseigner les informations suivantes :"
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)          ' Blanc
        .BackColor = RGB(30, 58, 138)            ' Bleu marine profond
        .BackStyle = fmBackStyleOpaque
        .TextAlign = fmTextAlignCenter
    End With
    
    
'    ' Initialiser le checkbox (Coché par défaut)
'    With CheckBox1
'        .Value = False
'        .Caption = "Activer le contrôle des numéros PTC"
'        .Font.Name = "Segoe UI"
'        .Font.Size = 12
'        .Font.Bold = True
'        .ForeColor = RGB(150, 150, 150)          ' Gris 'RGB(30, 58, 138)            ' Bleu marine moderne
'        .BackColor = RGB(245, 248, 250)          ' Même que le fond
'        .BackStyle = fmBackStyleTransparent      ' Transparent
'    End With
'
'    ' Label de description
'    With lblDescription
'        .Caption = "Le contrôle de cohérence des numéros PTC sera ignoré"
'        .Font.Name = "Segoe UI"
'        .Font.Size = 10
'        .Font.Italic = True
'        .ForeColor = RGB(200, 0, 0)
'        .BackStyle = fmBackStyleTransparent
'        .WordWrap = True
'        .AutoSize = False
'        .TextAlign = fmTextAlignCenter
'    End With

'    ' Bouton Valider
'    With btnValider
'        .Caption = "Valider"
'        .Font.Name = "Segoe UI"
'        .Font.Size = 11
'        .Font.Bold = True
'        .BackColor = RGB(34, 197, 94)            ' Vert moderne (Tailwind green-500)
'        .Width = 100
'        .Height = 32
'    End With
'
'    ' Bouton Annuler
'    With btnAnnuler
'        .Caption = "Annuler"
'        .Font.Name = "Segoe UI"
'        .Font.Size = 11
'        .BackColor = RGB(255, 0, 0)
'        .Width = 100
'        .Height = 32
'    End With

End Sub

