VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEntete 
   Caption         =   "Données de l'entête"
   ClientHeight    =   14130
   ClientLeft      =   180
   ClientTop       =   705
   ClientWidth     =   24930
   OleObjectBlob   =   "frmEntete.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEntete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public Annule As Boolean
'
'Private Sub Label9_Click()
'
'End Sub
'
'Private Sub lblRefPresta_Click()
'
'End Sub
'
'Private Sub LabelLibelle_Click()
'
'End Sub
'
'Private Sub FrameFournitures_Click()
'
'End Sub
'
'Private Sub Label1_Click()
'
'End Sub
'
'Private Sub Label3_Click()
'
'End Sub
'
'Private Sub Label4_Click()
'
'End Sub
'
'Private Sub lblTelGestionnaire_Click()
'
'End Sub
'
'Private Sub lblTitre_Click()
'
'End Sub
'
'Private Sub btnValider_Click()
'    ' Récupérer l'état du checkbox
'    '    Me.ControlerPTC = CheckBox1.Value
'    Me.Annule = False
'    Me.Hide
'End Sub
'
'Private Sub btnAnnuler_Click()
'    Me.Annule = True
'    Me.Hide
'End Sub
'
'Private Sub ScrollBar1_Change()
'
'End Sub
'
''Private Sub UserForm_Click()
''
''End Sub
'Private Sub UserForm_Initialize()
'    ' Initialiser le formulaire
'    Me.Annule = False
'    Me.StartUpPosition = 2
'    '    optModification.Value = True
'
'    ' Configuration du UserForm
'    With Me
'        .BackColor = RGB(245, 248, 250)          ' Bleu-gris très clair
'        .Width = 805
'        .Height = 685
'    End With
'
'    ' Label de titre
'    With lblTitre
'        .Caption = "Veuillez renseigner les informations suivantes :"
'        .Font.Name = "Segoe UI"
'        .Font.Size = 14
'        .Font.Bold = True
'        .ForeColor = RGB(255, 255, 255)          ' Blanc
'        .BackColor = RGB(30, 58, 138)            ' Bleu marine profond
'        .BackStyle = fmBackStyleOpaque
'        .TextAlign = fmTextAlignCenter
'    End With
'
'    '        With Me.FrameFournitures
'    ''        .BackColor = RGB(255, 250, 220)
'    ''        .BorderColor = RGB(0, 120, 100)
'    '        .BackColor = RGB(255, 240, 230)
'    '        .BorderColor = RGB(0, 120, 100)
'    '        .SpecialEffect = fmSpecialEffectRaised
'    '    End With
'
'
'    '    ' Initialiser le checkbox (Coché par défaut)
'    '    With CheckBox1
'    '        .Value = False
'    '        .Caption = "Activer le contrôle des numéros PTC"
'    '        .Font.Name = "Segoe UI"
'    '        .Font.Size = 12
'    '        .Font.Bold = True
'    '        .ForeColor = RGB(150, 150, 150)          ' Gris 'RGB(30, 58, 138)            ' Bleu marine moderne
'    '        .BackColor = RGB(245, 248, 250)          ' Même que le fond
'    '        .BackStyle = fmBackStyleTransparent      ' Transparent
'    '    End With
'    '
'    '    ' Label de description
'    '    With lblDescription
'    '        .Caption = "Le contrôle de cohérence des numéros PTC sera ignoré"
'    '        .Font.Name = "Segoe UI"
'    '        .Font.Size = 10
'    '        .Font.Italic = True
'    '        .ForeColor = RGB(200, 0, 0)
'    '        .BackStyle = fmBackStyleTransparent
'    '        .WordWrap = True
'    '        .AutoSize = False
'    '        .TextAlign = fmTextAlignCenter
'    '    End With
'
'    '    ' Bouton Valider
'    '    With btnValider
'    '        .Caption = "Valider"
'    '        .Font.Name = "Segoe UI"
'    '        .Font.Size = 11
'    '        .Font.Bold = True
'    '        .BackColor = RGB(34, 197, 94)            ' Vert moderne (Tailwind green-500)
'    '        .Width = 100
'    '        .Height = 32
'    '    End With
'    '
'    '    ' Bouton Annuler
'    '    With btnAnnuler
'    '        .Caption = "Annuler"
'    '        .Font.Name = "Segoe UI"
'    '        .Font.Size = 11
'    '        .BackColor = RGB(255, 0, 0)
'    '        .Width = 100
'    '        .Height = 32
'    '    End With
'
'End Sub
'
'Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
'    ' Si l'utilisateur ferme avec la croix (X)
'    If CloseMode = vbFormControlMenu Then
'        Me.Annule = True
'        Me.Hide
'        Cancel = True
'    End If
'End Sub
'


'UserForm frmEntete - Code VBA

Public Annule As Boolean

Private Sub lblAdresseChantier_Click()

End Sub

Private Sub lblNomClient_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Annule = False
    Me.StartUpPosition = 0
    Me.left = Application.left + (Application.width - Me.width) / 2
    Me.top = Application.top + (Application.Height - Me.Height) / 2
    
    ' Configuration du UserForm
    With Me
        .BackColor = RGB(245, 248, 250)
        '        .width = 520
        '        .Height = 620
        .caption = "Informations du devis"
    End With
    
    ' Label de titre
    With lblTitre
        .caption = "Informations du devis"
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = RGB(30, 58, 138)
        .BackStyle = fmBackStyleOpaque
        .TextAlign = fmTextAlignCenter
        '        .width = 500
        .Height = 35
        .top = 5
        .left = 10
    End With
    
    '    With lblNomClient
    '        .ForeColor = RGB(55, 65, 81)
    ''        .BackColor = RGB(33, 92, 152)
    '        .BackStyle = fmBackStyleTransparent
    '    End With
        
        
    ' Positionner les contrôles
    Dim topPos As Long
    Dim leftLabel As Long
    Dim leftTextBox As Long
    Dim espacement As Long
    
    topPos = 50
    leftLabel = 15
    leftTextBox = 200
    espacement = 35
    
    ' Client
    '    ConfigurerLabel lblNomClient, "Nom du client * :", leftLabel, topPos
    '    ConfigurerTextBox txtNomClient, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblAdresseClient, "Adresse du client :", leftLabel, topPos
    '    ConfigurerTextBox txtAdresseClient, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblCpVille, "Code postal et ville :", leftLabel, topPos
    '    ConfigurerTextBox txtCpVille, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblRefclient, "Référence client :", leftLabel, topPos
    '    ConfigurerTextBox txtRefclient, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblRefUEBeep, "Réf UEX * :", leftLabel, topPos
    '    ConfigurerTextBox txtRefUEBeep, leftTextBox, topPos
    '
    '    ' Gestionnaire
    '    topPos = topPos + espacement + 10
    '    ConfigurerLabel lblGestionnaire, "Gestionnaire :", leftLabel, topPos
    '    ConfigurerTextBox txtGestionnaire, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblTelGestionnaire, "Tél. gestionnaire :", leftLabel, topPos
    '    ConfigurerTextBox txtTelGestionnaire, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblMailGestionnaire, "Mail gestionnaire :", leftLabel, topPos
    '    ConfigurerTextBox txtMailGestionnaire, leftTextBox, topPos
    '
    '    ' Chantier
    '    topPos = topPos + espacement + 10
    '    ConfigurerLabel lblAdresseChantier, "Adresse chantier * :", leftLabel, topPos
    '    ConfigurerTextBox txtAdresseChantier, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblCpChantier, "Code postal * :", leftLabel, topPos
    '    ConfigurerTextBox txtCpChantier, leftTextBox, topPos, 100
    '
    '    ConfigurerLabel lblVilleChantier, "Ville * :", leftTextBox + 110, topPos
    '    ConfigurerTextBox txtVilleChantier, leftTextBox + 160, topPos, 140
    '
    '    topPos = topPos + espacement
    ''    ConfigurerLabel lblEmplTravaux, "Emplacement travaux :", leftLabel, topPos
    ''    ConfigurerTextBox txtEmplTravaux, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblPresentation, "Présentation projet :", leftLabel, topPos
    '    ConfigurerTextBox txtPresentation, leftTextBox, topPos
    '
    '    topPos = topPos + espacement
    '    ConfigurerLabel lblDesignation, "Description :", leftLabel, topPos
    '    ConfigurerTextBox txtDesignation, leftTextBox, topPos
    
    ' Boutons
    topPos = topPos + espacement + 10
    
    With btnValider
        .caption = "Valider"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .BackColor = RGB(34, 197, 94)
        .ForeColor = RGB(255, 255, 255)
        '        .width = 100
        '        .Height = 32
        '        .top = topPos
        '        .left = 150
    End With
    
    With btnAnnuler
        .caption = "Annuler"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .BackColor = RGB(239, 68, 68)
        .ForeColor = RGB(255, 255, 255)
        '        .width = 100
        '        .Height = 32
        '        .top = topPos
        '        .left = 270
    End With
End Sub

'Private Sub ConfigurerLabel(ctrl As MSForms.Label, caption As String, left As Long, top As Long)
'    With ctrl
'        .caption = caption
'        .Font.Name = "Segoe UI"
'        .Font.Size = 10
'        .ForeColor = RGB(55, 65, 81)
'        .BackStyle = fmBackStyleTransparent
'        .width = 180
'        .Height = 18
'        .top = top
'        .left = left
'    End With
'End Sub
'
'Private Sub ConfigurerTextBox(ctrl As MSForms.TextBox, left As Long, top As Long, Optional width As Long = 300)
'    With ctrl
'        .Font.Name = "Segoe UI"
'        .Font.Size = 10
'        .width = width
'        .Height = 22
'        .top = top
'        .left = left
'        .BackColor = RGB(255, 255, 255)
'        .BorderStyle = fmBorderStyleSingle
'        .BorderColor = RGB(209, 213, 219)
'    End With
'End Sub

Private Sub btnValider_Click()
    ' Vérification des champs obligatoires
    If Trim(txtNomClient.Value) = "" Then
        MsgBox "Veuillez saisir le nom du client.", vbExclamation
        txtNomClient.SetFocus
        Exit Sub
    End If
    
    If Trim(txtRefUEBeep.Value) = "" Then
        MsgBox "Veuillez saisir la référence UEX.", vbExclamation
        txtRefUEBeep.SetFocus
        Exit Sub
    End If
    
    If Trim(txtAdresseChantier.Value) = "" Then
        MsgBox "Veuillez saisir l'adresse du chantier.", vbExclamation
        txtAdresseChantier.SetFocus
        Exit Sub
    End If
    
    If Trim(txtCpChantier.Value) = "" Then
        MsgBox "Veuillez saisir le code postal du chantier.", vbExclamation
        txtCpChantier.SetFocus
        Exit Sub
    End If
    
    If Trim(txtVilleChantier.Value) = "" Then
        MsgBox "Veuillez saisir la ville du chantier.", vbExclamation
        txtVilleChantier.SetFocus
        Exit Sub
    End If
    
    Me.Annule = False
    Me.Hide
End Sub

Private Sub btnAnnuler_Click()
    Me.Annule = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Annule = True
        Me.Hide
        Cancel = True
    End If
End Sub

