VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDevisModification 
   Caption         =   "UserForm1"
   ClientHeight    =   12585
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   18675
   OleObjectBlob   =   "frmDevisModification.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDevisModification"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UserForm frmDevisModification - Code VBA

Public Annule As Boolean
Public dictLignes As Object
Private compteurLignes As Long

Private Sub UserForm_Activate()
 
    With Application
        LargeurFenetre = .width
        HauteurFenetre = .Height
        PositionGauche = .left
        PositionHaut = .top
    End With
    With Me
        .left = (PositionGauche + LargeurFenetre) - ((LargeurFenetre + .width) / 2)
        .top = (PositionHaut + HauteurFenetre) - ((HauteurFenetre + .Height) / 2)
    End With
 
End Sub


Private Sub UserForm_Initialize()
    Me.Annule = False
'    Me.StartUpPosition = 0
'    Me.left = Application.left + (Application.width - Me.width) / 2
'    Me.top = Application.top + (Application.Height - Me.Height) / 2
    
    ' Initialiser le dictionnaire
    Set dictLignes = CreateObject("Scripting.Dictionary")
    compteurLignes = 0
    
    ' Configuration du UserForm
    With Me
        .BackColor = RGB(245, 248, 250)
        .width = 1000
        .Height = 650
        .caption = "Modification des prix - Saisie manuelle"
    End With
    
    ' Titre
    With lblTitre
        .caption = "Saisie des lignes de devis"
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = RGB(30, 58, 138)
        .TextAlign = fmTextAlignCenter
        .width = 630
        .Height = 35
        .top = 5
        .left = 10
    End With
    
    ' Frame de saisie
    With frameSaisie
        .caption = "Nouvelle ligne"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(30, 58, 138)
        .width = 630
        .Height = 180
        .top = 50
        .left = 10
    End With
    
    Dim topPos As Long
    Dim leftLabel As Long
    Dim leftControl As Long
    
    topPos = 75
    leftLabel = 25
    leftControl = 150
    
    ' Désignation
    With lblDesignation
        .caption = "Désignation :"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(55, 65, 81)
        .top = topPos
        .left = leftLabel
        .width = 120
        .Height = 150
    End With
    
    With txtDesignation
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        '        .width = 450
        .Height = 22
        .top = topPos
        .left = leftControl
    End With
    
    topPos = topPos + 35
    
    ' Quantité
    With lblQuantite
        .caption = "Quantité :"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(55, 65, 81)
        .top = topPos
        .left = leftLabel
        .width = 120
    End With
    
    With txtQuantite
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .width = 100
        .Height = 22
        .top = topPos
        .left = leftControl
        .Value = "1"
    End With
    
    ' Prix unitaire
    With lblPrixUnitaire
        .caption = "Prix unitaire HT :"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(55, 65, 81)
        .top = topPos
        .left = 300
        .width = 120
    End With
    
    With txtPrixUnitaire
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .width = 100
        .Height = 22
        .top = topPos
        .left = 420
        .Value = "0"
    End With
    
    topPos = topPos + 35
    
    ' TVA
    With lblTVA
        .caption = "TVA % :"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(55, 65, 81)
        .top = topPos
        .left = leftLabel
        .width = 120
    End With
    
    With cboTVA
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .width = 100
        .Height = 22
        .top = topPos
        .left = leftControl
        .AddItem "5.5"
        .AddItem "10"
        .AddItem "20"
        .ListIndex = 1                           ' 10% par défaut
    End With
    
    ' Bouton Ajouter ligne
    With btnAjouterLigne
        .caption = "Ajouter la ligne"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = RGB(59, 130, 246)
        .ForeColor = RGB(255, 255, 255)
        .width = 130
        .Height = 32
        '        .top = topPos
        '        .left = 390
    End With
    
    ' Liste des lignes ajoutées
    With lblLignesAjoutees
        .caption = "Lignes ajoutées au devis :"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(30, 58, 138)
        .top = 245
        .left = 15
        .width = 630
    End With
    
    With lstLignes
        .Font.Name = "Consolas"
        .Font.Size = 9
        .width = 630
        .Height = 170
        .top = 270
        .left = 10
    End With
    
    ' Bouton Supprimer
    With btnSupprimerLigne
        .caption = "Supprimer la ligne"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .BackColor = RGB(239, 68, 68)
        .ForeColor = RGB(255, 255, 255)
        .width = 130
        .Height = 32
        '        .top = 450
        '        .left = 540
    End With
    
    ' Boutons principaux
    With btnValider
        .caption = "Générer le devis"
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = True
        '        .BackColor = RGB(34, 197, 94)
        '        .ForeColor = RGB(255, 255, 255)
        .width = 150
        .Height = 35
        .top = 490
        .left = 350
    End With
    
    With btnAnnuler
        .caption = "Annuler"
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = True
        '        .BackColor = RGB(239, 68, 68)
        '        .ForeColor = RGB(255, 255, 255)
        .width = 150
        .Height = 35
        .top = 490
        .left = 150
    End With
End Sub

Private Sub btnAjouterLigne_Click()
    Dim designation As String
    Dim quantite As Double
    Dim prixUnitaire As Double
    Dim tva As Double
    Dim montantHT As Double
    Dim cle As String
    
    ' Validation des champs
    designation = Trim(txtDesignation.Value)
    If designation = "" Then
        MsgBox "Veuillez saisir une désignation.", vbExclamation
        txtDesignation.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtQuantite.Value) Then
        MsgBox "Veuillez saisir une quantité valide.", vbExclamation
        txtQuantite.SetFocus
        Exit Sub
    End If
    quantite = CDbl(txtQuantite.Value)
    
    If Not IsNumeric(txtPrixUnitaire.Value) Then
        MsgBox "Veuillez saisir un prix unitaire valide.", vbExclamation
        txtPrixUnitaire.SetFocus
        Exit Sub
    End If
    prixUnitaire = CDbl(txtPrixUnitaire.Value)
    
    If cboTVA.ListIndex < 0 Then
        MsgBox "Veuillez sélectionner un taux de TVA.", vbExclamation
        cboTVA.SetFocus
        Exit Sub
    End If
    tva = CDbl(cboTVA.Value)
    
    montantHT = quantite * prixUnitaire
    
    ' Ajouter au dictionnaire
    compteurLignes = compteurLignes + 1
    cle = "Ligne" & compteurLignes
    
    Set dictLignes(cle) = CreateObject("Scripting.Dictionary")
    dictLignes(cle)("designation") = designation
    dictLignes(cle)("quantite") = quantite
    dictLignes(cle)("prix") = prixUnitaire
    dictLignes(cle)("tva") = tva
    
    ' Ajouter à la liste
    Dim ligneAffichage As String
    ligneAffichage = left(designation & Space(40), 40) & " | " & _
                     "Qté: " & Format(quantite, "0.00") & " | " & _
                     "PU: " & Format(prixUnitaire, "#,##0.00") & " € | " & _
                     "TVA: " & Format(tva, "0.0") & "% | " & _
                     "Total: " & Format(montantHT, "#,##0.00") & " €"
    
    lstLignes.AddItem ligneAffichage
    
    ' Réinitialiser les champs
    txtDesignation.Value = ""
    txtQuantite.Value = "1"
    txtPrixUnitaire.Value = "0"
    cboTVA.ListIndex = 1
    txtDesignation.SetFocus
End Sub

Private Sub btnSupprimerLigne_Click()
    If lstLignes.ListIndex >= 0 Then
        Dim index As Long
        index = lstLignes.ListIndex
        
        ' Supprimer du dictionnaire
        Dim cle As String
        cle = "Ligne" & (index + 1)
        If dictLignes.Exists(cle) Then
            dictLignes.Remove cle
        End If
        
        ' Supprimer de la liste
        lstLignes.RemoveItem index
        
        ' Réorganiser les clés du dictionnaire
        Call ReorganiserDictionnaire
    End If
End Sub

Private Sub ReorganiserDictionnaire()
    Dim nouveauDict As Object
    Dim compteur As Long
    Dim cle As Variant
    Dim nouvelleCle As String
    
    Set nouveauDict = CreateObject("Scripting.Dictionary")
    compteur = 0
    
    For Each cle In dictLignes.Keys
        compteur = compteur + 1
        nouvelleCle = "Ligne" & compteur
        Set nouveauDict(nouvelleCle) = dictLignes(cle)
    Next cle
    
    Set dictLignes = nouveauDict
    compteurLignes = compteur
End Sub

Private Sub btnValider_Click()
    If dictLignes.Count = 0 Then
        MsgBox "Veuillez ajouter au moins une ligne au devis.", vbExclamation
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

' Événements vides pour éviter les erreurs
Private Sub frameSaisie_Click()
End Sub

Private Sub lblTitre_Click()
End Sub


