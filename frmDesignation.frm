VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDesignation 
   Caption         =   "UserForm1"
   ClientHeight    =   6645
   ClientLeft      =   180
   ClientTop       =   675
   ClientWidth     =   12105
   OleObjectBlob   =   "frmDesignation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDesignation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' UserForm : frmDesignation
' Description : Formulaire pour choisir le mode de facturation (Détaillé ou Modification)
'========================================================================================

Public Annule As Boolean

Private Sub UserForm_Initialize()
    Me.Annule = False
    Me.StartUpPosition = 0
    Me.left = Application.left + (Application.width - Me.width) / 2
    Me.top = Application.top + (Application.Height - Me.Height) / 2
    
    ' ==================== CONFIGURATION DU USERFORM ====================
    With Me
        .BackColor = RGB(245, 248, 250)          ' Bleu-gris très clair
        
'        .width = 600
'        .Height = 300
        .caption = "Choix du mode de facturation"
    End With
    
    ' ==================== LABEL DE TITRE ====================
    With lblTitre
        .caption = "Choisissez le mode de facturation"
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)          ' Blanc
        .BackColor = RGB(30, 58, 138)            ' Bleu marine profond
        .BackStyle = fmBackStyleOpaque
        .TextAlign = fmTextAlignCenter
        .width = 580
        .Height = 30
        .top = 10
        .left = 10
    End With
    
    ' ==================== FRAME FOURNITURES ====================
    With FrameFournitures
        .caption = "Fournitures"
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.Bold = True
        .ForeColor = RGB(30, 58, 138)
        .BackColor = RGB(218, 233, 248)
'        .BackStyle = fmBackStyleOpaque
        .width = 580
        .Height = 120
        .top = 70
        .left = 10
    End With
    
    ' ==================== OPTION DÉTAILLÉ ====================
    With optDetaille
        .caption = "Détaillé"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = RGB(249, 115, 22)           ' Orange vif
        .BackStyle = fmBackStyleOpaque
        .width = 230
        .Height = 60
        .top = 30
        .left = 50
        .Value = True                             ' Sélectionné par défaut
    End With
    
    ' ==================== OPTION MODIFICATION ====================
    With optModification
        .caption = "Modification des prix"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = RGB(156, 163, 175)          ' Gris (non sélectionné)
        .BackStyle = fmBackStyleOpaque
        .width = 230
        .Height = 60
        .top = 30
        .left = 310
        .Value = False
    End With
    
    ' ==================== BOUTON VALIDER ====================
    With btnValider
        .caption = "Valider"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .BackColor = RGB(34, 197, 94)            ' Vert moderne
        .ForeColor = RGB(255, 255, 255)
        .width = 120
        .Height = 35
        .top = 220
        .left = 350
    End With
    
    ' ==================== BOUTON ANNULER ====================
    With btnAnnuler
        .caption = "Annuler"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .BackColor = RGB(239, 68, 68)            ' Rouge moderne
        .ForeColor = RGB(255, 255, 255)
        .width = 120
        .Height = 35
        .top = 220
        .left = 100
    End With
End Sub

'========================================================================================
' Événement : Clic sur l'option Détaillé
'========================================================================================
Private Sub optDetaille_Click()
    ' Mettre en surbrillance l'option sélectionnée (bleu)
    optDetaille.BackColor = RGB(30, 58, 138)
    ' Griser l'option non sélectionnée
    optModification.BackColor = RGB(156, 163, 175)
End Sub

'========================================================================================
' Événement : Clic sur l'option Modification
'========================================================================================
Private Sub optModification_Click()
    ' Mettre en surbrillance l'option sélectionnée (bleu)
    optModification.BackColor = RGB(30, 58, 138)
    ' Griser l'option non sélectionnée
    optDetaille.BackColor = RGB(156, 163, 175)
End Sub

'========================================================================================
' Événement : Clic sur le bouton Valider
'========================================================================================
Private Sub btnValider_Click()
    ' Vérifier qu'une option est sélectionnée
    If Not optDetaille.Value And Not optModification.Value Then
        MsgBox "Veuillez sélectionner un mode de facturation.", vbExclamation, "Sélection requise"
        Exit Sub
    End If
    
    ' L'opération n'est pas annulée
    Me.Annule = False
    Me.Hide
End Sub

'========================================================================================
' Événement : Clic sur le bouton Annuler
'========================================================================================
Private Sub btnAnnuler_Click()
    Me.Annule = True
    Me.Hide
End Sub

'========================================================================================
' Événement : Fermeture du formulaire avec la croix (X)
'========================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Annule = True
        Me.Hide
        Cancel = True
    End If
End Sub

'========================================================================================
' Événements vides pour éviter les erreurs
'========================================================================================
Private Sub FrameFournitures_Click()
    ' Nécessaire pour éviter les erreurs lors du clic sur le Frame
End Sub

Private Sub lblTitre_Click()
    ' Nécessaire pour éviter les erreurs lors du clic sur le Label
End Sub
