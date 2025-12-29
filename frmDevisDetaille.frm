VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDevisDetaille 
   Caption         =   "UserForm1"
   ClientHeight    =   13425
   ClientLeft      =   165
   ClientTop       =   690
   ClientWidth     =   28275
   OleObjectBlob   =   "frmDevisDetaille.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDevisDetaille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'========================================================================================
' UserForm : frmDevisDetaille
' Description : Formulaire pour sélectionner les fournitures et main d'œuvre en mode détaillé
'========================================================================================

Public Annule As Boolean
Public dictFournitures As Object
Public dictMainOeuvre As Object
Dim ws As Worksheet
Dim derniereLigne As Long

Private Sub lblElementsAjoutes_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.Annule = False
    Me.StartUpPosition = 0
    Me.left = Application.left + (Application.width - Me.width) / 2
    Me.top = Application.top + (Application.Height - Me.Height) / 2
    
    ' Initialiser les dictionnaires
    Set dictFournitures = CreateObject("Scripting.Dictionary")
    Set dictMainOeuvre = CreateObject("Scripting.Dictionary")
    
    ' Configuration du UserForm
    With Me
        .BackColor = RGB(245, 248, 250)
        .width = 1000
        .Height = 800
        .caption = "Devis détaillé - Sélection des éléments"
    End With
    
    ' ==================== TITRE ====================
    With lblTitre
        .caption = "Sélection des fournitures et main d'œuvre"
        .Font.Name = "Segoe UI"
        .Font.Size = 14
        .Font.Bold = True
        .ForeColor = RGB(255, 255, 255)
        .BackColor = RGB(30, 58, 138)
        .TextAlign = fmTextAlignCenter
        .width = 680
        .Height = 35
        .top = 5
        .left = 10
    End With
    
    ' Charger les listes depuis les feuilles Tarification
    Call ChargerListeFournitures
    Call ChargerListeMainOeuvre
    
    ' ==================== LISTE FOURNITURES ====================
    With lstFournitures
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        '        .width = 320
        '        .Height = 250
        '        .top = 90
        '        .left = 15
        .MultiSelect = fmMultiSelectMulti
    End With
    
    With lblFournitures
        .caption = "Fournitures disponibles :"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(30, 58, 138)
        '        .top = 65
        '        .left = 15
        '        .width = 320
    End With
    
    With lblQteFournitures
        .caption = "Quantité :"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        '        .top = 350
        '        .left = 15
    End With
    
    With txtQteFournitures
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        '        .width = 100
        '        .top = 348
        '        .left = 85
        .Value = "1"
    End With
    
    With btnAjouterFourniture
        .caption = "Ajouter"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = RGB(59, 130, 246)
        .ForeColor = RGB(255, 255, 255)
        '        .width = 80
        '        .Height = 28
        '        .top = 345
        '        .left = 200
    End With
    
    ' ==================== LISTE MAIN D'ŒUVRE ====================
    With lstMainOeuvre
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        '        .width = 320
        '        .Height = 250
        '        .top = 90
        '        .left = 355
        .MultiSelect = fmMultiSelectMulti
    End With
    
    With lblMainOeuvre
        .caption = "Main d'œuvre disponible :"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(30, 58, 138)
        '        .top = 65
        '        .left = 355
        '        .width = 320
    End With
    
    With lblHeuresMainOeuvre
        .caption = "Heures :"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        '        .top = 350
        '        .left = 355
    End With
    
    With txtHeuresMainOeuvre
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        '        .width = 100
        '        .top = 348
        '        .left = 420
        .Value = "1"
    End With
    
    With btnAjouterMainOeuvre
        .caption = "Ajouter"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = RGB(59, 130, 246)
        .ForeColor = RGB(255, 255, 255)
        '        .width = 80
        '        .Height = 28
        '        .top = 345
        '        .left = 540
    End With
    
    ' ==================== LISTE ÉLÉMENTS AJOUTÉS ====================
    With lblElementsAjoutes
        .caption = "Éléments ajoutés au devis :"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .ForeColor = RGB(30, 58, 138)
        '        .top = 395
        '        .left = 15
        '        .width = 660
    End With
    
    With lstElementsAjoutes
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        '        .width = 660
        '        .Height = 100
        '        .top = 420
        '        .left = 15
    End With
    
    With btnSupprimerElement
        .caption = "Supprimer"
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .BackColor = RGB(239, 68, 68)
        .ForeColor = RGB(255, 255, 255)
        '        .width = 100
        '        .Height = 28
        '        .top = 530
        '        .left = 575
    End With
    
    ' ==================== BOUTONS PRINCIPAUX ====================
    With btnValider
        .caption = "Générer le devis"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .Font.Bold = True
        .BackColor = RGB(34, 197, 94)
        .ForeColor = RGB(255, 255, 255)
        '        .width = 150
        '        .Height = 35
        '        .top = 565
        '        .left = 250
    End With
    
    With btnAnnuler
        .caption = "Annuler"
        .Font.Name = "Segoe UI"
        .Font.Size = 11
        .BackColor = RGB(239, 68, 68)
        .ForeColor = RGB(255, 255, 255)
        '        .width = 100
        '        .Height = 35
        '        .top = 565
        '        .left = 420
    End With
End Sub

'========================================================================================
' Chargement de la liste des fournitures depuis les feuilles Tarification
'========================================================================================
Private Sub ChargerListeFournitures()
    '    Dim ws As Worksheet
    ''    Dim derniereLigne As Long
    Dim i As Long
    Dim item As String
    
    lstFournitures.Clear
    
    ' ========== PLOMBERIE ==========
    Set ws = wsTarifPlomberie
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 4 To derniereLigne
        If Trim(ws.Cells(i, 1).Value) <> "" Or Trim(ws.Cells(i, 2).Value) <> "" Then
            item = Trim(ws.Cells(i, 1).Value) & " " & Trim(ws.Cells(i, 2).Value)
            lstFournitures.AddItem "[PLOMB] " & item & " - " & Format(ws.Cells(i, 5).Value, "#,##0.00") & " €"
        End If
    Next i
    
    ' ========== CHAUFFAGE ==========
    Set ws = wsTarifChauffage
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 4 To derniereLigne
        If Trim(ws.Cells(i, 1).Value) <> "" Or Trim(ws.Cells(i, 2).Value) <> "" Then
            item = Trim(ws.Cells(i, 1).Value) & " " & Trim(ws.Cells(i, 2).Value)
            lstFournitures.AddItem "[CHAUF] " & item & " - " & Format(ws.Cells(i, 5).Value, "#,##0.00") & " €"
        End If
    Next i
    
    ' ========== COMPTEURS D'EAU ==========
    Set ws = wsTarifClient
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 4 To derniereLigne
        If Trim(ws.Cells(i, 1).Value) <> "" Or Trim(ws.Cells(i, 2).Value) <> "" Then
            item = Trim(ws.Cells(i, 1).Value) & " " & Trim(ws.Cells(i, 2).Value)
            lstFournitures.AddItem "[COMPT] " & item & " - " & Format(ws.Cells(i, 5).Value, "#,##0.00") & " €"
        End If
    Next i
    
        ' ========== VANNES ==========
        Set ws = wsTarifVenteDeVannes
        derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For i = 4 To derniereLigne
            If Trim(ws.Cells(i, 1).Value) <> "" Or Trim(ws.Cells(i, 2).Value) <> "" Then
                item = Trim(ws.Cells(i, 1).Value) & " " & Trim(ws.Cells(i, 2).Value) & " Ø" & ws.Cells(i, 3).Value
                lstFournitures.AddItem "[VANNE] " & item & " - " & Format(ws.Cells(i, 5).Value, "#,##0.00") & " €"
            End If
        Next i
End Sub

'========================================================================================
' Chargement de la liste de la main d'œuvre depuis les feuilles Tarification
'========================================================================================
Private Sub ChargerListeMainOeuvre()
    Dim ws As Worksheet
    Dim derniereLigne As Long
    Dim i As Long
    Dim item As String
    
    lstMainOeuvre.Clear
    
    ' ========== TARIF GÉNÉRIQUE ==========
    Set ws = wsTarifGenerique
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 7 To derniereLigne
        If Trim(ws.Cells(i, 1).Value) <> "" Or Trim(ws.Cells(i, 2).Value) <> "" Then
            item = Trim(ws.Cells(i, 1).Value) & " " & Trim(ws.Cells(i, 2).Value)
            lstMainOeuvre.AddItem item & " - " & Format(ws.Cells(i, 3).Value, "#,##0.00") & " €/h"
        End If
    Next i
    
    ' ========== TARIF PASSAGE SUPPLÉMENTAIRE ==========
    Set ws = wsTarifPassage
    derniereLigne = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 4 To derniereLigne
        If Trim(ws.Cells(i, 1).Value) <> "" Or Trim(ws.Cells(i, 2).Value) <> "" Then
            item = Trim(ws.Cells(i, 1).Value) & " " & Trim(ws.Cells(i, 2).Value)
            lstMainOeuvre.AddItem "[PASSAGE] " & item & " - " & Format(ws.Cells(i, 5).Value, "#,##0.00") & " €"
        End If
    Next i
End Sub

'========================================================================================
' Ajouter les fournitures sélectionnées au devis
'========================================================================================
Private Sub btnAjouterFourniture_Click()
    Dim i As Long
    Dim item As String
    Dim prix As Double
    Dim quantite As Long
    
    If txtQteFournitures.Value = "" Or Not IsNumeric(txtQteFournitures.Value) Then
        MsgBox "Veuillez saisir une quantité valide.", vbExclamation
        Exit Sub
    End If
    
    quantite = CLng(txtQteFournitures.Value)
    
    For i = 0 To lstFournitures.ListCount - 1
        If lstFournitures.Selected(i) Then
            item = lstFournitures.List(i)
            prix = ExtrairePrix(item)
            
            If Not dictFournitures.Exists(item) Then
                dictFournitures.Add item, CreateObject("Scripting.Dictionary")
                dictFournitures(item)("quantite") = quantite
                dictFournitures(item)("prix") = prix
                lstElementsAjoutes.AddItem "[F] " & item & " x" & quantite
            End If
        End If
    Next i
    
    txtQteFournitures.Value = "1"
End Sub

'========================================================================================
' Ajouter la main d'œuvre sélectionnée au devis
'========================================================================================
Private Sub btnAjouterMainOeuvre_Click()
    Dim i As Long
    Dim item As String
    Dim prix As Double
    Dim heures As Double
    
    If txtHeuresMainOeuvre.Value = "" Or Not IsNumeric(txtHeuresMainOeuvre.Value) Then
        MsgBox "Veuillez saisir un nombre d'heures valide.", vbExclamation
        Exit Sub
    End If
    
    heures = CDbl(txtHeuresMainOeuvre.Value)
    
    For i = 0 To lstMainOeuvre.ListCount - 1
        If lstMainOeuvre.Selected(i) Then
            item = lstMainOeuvre.List(i)
            prix = ExtrairePrix(item)
            
            If Not dictMainOeuvre.Exists(item) Then
                dictMainOeuvre.Add item, CreateObject("Scripting.Dictionary")
                dictMainOeuvre(item)("heures") = heures
                dictMainOeuvre(item)("prix") = prix
                lstElementsAjoutes.AddItem "[MO] " & item & " x" & heures & "h"
            End If
        End If
    Next i
    
    txtHeuresMainOeuvre.Value = "1"
End Sub

'========================================================================================
' Extraire le prix d'une ligne de texte
'========================================================================================
'Private Function ExtrairePrix(texte As String) As Double
'    Dim pos As Long
'    Dim prixStr As String
'
'    pos = InStrRev(texte, " - ")
'    If pos > 0 Then
'        prixStr = Mid(texte, pos + 3)
'        prixStr = Replace(prixStr, " €", "")
'        prixStr = Replace(prixStr, " €/h", "")
'        prixStr = Replace(prixStr, ",", ".")
'        prixStr = Replace(prixStr, " ", "")
'        ExtrairePrix = CDbl(prixStr)
'    End If
'End Function

Private Function ExtrairePrix(texte As String) As Double
    On Error GoTo GestionErreur
    
    Dim pos As Long
    Dim prixStr As String
    Dim i As Integer
    Dim resultat As String
    
    ' Initialiser la valeur de retour
    ExtrairePrix = 0
    
    ' Trouver la position du séparateur " - "
    pos = InStrRev(texte, " - ")
    If pos = 0 Then Exit Function
    
    ' Extraire la partie prix
    prixStr = Mid(texte, pos + 3)
    
    ' Nettoyer la chaîne en gardant uniquement chiffres, point et virgule
    resultat = ""
    For i = 1 To Len(prixStr)
        Select Case Mid(prixStr, i, 1)
        Case "0" To "9", ".", ","
            resultat = resultat & Mid(prixStr, i, 1)
        End Select
    Next i
    
    ' Remplacer la virgule par un point (séparateur décimal)
    resultat = Replace(resultat, ",", ".")
    
    ' Vérifier qu'on a bien une valeur
    If Len(resultat) > 0 And IsNumeric(resultat) Then
        ExtrairePrix = CDbl(resultat)
    End If
    
    Exit Function
    
GestionErreur:
    Debug.Print "Erreur ExtrairePrix : " & Err.Description & " - Texte : " & texte
    ExtrairePrix = 0
End Function

'========================================================================================
' Supprimer un élément de la liste
'========================================================================================
Private Sub btnSupprimerElement_Click()
    If lstElementsAjoutes.ListIndex >= 0 Then
        lstElementsAjoutes.RemoveItem lstElementsAjoutes.ListIndex
    End If
End Sub

'========================================================================================
' Valider et générer le devis
'========================================================================================
Private Sub btnValider_Click()
    If dictFournitures.Count = 0 And dictMainOeuvre.Count = 0 Then
        MsgBox "Veuillez ajouter au moins un élément au devis.", vbExclamation
        Exit Sub
    End If
    
    Me.Annule = False
    Me.Hide
End Sub

'========================================================================================
' Annuler l'opération
'========================================================================================
Private Sub btnAnnuler_Click()
    Me.Annule = True
    Me.Hide
End Sub

'========================================================================================
' Gestion de la fermeture du formulaire
'========================================================================================
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Me.Annule = True
        Me.Hide
        Cancel = True
    End If
End Sub

