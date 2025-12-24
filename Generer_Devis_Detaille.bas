Attribute VB_Name = "Generer_Devis_Detaille"
'========================================================================================
' Module : Génération du Devis en Mode Détaillé
' Description : Génère un devis avec fournitures, main d'œuvre et déplacement détaillés
'========================================================================================

Sub GenererDevisDetaille()
    Dim ligneDebut As Long
    Dim ligneActuelle As Long
    
    ligneDebut = 26
    ligneActuelle = ligneDebut
    
    ' ========== Afficher le formulaire de sélection détaillée ==========
    frmDevisDetaille.Annule = True
    frmDevisDetaille.Show
    
    If frmDevisDetaille.Annule = True Then
        MsgBox "Opération annulée par l'utilisateur.", vbInformation
        Unload frmDevisDetaille
        Exit Sub
    End If
    
    ' ========== Créer les en-têtes du tableau ==========
    Call CreerEntetesTableauDetaille(ligneDebut)
    
    ligneActuelle = ligneDebut + 2
    
    ' ========== Ajouter la description ==========
    With wsDevis
        .Cells(ligneActuelle, 1).Value = descriptionDesignation
        .Cells(ligneActuelle, 1).Font.Bold = True
        .Cells(ligneActuelle, 1).Font.Size = 11
        .Cells(ligneActuelle, 1).Font.Color = RGB(30, 58, 138)
        ligneActuelle = ligneActuelle + 1
    End With
    
    ' ========== Variables pour les totaux ==========
    Dim totalFournitures As Double
    Dim totalMainOeuvre As Double
    Dim totalDeplacement As Double
    Dim totalHT As Double
    Dim montantTVA As Double
    Dim totalTTC As Double
    
    totalFournitures = 0
    totalMainOeuvre = 0
    totalDeplacement = 0
    
    ' ========== Ajouter les fournitures sélectionnées ==========
    If frmDevisDetaille.dictFournitures.Count > 0 Then
        ligneActuelle = AjouterFournitures(ligneActuelle, totalFournitures)
    End If
    
    ' ========== Ajouter la main d'œuvre ==========
    If frmDevisDetaille.dictMainOeuvre.Count > 0 Then
        ligneActuelle = AjouterMainOeuvre(ligneActuelle, totalMainOeuvre)
    End If
    
    ' ========== Ajouter le déplacement ==========
    ligneActuelle = AjouterDeplacement(ligneActuelle, totalDeplacement)
    
    ' ========== Ligne de séparation ==========
    ligneActuelle = ligneActuelle + 1
    
    ' ========== Calcul des totaux ==========
    totalHT = totalFournitures + totalMainOeuvre + totalDeplacement
    montantTVA = totalHT * 0.1 ' TVA 10%
    totalTTC = totalHT + montantTVA
    
    ' ========== Afficher les totaux ==========
    Call AfficherTotaux(ligneActuelle, totalHT, montantTVA, totalTTC)
    
    Unload frmDevisDetaille
End Sub

'========================================================================================
' Créer les en-têtes du tableau détaillé
'========================================================================================
Sub CreerEntetesTableauDetaille(ligne As Long)
    With wsDevis
        ' ========== En-têtes ==========
        .Cells(ligne, 1).Value = "Désignation"
        .Cells(ligne, 2).Value = "Fournitures"
        .Cells(ligne, 3).Value = "Main d'œuvre"
        .Cells(ligne, 4).Value = "Déplacement"
        .Cells(ligne, 5).Value = "Total HT"
        
        ' ========== Mise en forme des en-têtes ==========
        With .Range(.Cells(ligne, 1), .Cells(ligne, 5))
            .Font.Bold = True
            .Font.Size = 11
            .Font.Color = RGB(255, 255, 255)
            .Interior.Color = RGB(79, 129, 189)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
        
        ' ========== Largeur des colonnes ==========
        .Columns("A:A").ColumnWidth = 50
        .Columns("B:B").ColumnWidth = 18
        .Columns("C:C").ColumnWidth = 18
        .Columns("D:D").ColumnWidth = 18
        .Columns("E:E").ColumnWidth = 18
    End With
End Sub

'========================================================================================
' Ajouter les fournitures au tableau
'========================================================================================
Function AjouterFournitures(ligneDebut As Long, ByRef total As Double) As Long
    Dim ligne As Long
    Dim item As Variant
    Dim designation As String
    Dim prix As Double
    Dim quantite As Long
    Dim montant As Double
    
    ligne = ligneDebut
    
    ' ========== Parcourir le dictionnaire des fournitures ==========
    For Each item In frmDevisDetaille.dictFournitures.Keys
        With wsDevis
            ' Extraire la désignation (enlever le préfixe [PLOMB], [CHAUF], etc.)
            designation = item
            If InStr(designation, "]") > 0 Then
                designation = Trim(Mid(designation, InStr(designation, "]") + 1))
            End If
            
            ' Extraire le prix de la désignation
            If InStr(designation, " - ") > 0 Then
                designation = left(designation, InStr(designation, " - ") - 1)
            End If
            
            .Cells(ligne, 1).Value = designation
            .Cells(ligne, 1).Font.Size = 10
            
            quantite = frmDevisDetaille.dictFournitures(item)("quantite")
            prix = frmDevisDetaille.dictFournitures(item)("prix")
            montant = prix * quantite
            
            .Cells(ligne, 2).Value = Format(montant, "#,##0.00") & " €"
            .Cells(ligne, 2).HorizontalAlignment = xlRight
            .Cells(ligne, 2).Font.Size = 10
            
            total = total + montant
            
            ' ========== Bordures ==========
            .Range(.Cells(ligne, 1), .Cells(ligne, 5)).Borders.LineStyle = xlContinuous
            .Range(.Cells(ligne, 1), .Cells(ligne, 5)).Borders.Color = RGB(200, 200, 200)
            
            ligne = ligne + 1
        End With
    Next item
    
    AjouterFournitures = ligne
End Function

'========================================================================================
' Ajouter la main d'œuvre au tableau
'========================================================================================
Function AjouterMainOeuvre(ligneDebut As Long, ByRef total As Double) As Long
    Dim ligne As Long
    Dim item As Variant
    Dim designation As String
    Dim prix As Double
    Dim heures As Double
    Dim montant As Double
    
    ligne = ligneDebut
    
    ' ========== Parcourir le dictionnaire de la main d'œuvre ==========
    For Each item In frmDevisDetaille.dictMainOeuvre.Keys
        With wsDevis
            ' Extraire la désignation
            designation = item
            If InStr(designation, "]") > 0 Then
                designation = Trim(Mid(designation, InStr(designation, "]") + 1))
            End If
            
            ' Extraire le prix de la désignation
            If InStr(designation, " - ") > 0 Then
                designation = left(designation, InStr(designation, " - ") - 1)
            End If
            
            .Cells(ligne, 1).Value = designation
            .Cells(ligne, 1).Font.Size = 10
            
            heures = frmDevisDetaille.dictMainOeuvre(item)("heures")
            prix = frmDevisDetaille.dictMainOeuvre(item)("prix")
            montant = prix * heures
            
            .Cells(ligne, 3).Value = Format(montant, "#,##0.00") & " €"
            .Cells(ligne, 3).HorizontalAlignment = xlRight
            .Cells(ligne, 3).Font.Size = 10
            
            total = total + montant
            
            ' ========== Bordures ==========
            .Range(.Cells(ligne, 1), .Cells(ligne, 5)).Borders.LineStyle = xlContinuous
            .Range(.Cells(ligne, 1), .Cells(ligne, 5)).Borders.Color = RGB(200, 200, 200)
            
            ligne = ligne + 1
        End With
    Next item
    
    AjouterMainOeuvre = ligne
End Function

'========================================================================================
' Ajouter le déplacement au tableau
'========================================================================================
Function AjouterDeplacement(ligneDebut As Long, ByRef total As Double) As Long
    Dim ligne As Long
    Dim prixDeplacement As Double
    Dim tvaDeplacement As Double
    
    ligne = ligneDebut
    
    ' ========== Récupérer le prix du déplacement depuis Tarif générique 2025 ==========
    On Error Resume Next
    prixDeplacement = wsTarifGenerique.Cells(4, 5).Value ' Colonne E, ligne 4
    tvaDeplacement = wsTarifGenerique.Cells(4, 4).Value ' Colonne D, ligne 4
    On Error GoTo 0
    
    If prixDeplacement = 0 Then
        prixDeplacement = 50 ' Valeur par défaut si non trouvée
    End If
    
    With wsDevis
        .Cells(ligne, 1).Value = "Déplacement"
        .Cells(ligne, 1).Font.Size = 10
        
        .Cells(ligne, 4).Value = Format(prixDeplacement, "#,##0.00") & " €"
        .Cells(ligne, 4).HorizontalAlignment = xlRight
        .Cells(ligne, 4).Font.Size = 10
        
        total = prixDeplacement
        
        ' ========== Bordures ==========
        .Range(.Cells(ligne, 1), .Cells(ligne, 5)).Borders.LineStyle = xlContinuous
        .Range(.Cells(ligne, 1), .Cells(ligne, 5)).Borders.Color = RGB(200, 200, 200)
        
        ligne = ligne + 1
    End With
    
    AjouterDeplacement = ligne
End Function

'========================================================================================
' Afficher les totaux HT, TVA et TTC
'========================================================================================
Sub AfficherTotaux(ligne As Long, totalHT As Double, montantTVA As Double, totalTTC As Double)
    With wsDevis
        ' ========== Ligne vide ==========
        ligne = ligne + 1
        
        ' ========== Total HT ==========
        .Cells(ligne, 4).Value = "Total HT :"
        .Cells(ligne, 4).Font.Bold = True
        .Cells(ligne, 4).Font.Size = 11
        .Cells(ligne, 4).HorizontalAlignment = xlRight
        
        .Cells(ligne, 5).Value = Format(totalHT, "#,##0.00") & " €"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).Font.Size = 11
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        
        ligne = ligne + 1
        
        ' ========== TVA 10% ==========
        .Cells(ligne, 4).Value = "TVA 10% :"
        .Cells(ligne, 4).Font.Bold = True
        .Cells(ligne, 4).Font.Size = 11
        .Cells(ligne, 4).HorizontalAlignment = xlRight
        
        .Cells(ligne, 5).Value = Format(montantTVA, "#,##0.00") & " €"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).Font.Size = 11
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        
        ligne = ligne + 1
        
        ' ========== Total TTC ==========
        .Cells(ligne, 4).Value = "TOTAL TTC :"
        .Cells(ligne, 4).Font.Bold = True
        .Cells(ligne, 4).Font.Size = 12
        .Cells(ligne, 4).HorizontalAlignment = xlRight
        .Cells(ligne, 4).Font.Color = RGB(0, 0, 255)
        
        .Cells(ligne, 5).Value = Format(totalTTC, "#,##0.00") & " €"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).Font.Size = 12
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        .Cells(ligne, 5).Font.Color = RGB(0, 0, 255)
        
        ' ========== Bordure pour le total TTC ==========
        With .Range(.Cells(ligne, 4), .Cells(ligne, 5))
'            .Borders(xlEdgeTop).LineStyle = xlContinuous
'            .Borders(xlEdgeTop).Weight = xlThick
'            .Borders(xlEdgeBottom).LineStyle = xlDouble
'            .Borders(xlEdgeBottom).Weight = xlThick
            .Interior.Color = RGB(217, 217, 217)
        End With
        
        ' ========== Texte de fin ==========
        ligne = ligne + 3
        .Cells(ligne, 1).Value = "Conditions de règlement : A réception de la facture"
        .Cells(ligne, 1).Font.Italic = True
        .Cells(ligne, 1).Font.Size = 16
        .Cells(ligne, 1).Font.Name = "Times New Roman"
        
        ligne = ligne + 1
        .Cells(ligne, 1).Value = "Mode de règlement : chèque ou virement."
        .Cells(ligne, 1).Font.Italic = True
        .Cells(ligne, 1).Font.Bold = True
        .Cells(ligne, 1).Font.Size = 16
        .Cells(ligne, 1).Font.Name = "Times New Roman"
'        .Cells(ligne, 1).Font.Color = RGB(100, 100, 100)
        
         ligne = ligne + 1
        .Cells(ligne, 1).Value = "Ce devis est valable 30 jours à compter de sa date de réalisation."
        .Cells(ligne, 1).Font.Italic = True
        .Cells(ligne, 1).Font.Bold = True
        .Cells(ligne, 1).Font.Size = 16
        .Cells(ligne, 1).Font.Name = "Times New Roman"
'        .Cells(ligne, 1).Font.Color = RGB(100, 100, 100)

        ligne = ligne + 4
        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
        .Merge
        .Value = "Si ce devis vous convient, veuillez nous le retourner signé précédé de la mention: "
        .Font.Italic = True
        .Font.Bold = True
        .Font.Size = 24
        .Font.Name = "Times New Roman"
        End With
    End With
End Sub

'    Siège social : 27 rue Carnot 91300 MASSY
'   Tél standard : 01 64 54 27 99
'Siret : 582 017 810 00414    S.N.C au Capital de 3 034 169 euros
'RCS Evry - NAF 7739Z
'N° intracommunautaire : FR 92582017810      www.istablog.fr   www.ista.fr
