Attribute VB_Name = "Generer_Devis_Detaille"
'----------------------------------------------------------------------------------------
' Module : Génération du Devis en Mode Détaillé
' Description : Génère un devis avec fournitures, main d'œuvre et déplacement détaillés
'----------------------------------------------------------------------------------------

Sub GenererDevisDetaille()
    Dim ligneDebut As Long
    Dim ligneActuelle As Long
    
    ligneDebut = 25
    ligneActuelle = ligneDebut
    
    ' ---------- Afficher le formulaire de sélection détaillée ----------
    frmDevisDetaille.Annule = True
    frmDevisDetaille.Show

    If frmDevisDetaille.Annule Then
        MsgBox "Opération annulée par l'utilisateur.", vbInformation
        Unload frmDevisDetaille
        Exit Sub
    End If
    
    ' ---------- Créer les en-têtes du tableau ----------
    Call CreerEntetesTableauDetaille(ligneDebut)
    
    ligneActuelle = ligneDebut + 1
    
    ' ---------- Ajouter la description ----------
    With wsDevis
        .Cells(ligneActuelle, 1).Value = descriptionDesignation
        .Cells(ligneActuelle, 1).Font.Bold = True
        .Cells(ligneActuelle, 1).Font.Size = 16
        .Cells(ligneActuelle, 1).Font.Color = RGB(30, 58, 138)
        ligneActuelle = ligneActuelle + 1
    End With
    
    ' ---------- Variables pour les totaux ----------
    Dim totalHT As Double
    Dim montantTVA As Double
    Dim totalTTC As Double
    Dim totalTVA As Double
    
    totalHT = 0
    totalTVA = 0
    
    ' Ajouter les lignes saisies
    If frmDevisDetaille.dictFournitures.Count > 0 Or frmDevisDetaille.dictMainOeuvre.Count > 0 Then
        ligneActuelle = AjouterLignesDetaille(ligneActuelle, totalHT, totalTVA)
    End If
    
    ' Compléter le tableau jusqu'à 15 lignes minimum
    Dim ligneFinTableau As Long
    ligneFinTableau = ligneDebut + 16            ' En-tête + 15 lignes de contenu
    
    ' Compléter avec des lignes vides si nécessaire AVEC FORMATAGE
    Do While ligneActuelle < ligneFinTableau
        With wsDevis
            ' Appliquer le même formatage que les lignes de données
            With .Range(.Cells(ligneActuelle, 1), .Cells(ligneActuelle, 6))
                .Font.Name = "Arial"
                .Font.Size = 20
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
            .Rows(ligneActuelle).RowHeight = 20
        End With
        ligneActuelle = ligneActuelle + 1
    Loop
    
    ' Appliquer les bordures au tableau complet
    Call AppliquerBorduresTableau(ligneDebut, ligneFinTableau)
    
    ' ========== Calcul du total TTC ==========
    totalTTC = totalHT + totalTVA
    
    ' ========== Afficher les totaux ==========
    Call AfficherTotauxDetaille(ligneFinTableau, totalHT, totalTVA, totalTTC)
    
    Unload frmDevisDetaille
End Sub

'----------------------------------------------------------------------------------------
' Créer les en-têtes du tableau détaillé
'----------------------------------------------------------------------------------------
Sub CreerEntetesTableauDetaille(ligne As Long)
    With wsDevis
        ' En-têtes
        .Cells(ligne, 1).Value = "Désignation"
        .Cells(ligne, 2).Value = "Qté"
        .Cells(ligne, 3).Value = "Prix unitaire"
        .Cells(ligne, 4).Value = "Total HT"
        .Cells(ligne, 5).Value = "TVA"
        .Cells(ligne, 6).Value = "Total TTC"
        
        ' Mise en forme des en-têtes
        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
            .Font.Bold = True
            .Font.Color = RGB(0, 0, 0)
            .Font.Name = "Arial"
            .Font.Size = 20
            .Interior.Color = RGB(237, 242, 247)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        .Rows(ligne).RowHeight = 26.25
    End With
End Sub

'========================================================================================
' Ajouter les lignes détaillées au tableau
'========================================================================================
Function AjouterLignesDetaille(ligneDebut As Long, ByRef totalHT As Double, ByRef totalTVA As Double) As Long
    Dim ligne As Long
    Dim item As Variant
    Dim designation As String
    Dim categorie As String
    Dim prix As Double
    Dim quantite As Double
    Dim montantHT As Double
    Dim tva As Double
    Dim montantTVA As Double
    Dim montantTTC As Double
    
    ligne = ligneDebut
    
    ' ========== Traiter les fournitures ==========
    For Each item In frmDevisDetaille.dictFournitures.Keys
        designation = item
        categorie = "Fournitures"
        
        ' Extraire la désignation propre
        If InStr(designation, "]") > 0 Then
            designation = Trim(Mid(designation, InStr(designation, "]") + 1))
        End If
        If InStr(designation, " - ") > 0 Then
            designation = left(designation, InStr(designation, " - ") - 1)
        End If
        
        designation = categorie & " - " & designation
        quantite = frmDevisDetaille.dictFournitures(item)("quantite")
        prix = frmDevisDetaille.dictFournitures(item)("prix")
        tva = 10
        
        montantHT = prix * quantite
        montantTVA = montantHT * (tva / 100)
        montantTTC = montantHT + montantTVA
        
        With wsDevis
            .Cells(ligne, 1).Value = designation
            .Cells(ligne, 1).WrapText = True
            .Cells(ligne, 2).Value = quantite
            .Cells(ligne, 3).Value = Format(prix, "#,##0.00") & " €"
            .Cells(ligne, 4).Value = Format(montantHT, "#,##0.00") & " €"
            .Cells(ligne, 5).Value = tva & " %"
            .Cells(ligne, 6).Value = Format(montantTTC, "#,##0.00") & " €"
            
            ' Mise en forme complète de la ligne
            With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
                .Font.Name = "Arial"
                .Font.Size = 20
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
            
            ' Alignement spécifique pour la désignation
            .Cells(ligne, 1).HorizontalAlignment = xlLeft
            
            totalHT = totalHT + montantHT
            totalTVA = totalTVA + montantTVA
            ligne = ligne + 1
        End With
    Next item
    
    ' ========== Traiter la main d'œuvre ==========
    For Each item In frmDevisDetaille.dictMainOeuvre.Keys
        designation = item
        categorie = "Main d'œuvre"
        
        If InStr(designation, "]") > 0 Then
            designation = Trim(Mid(designation, InStr(designation, "]") + 1))
        End If
        If InStr(designation, " - ") > 0 Then
            designation = left(designation, InStr(designation, " - ") - 1)
        End If
        
        designation = categorie & " - " & designation
        quantite = frmDevisDetaille.dictMainOeuvre(item)("heures")
        prix = frmDevisDetaille.dictMainOeuvre(item)("prix")
        tva = 10
        
        montantHT = prix * quantite
        montantTVA = montantHT * (tva / 100)
        montantTTC = montantHT + montantTVA
        
        With wsDevis
            .Cells(ligne, 1).Value = designation
            .Cells(ligne, 1).WrapText = True
            .Cells(ligne, 2).Value = quantite
            .Cells(ligne, 3).Value = Format(prix, "#,##0.00") & " €/h"
            .Cells(ligne, 4).Value = Format(montantHT, "#,##0.00") & " €"
            .Cells(ligne, 5).Value = tva & " %"
            .Cells(ligne, 6).Value = Format(montantTTC, "#,##0.00") & " €"
            
            ' Mise en forme complète de la ligne
            With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
                .Font.Name = "Arial"
                .Font.Size = 20
                .VerticalAlignment = xlCenter
                .HorizontalAlignment = xlCenter
            End With
            
            ' Alignement spécifique pour la désignation
            .Cells(ligne, 1).HorizontalAlignment = xlLeft
            
            totalHT = totalHT + montantHT
            totalTVA = totalTVA + montantTVA
            ligne = ligne + 1
        End With
    Next item
    
    ' ========== Ajouter le déplacement ==========
    Dim prixDeplacement As Double
    On Error Resume Next
    prixDeplacement = wsTarifGenerique.Cells(4, 5).Value
    On Error GoTo 0
    
    If prixDeplacement = 0 Then prixDeplacement = 50
    
    tva = 10
    montantHT = prixDeplacement
    montantTVA = montantHT * (tva / 100)
    montantTTC = montantHT + montantTVA
    
    With wsDevis
        .Cells(ligne, 1).Value = "Déplacement"
        .Cells(ligne, 2).Value = 1
        .Cells(ligne, 3).Value = Format(prixDeplacement, "#,##0.00") & " €"
        .Cells(ligne, 4).Value = Format(montantHT, "#,##0.00") & " €"
        .Cells(ligne, 5).Value = tva & " %"
        .Cells(ligne, 6).Value = Format(montantTTC, "#,##0.00") & " €"
        
        ' Mise en forme complète de la ligne
        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
            .Font.Name = "Arial"
            .Font.Size = 20
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
        .Cells(ligne, 1).HorizontalAlignment = xlLeft
        
        totalHT = totalHT + montantHT
        totalTVA = totalTVA + montantTVA
        ligne = ligne + 1
    End With
    
    AjouterLignesDetaille = ligne
End Function

'Sub GenererDevisDetaille()
'    Dim ligneDebut As Long
'    Dim ligneActuelle As Long
'
'    ligneDebut = 25
'    ligneActuelle = ligneDebut
'
'    ' ---------- Afficher le formulaire de sélection détaillée ----------
'    frmDevisDetaille.Annule = True
'    frmDevisDetaille.Show
'
'    If frmDevisDetaille.Annule Then
'        MsgBox "Opération annulée par l'utilisateur.", vbInformation
'        Unload frmDevisDetaille
'        Exit Sub
'    End If
'
'    ' ---------- Créer les en-têtes du tableau ----------
'    Call CreerEntetesTableauDetaille(ligneDebut)
'
'    ligneActuelle = ligneDebut + 1
'
'    ' ---------- Ajouter la description ----------
'    With wsDevis
'        .Cells(ligneActuelle, 1).Value = descriptionDesignation
'        .Cells(ligneActuelle, 1).Font.Bold = True
'        .Cells(ligneActuelle, 1).Font.Size = 16
'        .Cells(ligneActuelle, 1).Font.Color = RGB(30, 58, 138)
'        ligneActuelle = ligneActuelle + 1
'    End With
'
'    ' ---------- Variables pour les totaux ----------
'    '    Dim totalFournitures As Double
'    '    Dim totalMainOeuvre As Double
'    '    Dim totalDeplacement As Double
'    Dim totalHT As Double
'    Dim montantTVA As Double
'    Dim totalTTC As Double
'    Dim totalTVA As Double
'
'    '    totalFournitures = 0
'    '    totalMainOeuvre = 0
'    '    totalDeplacement = 0
'    totalHT = 0
'    totalTVA = 0
'
'       ' Ajouter les lignes saisies
'    If frmDevisDetaille.dictFournitures.Count > 0 Then
'        ligneActuelle = AjouterLignesDetaille(ligneActuelle, totalHT, totalTVA)
'    End If
''    ligneActuelle = AjouterLignesDetaille(ligneActuelle, totalHT, totalTVA)
'
'     ' Compléter le tableau jusqu'à 15 lignes minimum
'    Dim ligneFinTableau As Long
'    ligneFinTableau = ligneDebut + 16  ' En-tête + 15 lignes de contenu
'
'    ' Compléter avec des lignes vides si nécessaire
'    Do While ligneActuelle < ligneFinTableau
'        With wsDevis
''            .Range(.Cells(ligneActuelle, 1), .Cells(ligneActuelle, 6)).Borders.LineStyle = xlContinuous
'            .Rows(ligneActuelle).RowHeight = 20
'        End With
'        ligneActuelle = ligneActuelle + 1
'    Loop
'
'     ' Appliquer les bordures au tableau complet
'    Call AppliquerBorduresTableau(ligneDebut, ligneFinTableau)
'
'    ' ========== Calcul du total TTC ==========
'    totalTTC = totalHT + totalTVA
'
'    ' ========== Afficher les totaux ==========
'    Call AfficherTotauxDetaille(ligneActuelle, totalHT, totalTVA, totalTTC)
'
'    Unload frmDevisDetaille
'End Sub
'
''----------------------------------------------------------------------------------------
'' Créer les en-têtes du tableau détaillé
''----------------------------------------------------------------------------------------
'Sub CreerEntetesTableauModification(ligne As Long)
'    With wsDevis
'        ' En-têtes
'        .Cells(ligne, 1).Value = "Désignation"
'        .Cells(ligne, 2).Value = "Qté"
'        .Cells(ligne, 3).Value = "Prix unitaire(€)"
'        .Cells(ligne, 4).Value = "Total HT(€)"
'        .Cells(ligne, 5).Value = "TVA"
'        .Cells(ligne, 6).Value = "Total TTC(€)"
'
'        ' Mise en forme des en-têtes
'        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
'            .Font.Bold = True
'            .Font.Color = RGB(0, 0, 0)
'            .Font.Name = "Arial"
'            .Font.Size = 20
'            .Interior.Color = RGB(237, 242, 247)
'            .HorizontalAlignment = xlCenter
'            .VerticalAlignment = xlCenter
'            .Borders.LineStyle = xlContinuous
'            .Borders.Weight = xlMedium
'        End With
'
'        .Rows(ligne).RowHeight = 26.25
'    End With
'End Sub
'
''========================================================================================
'' Ajouter les lignes détaillées au tableau
''========================================================================================
'Function AjouterLignesDetaille(ligneDebut As Long, ByRef totalHT As Double, ByRef totalTVA As Double) As Long
'    Dim ligne As Long
'    Dim item As Variant
'    Dim designation As String
'    Dim categorie As String
'    Dim prix As Double
'    Dim quantite As Double
'    Dim montantHT As Double
'    Dim tva As Double
'    Dim montantTVA As Double
'    Dim montantTTC As Double
'
'    ligne = ligneDebut
'
'    ' ========== Traiter les fournitures ==========
'    For Each item In frmDevisDetaille.dictFournitures.Keys
'        designation = item
'        categorie = "Fournitures"
'
'        ' Extraire la désignation propre
'        If InStr(designation, "]") > 0 Then
'            designation = Trim(Mid(designation, InStr(designation, "]") + 1))
'        End If
'        If InStr(designation, " - ") > 0 Then
'            designation = left(designation, InStr(designation, " - ") - 1)
'        End If
'
'        designation = categorie & " - " & designation
'        quantite = frmDevisDetaille.dictFournitures(item)("quantite")
'        prix = frmDevisDetaille.dictFournitures(item)("prix")
'        tva = 10
'
'        montantHT = prix * quantite
'        montantTVA = montantHT * (tva / 100)
'        montantTTC = montantHT + montantTVA
'
'        With wsDevis
'            .Cells(ligne, 1).Value = designation
'            .Cells(ligne, 1).WrapText = True
'            .Cells(ligne, 2).Value = quantite
'            .Cells(ligne, 2).HorizontalAlignment = xlCenter
'            .Cells(ligne, 3).Value = Format(prix, "#,##0.00") & " €"
'            .Cells(ligne, 3).HorizontalAlignment = xlCenter
'            .Cells(ligne, 4).Value = Format(montantHT, "#,##0.00") & " €"
'            .Cells(ligne, 4).HorizontalAlignment = xlCenter
'            .Cells(ligne, 5).Value = tva & " %"
'            .Cells(ligne, 5).HorizontalAlignment = xlCenter
'            .Cells(ligne, 6).Value = Format(montantTTC, "#,##0.00") & " €"
'            .Cells(ligne, 6).HorizontalAlignment = xlCenter
'
'            ' Bordures et mise en forme
'            With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
'                .Borders.LineStyle = xlContinuous
'                .Font.Name = "Arial"
'                .Font.Size = 20
'                .VerticalAlignment = xlCenter
'                .HorizontalAlignment = xlCenter
'            End With
'
'            ' Alignement spécifique pour la désignation
'            .Cells(ligne, 1).HorizontalAlignment = xlLeft
'
'
'            totalHT = totalHT + montantHT
'            totalTVA = totalTVA + montantTVA
'            ligne = ligne + 1
'        End With
'    Next item
'
'    ' ========== Traiter la main d'œuvre ==========
'    For Each item In frmDevisDetaille.dictMainOeuvre.Keys
'        designation = item
'        categorie = "Main d'œuvre"
'
'        If InStr(designation, "]") > 0 Then
'            designation = Trim(Mid(designation, InStr(designation, "]") + 1))
'        End If
'        If InStr(designation, " - ") > 0 Then
'            designation = left(designation, InStr(designation, " - ") - 1)
'        End If
'
'        designation = categorie & " - " & designation
'        quantite = frmDevisDetaille.dictMainOeuvre(item)("heures")
'        prix = frmDevisDetaille.dictMainOeuvre(item)("prix")
'        tva = 10
'
'        montantHT = prix * quantite
'        montantTVA = montantHT * (tva / 100)
'        montantTTC = montantHT + montantTVA
'
'        With wsDevis
'            .Cells(ligne, 1).Value = designation
'            .Cells(ligne, 1).WrapText = True
'            .Cells(ligne, 2).Value = quantite
'            .Cells(ligne, 2).HorizontalAlignment = xlCenter
'            .Cells(ligne, 3).Value = Format(prix, "#,##0.00") & " €/h"
'            .Cells(ligne, 3).HorizontalAlignment = xlCenter
'            .Cells(ligne, 4).Value = Format(montantHT, "#,##0.00") & " €"
'            .Cells(ligne, 4).HorizontalAlignment = xlCenter
'            .Cells(ligne, 5).Value = tva & " %"
'            .Cells(ligne, 5).HorizontalAlignment = xlCenter
'            .Cells(ligne, 6).Value = Format(montantTTC, "#,##0.00") & " €"
'            .Cells(ligne, 6).HorizontalAlignment = xlCenter
'
'            .Range(.Cells(ligne, 1), .Cells(ligne, 6)).Borders.LineStyle = xlContinuous
'
'            totalHT = totalHT + montantHT
'            totalTVA = totalTVA + montantTVA
'            ligne = ligne + 1
'        End With
'    Next item
'
'    ' ========== Ajouter le déplacement ==========
'    Dim prixDeplacement As Double
'    On Error Resume Next
'    prixDeplacement = wsTarifGenerique.Cells(4, 5).Value
'    On Error GoTo 0
'
'    If prixDeplacement = 0 Then prixDeplacement = 50
'
'    tva = 10
'    montantHT = prixDeplacement
'    montantTVA = montantHT * (tva / 100)
'    montantTTC = montantHT + montantTVA
'
'    With wsDevis
'        .Cells(ligne, 1).Value = "Déplacement"
'        .Cells(ligne, 2).Value = 1
'        .Cells(ligne, 2).HorizontalAlignment = xlCenter
'        .Cells(ligne, 3).Value = Format(prixDeplacement, "#,##0.00") & " €"
'        .Cells(ligne, 3).HorizontalAlignment = xlCenter
'        .Cells(ligne, 4).Value = Format(montantHT, "#,##0.00") & " €"
'        .Cells(ligne, 4).HorizontalAlignment = xlCenter
'        .Cells(ligne, 5).Value = tva & " %"
'        .Cells(ligne, 5).HorizontalAlignment = xlCenter
'        .Cells(ligne, 6).Value = Format(montantTTC, "#,##0.00") & " €"
'        .Cells(ligne, 6).HorizontalAlignment = xlCenter
'
'        .Range(.Cells(ligne, 1), .Cells(ligne, 6)).Borders.LineStyle = xlContinuous
'
'        totalHT = totalHT + montantHT
'        totalTVA = totalTVA + montantTVA
'        ligne = ligne + 1
'    End With
'
'    AjouterLignesDetaille = ligne
'End Function
'
'Sub AppliquerBorduresTableau(ligneDebut As Long, ligneFin As Long)
'    ' Cette routine applique les bordures uniquement à l'extérieur du tableau
'    ' et entre les colonnes (bordures verticales internes)
'
'    With wsDevis.Range(wsDevis.Cells(ligneDebut, 1), wsDevis.Cells(ligneFin - 1, 6))
'        ' Supprimer toutes les bordures existantes
'        .Borders.LineStyle = xlNone
'
'        ' Bordure extérieure (cadre)
'        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
'
'        ' Bordures verticales internes pour délimiter les colonnes
'        .Borders(xlInsideVertical).LineStyle = xlContinuous
'        .Borders(xlInsideVertical).Weight = xlThin
'
'        ' Bordure horizontale uniquement sous l'en-tête (entre ligne 27 et 28)
'        wsDevis.Range(wsDevis.Cells(ligneDebut, 1), wsDevis.Cells(ligneDebut, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
'        wsDevis.Range(wsDevis.Cells(ligneDebut, 1), wsDevis.Cells(ligneDebut, 6)).Borders(xlEdgeBottom).Weight = xlMedium
'    End With
'End Sub
'
'========================================================================================
' Afficher les totaux HT, TVA et TTC
'========================================================================================
Sub AfficherTotauxDetaille(ligneFinTableau As Long, totalHT As Double, montantTVA As Double, totalTTC As Double)
    Dim ligne As Long
    ligne = ligneFinTableau + 2

    With wsDevis
        ligne = ligne + 1

       ' Total HT, TVA et Total TTC sur 3 lignes
        With .Range(.Cells(ligne, 4), .Cells(ligne, 5))
            .Merge
            .Value = "Total HT"
            .Font.Bold = True
            .Font.Name = "Arial"
            .Font.Size = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With

        .Cells(ligne, 6).Value = Format(totalHT, "#,##0.00") & " €"
        .Cells(ligne, 6).Font.Bold = True
        .Cells(ligne, 6).Font.Name = "Arial"
        .Cells(ligne, 6).Font.Size = 20
        .Cells(ligne, 6).HorizontalAlignment = xlCenter
        .Cells(ligne, 6).Borders.LineStyle = xlContinuous

        ' Conditions de règlement
        With .Range(.Cells(ligne, 1), .Cells(ligne, 3))
            .Merge
            .Value = "Conditions de règlement : A réception de la facture"
            .Font.Italic = True
            .Font.Size = 18 '16
            .Font.Name = "Arial"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        .Rows(ligne).RowHeight = 26.25

        ligne = ligne + 1

        ' TVA
        With .Range(.Cells(ligne, 4), .Cells(ligne, 5))
            .Merge
            .Value = "TVA"
            .Font.Bold = True
            .Font.Name = "Arial"
            .Font.Size = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        .Cells(ligne, 6).Value = Format(montantTVA, "#,##0.00") & " €"
        .Cells(ligne, 6).Font.Bold = True
        .Cells(ligne, 6).Font.Name = "Arial"
        .Cells(ligne, 6).Font.Size = 20
        .Cells(ligne, 6).HorizontalAlignment = xlCenter
        .Cells(ligne, 6).Borders.LineStyle = xlContinuous

        ' Mode de règlement
        With .Range(.Cells(ligne, 1), .Cells(ligne, 3))
            .Merge
            .Value = "Mode de règlement : chèque ou virement"
            .Font.Italic = True
            .Font.Bold = True
            .Font.Size = 18 '16
            .Font.Name = "Arial"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        .Rows(ligne).RowHeight = 26.25

        ligne = ligne + 1

        ' Total TTC
        With .Range(.Cells(ligne, 4), .Cells(ligne, 5))
            .Merge
            .Value = "TOTAL TTC"
            .Font.Bold = True
            .Font.Name = "Arial"
            .Font.Size = 20
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Interior.Color = RGB(237, 242, 247)
            .Borders.LineStyle = xlContinuous
        End With
        .Cells(ligne, 6).Value = Format(totalTTC, "#,##0.00") & " €"
        .Cells(ligne, 6).Font.Bold = True
        .Cells(ligne, 6).Font.Name = "Arial"
        .Cells(ligne, 6).Font.Size = 20
        .Cells(ligne, 6).HorizontalAlignment = xlCenter
        .Cells(ligne, 6).Interior.Color = RGB(237, 242, 247)
        .Cells(ligne, 6).Borders.LineStyle = xlContinuous

        ' Validité du devis
        With .Range(.Cells(ligne, 1), .Cells(ligne, 3))
            .Merge
            .Value = "Ce devis est valable 30 jours à compter de sa date de réalisation"
            .Font.Italic = True
            .Font.Bold = True
            .Font.Size = 18 '16
            .Font.Name = "Arial"
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = xlCenter
        End With
        .Rows(ligne).RowHeight = 26.25

        ' Positionner les éléments en bas de page (ligne 50 environ)
        ligne = 50

        ' "Si ce devis vous convient..."
        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
            .Merge
            .Value = "Si ce devis vous convient, veuillez nous le retourner signé précédé de la mention:"
            .Font.Italic = True
            .Font.Bold = True
            .Font.Size = 22
            .Font.Name = "Times New Roman"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Rows(ligne).RowHeight = 30

        ligne = ligne + 1

        ' "Bon pour accord..."
        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
            .Merge
            .Value = "Bon pour accord et exécution des travaux"
            .Font.Italic = True
            .Font.Bold = True
            .Font.Size = 22
            .Font.Name = "Times New Roman"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Rows(ligne).RowHeight = 30

        ligne = ligne + 2

        ' Date et Signature
        .Cells(ligne, 1).Value = "Date"
        .Cells(ligne, 1).Font.Italic = True
        .Cells(ligne, 1).Font.Size = 20
        .Cells(ligne, 1).Font.Name = "Times New Roman"

        .Cells(ligne, 5).Value = "Signature"
        .Cells(ligne, 5).Font.Italic = True
        .Cells(ligne, 5).Font.Size = 20
        .Cells(ligne, 5).Font.Name = "Times New Roman"

        ligne = ligne + 1
        .Rows(ligne).RowHeight = 80

        ligne = ligne + 2

        ' Siège social
        With .Range(.Cells(ligne, 1), .Cells(ligne + 4, 6))
            .Merge
            .Value = "Siège social : 27 rue Carnot 91300 MASSY" & vbCrLf & _
                     "Tél standard : 01 64 54 27 99" & vbCrLf & _
                     "Siret : 582 017 810 00414    S.N.C au Capital de 3 034 169 euros" & vbCrLf & _
                     "RCS Evry - NAF 7739Z" & vbCrLf & _
                     "N° intracommunautaire : FR 92582017810      www.istablog.fr   www.ista.fr"
            .Font.Size = 16
            .Font.Name = "Arial"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        .Rows(ligne).RowHeight = 87.75
    End With
End Sub

