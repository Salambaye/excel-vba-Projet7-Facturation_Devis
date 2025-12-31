Attribute VB_Name = "Generer_Devis_Modification"
'Module pour générer le devis en mode modification

Sub GenererDevisModification()
    Dim ligneDebut As Long
    Dim ligneActuelle As Long
    
    ligneDebut = 25
    ligneActuelle = ligneDebut
    
    ' Afficher le formulaire de modification
    frmDevisModification.Annule = True
    frmDevisModification.Show
    
    If frmDevisModification.Annule = True Then
        MsgBox "Opération annulée par l'utilisateur.", vbInformation
        Unload frmDevisModification
        Exit Sub
    End If
    
    ' Créer les en-têtes du tableau
    Call CreerEntetesTableauModification(ligneDebut)
    
    ligneActuelle = ligneDebut + 1
    
    ' Variables pour les totaux
    Dim totalHT As Double
    Dim montantTVA As Double
    Dim totalTTC As Double
    
    totalHT = 0
    
    ' Ajouter les lignes saisies
    If frmDevisModification.dictLignes.Count > 0 Then
        ligneActuelle = AjouterLignesModification(ligneActuelle, totalHT, montantTVA)
    End If
    
    ' Compléter le tableau jusqu'à 15 lignes minimum
    Dim ligneFinTableau As Long
    ligneFinTableau = ligneDebut + 16            ' En-tête + 15 lignes de contenu
    
    ' Compléter avec des lignes vides si nécessaire
    Do While ligneActuelle < ligneFinTableau
        With wsDevis
            '            .Range(.Cells(ligneActuelle, 1), .Cells(ligneActuelle, 6)).Borders.LineStyle = xlContinuous
            .Rows(ligneActuelle).RowHeight = 20
        End With
        ligneActuelle = ligneActuelle + 1
    Loop
    
    ' Appliquer les bordures au tableau complet
    Call AppliquerBorduresTableau(ligneDebut, ligneFinTableau)
    
    ' Calcul du total TTC
    totalTTC = totalHT + montantTVA
    
    ' Afficher les totaux
    Call AfficherTotauxModification(ligneFinTableau, totalHT, montantTVA, totalTTC)
    
    Unload frmDevisModification
End Sub

Sub CreerEntetesTableauModification(ligne As Long)
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
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlMedium
        End With
        
        .Rows(ligne).RowHeight = 26.25
    End With
End Sub

Function AjouterLignesModification(ligneDebut As Long, ByRef totalHT As Double, ByRef totalTVA As Double) As Long
    Dim ligne As Long
    Dim item As Variant
    Dim designation As String
    Dim quantite As Double
    Dim prixUnitaire As Double
    Dim tva As Double
    Dim montantHT As Double
    Dim montantTVA As Double
    Dim montantTTC As Double
    
    ligne = ligneDebut
    
    ' Parcourir le dictionnaire des lignes
    For Each item In frmDevisModification.dictLignes.Keys
        With wsDevis
            designation = frmDevisModification.dictLignes(item)("designation")
            quantite = frmDevisModification.dictLignes(item)("quantite")
            prixUnitaire = frmDevisModification.dictLignes(item)("prix")
            tva = frmDevisModification.dictLignes(item)("tva")
            
            montantHT = quantite * prixUnitaire
            montantTVA = montantHT * (tva / 100)
            montantTTC = montantHT + montantTVA
            
            ' Remplir la ligne
            .Cells(ligne, 1).Value = designation
                        .Cells(ligne, 1).WrapText = True
            .Cells(ligne, 2).Value = quantite
            .Cells(ligne, 3).Value = Format(prixUnitaire, "#,##0.00") & " €"
            .Cells(ligne, 4).Value = Format(montantHT, "#,##0.00") & " €"
            .Cells(ligne, 5).Value = tva & " %"
            .Cells(ligne, 6).Value = Format(montantTTC, "#,##0.00") & " €"
            
            ' Bordures et mise en forme
            With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
                .Borders.LineStyle = xlContinuous
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
    
    AjouterLignesModification = ligne
End Function

Sub AppliquerBorduresTableau(ligneDebut As Long, ligneFin As Long)
    ' Cette routine applique les bordures uniquement à l'extérieur du tableau
    ' et entre les colonnes (bordures verticales internes)
    
    With wsDevis.Range(wsDevis.Cells(ligneDebut, 1), wsDevis.Cells(ligneFin - 1, 6))
        ' Supprimer toutes les bordures existantes
        .Borders.LineStyle = xlNone
        
        ' Bordure extérieure (cadre)
        .BorderAround LineStyle:=xlContinuous, Weight:=xlMedium
        
        ' Bordures verticales internes pour délimiter les colonnes
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).Weight = xlThin
        
        ' Bordure horizontale uniquement sous l'en-tête (entre ligne 27 et 28)
        wsDevis.Range(wsDevis.Cells(ligneDebut, 1), wsDevis.Cells(ligneDebut, 6)).Borders(xlEdgeBottom).LineStyle = xlContinuous
        wsDevis.Range(wsDevis.Cells(ligneDebut, 1), wsDevis.Cells(ligneDebut, 6)).Borders(xlEdgeBottom).Weight = xlMedium
    End With
End Sub

Sub AfficherTotauxModification(ligneFinTableau As Long, totalHT As Double, montantTVA As Double, totalTTC As Double)
    Dim ligne As Long
    ligne = ligneFinTableau + 2
    
    With wsDevis
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

Sub FinaliserDevis()
    ' Zoom optimal
    wsDevis.Range("A1").Select
    ActiveWindow.Zoom = 55
    
    ' Sauvegarde du fichier
    Dim nomFichier As String
    nomFichier = "Devis_" & refUEBeep & "-" & Format(Now, "yyyymmdd") & "_" & nomClient & ".xlsx"
    cheminSortie = dossierSauvegarde & "\" & nomFichier
    
    On Error Resume Next
    wbDevis.SaveAs Filename:=cheminSortie, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de la sauvegarde : " & Err.Description, vbCritical
        Err.Clear
    End If
    On Error GoTo 0
End Sub


