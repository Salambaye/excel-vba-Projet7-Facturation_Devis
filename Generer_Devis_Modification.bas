Attribute VB_Name = "Generer_Devis_Modification"
'Module pour générer le devis en mode modification

Sub GenererDevisModification()
    Dim ligneDebut As Long
    Dim ligneActuelle As Long
    
    ligneDebut = 26
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
    
    ligneActuelle = ligneDebut + 2
    
    ' Ajouter la description
    With wsDevis
        .Cells(ligneActuelle, 1).Value = descriptionDesignation
        .Cells(ligneActuelle, 1).Font.Bold = True
        .Cells(ligneActuelle, 1).Font.Size = 11
        ligneActuelle = ligneActuelle + 1
    End With
    
    ' Variables pour les totaux
    Dim totalHT As Double
    Dim montantTVA As Double
    Dim totalTTC As Double
    
    totalHT = 0
    
    ' Ajouter les lignes saisies
    If frmDevisModification.dictLignes.Count > 0 Then
        ligneActuelle = AjouterLignesModification(ligneActuelle, totalHT, montantTVA)
    End If
    
    ' Ligne de séparation
    ligneActuelle = ligneActuelle + 1
    
    ' Calcul du total TTC
    totalTTC = totalHT + montantTVA
    
    ' Afficher les totaux
    Call AfficherTotauxModification(ligneActuelle, totalHT, montantTVA, totalTTC)
    
    Unload frmDevisModification
End Sub

Sub CreerEntetesTableauModification(ligne As Long)
    With wsDevis
        ' En-têtes
        .Cells(ligne, 1).Value = "Désignation"
        .Cells(ligne, 2).Value = "Quantité"
        .Cells(ligne, 3).Value = "Prix unitaire HT"
        .Cells(ligne, 4).Value = "TVA %"
        .Cells(ligne, 5).Value = "Total HT"
        
        ' Mise en forme des en-têtes
        With .Range(.Cells(ligne, 1), .Cells(ligne, 5))
            .Font.Bold = True
            .Font.Size = 11
            .Interior.Color = RGB(79, 129, 189)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
        
        ' Largeur des colonnes
        .Columns("A:A").ColumnWidth = 50
        .Columns("B:B").ColumnWidth = 12
        .Columns("C:C").ColumnWidth = 18
        .Columns("D:D").ColumnWidth = 10
        .Columns("E:E").ColumnWidth = 15
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
            
            ' Remplir la ligne
            .Cells(ligne, 1).Value = designation
            .Cells(ligne, 2).Value = quantite
            .Cells(ligne, 2).HorizontalAlignment = xlCenter
            .Cells(ligne, 3).Value = Format(prixUnitaire, "#,##0.00") & " €"
            .Cells(ligne, 3).HorizontalAlignment = xlRight
            .Cells(ligne, 4).Value = tva & " %"
            .Cells(ligne, 4).HorizontalAlignment = xlCenter
            .Cells(ligne, 5).Value = Format(montantHT, "#,##0.00") & " €"
            .Cells(ligne, 5).HorizontalAlignment = xlRight
            
            ' Bordures
            .Range(.Cells(ligne, 1), .Cells(ligne, 5)).Borders.LineStyle = xlContinuous
            
            totalHT = totalHT + montantHT
            totalTVA = totalTVA + montantTVA
            
            ligne = ligne + 1
        End With
    Next item
    
    AjouterLignesModification = ligne
End Function

Sub AfficherTotauxModification(ligne As Long, totalHT As Double, montantTVA As Double, totalTTC As Double)
    With wsDevis
        ' Ligne vide
        ligne = ligne + 1
        
        ' Total HT
        .Cells(ligne, 4).Value = "Total HT :"
        .Cells(ligne, 4).Font.Bold = True
        .Cells(ligne, 4).HorizontalAlignment = xlRight
        .Cells(ligne, 5).Value = Format(totalHT, "#,##0.00") & " €"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        
        ligne = ligne + 1
        
        ' TVA
        .Cells(ligne, 4).Value = "TVA :"
        .Cells(ligne, 4).Font.Bold = True
        .Cells(ligne, 4).HorizontalAlignment = xlRight
        .Cells(ligne, 5).Value = Format(montantTVA, "#,##0.00") & " €"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        
        ligne = ligne + 1
        
        ' Total TTC
        .Cells(ligne, 4).Value = "TOTAL TTC :"
        .Cells(ligne, 4).Font.Bold = True
        .Cells(ligne, 4).Font.Size = 12
        .Cells(ligne, 4).HorizontalAlignment = xlRight
        .Cells(ligne, 5).Value = Format(totalTTC, "#,##0.00") & " €"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).Font.Size = 12
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        
        ' Bordure pour le total TTC
        With .Range(.Cells(ligne, 4), .Cells(ligne, 5))
'            .Borders(xlEdgeTop).LineStyle = xlContinuous
'            .Borders(xlEdgeTop).Weight = xlThick
'            .Borders(xlEdgeBottom).LineStyle = xlDouble
            .Interior.Color = RGB(217, 217, 217)
        End With
        
        ' Texte de fin
'        ligne = ligne + 3
'        .Cells(ligne, 1).Value = "Conditions de règlement : A réception de la facture"
''        .Cells(ligne, 1).Font.Italic = True
'        .Cells(ligne, 1).Font.Size = 10
''        .Cells(ligne, 1).Font.Color = RGB(100, 100, 100)
'
'        ligne = ligne + 1
'        .Cells(ligne, 1).Value = "Mode de règlement : chèque ou virement."
''        .Cells(ligne, 1).Font.Italic = True
'        .Cells(ligne, 1).Font.Size = 10
''        .Cells(ligne, 1).Font.Color = RGB(100, 100, 100)
'
'         ligne = ligne + 1
'        .Cells(ligne, 1).Value = "Ce devis est valable 30 jours à compter de sa date de réalisation."
''        .Cells(ligne, 1).Font.Italic = True
'        .Cells(ligne, 1).Font.Size = 10
''        .Cells(ligne, 1).Font.Color = RGB(100, 100, 100)

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
        .Font.Size = 20
        .Font.Name = "Times New Roman"
        End With
    End With
 
End Sub

Sub FinaliserDevis()
    ' Ajustement automatique des lignes
    wsDevis.Rows.AutoFit
    
    ' Zoom optimal
    wsDevis.Range("A1").Select
    ActiveWindow.Zoom = 85
    
    ' Sauvegarde du fichier
    Dim nomFichier As String
    nomFichier = "Devis_" & Replace(nomClient, " ", "_") & "_" & Format(Now, "yyyymmdd_hhmmss") & ".xlsx"
    cheminSortie = dossierSauvegarde & "\" & nomFichier
    
    On Error Resume Next
    wbDevis.SaveAs Filename:=cheminSortie, FileFormat:=xlOpenXMLWorkbook
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de la sauvegarde : " & Err.Description, vbCritical
        Err.Clear
    End If
    On Error GoTo 0
End Sub


'Sub InsererImageSansChemin()
'    Dim wsSource As Worksheet
'    Dim wsDest As Worksheet
'    Dim img As Shape
'    Dim copieImg As Shape
'
'    ' Feuille où l'image est stockée (masquée)
'    Set wsSource = ThisWorkbook.Sheets("Images")
'
'    ' Feuille où on veut insérer l'image
'    Set wsDest = ThisWorkbook.Sheets("Macro")
'
'    ' Vérifier si l'image existe
'    On Error Resume Next
'    Set img = wsSource.Shapes("LogoStocke")
'    On Error GoTo 0
'
'    If img Is Nothing Then
'        MsgBox "L'image 'LogoStocke' n'existe pas dans la feuille 'Images'.", vbExclamation
'        Exit Sub
'    End If
'
'    ' Copier l'image depuis la feuille masquée
'    img.Copy
'
'    ' Coller dans la feuille de destination
'    wsDest.Paste
'    Set copieImg = wsDest.Shapes(wsDest.Shapes.Count)
'
'    ' Positionner et redimensionner
'    With copieImg
'        .Top = wsDest.Range("B2").Top
'        .Left = wsDest.Range("B2").Left
'        .LockAspectRatio = msoTrue
'        .Height = 50
'    End With
'
'    MsgBox "Image insérée avec succès sans chemin externe.", vbInformation
'End Sub
'
