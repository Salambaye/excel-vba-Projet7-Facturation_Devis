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
        .Cells(ligne, 4).Value = "Total HT"
        .Cells(ligne, 5).Value = "TVA %"
        .Cells(ligne, 6).Value = "Total TTC"
        
        ' Mise en forme des en-têtes
        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
            .Font.Bold = True
            .Interior.Color = RGB(79, 129, 189)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .Borders.LineStyle = xlContinuous
        End With
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
            .Cells(ligne, 2).HorizontalAlignment = xlCenter
            .Cells(ligne, 3).Value = Format(prixUnitaire, "#,##0.00") & " €"
            .Cells(ligne, 3).HorizontalAlignment = xlRight
            .Cells(ligne, 4).Value = Format(montantHT, "#,##0.00") & " €"
            .Cells(ligne, 4).HorizontalAlignment = xlRight
            .Cells(ligne, 5).Value = tva & " %"
            .Cells(ligne, 5).HorizontalAlignment = xlCenter
            .Cells(ligne, 6).Value = Format(montantTTC, "#,##0.00") & " €"
            .Cells(ligne, 6).HorizontalAlignment = xlRight
            
            ' Bordures
            .Range(.Cells(ligne, 1), .Cells(ligne, 6)).Borders.LineStyle = xlContinuous
            
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
        .Cells(ligne, 5).Value = "Total HT :"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        .Cells(ligne, 6).Value = Format(totalHT, "#,##0.00") & " €"
        .Cells(ligne, 6).Font.Bold = True
        .Cells(ligne, 6).HorizontalAlignment = xlRight
        
        ligne = ligne + 1
        
        ' TVA
        .Cells(ligne, 5).Value = "TVA :"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        .Cells(ligne, 6).Value = Format(montantTVA, "#,##0.00") & " €"
        .Cells(ligne, 6).Font.Bold = True
        .Cells(ligne, 6).HorizontalAlignment = xlRight
        
        ligne = ligne + 1
        
        ' Total TTC
        .Cells(ligne, 5).Value = "TOTAL TTC :"
        .Cells(ligne, 5).Font.Bold = True
        .Cells(ligne, 5).Font.Size = 12
        .Cells(ligne, 5).HorizontalAlignment = xlRight
        .Cells(ligne, 6).Value = Format(totalTTC, "#,##0.00") & " €"
        .Cells(ligne, 6).Font.Bold = True
        .Cells(ligne, 6).Font.Size = 12
        .Cells(ligne, 6).HorizontalAlignment = xlRight
        
        ' Bordure pour le total TTC
        With .Range(.Cells(ligne, 5), .Cells(ligne, 6))
            .Interior.Color = RGB(217, 217, 217)
        End With
        
       ' ---------- Texte de fin ----------
        ligne = ligne + 3
        .Cells(ligne, 1).Value = "Conditions de règlement : A réception de la facture"
        .Cells(ligne, 1).Font.Italic = True
        .Cells(ligne, 1).Font.Size = 16
        .Cells(ligne, 1).Font.Name = "Arial"
        Rows(ligne).RowHeight = 26.25

        ligne = ligne + 1
        .Cells(ligne, 1).Value = "Mode de règlement : chèque ou virement"
        .Cells(ligne, 1).Font.Italic = True
        .Cells(ligne, 1).Font.Bold = True
        .Cells(ligne, 1).Font.Size = 16
        .Cells(ligne, 1).Font.Name = "Arial"
        Rows(ligne).RowHeight = 26.25

        ligne = ligne + 1
        .Cells(ligne, 1).Value = "Ce devis est valable 30 jours à compter de sa date de réalisation"
        .Cells(ligne, 1).Font.Italic = True
        .Cells(ligne, 1).Font.Bold = True
        .Cells(ligne, 1).Font.Size = 16
        .Cells(ligne, 1).Font.Name = "Arial"
        Rows(ligne).RowHeight = 26.25

        ligne = ligne + 3
        Rows(ligne).RowHeight = 54.75

        ligne = ligne + 1
        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
            .Merge
            .Value = "Si ce devis vous convient, veuillez nous le retourner signé précédé de la mention:"
            .Font.Italic = True
            .Font.Bold = True
            .Font.Size = 24
            .Font.Name = "Times New Roman"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ligne = ligne + 1
        With .Range(.Cells(ligne, 1), .Cells(ligne, 6))
            .Merge
            .Value = " Bon pour accord et exécution des travaux"
            .Font.Italic = True
            .Font.Bold = True
            .Font.Size = 24
            .Font.Name = "Times New Roman"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ligne = ligne + 2
        .Cells(ligne, 1).Value = "Date"
        .Cells(ligne, 1).Font.Italic = True
        .Cells(ligne, 1).Font.Size = 20
        .Cells(ligne, 1).Font.Name = "Times New Roman"

        .Cells(ligne, 5).Value = "Signature"
        .Cells(ligne, 5).Font.Italic = True
        .Cells(ligne, 5).Font.Size = 20
        .Cells(ligne, 5).Font.Name = "Times New Roman"

        ligne = ligne + 1
        Rows(ligne).RowHeight = 123

        ligne = ligne + 2
        With .Range(.Cells(ligne, 1), .Cells(ligne + 4, 6))
            .Merge
            .Value = "Siège social : 27 rue Carnot 91300 MASSY" & vbCrLf & "Tél standard : 01 64 54 27 99" & vbCrLf & _
                    "Siret : 582 017 810 00414    S.N.C au Capital de 3 034 169 euros" & vbCrLf & _
                    "RCS Evry - NAF 7739Z" & vbCrLf & "N° intracommunautaire : FR 92582017810      www.istablog.fr   www.ista.fr"
            .Font.Size = 16
            .Font.Name = "Arial"
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With

        ligne = ligne + 4
        Rows(ligne).RowHeight = 87.75
    End With
End Sub

Sub FinaliserDevis()
    ' Ajustement automatique des lignes pour le retour à la ligne
    wsDevis.Rows.AutoFit
    
    ' Zoom optimal
    wsDevis.Range("A1").Select
    ActiveWindow.Zoom = 55
    
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
