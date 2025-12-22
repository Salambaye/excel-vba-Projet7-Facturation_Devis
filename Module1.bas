Attribute VB_Name = "Module1"
'Salamata Nourou MBAYE - 19/12/2025 - Version 1.0
'PROJET 7 - Facturation


' ____________Variables globales pour le fichier de sortie____________________

Dim wbDevis As Workbook
Dim wsDevis As Worksheet

Sub Facturation()

    '---------------------- Optimisation pour accélérer la macro --------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' ------------------------ Déclaration des variables -------------------------------------
    Dim fdlg As FileDialog
    Dim cheminFichier As String
    Dim cheminSortie As String
    Dim i, j As Long
    Dim ligneOut As Long
    
    Dim dossierSauvegarde As String
    Dim fdlgDossier As FileDialog
    
    Dim wbTarification As Workbook
    Dim wsTarifGenerique As Worksheet
    Dim wsTarifPlomberie As Worksheet
    Dim wsTarifChauffage As Worksheet
    Dim wsTarifVenteDeVannes As Worksheet
    Dim wsTarifClient As Worksheet
    Dim wsTarifPassage As Worksheet
    


    ' ------------------ Sélection fichier Tarification ( input ) -------------------------------------
    MsgBox "Sélectionner le fichier 'Tarification des prestations travaux'"
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "Sélection du fichier 'Tarification des prestations travaux'"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers Excel", "*.xlsx; *.xls; *.xlsm"
    fdlg.AllowMultiSelect = False

    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée par l'utilisateur.", vbInformation
        Exit Sub
    End If
    cheminFichier = fdlg.SelectedItems(1)

    ' --------------- Vérification du fichier -------------
    If Dir(cheminFichier) = "" Then
        MsgBox "Le fichier 'Tarification des prestations travaux' n'existe pas : " & cheminFichier, vbCritical
        GoTo Fin
        Exit Sub
    End If

    ' -------------------------- Sélection du dossier de sauvegarde du devis -----------------------------------
    MsgBox "Choisir le dossier de sauvegarde du devis créé"
    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlgDossier
        .Title = "Choisir le dossier de sauvegarde du devis"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\DESKTOP\"
    End With
    
    If fdlgDossier.Show <> -1 Then
        MsgBox "Sélection du dossier annulée par l'utilisateur.", vbInformation
        Exit Sub
    End If
    
    dossierSauvegarde = fdlgDossier.SelectedItems(1)
    
    ' Vérifier que le dossier existe et est accessible
    If Dir(dossierSauvegarde, vbDirectory) = "" Then
        MsgBox "Le dossier sélectionné n'est pas accessible : " & dossierSauvegarde, vbCritical
        Exit Sub
    End If
    
    ' -------- Ouvrir le fichier source (UpdateLinks:=0 désactive la boîte de dialogue de mise à jour)---------------
    On Error Resume Next
    Set wbTarification = Workbooks.Open(Filename:=cheminFichier, ReadOnly:=True, UpdateLinks:=0, IgnoreReadOnlyRecommended:=True)
    If Err.Number <> 0 Then
        MsgBox "Erreur lors de l'ouverture de Tarification : " & Err.Description, vbCritical
        GoTo Fin
    End If
    Err.Clear
    
    ' Références aux feuilles
    On Error Resume Next
    
    Set wsTarifGenerique = wbTarification.Worksheets("Tarif générique 2025 ")
    Set wsTarifPlomberie = wbTarification.Worksheets("Tarif travaux Plomberie")
    Set wsTarifChauffage = wbTarification.Worksheets("Tarif travaux Chauffage")
'    Set wsTarifVenteDeVannes = wbTarification.Worksheets("Tarif de vente de vannes")
    Set wsTarifClient = wbTarification.Worksheets("Tarif Client compteurs d'eau")
    Set wsTarifPassage = wbTarification.Worksheets("Tarif passage supplémentaire")
    
    On Error GoTo 0

    ' Vérification que toutes les feuilles existent
    If wsTarifGenerique Is Nothing Then
        MsgBox "La feuille 'Tarif générique 2025' n'existe pas dans Tarification", vbCritical
        GoTo Fin
    End If
    If wsTarifPlomberie Is Nothing Then
        MsgBox "La feuille 'Tarif travaux Plomberie' n'existe pas dans Tarification", vbCritical
        GoTo Fin
    End If
    If wsTarifChauffage Is Nothing Then
        MsgBox "La feuille 'Tarif travaux Chauffage' n'existe pas dans Tarification", vbCritical
        GoTo Fin
    End If
'    If wsTarifVenteDeVannes Is Nothing Then
'        MsgBox "La feuille 'Tarif de vente de vannes' n'existe pas dans Tarification", vbCritical
'        GoTo Fin
'    End If
    If wsTarifClient Is Nothing Then
        MsgBox "La feuille 'Tarif Client compteurs d'eau' n'existe pas dans Tarification", vbCritical
        GoTo Fin
    End If
    If wsTarifPassage Is Nothing Then
        MsgBox "La feuille 'Tarif passage supplémentaire' n'existe pas dans Tarification", vbCritical
        GoTo Fin
    End If

    Call InitialiserDevis


    '------------------------------- Message de fin de traitement --------------------------
    MsgBox "Traitement terminé", vbInformation

    ' Ouvrir le dossier contenant le fichier créé
    'Shell "explorer.exe /select,""" & cheminSortie & """", vbNormalFocus
    
    'Ouvrir directement le devis
    Dim MonApplication As Object
    Set MonApplication = CreateObject("Shell.Application")
    MonApplication.Open (cheminSortie)


Fin:

    ' ------------------------ Nettoyer la référence au dialog ------------------------------------
    Set fdlg = Nothing
    
    ' ----------------------------------- Restautrer les paramètres --------------------------------
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

End Sub

Sub InitialiserDevis()
    ' Créer le fichier de sortie
    Set wbDevis = Workbooks.Add
    
    ' Créer la feuille "launcher quotidien"
    Set wsDevis = wbDevis.Worksheets(1)
    wsDevis.Name = "Devis Travaux"
    wsDevis.Tab.Color = RGB(242, 206, 239)
    
    Call FormaterDevis
End Sub

Sub FormaterDevis()

    ' En-têtes
    With wsDevis
        .Cells(3, 3).Value = "Devis N° "
     
        .Cells(6, 1).Value = "Ista Comptage Immobilier Services"
        .Cells(7, 1).Value = "3 rue Christophe Colomb"
        .Cells(8, 1).Value = "91300 MASSY"
        
        .Cells(7, 4).Value = "Date : " & Format(Now, "dd/mm/yyyy")
        
        .Cells(11, 1).Value = "Dossier généré par : Olivier Contat"
        .Cells(12, 1).Value = "Téléphone : 06.73.47.65.06"
        .Cells(13, 1).Value = "Adresse mail : ocontat@ista.fr"
        
        .Cells(10, 4).Value = "Nom du client : "
        .Cells(11, 4).Value = "Adresse : "
        .Cells(12, 4).Value = "Code postal : "
        
        .Cells(16, 1).Value = "Référence client : "
        .Cells(17, 1).Value = "N/Référence UEX + BEEP : "
        
        .Cells(15, 4).Value = "Gestionnaire : "
        .Cells(16, 4).Value = "telephone gestionnaire : "
        .Cells(17, 4).Value = "mail gestionnaire"
        
        .Cells(19, 1).Value = "Adresse chantier : "
        .Cells(20, 1).Value = "Code postal : "
        .Cells(21, 1).Value = "Emplacement travaux : "
        
        .Cells(23, 1).Value = "Présentation du projet : "
    End With
    
    With wsDevis.Range("A1:A1")
        .Font.Name = "Calibri"
        .Font.Bold = True
        .Font.Size = 11
        .ColumnWidth = 75
        '        .Borders.LineStyle = xlContinuous
        '        .HorizontalAlignment = xlCenter
        '        .VerticalAlignment = xlCenter
    End With
End Sub

