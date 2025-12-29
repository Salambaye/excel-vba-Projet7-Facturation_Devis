Attribute VB_Name = "Module_principal"
'Salamata Nourou MBAYE - 29/12/2025 - Version 1.0
'PROJET 7 - Facturation
'Module_principal

' ____________Variables globales____________________
Public wbDevis As Workbook
Public wsDevis As Worksheet
Public wbTarification As Workbook

' Variables pour l'en-tête
Public nomClient As String
Public adresseClient As String
Public codePostalVilleClient As String
Public refClient As String
Public refUEBeep As String
Public gestionnaire As String
Public telGestionnaire As String
Public mailGestionnaire As String
Public emplacementTravaux As String
Public adresseChantier As String
Public codePostalChantier As String
Public villeChantier As String
Public presentationProjet As String
Public descriptionDesignation As String

' Variables pour les feuilles de tarification
Public wsTarifGenerique As Worksheet
Public wsTarifPlomberie As Worksheet
Public wsTarifChauffage As Worksheet
Public wsTarifVenteDeVannes As Worksheet
Public wsTarifClient As Worksheet
Public wsTarifPassage As Worksheet

' Variables pour le mode de facturation
Public modeDetaille As Boolean
Public cheminSortie As String
Public dossierSauvegarde As String
Public numeroDevis As String

Sub Facturation_Devis()
    '---------------------- Optimisation --------------------------
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False

    ' ------------------------ Déclaration des variables -------------------------------------
    Dim fdlg As FileDialog
    Dim cheminFichier As String
    Dim fdlgDossier As FileDialog
    
    ' ------------------ Sélection fichier Tarification -------------------------------------
    MsgBox "Sélectionner le fichier 'Tarification des prestations travaux'"
    Set fdlg = Application.FileDialog(msoFileDialogFilePicker)
    fdlg.Title = "Sélection du fichier 'Tarification des prestations travaux'"
    fdlg.Filters.Clear
    fdlg.Filters.Add "Fichiers Excel", "*.xlsx; *.xls; *.xlsm"
    fdlg.AllowMultiSelect = False

    If fdlg.Show <> -1 Then
        MsgBox "Sélection annulée par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    cheminFichier = fdlg.SelectedItems(1)

    ' --------------- Vérification du fichier -------------
    If Dir(cheminFichier) = "" Then
        MsgBox "Le fichier 'Tarification des prestations travaux' n'existe pas : " & cheminFichier, vbCritical
        GoTo Fin
    End If

    ' -------------------------- Sélection du dossier de sauvegarde -----------------------------------
    MsgBox "Choisir le dossier de sauvegarde du devis créé"
    Set fdlgDossier = Application.FileDialog(msoFileDialogFolderPicker)
    With fdlgDossier
        .Title = "Choisir le dossier de sauvegarde du devis"
        .AllowMultiSelect = False
        .InitialFileName = Environ("USERPROFILE") & "\DESKTOP\"
    End With
    
    If fdlgDossier.Show <> -1 Then
        MsgBox "Sélection du dossier annulée par l'utilisateur.", vbInformation
        GoTo Fin
    End If
    
    dossierSauvegarde = fdlgDossier.SelectedItems(1)
    
    If Dir(dossierSauvegarde, vbDirectory) = "" Then
        MsgBox "Le dossier sélectionné n'est pas accessible : " & dossierSauvegarde, vbCritical
        GoTo Fin
    End If
    
    ' -------- Ouvrir le fichier source ---------------
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
    '    Set wsTarifVenteDeVannes = wbTarification.Worksheets("Tarif vente de vannes")
    Set wsTarifVenteDeVannes = wbTarification.Worksheets("Tarif")
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
        If wsTarifVenteDeVannes Is Nothing Then
            MsgBox "La feuille 'Tarif de vente de vannes' n'existe pas dans Tarification", vbCritical
            GoTo Fin
        End If
    If wsTarifClient Is Nothing Then
        MsgBox "La feuille 'Tarif Client compteurs d'eau' n'existe pas dans Tarification", vbCritical
        GoTo Fin
    End If
    If wsTarifPassage Is Nothing Then
        MsgBox "La feuille 'Tarif passage supplémentaire' n'existe pas dans Tarification", vbCritical
        GoTo Fin
    End If
    
    '_________ Renseignements de l'en-tête via UserForm______
    frmEntete.Annule = True
    frmEntete.Show
    
    If frmEntete.Annule = True Then
        MsgBox "Opération annulée par l'utilisateur.", vbInformation
        Unload frmEntete
        GoTo Fin
    End If
    
    ' Récupération des données
    nomClient = Trim(frmEntete.txtNomClient.Value)
    adresseClient = Trim(frmEntete.txtAdresseClient.Value)
    codePostalVilleClient = Trim(frmEntete.txtCpVille.Value)
    refClient = Trim(frmEntete.txtRefclient.Value)
    refUEBeep = Trim(frmEntete.txtRefUEBeep.Value)
    gestionnaire = Trim(frmEntete.txtGestionnaire.Value)
    telGestionnaire = Trim(frmEntete.txtTelGestionnaire.Value)
    mailGestionnaire = Trim(frmEntete.txtMailGestionnaire.Value)
    emplacementTravaux = Trim(frmEntete.txtEmplTravaux.Value)
    adresseChantier = Trim(frmEntete.txtAdresseChantier.Value)
    codePostalChantier = Trim(frmEntete.txtCpChantier.Value)
    villeChantier = Trim(frmEntete.txtVilleChantier.Value)
    presentationProjet = Trim(frmEntete.txtPresentation.Value)
    descriptionDesignation = Trim(frmEntete.txtDesignation.Value)
    
    Unload frmEntete
    
    '_________ Choix du mode de facturation (Détaillé ou Modification)______
    frmDesignation.Annule = True
    frmDesignation.Show
    
    If frmDesignation.Annule = True Then
        MsgBox "Opération annulée par l'utilisateur.", vbInformation
        Unload frmDesignation
        GoTo Fin
    End If
    
    modeDetaille = frmDesignation.optDetaille.Value
    Unload frmDesignation
    
        ' Générer le numéro de devis
    numeroDevis = GenererNumeroDevis()
    
    
    '------------------- Initialisation du devis -----------------------------------
    Call InitialiserDevis
    
    '------------------- Génération du contenu selon le mode -------------------
    If modeDetaille Then
        Call GenererDevisDetaille
    Else
        Call GenererDevisModification
    End If
    
    '------------------- Finalisation du devis -------------------
    Call FinaliserDevis
    
    '------------------------------- Message de fin --------------------------
    MsgBox "Traitement terminé avec succès !" & vbCrLf & "Fichier : " & cheminSortie, vbInformation

    ' Ouvrir le devis
    Dim MonApplication As Object
    Set MonApplication = CreateObject("Shell.Application")
    MonApplication.Open (cheminSortie)

Fin:
    ' Nettoyage
    Set fdlg = Nothing
    Set fdlgDossier = Nothing
    
    If Not wbTarification Is Nothing Then wbTarification.Close False
    
    ' Restaurer les paramètres
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Function GenererNumeroDevis() As String
    Dim compteur As Long
    Dim wsCompteur As Worksheet
    Dim cheminCompteur As String
    Dim wbCompteur As Workbook
    
    ' Chemin du fichier compteur dans le même dossier que le classeur actuel
    cheminCompteur = ThisWorkbook.Path & "\CompteurDevis.txt"
    
    ' Lire le compteur depuis le fichier
    On Error Resume Next
    Open cheminCompteur For Input As #1
    Input #1, compteur
    Close #1
    On Error GoTo 0
    
    ' Si le fichier n'existe pas, initialiser à 1
    If compteur = 0 Then compteur = 1
    
    ' Générer le numéro
    GenererNumeroDevis = refUEBeep & "-" & Format(compteur, "0000")
    
    ' Incrémenter et sauvegarder
    compteur = compteur + 1
'    Open cheminCompteur For Output As #1
'    Print #1, compteur
'    Close #1
End Function

Sub InitialiserDevis()
    Dim wsSource As Worksheet
    Dim img As Shape
    Dim copieImg As Shape
    
    ' Créer le fichier de sortie
    Set wbDevis = Workbooks.Add
    Set wsDevis = wbDevis.Worksheets(1)
    wsDevis.Name = "Devis Travaux"
    wsDevis.Tab.Color = RGB(242, 206, 239)

    
    ' Copier le logo si disponible
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets("Images")
    If Not wsSource Is Nothing Then
        Set img = wsSource.Shapes("LogoIsta")
        If Not img Is Nothing Then
            img.Copy
            wsDevis.Paste
            Set copieImg = wsDevis.Shapes(wsDevis.Shapes.Count)
            With copieImg
                .top = wsDevis.Range("A2").top
                .left = wsDevis.Range("A2").left
                .LockAspectRatio = msoTrue
                .Height = 60
            End With
        End If
    End If
    On Error GoTo 0
    
    ' Configurer la mise en page pour A4 et plusieurs pages si nécessaire
    With wsDevis.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait
        .Zoom = False
           .FitToPagesWide = 1
        .FitToPagesTall = False
        ' Laisser Excel gérer le nombre de pages
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
    End With

    ' Ajustement automatique des colonnes et lignes
'    With wsDevis
'        .Columns("A:D").AutoFit
'        .Rows.AutoFit
'    End With
   
    Call FormaterEntete
End Sub

Sub FormaterEntete()
    With wsDevis
        ' En-tête du devis
'        .Range("C3:D3").Merge
'        .Range("C3:D3").Value = "Devis N° " & refUEBeep
'        .Range("C3:D3").Font.Bold = True
'        .Range("C3:D3").Font.Size = 36
'        .Range("C3:D3").Font.Name = "Aptos Narrow"
        .Cells(3, 3).Value = "Devis N° " & numeroDevis
        .Cells(3, 3).Font.Bold = True
        .Cells(3, 3).Font.Size = 36
        .Cells(3, 3).Font.Name = "Aptos Narrow"
        Rows("3:3").RowHeight = 46.5
        
        ' Informations Ista
        .Cells(6, 1).Value = "Ista Comptage Immobilier Services"
        .Cells(7, 1).Value = "3 rue Christophe Colomb"
        .Cells(8, 1).Value = "91300 MASSY"
        
        .Cells(7, 4).Value = "Date : " & Format(Now, "dd/mm/yyyy")
        .Cells(7, 4).Font.Name = "Arial"
        .Cells(7, 4).Font.Size = 20
        
        .Cells(11, 1).Value = "Dossier géré par : Olivier Contat"
        .Cells(12, 1).Value = "Téléphone : 06.73.47.65.06"
        .Cells(13, 1).Value = "Adresse mail : ocontat@ista.fr"
        
        ' Informations client
        .Cells(10, 4).Value = "Nom du client : " & nomClient
        .Cells(11, 4).Value = "Adresse : " & adresseClient
        .Cells(12, 4).Value = "Code postal et Ville : " & codePostalVilleClient
        
        ' Références
        .Cells(16, 1).Value = "Référence client : " & refClient
        .Cells(17, 1).Value = "N/Référence UEX : " & refUEBeep
        
        ' Gestionnaire
        .Cells(15, 4).Value = "Gestionnaire : " & gestionnaire
        .Cells(16, 4).Value = "Téléphone gestionnaire : " & telGestionnaire
        .Cells(17, 4).Value = "Mail gestionnaire : " & mailGestionnaire
        
        ' Adresse chantier
        .Cells(19, 1).Value = "Adresse chantier : " & adresseChantier
        .Cells(20, 1).Value = "Code postal et ville : " & codePostalChantier & " " & villeChantier
        .Cells(21, 1).Value = "Emplacement travaux : " & emplacementTravaux
        
         Rows("22:22").RowHeight = 33
         Rows("23:23").RowHeight = 51.75
         
        ' Présentation
        .Cells(23, 1).Value = "Présentation du projet : "
        .Cells(23, 1).Font.Bold = True
        .Cells(23, 1).Font.Underline = xlUnderlineStyleSingle
'        .Cells(23, 2).Value = presentationProjet
        .Range("B23:B23").Value = presentationProjet
        .Range("B23:E23").Font.Underline = xlUnderlineStyleSingle

'    .HorizontalAlignment = xlCenter
        .Range("B23:E23").HorizontalAlignment = xlCenterAcrossSelection
                 .Range("B23:E23").VerticalAlignment = xlBottom

        Rows("26:26").RowHeight = 26.25
        
        ' Mise en forme
        .Range("A6:A8").Font.Name = "Aptos Narrow"
        .Range("A6:A8").Font.Size = 16
        .Range("A10:F26").Font.Name = "Arial"
        .Range("A10:F26").Font.Size = 20
        .Range("A11:A21").Font.Italic = True
        
'         .Range("A6:A23").Font.Name = "Calibri"
'        .Range("A6:A23").Font.Size = 11
'
        ' Largeur des colonnes
        .Columns("A:A").ColumnWidth = 74.5
        .Columns("B:B").ColumnWidth = 9.25
        .Columns("C:C").ColumnWidth = 25.63
        .Columns("D:D").ColumnWidth = 17.13
        .Columns("E:E").ColumnWidth = 24.75
    End With
End Sub

