Attribute VB_Name = "Module_principal"
'TEST

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
                .top = wsDevis.Range("B2").top
                .left = wsDevis.Range("B2").left
                .LockAspectRatio = msoTrue
                .Height = 50
            End With
        End If
    End If
    On Error GoTo 0
   
    Call FormaterEntete
End Sub

Sub FormaterEntete()
    With wsDevis
        ' En-tête du devis
        .Cells(3, 3).Value = "Devis N° " & refUEBeep
        .Cells(3, 3).Font.Bold = True
        .Cells(3, 3).Font.Size = 14
        
        ' Informations Ista
        .Cells(6, 1).Value = "Ista Comptage Immobilier Services"
        .Cells(7, 1).Value = "3 rue Christophe Colomb"
        .Cells(8, 1).Value = "91300 MASSY"
        
        .Cells(7, 4).Value = "Date : " & Format(Now, "dd/mm/yyyy")
        
        .Cells(11, 1).Value = "Dossier généré par : Olivier Contat"
        .Cells(12, 1).Value = "Téléphone : 06.73.47.65.06"
        .Cells(13, 1).Value = "Adresse mail : ocontat@ista.fr"
        
        ' Informations client
        .Cells(10, 4).Value = "Nom du client : " & nomClient
        .Cells(11, 4).Value = "Adresse : " & adresseClient
        .Cells(12, 4).Value = "Code postal et Ville : " & codePostalVilleClient
        
        ' Références
        .Cells(16, 1).Value = "Référence client : " & refClient
        .Cells(17, 1).Value = "N/Référence UEX + BEEP : " & refUEBeep
        
        ' Gestionnaire
        .Cells(15, 4).Value = "Gestionnaire : " & gestionnaire
        .Cells(16, 4).Value = "Téléphone gestionnaire : " & telGestionnaire
        .Cells(17, 4).Value = "Mail gestionnaire : " & mailGestionnaire
        
        ' Adresse chantier
        .Cells(19, 1).Value = "Adresse chantier : " & adresseChantier
        .Cells(20, 1).Value = "Code postal et ville : " & codePostalChantier & " " & villeChantier
        .Cells(21, 1).Value = "Emplacement travaux : " & emplacementTravaux
        
        ' Présentation
        .Cells(23, 1).Value = "Présentation du projet : "
        .Cells(23, 2).Value = presentationProjet
        
        ' Mise en forme
        .Range("A6:A23").Font.Name = "Calibri"
        .Range("A6:A23").Font.Size = 11
        .Range("D10:D17").Font.Name = "Calibri"
        .Range("D10:D17").Font.Size = 11
        
        ' Largeur des colonnes
        .Columns("A:A").ColumnWidth = 50
        .Columns("D:D").ColumnWidth = 40
    End With
End Sub
