VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5280
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6630
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   27
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "btnOpen"
            Object.ToolTipText     =   "Ouvrir une session"
            Object.Tag             =   ""
            ImageIndex      =   16
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "btnSaves"
            Object.ToolTipText     =   "Assistant de sauvegarde"
            Object.Tag             =   ""
            ImageIndex      =   37
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnPrint"
            Object.ToolTipText     =   "Imprimer"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnFrs"
            Object.ToolTipText     =   "Gestion des fournisseurs"
            Object.Tag             =   ""
            ImageIndex      =   58
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnAP"
            Object.ToolTipText     =   "Affaires en projet"
            Object.Tag             =   ""
            ImageIndex      =   31
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnAR"
            Object.ToolTipText     =   "Affaires en réalisation"
            Object.Tag             =   ""
            ImageIndex      =   59
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnDiscip"
            Object.ToolTipText     =   "Liste des disciplines"
            Object.Tag             =   ""
            ImageIndex      =   57
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnEmp"
            Object.ToolTipText     =   "Liste des employés"
            Object.Tag             =   ""
            ImageIndex      =   53
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnPdt"
            Object.ToolTipText     =   "Liste des produits"
            Object.Tag             =   ""
            ImageIndex      =   60
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnSup"
            Object.ToolTipText     =   "Supports de classement"
            Object.Tag             =   ""
            ImageIndex      =   61
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnDoc"
            Object.ToolTipText     =   "Saisie d'un document"
            Object.Tag             =   ""
            ImageIndex      =   62
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnMvt"
            Object.ToolTipText     =   "Mouvement de documents"
            Object.Tag             =   ""
            ImageIndex      =   63
         EndProperty
         BeginProperty Button16 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button17 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnBegin"
            Object.ToolTipText     =   "Premier"
            Object.Tag             =   ""
            ImageIndex      =   40
         EndProperty
         BeginProperty Button18 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnPrevious"
            Object.ToolTipText     =   "Précédent"
            Object.Tag             =   ""
            ImageIndex      =   39
         EndProperty
         BeginProperty Button19 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnNext"
            Object.ToolTipText     =   "Suivant"
            Object.Tag             =   ""
            ImageIndex      =   38
         EndProperty
         BeginProperty Button20 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnEnd"
            Object.ToolTipText     =   "Dernier"
            Object.Tag             =   ""
            ImageIndex      =   41
         EndProperty
         BeginProperty Button21 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button22 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnFind"
            Object.ToolTipText     =   "Chercher"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button23 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnCalc"
            Object.ToolTipText     =   "Calculatrice"
            Object.Tag             =   ""
            ImageIndex      =   46
         EndProperty
         BeginProperty Button24 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "btnMainDB"
            Object.ToolTipText     =   "Utilitaires"
            Object.Tag             =   ""
            ImageIndex      =   24
         EndProperty
         BeginProperty Button25 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnHelp"
            Object.ToolTipText     =   "Sommaire de l'aide"
            Object.Tag             =   ""
            ImageIndex      =   25
         EndProperty
         BeginProperty Button26 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button27 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "btnFermer"
            Object.ToolTipText     =   "Quitter"
            Object.Tag             =   ""
            ImageIndex      =   22
         EndProperty
      EndProperty
      Begin VB.Label lblFont 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   135
         Left            =   8880
         TabIndex        =   1
         Top             =   120
         Width           =   495
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   2640
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   63
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":031A
            Key             =   "FindTable"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0634
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":094E
            Key             =   "Table"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0C68
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":0F82
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":129C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":15B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":18D0
            Key             =   "SortAsc"
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":1BEA
            Key             =   "New"
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":1F04
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":221E
            Key             =   "SortDes"
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2538
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2852
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2B6C
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":2E86
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":31A0
            Key             =   "Properties"
         EndProperty
         BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":34BA
            Key             =   "Archive"
         EndProperty
         BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":37D4
            Key             =   "Budget"
         EndProperty
         BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3AEE
            Key             =   "Pwd"
         EndProperty
         BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":3E08
            Key             =   "Facture"
         EndProperty
         BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4122
            Key             =   "Exit"
         EndProperty
         BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":443C
            Key             =   "Gestion"
         EndProperty
         BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4756
            Key             =   "MainDB"
         EndProperty
         BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4A70
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":4D8A
            Key             =   "Offre"
         EndProperty
         BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":50A4
            Key             =   "Personnel"
         EndProperty
         BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":53BE
            Key             =   "Virement"
         EndProperty
         BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":56D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":59F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":5D0C
            Key             =   "Affaire"
         EndProperty
         BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6026
            Key             =   "Param"
         EndProperty
         BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6340
            Key             =   "Calculette"
         EndProperty
         BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":665A
            Key             =   "Calendrier"
         EndProperty
         BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6974
            Key             =   "Prefer"
         EndProperty
         BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6C8E
            Key             =   "Appro"
         EndProperty
         BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":6FA8
            Key             =   "Saves"
         EndProperty
         BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":72C2
            Key             =   "Next"
         EndProperty
         BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":73D4
            Key             =   "Previous"
         EndProperty
         BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":74E6
            Key             =   "Begin"
         EndProperty
         BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":75F8
            Key             =   "End"
         EndProperty
         BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":770A
            Key             =   "Fermer"
         EndProperty
         BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":785C
            Key             =   "Filtre"
         EndProperty
         BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":7BAE
            Key             =   "Chercher"
         EndProperty
         BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":7F00
            Key             =   "Filtre2"
         EndProperty
         BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":821A
            Key             =   "Calc"
         EndProperty
         BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":8534
            Key             =   "Cheque"
         EndProperty
         BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":884E
            Key             =   "Conges"
         EndProperty
         BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":8B68
            Key             =   "Journal"
         EndProperty
         BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":8E82
            Key             =   "Livre"
         EndProperty
         BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":919C
            Key             =   "Rdv"
         EndProperty
         BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":94B6
            Key             =   "Calcul"
         EndProperty
         BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":97D0
            Key             =   "Salaries"
         EndProperty
         BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":9AEA
            Key             =   "F2x"
         EndProperty
         BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":9E04
            Key             =   "Variables"
         EndProperty
         BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":A11E
            Key             =   ""
         EndProperty
         BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":A230
            Key             =   "Cumuls"
         EndProperty
         BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":A54A
            Key             =   "Frs"
         EndProperty
         BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":A864
            Key             =   ""
         EndProperty
         BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":AB7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":AE98
            Key             =   ""
         EndProperty
         BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":B1B2
            Key             =   ""
         EndProperty
         BeginProperty ListImage63 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "MDIForm1.frx":B4CC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu xx 
      Caption         =   "xx"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    MakeToolBarFlat Me.Toolbar1
End Sub
