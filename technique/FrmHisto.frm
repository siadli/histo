VERSION 5.00
Begin VB.Form FrmHisto 
   BorderStyle     =   0  'None
   ClientHeight    =   8610
   ClientLeft      =   5640
   ClientTop       =   2850
   ClientWidth     =   9345
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   9345
   ShowInTaskbar   =   0   'False
   Begin VB.Frame CadreSupp 
      Caption         =   "Supprimer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2175
      Left            =   2520
      TabIndex        =   35
      Top             =   3360
      Width           =   5295
      Begin VB.CommandButton Annule 
         Caption         =   "&Annule"
         Height          =   615
         Left            =   3000
         TabIndex        =   37
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton CdeSpph 
         Caption         =   "&Ok"
         Height          =   615
         Left            =   600
         TabIndex        =   36
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "D�sirez-vous supprimer cet Historique ?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   38
         Top             =   600
         Width           =   4815
      End
   End
   Begin VB.Data DataHisto 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Data DataHistoTravo 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   405
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Data DataTravo 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.Frame CadreHisto 
      Caption         =   "Historique"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   18
      Top             =   240
      Width           =   8415
      Begin VB.CommandButton CdeModifierh 
         Caption         =   "Modifier"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   24
         Top             =   5640
         Width           =   1935
      End
      Begin VB.CommandButton CdeSupprimeh 
         Caption         =   "&Supprime"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         MaskColor       =   &H8000000D&
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   6480
         Width           =   1935
      End
      Begin VB.CommandButton CdeAnnuleh 
         Caption         =   "&Annule"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   22
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox TxtEffectu�h 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4455
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   3240
         Width           =   5415
      End
      Begin VB.Frame Framedate 
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   3495
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   2055
         Begin VB.ListBox ListeDateh 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2910
            Left            =   120
            TabIndex        =   20
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Label LblD�biteurh 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4320
         TabIndex        =   34
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label LblClient 
         Caption         =   "Ancien Client :"
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   2760
         TabIndex        =   33
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Kms"
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   6
         Left            =   2760
         TabIndex        =   32
         Top             =   2400
         Width           =   735
      End
      Begin VB.Label LblKilom�treh 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3840
         TabIndex        =   31
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Emp."
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   5
         Left            =   2760
         TabIndex        =   30
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label LblEmploy�h 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3840
         TabIndex        =   29
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "N�:"
         ForeColor       =   &H00C00000&
         Height          =   435
         Index           =   4
         Left            =   2760
         TabIndex        =   28
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label LblNumVoitureh 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3840
         TabIndex        =   27
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Ima."
         ForeColor       =   &H00C00000&
         Height          =   435
         Index           =   3
         Left            =   4920
         TabIndex        =   26
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label LblImmah 
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5760
         TabIndex        =   25
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame CadreTravaux 
      Caption         =   "Modification Historique"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8415
      Begin VB.TextBox TxtDateTravo 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   960
         MaxLength       =   10
         TabIndex        =   10
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtEmploy�Travo 
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   960
         MaxLength       =   35
         TabIndex        =   9
         Text            =   "Bernard SINIBALDI"
         Top             =   960
         Width           =   2772
      End
      Begin VB.Frame CadreEffectu� 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4815
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox TxtModifi� 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2775
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   7
            Top             =   1080
            Width           =   4695
         End
         Begin VB.CommandButton CdeAnnuleEffectu� 
            Caption         =   "&Annule"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1440
            TabIndex        =   6
            Top             =   4080
            Width           =   1335
         End
         Begin VB.TextBox TxtKilom�tre 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   5
            Top             =   240
            Width           =   1455
         End
         Begin VB.CommandButton CdeModifier 
            Caption         =   "Modifier"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2760
            TabIndex        =   4
            Top             =   4080
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Kilom�tre:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.ListBox ListeEmploy� 
         Appearance      =   0  'Flat
         Height          =   930
         Left            =   960
         TabIndex        =   2
         Top             =   1320
         Visible         =   0   'False
         Width           =   2772
      End
      Begin VB.TextBox LblAncienKM 
         Appearance      =   0  'Flat
         BackColor       =   &H80000010&
         Height          =   420
         Left            =   3960
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox CommonDialog1 
         Height          =   480
         Left            =   3840
         ScaleHeight     =   420
         ScaleWidth      =   1140
         TabIndex        =   39
         Top             =   360
         Width           =   1200
      End
      Begin VB.Frame CadrePr�voir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   240
         TabIndex        =   11
         Top             =   2520
         Visible         =   0   'False
         Width           =   5775
         Begin VB.TextBox TxtPr�voir 
            Height          =   2535
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   240
            Width           =   4095
         End
         Begin VB.CommandButton CdeOk 
            Caption         =   "&Ok"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   13
            Top             =   3000
            Width           =   972
         End
         Begin VB.CommandButton CdeAnnule 
            Caption         =   "&Annule"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1080
            TabIndex        =   12
            Top             =   3000
            Width           =   1095
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Resp.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   615
      End
      Begin VB.Label LblNumVoiture 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2640
         TabIndex        =   15
         Top             =   480
         Width           =   735
      End
      Begin VB.Image Travo1 
         Height          =   1260
         Left            =   6000
         Picture         =   "FrmHisto.frx":0000
         Top             =   360
         Width           =   1200
      End
      Begin VB.Image Travo2 
         Height          =   1260
         Left            =   6000
         Picture         =   "FrmHisto.frx":1E82
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.Menu MnuFichier 
      Caption         =   "&Fichier"
      Begin VB.Menu MnuQuitte 
         Caption         =   "&Quitte"
         Shortcut        =   ^Q
      End
   End
End
Attribute VB_Name = "FrmHisto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

Dim Alert
   
' LblNumVoiture = FrmV�hicule.LblNumVoiture
   ' LblNumVoitureh = FrmV�hicule.LblNumVoiture
   
   LblNumVoiture = "1"
    LblNumVoitureh = "1"
    
  'Historique
      CadreTravaux.Visible = False
      CadreHisto.Visible = True
      CadreSupp.Visible = False
     
     '''''''''''''''''''''''''
   
'     LblImmah = FrmV�hicule.LblImma
     LblImmah = "7441 TP 25"
     
         ListeDateh.Clear
       DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
       
         DataHisto.RecordSource = "select * from histo where num�ro=" & Val(LblNumVoiture) & " order by kilom�tre desc"
        
         DataHisto.Refresh
         Do While DataHisto.Recordset.EOF = False
          ListeDateh.AddItem DataHisto.Recordset("date")
      
          DataHisto.Recordset.MoveNext
         Loop
     
        FrmHisto.Caption = "Historique: voiture "
     
     
End Sub


Private Sub Form_Unload(Cancel As Integer)

 DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
         DataHisto.RecordSource = "select * from histo where num�ro=" & Val(LblNumVoiture) & " order by date desc"
         DataHisto.Refresh
        
        
         Do While DataHisto.Recordset.EOF = False
       
       


         Exit Do
Loop



Unload FrmHisto
End Sub
Private Sub Annule_Click()
CadreTravaux.Visible = False
      CadreHisto.Visible = True
      CadreSupp.Visible = False
End Sub

Private Sub CdeAnnule_Click()
    Unload FrmTravaux
     FrmV�hicule.WindowState = vbMaximized
End Sub

Private Sub CdeAnnuleEffectu�_Click()
   CadreTravaux.Visible = False
      CadreHisto.Visible = True
End Sub

Private Sub CdeModifier_Click()
 DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
        DataHisto.RecordSource = "select * from histo where date=" & Format(TxtDateTravo) & ""
        DataHisto.Refresh
        DataHisto.Database.Execute "delete  * from histo where date=" & Format(TxtDateTravo) & ""
        
        
        
         DataTravo.DatabaseName = Chemin & "\technique\technique.mdb"
         DataTravo.RecordSource = "select * from Histo"
         DataTravo.Refresh
         DataTravo.Recordset.AddNew
         DataTravo.Recordset("num�ro") = Val(Me.LblNumVoiture)
         DataTravo.Recordset("pr�voir") = Me.TxtModifi�
         DataTravo.Recordset("date") = Me.TxtDateTravo
         DataTravo.Recordset("Employ�") = Me.TxtEmploy�Travo
         DataTravo.Recordset("d�biteur") = FrmV�hicule.TxtD�biteur
         
        
         DataTravo.Recordset("Kilom�tre") = Me.TxtKilom�tre
         
        DataTravo.Recordset.Update
         
        CadreTravaux.Visible = False
        CadreHisto.Visible = True
        TxtEffectu�h.Refresh
        
End Sub

Private Sub CdeSpph_Click()
DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
        DataHisto.RecordSource = "select * from histo where date=" & Format(ListeDateh) & ""
        DataHisto.Refresh
        'DataHisto.Database.Execute "delete  * from histo where date=" & Format(ListeDateh) & ""
        DataHisto.Database.Execute "delete  * from histo where date=" & Format(ListeDateh) & " and Num�ro = " & Val(LblNumVoitureh) & ""
        
        CdeSupprimeh.Enabled = False
        
         ListeDateh.Clear
         DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
         DataHisto.RecordSource = "select * from histo where num�ro=" & Val(LblNumVoitureh) & " order by date desc"
         DataHisto.Refresh
         Do While DataHisto.Recordset.EOF = False
          ListeDateh.AddItem DataHisto.Recordset("date")
         
          DataHisto.Recordset.MoveNext
         Loop
         LblD�biteurh = ""
        LblEmploy�h = ""
        LblKilom�treh = "":
        TxtEffectu�h = "":
        
       CadreTravaux.Visible = False
      CadreHisto.Visible = True
      CadreSupp.Visible = False
       
End Sub

Private Sub ListeEmploy�_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
     TxtEmploy�Travo.SetFocus
     ListeEmploy�.Visible = False
    End If

    If KeyAscii = 13 Then
     i = ListeEmploy�.ListIndex
     If i < 0 Then i = 0
     TxtEmploy�Travo = ListeEmploy�.List(i)
     TxtEmploy�Travo.SetFocus
     ListeEmploy�.Visible = False
    End If


End Sub

Private Sub ListeEmploy�_LostFocus()
    ListeEmploy�.Visible = False
End Sub

Private Sub MnuAjoutTexte_Click()
    Reponse = InputBox("Entrez le texte standard.", "Ajout d'un texte", "Texte standard")
    If Len(Reponse) = 0 Then Exit Sub
         DataTravo.DatabaseName = Chemin & "\technique\technique.mdb"
         DataTravo.RecordSource = "select * from texte"
         DataTravo.Refresh
         DataTravo.Recordset.AddNew
         DataTravo.Recordset("texte") = Reponse
         DataTravo.Recordset.Update
        

End Sub

Private Sub MnuCr�erEmploy�_Click()
    Reponse = InputBox("Entrez le Pr�nom et le nom de l'employ�.", "Cr�ation fiche employ�", "Pr�nom NOM")
    If Len(Reponse) = 0 Then Exit Sub
         DataTravo.DatabaseName = Chemin & "\technique\technique.mdb"
         DataTravo.RecordSource = "select * from Codeemploy� "
         DataTravo.Refresh
         DataTravo.Recordset.AddNew
         DataTravo.Recordset("D�signation") = Reponse
         DataTravo.Recordset.Update

End Sub

Private Sub MnuQuitte_Click()

 DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
         DataHisto.RecordSource = "select * from histo where num�ro=" & Val(LblNumVoiture) & " order by date desc"
         DataHisto.Refresh
        
        
         Do While DataHisto.Recordset.EOF = False
       
       If IsNull(DataHisto.Recordset("Kilom�tre")) = False Then
        FrmV�hicule.LbldKmvisite.Caption = DataHisto.Recordset("Kilom�tre")
       Else
        LbldKmvisite = ""
       End If


         Exit Do
Loop

    Unload FrmHisto
    FrmV�hicule.WindowState = vbMaximized
End Sub



Private Sub TxtDateTravo_GotFocus()
    Surbrillance
End Sub

Private Sub TxtDateTravo_LostFocus()
   If TxtDateTravo <> "" Then
    If IsDate(TxtDateTravo) = False Then
     MsgBox "Erreur dans la date", vbCritical, "Erreur"
     TxtDateTravo.SetFocus
    Else
     TxtDateTravo = Format(TxtDateTravo, "dd.mm.yyyy")
    End If
   End If

End Sub






Private Sub TxtEmploy�Travo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then
        KeyAscii = 0
         ListeEmploy�.Clear
         DataTravo.DatabaseName = Chemin & "\technique\technique.mdb"
         DataTravo.RecordSource = "select * from Codeemploy� order by d�signation"
         DataTravo.Refresh
         Do While DataTravo.Recordset.EOF = False
          ListeEmploy�.AddItem DataTravo.Recordset("D�signation")
          DataTravo.Recordset.MoveNext
         Loop
         ListeEmploy�.Visible = True
         ListeEmploy�.SetFocus
    End If
End Sub

Private Sub TxtKilom�tre_KeyPress(KeyAscii As Integer)
    KeyAscii = ControleAscii(KeyAscii, 1)
End Sub

Private Sub TxtModifi�_Change()
CdeModifier.Enabled = True
End Sub

Private Sub CdeAnnuleh_Click()

DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
         DataHisto.RecordSource = "select * from histo where num�ro=" & Val(LblNumVoiture) & " order by date desc"
         DataHisto.Refresh
        
        
         Do While DataHisto.Recordset.EOF = False
       
       If IsNull(DataHisto.Recordset("Kilom�tre")) = False Then
        FrmV�hicule.LbldKmvisite.Caption = DataHisto.Recordset("Kilom�tre")
       Else
        LbldKmvisite = ""
       End If


         Exit Do
Loop


  FrmV�hicule.ListeNumV�hicule.List(k) = LblNumVoiture
Unload FrmHisto

 
End Sub

Private Sub CdeSupprimeh_Click()
       ' R�ponse = MsgBox("D�sirez-vous supprimer cet historique ?", vbCritical + vbYesNo, "Suppression")
       ' If R�ponse = vbNo Then Exit Sub
       ' DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
       ' DataHisto.RecordSource = "select * from histo where date=" & Format(ListeDateh) & ""
      '  DataHisto.Refresh
       ' DataHisto.Database.Execute "delete  * from histo where date=" & Format(ListeDateh) & ""
       ' CdeSupprimeh.Enabled = False
        
        ' ListeDateh.Clear
       '  DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
       '  DataHisto.RecordSource = "select * from histo where num�ro=" & Val(LblNumVoiture) & " order by date desc"
       '  DataHisto.Refresh
        ' Do While DataHisto.Recordset.EOF = False
       '   ListeDateh.AddItem DataHisto.Recordset("date")
         
        '  DataHisto.Recordset.MoveNext
       '  Loop
       '  LblD�biteurh = ""
        ' LblEmploy�h = ""
       '  LblKilom�treh = "":
         
        CadreTravaux.Visible = False
      CadreHisto.Visible = False
      CadreSupp.Visible = True
         
End Sub
Private Sub ListeDateh_Click()
         DataHisto.DatabaseName = Chemin & "\technique\technique.mdb"
       DataHisto.RecordSource = "select * from histo where num�ro=" & Val(LblNumVoiture) & " and date =" & Format(ListeDateh) & ""
        
         
   '   "vb6 "  DataHisto.RecordSource = "select * from histo where date = # " & Format(ListeDateh, "dd.mm.yyyy") & " # "
        
         DataHisto.Refresh
         Do While DataHisto.Recordset.EOF = False
          LblEmploy�h = DataHisto.Recordset("employ�")
          LblKilom�treh = DataHisto.Recordset("kilom�tre")
          TxtEffectu�h = DataHisto.Recordset("pr�voir")
          
          If IsNull(DataHisto.Recordset("D�biteur")) = False Then
           LblD�biteurh = DataHisto.Recordset("D�biteur")
           
           
          Else
           LblD�biteurh = ""
          End If
          
          Exit Do
         Loop
         CdeSupprimeh.Enabled = True
         CdeModifierh.Enabled = True
End Sub

Private Sub MnuQuitteh_Click()
    Unload FrmHisto
End Sub

Private Sub CdeModifierh_Click()


     TxtDateTravo = ListeDateh

TxtKilom�tre = LblKilom�treh

  TxtModifi� = TxtEffectu�h & vbCrLf
  CadreEffectu�.Visible = True
    
    ListeDateh.Refresh
    
  
    
    
    'Modification des travaux effectu�s

      CadreTravaux.Visible = True
      CadreHisto.Visible = False
    Travo1.Visible = True
    Travo2.Visible = False
  CadreEffectu�.Visible = True
  TxtModifi�.Visible = True
  CdeModifier.Enabled = False
  
  TxtKilom�tre.Locked = True
  TxtDateTravo.Locked = True
  
 
   
    
     CdeModifier.Visible = True
     
    
End Sub


