VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{4777436C-EB8C-4596-98A8-EBCF98683969}#1.8#0"; "FlexCell.ocx"
Begin VB.Form frmptautomation_medak 
   Caption         =   "frmptautomation"
   ClientHeight    =   10470
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14730
   LinkTopic       =   "Form4"
   ScaleHeight     =   10470
   ScaleWidth      =   14730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5535
      Left            =   120
      TabIndex        =   105
      Top             =   4680
      Width           =   11295
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2595
         ItemData        =   "Frmautomation_medak1.frx":0000
         Left            =   120
         List            =   "Frmautomation_medak1.frx":0007
         TabIndex        =   119
         Top             =   480
         Width           =   4845
      End
      Begin VB.TextBox Rem1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   118
         Top             =   4320
         Width           =   4935
      End
      Begin VB.TextBox txtFlawMsg 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Height          =   480
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   117
         Top             =   3720
         Width           =   10575
      End
      Begin VB.TextBox txtbreakidno 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   116
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtTotPT 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   115
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtnextAllocateidLen 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   8400
         Locked          =   -1  'True
         TabIndex        =   114
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtNextAllocateID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   2160
         Width           =   2025
      End
      Begin VB.TextBox txtOriginGap 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   112
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtLastBarcodeID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   111
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox balqty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6360
         Locked          =   -1  'True
         TabIndex        =   110
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtDrawBarcodeId 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   9000
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox ptype 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6960
         Locked          =   -1  'True
         TabIndex        =   108
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox drqty 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DragMode        =   1  'Automatic
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   107
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   4920
         Width           =   8535
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "PT LOG"
         Height          =   195
         Left            =   120
         TabIndex        =   135
         Top             =   240
         Width           =   585
      End
      Begin VB.Label Label4 
         Caption         =   "Next BFD/Lump Position"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   134
         Top             =   3480
         Width           =   2835
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Last BreakNo."
         Height          =   195
         Index           =   0
         Left            =   7800
         TabIndex        =   133
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Total PT Done"
         Height          =   195
         Index           =   8
         Left            =   5160
         TabIndex        =   132
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "BE Scrap"
         Height          =   375
         Index           =   1
         Left            =   7800
         TabIndex        =   131
         Top             =   840
         Width           =   930
      End
      Begin VB.Label txtBS 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   9000
         TabIndex        =   130
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Next  ID"
         Height          =   195
         Index           =   7
         Left            =   5160
         TabIndex        =   129
         Top             =   2280
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Origin Gap"
         Height          =   195
         Left            =   5160
         TabIndex        =   128
         Top             =   1680
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Balance Length"
         Height          =   195
         Index           =   1
         Left            =   5160
         TabIndex        =   127
         Top             =   840
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Draw BarcodeId"
         Height          =   195
         Left            =   7800
         TabIndex        =   126
         Top             =   1305
         Width           =   1155
      End
      Begin VB.Label Label15 
         Caption         =   "Machine No. for Draw ID"
         Height          =   375
         Left            =   120
         TabIndex        =   125
         Top             =   4920
         Width           =   1935
      End
      Begin VB.Label Label13 
         Caption         =   "DrQty"
         Height          =   255
         Left            =   5160
         TabIndex        =   124
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Last ID"
         Height          =   255
         Left            =   7800
         TabIndex        =   123
         Top             =   1755
         Width           =   1095
      End
      Begin VB.Label lblMSStatus 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   5160
         TabIndex        =   122
         Top             =   2640
         Width           =   5535
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "                "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   5160
         TabIndex        =   121
         Top             =   3120
         Width           =   5535
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Fault at"
         Height          =   195
         Left            =   5040
         TabIndex        =   120
         Top             =   4440
         Width           =   525
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   120
      TabIndex        =   40
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame2 
         Caption         =   "(B-BFD,L-Lumps,C-SCD,S-Bot End,M-Multiple End,D-Scratch)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   3600
         TabIndex        =   75
         Top             =   120
         Width           =   5715
         Begin VB.CheckBox ChkScratch 
            Caption         =   "Scratch Rejection"
            Height          =   285
            Left            =   120
            TabIndex        =   82
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox Remark 
            Height          =   375
            Left            =   2520
            MaxLength       =   15
            TabIndex        =   81
            Top             =   840
            Width           =   2895
         End
         Begin VB.ComboBox DRReason 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Frmautomation_medak1.frx":0012
            Left            =   3360
            List            =   "Frmautomation_medak1.frx":0014
            TabIndex        =   80
            Top             =   240
            Width           =   1815
         End
         Begin VB.CheckBox chkMERej 
            Caption         =   "Mulitple End Rejection"
            Height          =   285
            Left            =   120
            TabIndex        =   79
            Top             =   960
            Width           =   1935
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Balance Draw Rejection"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   600
            Width           =   2055
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Rejection"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   300
            Width           =   975
         End
         Begin VB.ComboBox cmbmutiend 
            Height          =   315
            ItemData        =   "Frmautomation_medak1.frx":0016
            Left            =   2520
            List            =   "Frmautomation_medak1.frx":002C
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   1560
            Visible         =   0   'False
            Width           =   2775
         End
         Begin VB.Label Label4 
            Caption         =   "Remark for draw rejection"
            Height          =   255
            Index           =   0
            Left            =   2520
            TabIndex        =   85
            Top             =   600
            Width           =   1935
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Reason for DrRejection"
            Height          =   195
            Left            =   1560
            TabIndex        =   84
            Top             =   300
            Width           =   1665
         End
         Begin VB.Label LBLMULTIEND 
            Caption         =   "Reason for Multiple End"
            Height          =   255
            Left            =   2520
            TabIndex        =   83
            Top             =   1320
            Visible         =   0   'False
            Width           =   1935
         End
      End
      Begin VB.ComboBox cboShiftIncharge 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":0079
         Left            =   1560
         List            =   "Frmautomation_medak1.frx":007B
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   4200
         Width           =   1815
      End
      Begin VB.ComboBox txtScReason 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":007D
         Left            =   7320
         List            =   "Frmautomation_medak1.frx":008D
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   3600
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Check PTID"
         Height          =   435
         Left            =   9360
         TabIndex        =   72
         Top             =   1440
         Width           =   1665
      End
      Begin VB.CommandButton cmdUnloadME 
         Caption         =   "Unload for ME"
         Height          =   435
         Left            =   9360
         TabIndex        =   71
         Top             =   2400
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.CommandButton cmdAllocatedID 
         Caption         =   "Get allocated ID"
         Height          =   435
         Left            =   9330
         TabIndex        =   70
         Top             =   480
         Width           =   1665
      End
      Begin VB.CommandButton cmdMELen 
         Caption         =   "M End Len Calc "
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9345
         TabIndex        =   69
         Top             =   960
         Width           =   1665
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "&Exit"
         Height          =   435
         Left            =   9360
         TabIndex        =   68
         Tag             =   "NA"
         Top             =   1920
         Width           =   1665
      End
      Begin VB.Frame Frame5 
         Caption         =   "TimeLoss"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   3600
         TabIndex        =   61
         Top             =   2040
         Width           =   5475
         Begin VB.ComboBox txtTLReason 
            Height          =   315
            ItemData        =   "Frmautomation_medak1.frx":00BB
            Left            =   1560
            List            =   "Frmautomation_medak1.frx":00BD
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   720
            Width           =   2775
         End
         Begin VB.TextBox txtSpeedLoss 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   3240
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   63
            Top             =   270
            Width           =   1095
         End
         Begin VB.TextBox txtTimeLoss 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   1080
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   62
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label14 
            Caption         =   "Time Loss Reason"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Speed Loss"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   66
            Top             =   360
            Width           =   840
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "TimeLoss"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   65
            Top             =   360
            Width           =   675
         End
      End
      Begin VB.TextBox txtBrkCntLen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   60
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox txtBarcodeID 
         BackColor       =   &H0080FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   59
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox cboColor 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":00BF
         Left            =   1560
         List            =   "Frmautomation_medak1.frx":00CC
         TabIndex        =   58
         Text            =   "cboColor"
         Top             =   3820
         Width           =   1815
      End
      Begin VB.TextBox PTID 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   1530
         TabIndex        =   57
         Top             =   1180
         Width           =   1815
      End
      Begin VB.ComboBox ptno 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":00E6
         Left            =   1560
         List            =   "Frmautomation_medak1.frx":00E8
         TabIndex        =   56
         Top             =   3060
         Width           =   1815
      End
      Begin VB.ComboBox opr 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":00EA
         Left            =   1560
         List            =   "Frmautomation_medak1.frx":00EC
         Sorted          =   -1  'True
         TabIndex        =   55
         Text            =   "AAA"
         Top             =   3440
         Width           =   1815
      End
      Begin VB.ComboBox drawid 
         DataField       =   "ID"
         DataSource      =   "Adodc1"
         Enabled         =   0   'False
         Height          =   315
         Left            =   1542
         Sorted          =   -1  'True
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   800
         Width           =   1815
      End
      Begin VB.TextBox ptlen 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1560
         TabIndex        =   53
         Top             =   2210
         Width           =   1815
      End
      Begin VB.TextBox txtMultimodeID 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   52
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox txtScr 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":00EE
         Left            =   3600
         List            =   "Frmautomation_medak1.frx":00F0
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   4200
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CheckBox PTscrap 
         Caption         =   "PT Scrap"
         Height          =   255
         Left            =   3600
         TabIndex        =   50
         Top             =   3360
         Width           =   975
      End
      Begin VB.CheckBox chkpmd 
         Caption         =   "ZTPMD"
         Height          =   195
         Left            =   4800
         TabIndex        =   49
         Top             =   3360
         Width           =   1095
      End
      Begin VB.ComboBox txtBobType 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":00F2
         Left            =   4920
         List            =   "Frmautomation_medak1.frx":00F4
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CheckBox chkDeviationSpol 
         Caption         =   "Deviation Spool"
         Height          =   375
         Left            =   9360
         TabIndex        =   47
         Top             =   2880
         Width           =   1455
      End
      Begin VB.TextBox txtSpoolNo 
         Height          =   375
         Left            =   9360
         TabIndex        =   46
         Top             =   3480
         Width           =   1695
      End
      Begin VB.ComboBox txtBobNo 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":00F6
         Left            =   6240
         List            =   "Frmautomation_medak1.frx":0103
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   4200
         Width           =   1095
      End
      Begin VB.ComboBox txtVibration 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":0110
         Left            =   7560
         List            =   "Frmautomation_medak1.frx":011A
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   4200
         Width           =   1095
      End
      Begin VB.ComboBox txtDancer 
         Height          =   315
         ItemData        =   "Frmautomation_medak1.frx":0127
         Left            =   8880
         List            =   "Frmautomation_medak1.frx":0131
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   4200
         Width           =   1095
      End
      Begin VB.CheckBox chkdocsample 
         Caption         =   "DOC "
         Height          =   255
         Left            =   3600
         TabIndex        =   42
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox txtProducttype 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   315
         Left            =   10080
         TabIndex        =   41
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Shift Incharge"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   104
         Top             =   4200
         Width           =   990
      End
      Begin VB.Label lblErrMsg 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   6660
         TabIndex        =   103
         Top             =   4680
         Width           =   45
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "PTScrap Reason"
         Height          =   195
         Left            =   6000
         TabIndex        =   102
         Top             =   3720
         Width           =   1230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "BrkCntLen"
         Height          =   255
         Left            =   120
         TabIndex        =   101
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Barcode"
         Height          =   255
         Left            =   120
         TabIndex        =   100
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Bobbin Color"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   99
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Machine Number"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   98
         Top             =   3150
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Operator Name"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   97
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PT Length"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   96
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "PT ID"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   95
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Draw ID"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   94
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "For Multomode"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   93
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Scratches"
         Height          =   255
         Left            =   3600
         TabIndex        =   92
         Top             =   3960
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Bobbin Type"
         Height          =   195
         Index           =   4
         Left            =   4920
         TabIndex        =   91
         Top             =   3960
         Width           =   900
      End
      Begin VB.Label Label20 
         Caption         =   "MainSpool No"
         Height          =   255
         Left            =   9360
         TabIndex        =   90
         Top             =   3240
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Bobbin No"
         Height          =   375
         Index           =   1
         Left            =   6240
         TabIndex        =   89
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label21 
         Caption         =   "Payoff Vibration"
         Height          =   255
         Left            =   7560
         TabIndex        =   88
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label22 
         Caption         =   "Dancer Vibration"
         Height          =   255
         Left            =   8880
         TabIndex        =   87
         Top             =   3960
         Width           =   1335
      End
      Begin VB.Label Label23 
         Caption         =   "Product Type"
         Height          =   255
         Left            =   10080
         TabIndex        =   86
         Top             =   3960
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   11760
      TabIndex        =   21
      Top             =   7680
      Visible         =   0   'False
      Width           =   5115
      Begin VB.TextBox txtFactor1 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton CmdSAP 
         Caption         =   "SAP Entry"
         Enabled         =   0   'False
         Height          =   345
         Left            =   120
         TabIndex        =   31
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox date3 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         HideSelection   =   0   'False
         Left            =   1920
         TabIndex        =   30
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox time3 
         BackColor       =   &H80000000&
         Enabled         =   0   'False
         Height          =   375
         HideSelection   =   0   'False
         Left            =   1920
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Frame Frame 
         Height          =   1095
         Left            =   1560
         TabIndex        =   23
         Top             =   1440
         Visible         =   0   'False
         Width           =   1815
         Begin VB.CheckBox chkFlaw 
            Alignment       =   1  'Right Justify
            Caption         =   "Flaw"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1005
         End
         Begin VB.TextBox txtFlawNo 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1410
            TabIndex        =   26
            Top             =   240
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.CheckBox chkME 
            Alignment       =   1  'Right Justify
            Caption         =   "ME"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   25
            Top             =   630
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.TextBox txtCuttingLen 
            Height          =   315
            Left            =   1380
            TabIndex        =   24
            Top             =   630
            Visible         =   0   'False
            Width           =   315
         End
         Begin VB.Label lblDateDiff 
            AutoSize        =   -1  'True
            Caption         =   "0"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   1080
            Width           =   90
         End
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2760
         TabIndex        =   22
         Top             =   960
         Width           =   2295
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FG1 
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   1920
         Visible         =   0   'False
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         Cols            =   10
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FG2 
         Height          =   255
         Left            =   840
         TabIndex        =   34
         Top             =   1920
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   450
         _Version        =   393216
         Cols            =   10
         _NumberOfBands  =   1
         _Band(0).Cols   =   10
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   38501
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Location of Pitchchange with reason"
         Height          =   195
         Left            =   3360
         TabIndex        =   37
         Top             =   240
         Visible         =   0   'False
         Width           =   2355
      End
      Begin VB.Label Label9 
         Caption         =   "Location of Flaws"
         Height          =   315
         Left            =   2760
         TabIndex        =   36
         Top             =   600
         Width           =   2175
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   18360
      Top             =   5760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   18360
      Top             =   4800
   End
   Begin VB.Timer Timer3 
      Interval        =   400
      Left            =   18360
      Top             =   5280
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1245
      Left            =   11640
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   "Bobbin Out Entry"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   11520
      TabIndex        =   13
      Top             =   0
      Width           =   3375
      Begin VB.ListBox lstOutScanPending 
         Height          =   840
         Left            =   1320
         TabIndex        =   16
         Top             =   690
         Width           =   1815
      End
      Begin VB.CommandButton cmdOutScanPending 
         Caption         =   "Out Scan Pending List"
         Height          =   435
         Left            =   1320
         TabIndex        =   15
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox txtPTOut 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Out Scan Pending List"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   10
         Left            =   240
         TabIndex        =   18
         Top             =   750
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "OUT Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.TextBox txtCtrlMsg 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   16440
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Log Sheet"
      Height          =   495
      Left            =   13200
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   17040
      TabIndex        =   9
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   11520
      TabIndex        =   8
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   11520
      TabIndex        =   7
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   15480
      TabIndex        =   4
      Text            =   "Text4"
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MBend samples"
      Height          =   495
      Left            =   15480
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtCoatType 
      Enabled         =   0   'False
      Height          =   375
      Left            =   11880
      TabIndex        =   2
      Top             =   7200
      Width           =   2535
   End
   Begin VB.CheckBox checkflawnew 
      Caption         =   "Check Flaw"
      Height          =   375
      Left            =   11520
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
   Begin FlexCell.Grid Grid5 
      Height          =   2655
      Left            =   16920
      TabIndex        =   1
      Top             =   1440
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   4683
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin FlexCell.Grid Grid4 
      Height          =   1695
      Left            =   15840
      TabIndex        =   5
      Top             =   5400
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   2990
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin FlexCell.Grid Grid3 
      Height          =   1935
      Left            =   15120
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   3413
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin FlexCell.Grid Grid1 
      Height          =   1575
      Left            =   12600
      TabIndex        =   11
      Top             =   2400
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2778
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   10110
      Width           =   14730
      _ExtentX        =   25982
      _ExtentY        =   635
      _Version        =   327682
      Appearance      =   1
   End
   Begin FlexCell.Grid grdFlaw 
      Height          =   1755
      Left            =   11520
      TabIndex        =   38
      Top             =   5400
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   3096
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin FlexCell.Grid Grid2 
      Height          =   1755
      Left            =   15120
      TabIndex        =   39
      Top             =   2400
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   3096
      Cols            =   5
      DefaultFontSize =   8.25
      Rows            =   30
   End
   Begin VB.Label LBLSAMPLENO 
      Height          =   255
      Left            =   19320
      TabIndex        =   136
      Top             =   960
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmptautomation_medak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Public cn As New ADODB.Connection
Public cmd As New ADODB.Command
Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Public DATE1 As Date, date2 As Date, TIME1, TIME2
Dim McStdate As Date, McStTIME
Public DateDrrej As Date, TimedrRej
Public DTSCR As Boolean
Dim datedr As Date, PRDate As Date
Dim FStr, FolderPath, ThisFolder As Folder, f As Files, F1 As File
Dim nbreak As Single
Dim nbreak1 As Single
Dim CHK As Boolean
Dim protype
Dim TagDrrej As Single, BrkCnt, TypeOfFlaw, TypeOfRemark, ActBrk
Dim Factor As Double, Stag
Dim RS11 As New ADODB.Recordset
Dim rsbar As New ADODB.Recordset
Dim Minlen As Integer
Dim RsRel As New ADODB.Recordset
Dim RsPT As New ADODB.Recordset
Dim OHSamFlag As Boolean
Dim ChinaPTMSFlag As Boolean
Dim MatNo
Dim CoreSinterMcno
Dim CoreDepMcno
Dim SocDepMcno
Dim scrqty As Integer
Dim drrej
Dim SSRC_Coat_Flag As Boolean
Dim FSYS As New FileSystemObject
Dim INSTREAM As TextStream
Dim LUVANTIX_Coat_Flag As Boolean
Dim PTLoginRecFlag As Boolean
Dim PT50Flag  As Boolean
Dim McPath
Dim iPTContBrkCnt As Integer
Dim iPTContBrkFlag As Integer
'---------------------Create Boolean Variables for Integrated cutting logic Sunil 10.03.2007---------
Dim flag_rejection As Boolean
Dim pref_stretch As Boolean  ''added for pref stretch
Dim flag_china As Boolean
Dim Flag_Zgade As Boolean
Dim BendliteTxt As String
Dim DefaultTxt As String
Dim ChinaTxt As String
Dim PhyCheck As Boolean
Dim checkbarcodeid As String
Dim srnogotf
Dim NextFault
Dim FlagPMD As Boolean
Dim HoldForInst As Boolean
Dim StandardLen As Double
Dim Go_for_Sample As Integer
Dim S1, S2, S3, S4, S5, S6, S7, S8, S9, S10, SampleCnt As Integer
Dim FlagScratch As Boolean
Dim PTMode As String
Dim NextFlawPosGlobal
Dim plant
Dim totscrap, totrej As Integer
Dim finalstatus As Boolean
Dim lockbarcedflag As Boolean
Dim datasaved As Boolean
Dim bobbinprefend As Integer  '''this is for preform end
'''CODE ADDED BY DINESH ON REF MAYUR SIR ON 30032018
'''=================================================
Dim BOBBINLENGTH As Double
'''===========================================
Dim medakfilename As String
Dim LINECOUNTER As Integer
Dim LENGHGET As Boolean

''code added by dinesh for variable declaratio of option explicit
'Dim stpos, Edpos As Double
'Dim sErrorDescription As String
'Dim TwNo As Integer
'Dim SetLen As Double
'Dim TestFlag As Boolean


Private Sub Check1_Click()
'    If Check1.Value = 1 And Trim(REMARK) = "" Then
'        MsgBox "Enter remark for draw rejection", vbCritical, "Error"
'        Check1 = 0
'        Exit Sub
'    End If
   On Error GoTo Check1_Click_Error:
   
If Check1.value = 1 Then
    'ChkFlaw
'    TypeOfFlaw = ""
'    TypeOfRemark = ""
'    DRReason.Text = ""
'    Remark.Text = ""
    If Check2.value = 1 Then
        TypeOfFlaw = "S"
        TypeOfRemark = "BESCRAP"
    End If
    If TypeOfFlaw = "" Or TypeOfFlaw = "U" Then
    
        stpos = Val(balqty) / 1000 - 2
        EdPos = Val(balqty) / 1000 + 2
        RS.Open "select * from DRAWflaw where FID = '" & drawid.Text & "' and  position >'" & stpos & "' and position <'" & EdPos & "'", cn, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            Do While RS.EOF = False
'                If rs("REMARK") <> "AIRLINE" Then   ''''''''''''ADDED BY ABDUL ON 12.07.2013 AS PER DILIP/ASHISH
                    If RS("REMARK") = "LUMPS" Or Mid(RS("REMARK"), 1, 1) = "L" Then
                        TypeOfFlaw = "L"
                        TypeOfRemark = "LUMPS"
                    End If
                    If RS("REMARK") = "BFD" Or Mid(RS("REMARK"), 1, 1) = "B" Then
                        TypeOfFlaw = "B"
                        TypeOfRemark = "BFD"
                    End If
                    If RS("REMARK") = "AIRLINE" Or Mid(RS("REMARK"), 1, 1) = "A" Then
                        TypeOfFlaw = "A"
                        TypeOfRemark = "AIRLINE"
                    End If
                    If RS("REMARK") = "SCD" Or RS("REMARK") = "CD2" Then
                        TypeOfFlaw = "C"
                        TypeOfRemark = "SCD"
                    End If
                    If RS("REMARK") = "PITCHCHANGE" Then
                        If TypeOfFlaw = "" Then
                            TypeOfFlaw = "P"
                            TypeOfRemark = "PITCHCHANGE"
                        End If
                    End If
'                End If
                RS.MoveNext
            Loop
        End If
        RS.Close
        If TypeOfFlaw = "" Then
            TypeOfFlaw = "U" 'For PT Machine Fault
            TypeOfRemark = "UNKNOWN"
        Else
            frmptautomation.DRReason = TypeOfFlaw
            Remark.Text = TypeOfRemark
        End If
    Else
        frmptautomation.DRReason = TypeOfFlaw
        Remark.Text = TypeOfRemark
    End If
End If
    
If Check1.value = 0 Then
    DRReason.Text = ""
    Remark.Text = ""
'    TypeOfFlaw = ""
'    TypeOfRemark = ""
End If

'If TypeOfFlaw <> "" Then DRReason = TypeOfFlaw
'DRReason = Trim(UCase$(DRReason.Text))
'If Check1.Value = 1 And Not (DRReason = "L" Or DRReason = "B" Or DRReason = "S" Or DRReason = "C" Or DRReason = "O") Then
'    MsgBox "Enter correct Reason for DrRejection (F-flaw,B-BFD,S-Bottom End Scrap)", vbCritical, "Error"
'    Check1 = 0
'    Exit Sub
'End If
If Check1.value = 1 Then PTscrap.value = 0
   
   Exit Sub

Check1_Click_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Check1_Click" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Check1_Click", Erl)
   err.clear
 
End Sub

Private Sub Check2_Click()
   On Error GoTo Check2_Click_Error:
   
 If Check2 = 1 And Val(balqty.Text) > 0 And Val(balqty.Text) < 30000 Then
    ptlen = Val(balqty.Text)
    Check1.value = 1
    TypeOfFlaw = "S"
    TypeOfRemark = "BESCRAP"
    DRReason.Text = "S"
    Remark.Text = "BESCRAP"
    Storedata
 End If
 Check2 = 0
   
   Exit Sub

Check2_Click_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Check2_Click" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Check2_Click", Erl)
   err.clear
End Sub
Private Sub Set_Length_Automatically(Fault As Double)

    Dim RelSample As Boolean
    Dim rs1 As New ADODB.Recordset
    On Error GoTo Set_Length_Automatically_Error:
    Dim CutFlag As Integer
    
    twno = Val(Mid(drawid.Text, 9, 2))
    Go_for_Sample = 1
    CutFlag = 0
    StandardLen = 50.85  ' Default Standard len
    If Mid(drawid.Text, 1, 1) = "K" Or Mid(drawid.Text, 1, 1) = "D" Or Mid(drawid.Text, 1, 1) = "L" Or Mid(drawid.Text, 1, 1) = "Q" Or Mid(drawid.Text, 1, 1) = "A" Then    ''Q TYPE ADDED ON 20.12.2013
       StandardLen = 26   'for K Type Sathis Ram /Awantika 13102010,e type from hrikesh,dof arun 27/11/2012
    Else
        rs1.Open "select * from cutting_logic where  twno='" & twno & "' ", cn, adOpenKeyset, adLockReadOnly
            If rs1.RecordCount > 0 Then
                If plant = "WALUJ" Then
                    If rs1("LENGTH") = 48.8 Then StandardLen = 49.15  '24.65 chg as per PCR PT-24691,By Mohit,2013-06-25,24.75 to 49.15
                    If rs1("LENGTH") = 50.4 Then StandardLen = 50.85    '50.75 to 50.85 as per milind sir mail on 10.03.2014
                    If rs1("LENGTH") = 24.4 Then StandardLen = 24.75  '24.65 chg as per PCR PT-24691,By Mohit,2013-06-25
                    If rs1("LENGTH") = 25.2 Then StandardLen = 25.6     '25.4 to 25.6 as per milind sir mail on 10.03.2014
                    If rs1("LENGTH") = 54.6 Then StandardLen = 59.2    ''added for 58.8 pt len by abdul on 29.01.2014
                    If rs1("LENGTH") = 58.8 Then StandardLen = 59.2    ''added for 58.8 pt len by abdul on 29.01.2014
                Else
                            
                    If rs1("LENGTH") = 48.8 Then StandardLen = 49.3
                    If rs1("LENGTH") = 50.4 Then StandardLen = 50.85    '50.75 to 50.85 as per milind sir mail on 10.03.2014
                    If rs1("LENGTH") = 24.4 Then StandardLen = 24.8
                    If rs1("LENGTH") = 25.2 Then StandardLen = 25.6     '25.4 to 25.6 as per milind sir mail on 10.03.2014
                    If rs1("LENGTH") = 54.6 Then StandardLen = 59.2     ''added for 58.8 pt len by abdul on 29.01.2014
                    If rs1("LENGTH") = 58.8 Then StandardLen = 59.2     ''added for 58.8 pt len by abdul on 29.01.2014
                End If
            End If
        rs1.Close
    End If
        
''''''''''''''''''''''''''''''''''''''''''ADDED BY ABDUL ON 05.11.2015 AS PER MILIND SIR
    Dim rsGrcomp As New ADODB.Recordset
    
    If rsGrcomp.State = 1 Then rsGrcomp.Close
    rsGrcomp.Open "select product_type from draw where preid='" & Mid(drawid.Text, 1, 8) & "'", cn, adOpenKeyset, adLockReadOnly
    If rsGrcomp.RecordCount > 0 Then
      If IsNull(rsGrcomp("PRODUCT_TYPE")) = False Then
          protype = rsGrcomp("PRODUCT_TYPE")
      End If
    End If
    rsGrcomp.Close
    
    If protype = "" Then
    If rsGrcomp.State = 1 Then rsGrcomp.Close
    rsGrcomp.Open "select PRODUCT_TYPE from coredep where preid = '" & Mid(drawid.Text, 1, 7) & "'  AND PRODUCT_TYPE IS NOT NULL ", cn, adOpenKeyset, adLockReadOnly
        If rsGrcomp.RecordCount > 0 Then
            protype = rsGrcomp(0)
        End If
    rsGrcomp.Close
    End If
    If protype = "" Then
    If rsGrcomp.State = 1 Then rsGrcomp.Close
    rsGrcomp.Open "select PROD_TYPE from RIC_PREFORM_ASSEMBLY where preid = '" & Mid(drawid.Text, 1, 8) & "'  AND PROD_TYPE IS NOT NULL ", cn, adOpenKeyset, adLockReadOnly
        If rsGrcomp.RecordCount > 0 Then
            protype = rsGrcomp(0)
        End If
    rsGrcomp.Close
    End If
    
    If plant = "SHENDRA" And protype = "L1" Then
        StandardLen = 24.8
    End If
''''''''''''''''''''''''''''''''''''''''''
     
     If NextFlawPosGlobal < 2 Then
        Go_for_Sample = 0
     End If


     If plant = "WALUJ" Then
         If Mid(txtFlawMsg.Text, 1, 3) = "CUT" Then
                  SetLen = Val(txtCuttingLen.Text)
                  Go_for_Sample = 0
                  CutFlag = 1
         ElseIf Mid(txtFlawMsg.Text, 1, 4) = "NEXT" Then
                If NextFlawPosGlobal < 4.6 Then       '(4.2 + 0.5 +1.5)  < 4.6 no usaable length
                    SetLen = 4.6     ' scope of improvement
                ElseIf NextFlawPosGlobal >= StandardLen Then   ' team will let us know if std len is 25.2 and next flaw is >50.4 still dey want to set 25.2 if not logic will change
                    SetLen = StandardLen
                Else
                    SetLen = (NextFlawPosGlobal) - 0.15
                End If
         Else
            SetLen = StandardLen
            If FlagScratch = True And Val(txtTotPT.Text) = 0 Then
              SetLen = 4.2
            End If
            
            If (Val(balqty.Text) / 1000 - Val(txtBS.Caption) / 1000 <= StandardLen) And (Val(balqty.Text) / 1000 - Val(txtBS.Caption) / 1000 >= 4.2) Then
                SetLen = Val(balqty.Text) / 1000 - Val(txtBS.Caption) / 1000
            End If
         End If
    Else
         If Mid(txtFlawMsg.Text, 1, 3) = "CUT" Then
            SetLen = Val(txtCuttingLen.Text)
            Go_for_Sample = 0
            CutFlag = 1
         Else
            SetLen = StandardLen
            If FlagScratch = True And Val(txtTotPT.Text) = 0 Then
              SetLen = 4.2
            End If
            If (Val(balqty.Text) / 1000 - Val(txtBS.Caption) / 1000 <= StandardLen) And (Val(balqty.Text) / 1000 - Val(txtBS.Caption) / 1000 >= 4.2) Then
                SetLen = Val(balqty.Text) / 1000 - Val(txtBS.Caption) / 1000
            End If
            If NextFlawPosGlobal < 26 Then
                SetLen = 25.6
            ElseIf NextFlawPosGlobal > 38.6 And NextFlawPosGlobal < 50.85 Then
                SetLen = 25.6
            End If
        End If
    End If


    If iPTContBrkFlag = 1 Or iPTContBrkFlag = 2 Then  'SUNIL 17042008
        SetLen = 4.2
        Go_for_Sample = 0 'Reliability Sample Can not be taken
    End If

    If FlagPMD = True Then
        SetLen = 0.22
        Go_for_Sample = 0 'Reliability Sample Can not be taken
        FlagPMD = False
    End If

    '--------------------If no cont. break, no cutting,no ZTPMD then check for Reliability Sample


    If plant = "WALUJ" Then
        If NextFlawPosGlobal > 12.6 Then
            RS.Open "Select prremarks from DRAW  where FID='" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
                If RS.RecordCount > 0 Then
                    If RS("prremarks") = "PREFORM_END" Then
                        rs1.Open "Select OKLEN from PT  WHERE  FID= '" & drawid.Text & "' and OKLEN > 0", cn, adOpenKeyset, adLockReadOnly
                            If rs1.RecordCount = 0 Then
                                SetLen = 12.75
                                Go_for_Sample = 0
                            End If
                        rs1.Close
                    End If
                End If
            RS.Close
        End If
    
             ''    'As per PCR,by Mohit on 29-05-2013
        If Len(PTID.Text) = 12 Then
            If NextFlawPosGlobal > 12.6 Then
                TestFlag = True
                ''**logic written by ankit on 05.02.2014
                If drawid.Text = "" Then Exit Sub
                '''
                rs1.Open "Select * from PT WHERE  PTID like '" & drawid.Text & "%' AND OKLEN>0 AND SUBSTR(PTID, 12, 1) ='A'  ", cn, adOpenKeyset, adLockReadOnly
                    If rs1.RecordCount > 0 Then
                        TestFlag = False
                    Else
                       SetLen = 17      'changed from 12.75 to 17
                       Go_for_Sample = 0
                    End If
                rs1.Close
            End If
        End If
    End If
      '------------------------New code-By Mohit-26-06-2013
     If plant = "WALUJ" Then
          If Len(PTID.Text) = 12 Then
             newID = Mid(PTID.Text, 12, 1)
             If (newID = "A") Or (newID = "B" And finalstatus = True) Then
                 If Val(SetLen) > 26 Then stlen = 25.4
             End If
          End If
          If (Val(balqty) / 1000) < 76 Then
             If Val(SetLen) > 26 Then SetLen = 25.4
          End If

         If SetLen > 38.2 And SetLen < 50.45 Then        ''logic changed from 49.15 to 50.45 as per milind sir mail on 10.03.2014
               If (SetLen - 25.45) >= 12.7 Then
                    SetLen = 25.6                        ''25.4 to 25.6 on 10.03.2014 as per milind sir mail
               End If
         End If
     Else
         If SetLen > 38.2 And SetLen < 50.45 Then        ''logic changed from 49.15 to 50.45 as per milind sir mail on 10.03.2014
             If (SetLen - 25.45) >= 12.7 Then        'changed  from 4.3 to 12.7 on 23.10.2013
                 SetLen = 25.6               ''from 24.8 to 25.45 on 28.02.2014,25.45 to 25.6 on 10.03.2014 as per milind sir
             End If
         End If
     End If
    
    '''----------------------------------------------------------------------------
    '''MsgBox "HOROUS Length shluld be-" & SetLen
    '''-------------------------------------------------------
    
    If plant = "SHENDRA" Then
        If CutFlag = 0 Then
            If balqty.Text = drqty.Text Then
                SetLen = 12.75
                Go_for_Sample = 0
            End If
            If rs1.State = 1 Then rs1.Close
            rs1.Open "Select * from PT WHERE  PTID like '" & drawid.Text & "%' AND OKLEN>0 AND SUBSTR(PTID, 12, 1) ='A'  ", cn, adOpenKeyset, adLockReadOnly
            If rs1.RecordCount > 0 Then
                'Record exists
            Else
               SetLen = 12.75
               Go_for_Sample = 0
            End If
            If rs1.State = 1 Then rs1.Close
        End If
    End If
    '---------------------------------------------------
    If Go_for_Sample = 1 Then
        If Trim(txtBarcodeID.Text) <> "" Then
            If TagDrrej = 0 Then
                RelSample = Check_Sample_For_Reliability(txtBarcodeID.Text, PTID.Text)
                If RelSample = True Then
                    If plant = "WALUJ" Then
                        SetLen = 0.2
                        MsgBox "Please Take 250 mts Sample for Reliability", vbInformation
                    Else
                        SetLen = 0.5        '0.5 to 0.2,0.2 to 0.5
                        MsgBox "Please Take 500 mts Sample for Reliability/DOC testing", vbInformation
                    End If
                End If
            End If
        End If
    End If
  ''==================================================================================================================
  ''CODE ADDED BY dinesh on ref yogesh sir & mayur sir on 20092017 regarding preform end 12.6,12.6,25.2
  ''Old code this can use for handle portion cutting instruction in future
  ''==================================================================================================================
  
''  If protype = "E1" Then
''    If SetLen >= 12.6 Then
''
''    'If NextFlawPosGlobal > 12.6 Then
''    Dim Preformcounter As String
''    Dim rscnt As New ADODB.Recordset
''      If rs.State = 1 Then
''       rs.Close
''       Set rs = Nothing
''      End If
''       rs.Open "Select prremarks from DRAW  where FID='" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
''                If rs.RecordCount > 0 Then
''                    If rs("prremarks") = "PREFORM_END" Then
''                        If rs1.State = 1 Then
''                         rs1.Close
''                         Set rs1 = Nothing
''                        End If
''                    bobbinprefend = 0
''                     rscnt.Open "Select OKLEN from PT  WHERE  FID= '" & drawid.Text & "' and OKLEN > 12", cn, adOpenKeyset, adLockReadOnly
''                     If rscnt.EOF = False Then
''                     bobbinprefend = rscnt.RecordCount
''
''                     End If
''
''
''                   Preformcounter = getfid_counterpostion
''                   If Preformcounter < 50 Then
''                    If bobbinprefend < 2 Then
''
''                         If Val(Val(Preformcounter) + 13) <= 52 Then
''                          SetLen = 13
''                         ' bobbinprefend = bobbinprefend + 1
''                         Else
''                          SetLen = Val(50 - Val(Preformcounter))
''                           ' bobbinprefend = bobbinprefend + 1
''                         End If
''
''
''
''                    End If
''                     If bobbinprefend = 2 Then
''
''                         If Val(Val(Preformcounter) + 25.75) <= 52 Then
''                          SetLen = 25.75
''                          'bobbinprefend = bobbinprefend + 1
''                          Else
''                          SetLen = Val(50 - Val(Preformcounter))
''                           ' bobbinprefend = bobbinprefend + 1
''                         End If
''
''
''                    End If
''                   End If
''
''
''
''                         If rs1.State = 1 Then
''                         rs1.Close
''                         Set rs1 = Nothing
''                        End If
''
''                    End If
''                End If
''
''             If rs.State = 1 Then
''             rs.Close
''             Set rs = Nothing
''            End If
''        End If
''    End If
''
 ''''==================================================
    
     ''==================================================================================================================
  ''CODE ADDED BY dinesh on ref yogesh sir & mayur sir on 25092017 regarding preform end 12.6,12.6,25.2
  ''Old code this can use for handle portion cutting instruction in future
  ''==================================================================================================================
       Dim rscnt As New ADODB.Recordset
     If protype = "E1" Then
      If SetLen >= 12.6 Then
            If Val(txtFactor1.Text) < 100 Then
            If (Val(balqty) / 1000) < 76 Then
              SetLen = 50.85
                  If (Val(balqty) / 1000) < 76 Then
                  If Val(SetLen) > 26 Then SetLen = 25.4
                    If Val(Val((Val(balqty) / 1000)) - 25.4) > 30 Then
                       
                       Else
                        SetLen = 25.75
                    End If
                  End If
                  
                
                    If (Val(balqty) / 1000) <= 50 Then
                     rscnt.Open "Select OKLEN from PT  WHERE  FID= '" & drawid.Text & "' and OKLEN > 4.2 and FACTOR<=50", cn, adOpenKeyset, adLockReadOnly
                    ' If rscnt.EOF = False Then
                     
                     
                        bobbinprefend = rscnt.RecordCount
                                        If bobbinprefend = 0 Then
                                         
                                                    SetLen = 25.75
                                                 If Val(Val((Val(balqty) / 1000)) - Val(25.75) - Val(Val(txtBS.Caption) / 1000)) > 0 Then
                                                   SetLen = 25.75
                                                   
                                                 
                                                 Else
                                                   If Val(Val(Val(balqty) / 1000) - Val(Val(txtBS.Caption) / 1000)) > 0 Then
                                                      SetLen = Val(Val(Val(balqty) / 1000) - Val(Val(txtBS.Caption) / 1000))
                                                    Else
                                                          SetLen = Val(Val(txtBS.Caption) / 1000)
                                                         
                                                    End If
                                                End If
                ''
                                                If Mid(txtFlawMsg.Text, 1, 4) = "NEXT" Then
                                                 ''code newly added by dinesh on ref yojna mam 28092017
                                                        If NextFlawPosGlobal < 4.6 Then       '(4.2 + 0.5 +1.5)  < 4.6 no usaable length
                                                            SetLen = 4.6     ' scope of improvement
                                                        ElseIf NextFlawPosGlobal >= SetLen Then   ' team will let us know if std len is 25.2 and next flaw is >50.4 still dey want to set 25.2 if not logic will change
                                                            SetLen = SetLen
                                                        Else
                                                            SetLen = (NextFlawPosGlobal) - 0.15
                                                        End If
                                                    End If
                                        End If
                                        
                                
                                 If bobbinprefend >= 1 Then
                                   '''code modified by dinesh on ref mayur sir & yojna mam on 05/10/2017
                                     SetLen = 13
                                     If Val(Val((Val(balqty) / 1000)) - Val(13) - Val(Val(txtBS.Caption) / 1000)) > 0 Then
                                       SetLen = 13
                                     
                                     Else
                                      
                                            If Val(Val(Val(balqty) / 1000) - Val(Val(txtBS.Caption) / 1000)) > 0 Then
                                              SetLen = Val(Val(Val(balqty) / 1000) - Val(Val(txtBS.Caption) / 1000))
                                              Else
                                                  SetLen = Val(Val(txtBS.Caption) / 1000)
                                                 
                                            End If
                                    End If
                                    
                                         If Mid(txtFlawMsg.Text, 1, 4) = "NEXT" Then
                                                 ''code newly added by dinesh on ref yojna mam 28092017
                                                        If NextFlawPosGlobal < 4.6 Then       '(4.2 + 0.5 +1.5)  < 4.6 no usaable length
                                                            SetLen = 4.6     ' scope of improvement
                                                        ElseIf NextFlawPosGlobal >= SetLen Then   ' team will let us know if std len is 25.2 and next flaw is >50.4 still dey want to set 25.2 if not logic will change
                                                            SetLen = SetLen
                                                        Else
                                                            SetLen = (NextFlawPosGlobal) - 0.15  '''getting between data
                                                        End If
                                         End If
        ''
                                End If
                        
                         If rscnt.State = 1 Then rscnt.Close
                    ' End If
                    
                     
                    End If
                    
                End If
            End If
        End If
     End If
     
     
          Text3.Text = StandardLen
          Text2.Text = SetLen
          
       '''''CODE ADDED BY DINESH for e2 cuutting length product type as per milind sir 13022018
           ' If UCase(plant) = "WALUJ" Then
              Dim rsCut As New ADODB.Recordset
              If rsCut.State = 1 Then
                rsCut.Close
                Set rsCut = Nothing
               End If
            rsCut.Open "select length,PRODUCT_TYPE from PT_cutting_logic where plant='" & UCase(plant) & "' and PRODUCT_TYPE= '" & UCase(protype) & "' and " & SetLen & ">= MIN_VAL  and " & SetLen & "<= MAX_VAL order by EDATE desc", cn, adOpenDynamic, adLockOptimistic
                If rsCut.EOF = False Then
                       If rsCut(0) <= BOBBINLENGTH Then
                            Text3.Text = StandardLen
                            If IsNull(rsCut(0)) = False Then Text2.Text = rsCut(0)
                            If IsNull(rsCut(0)) = False Then SetLen = rsCut(0)
                        Else

                            Text3.Text = StandardLen
                            Text2.Text = SetLen

                        End If
                   Else
                    Text3.Text = StandardLen
                    Text2.Text = SetLen

                   End If
               If rsCut.State = 1 Then
                rsCut.Close
                Set rsCut = Nothing
               End If
               
'
'            Else
'
'                 Text3.Text = StandardLen
'                 Text2.Text = SetLen
'            End If


'''============================================================

'''============================================================
'''code added by dinesh on ref bajirao sir on 25042018
            If UCase(plant) = "SHENDRA" Then
              Call Ptflawallocation
            End If

'''============================================================
 '''CODE COMMENTED BY DINESH FOR MEDAK TESTING ON  30062018
   '' Set_Horus_Parameters (SetLen)

    If plant = "SHENDRA" Then
        Text1.Visible = True
        Text1.BackColor = vbGreen
        Text1.ForeColor = vbRed
        If SetLen > 26 Then
            Text1.Text = "USE 50KM BOBBIN"
            MsgBox "", vbInformation
        Else
            Text1.Text = "USE 25KM BOBBIN"
            MsgBox "", vbInformation
        End If
        Text1.Text = ""
        Text1.ForeColor = &H80000008
        Text1.Visible = False
    End If
    Exit Sub

Set_Length_Automatically_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Set_Length_Automatically" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Set_Length_Automatically", Erl)
   err.clear
End Sub

Function getfid_counterpostion()
Dim rsdk As New ADODB.Recordset
Dim rsdk1 As New ADODB.Recordset
Dim rsdr As New ADODB.Recordset
Dim rsTot As New ADODB.Recordset

Dim CntTotal As Single, limit1 As Single, limit2 As Single, drlen As Single
StDate = getcurdate
drawid.Text = UCase$(Trim(drawid.Text))


rsTot.Open "select fid from draw where fid like '" & drawid.Text & "%'", cn, adOpenKeyset, adLockReadOnly
Do While rsTot.EOF = False
rsdk1.Open "select drlen from draw where fid = '" & rsTot(0) & "'", cn, adOpenKeyset, adLockReadOnly
If rsdk1.RecordCount > 0 Then
    drlen = rsdk1(0)
End If
rsdk1.Close

st = "select * from PT where fid =  '" & rsTot(0) & "' ORDER BY prodDATE"
st1 = "select * from drawflaw where fid = '" & rsTot(0) & "' order by position desc"
'Screen.MousePointer = vbHourglass
rsdr.Open st1, cn, adOpenKeyset, adLockReadOnly
rsdk1.Open "SELECT * FROM MATMOVE WHERE ID LIKE '" & rsTot(0) & "%' AND PROCESS = 'PT' AND MOVETYPE = '101'", cn, adOpenKeyset, adLockReadOnly

Tag = 1
If rsdr.EOF = True Then Tag = 0


rsdk.Open st, cn, adOpenKeyset, adLockReadOnly
limit1 = drlen
limit2 = drlen
    
   

If rsdk.RecordCount > 0 Then
rsdk.MoveFirst
End If
CntTotal = 0
For i = 0 To rsdk.RecordCount - 1
  
   CNT = 1
   
   If IsNull(rsdk("DRREJ1")) = False Then drrej2 = rsdk("DRREJ1")
   drrej1 = rsdk("DRREJ") - drrej2
   'For draw rejection BEFORE ok LENGTH
   If drrej1 > 0 Then
   
   
   CntTotal = CntTotal + drrej1
   
    If CNT = 1 Then
       
    Else
       
    End If
    If rsdk1.EOF = False Then
    rsdk1.MoveFirst
    For k = 0 To rsdk1.RecordCount - 1
    If rsdk1("id") = rsdk("PTID") Then
  
  
    Exit For
    End If
    rsdk1.MoveNext
    Next k
    End If

    
   CNT = CNT + 1
   
   

   limit1 = drlen - CntTotal + drrej1
   limit2 = drlen - CntTotal
   Do While rsdr.EOF = False
   If (rsdr(2) >= limit2 And rsdr(2) <= limit1) Then
    
     rsdr.MoveNext
    
   Else
     Exit Do
   End If
   Loop

   
   End If

   If rsdk("scrap1") > 0 Then
  
   CntTotal = CntTotal + rsdk("scrap1")
    S1 = S1 & CntTotal & vbTab
    If CNT = 1 Then
      
    Else
       
    End If
   
    
    If rsdk1.EOF = False Then
    rsdk1.MoveFirst
    For k = 0 To rsdk1.RecordCount - 1
    If rsdk1("id") = rsdk("PTID") Then
   
    If rsdk("Process") = "PT" Then
       
    ElseIf rsdk("Process") = "PT_MAN" Then
       
    End If
    Exit For
    End If
    rsdk1.MoveNext
    Next k
    End If
    
   CNT = CNT + 1
   
    
   limit1 = drlen - CntTotal + rsdk("scrap1")
   limit2 = drlen - CntTotal
   Do While rsdr.EOF = False
    If (rsdr(2) >= limit2 And rsdr(2) <= limit1) Then
     
     rsdr.MoveNext
    Else
     Exit Do
    End If
   Loop

  
   End If
   'For OK LENgth
   If rsdk("oklen") > 0 Then
 
   CntTotal = CntTotal + rsdk("oklen")
   
    If CNT = 1 Then
       
    Else
        
    End If
   CNT = CNT + 1
   Dim fid1 As String
   fid1 = rsdk("ptid")
   rsbar.Open "select barcodeid from ptbarcode where fid='" & fid1 & "'", cn, adOpenDynamic, adLockReadOnly
   If rsbar.EOF = False Then
       ' txtBarcode = rsbar(0)
   End If
   rsbar.Close
   
    
    rsdk1.MoveFirst
    For k = 0 To rsdk1.RecordCount - 1
    If rsdk1("id") = rsdk("PTID") Then
  
    If rsdk("Process") = "PT" Then
   
    ElseIf rsdk("Process") = "PT_MAN" Then
      
    End If
    Exit For
    End If
    rsdk1.MoveNext
    Next k
  

       
   limit1 = drlen - CntTotal + rsdk("oklen")
   limit2 = drlen - CntTotal
   Do While rsdr.EOF = False
   If rsdr(2) >= limit2 And rsdr(2) <= limit1 Then
    
     rsdr.MoveNext
   Else
     Exit Do
   End If
   Loop

    
   End If
   'For draw rejection AFTER ok LENGTH
   If drrej2 > 0 Then
   
  
   CntTotal = CntTotal + drrej2
   
    If CNT = 1 Then
     
    Else
     
    End If
    If rsdk1.EOF = False Then
    rsdk1.MoveFirst
    For k = 0 To rsdk1.RecordCount - 1
    If rsdk1("id") = rsdk("PTID") Then
   
    If rsdk("Process") = "PT" Then
      
    ElseIf rsdk("Process") = "PT_MAN" Then
      
    End If
    Exit For
    End If
    rsdk1.MoveNext
    Next k
    End If
    
   
 
   CNT = CNT + 1
   limit1 = drlen - CntTotal + drrej2
   limit2 = drlen - CntTotal
   Do While rsdr.EOF = False
   If (rsdr(2) >= limit2 And rsdr(2) <= limit1) Then
   
     rsdr.MoveNext
    
   Else
     Exit Do
   End If
   Loop

   
   End If
   
   
   'For Scrap2 (after ok length)
   If rsdk("scrap2") > 0 Then

   CntTotal = CntTotal + rsdk("scrap2")
   
    If CNT = 1 Then
    
    Else
       
    End If
   
     
    If rsdk1.EOF = False Then
    rsdk1.MoveFirst
    For k = 1 To rsdk1.RecordCount - 1
    If rsdk1("id") = rsdk("PTID") Then
    
    Exit For
    End If
    rsdk1.MoveNext
    Next k
    End If
     
   CNT = CNT + 1
   
   
   limit1 = drlen - CntTotal + rsdk("scrap2")
   limit2 = drlen - CntTotal
   Do While rsdr.EOF = False
   If (rsdr(2) >= limit2 And rsdr(2) <= limit1) Then
    
     rsdr.MoveNext
   Else
    Exit Do
   End If
   Loop

   
   End If
   
rsdk.MoveNext
Next i
rsdk1.Close
rsdk.Close
rsdr.Close
rsTot.MoveNext
Loop
rsTot.Close

getfid_counterpostion = CntTotal
End Function

Private Sub checkflawnew_Click()
   chkFlaw.value = checkflawnew.value
End Sub

Private Sub chkMERej_Click()
    If chkMERej.value = 1 Then
        DRReason.Text = "M"
        TypeOfFlaw = "M"
        TypeOfRemark = "ME"
        ChkScratch.value = 0
        Check1.value = 1
        cmbmutiend.Visible = True
        LBLMULTIEND.Visible = True
    Else
        Check1.value = 0
        DRReason.Text = ""
        cmbmutiend.Visible = False
        LBLMULTIEND.Visible = False
    End If
End Sub

Private Sub ChkScratch_Click()
    If ChkScratch.value = 1 Then
        DRReason.Text = "D"
        TypeOfFlaw = "D"
        TypeOfRemark = "SCR"
        chkMERej.value = 0
        Check1.value = 1
    Else
        Check1.value = 0
        DRReason.Text = ""
    End If

End Sub
Private Sub cmdAllocatedID_Click()
    PTLoginRecFlag = False
    Call PTAllocation
    PTLoginRecFlag = True
    
    Dim RS As New ADODB.Recordset
    RS.Open "SELECT PDATE FROM DRAW WHERE FID = '" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            If Format(RS(0), "yyyy-mm-dd hh:mm:ss") <= "2010-11-01 11:00:00" Then
                RS.Close
                Exit Sub
            End If
        End If
    RS.Close
    
    '''''''''''''added by abdul on 20.03.2014
    Dim rs12 As New ADODB.Recordset
    If plant = "SHENDRA" Then
    rs12.Open "select * from coredep where preid='" & Mid(PTID.Text, 1, 7) & "'", cn, adOpenKeyset, adLockReadOnly
    nrec = rs12.RecordCount
    If nrec > 0 Then
        If IsNull(rs12("PRODUCT_TYPE_E")) = False Then X = rs12("PRODUCT_TYPE_E")
        If X = "E2" Then
            MsgBox "This is 200 micron preform its PT Line speed should be 1800 m/min.Please contact Shift Incharge. ", vbCritical
        End If
    End If
    rs12.Close
    End If
    
    
    RS.Open "SELECT * FROM ora_global_params WHERE PARAMETER = 'CHECK_PMD_TESTING'", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
    Check_PMd_Testing = RS("value")
    End If
    RS.Close
    
    If Check_PMd_Testing = 1 Then
        If drawid.Text <> "" Then
            If rsdate.State = 1 Then rsdate.Close
            rsdate.Open "select SYSDATE-5 FROM DUAL", cn
            fromdate = Format(rsdate(0), "yyyy-mm-dd 00:00:00")
            rsdate.Close
            TowerNo = Val(Mid(drawid.Text, 9, 2))
            RS.Open "Select * from draw where pdate  >= '" & fromdate & "' and twno = '" & TowerNo & "' order by pdate desc", cn, adOpenKeyset, adLockReadOnly
            If RS.RecordCount > 0 Then
                For i = 0 To RS.RecordCount - 1
                      If RS("fid") = drawid.Text Then
                          RS.MoveNext
                          If RS.EOF = False Then RS.MoveNext '3rd no spool
                          If RS.EOF = False Then RS.MoveNext '3rd no spool
                          If RS.EOF = False Then Prevspool = RS("FID")
                          RS.Close
                          Exit For
                      Else
                          RS.MoveNext
                      End If
                Next i
            Else
                RS.Close
                Exit Sub
            End If
              
            If Prevspool = "" Then
              Exit Sub
            End If
        
            'Check for Samples in PT Sample Table
            RS.Open "select count(*) from pt_inst_ztpmd where substr(fid,1,11) = '" & Prevspool & "'", cn, adOpenKeyset, adLockReadOnly
            If RS.RecordCount > 0 Then
                ptSample = RS(0)
            Else
                MsgBox "No Samples have been collected from spool " & Prevspool
                RS.Close
                Exit Sub
            End If
            RS.Close
        
          'List of ZTPMD IDs
          IDStr = ""
          RS.Open "select * from pt_inst_ztpmd where substr(fid,1,11) = '" & Prevspool & "'", cn, adOpenKeyset, adLockReadOnly
          If RS.RecordCount > 0 Then
              While RS.EOF = False
                  If IDStr = "" Then
                      IDStr = RS("ZTPMDID")
                  Else
                      IDStr = IDStr & "," & RS("ZTPMDID")
                  End If
                  RS.MoveNext
              Wend
          End If
          RS.Close
      
        
            'Check for Testing Records
'620               RS.Open "SELECT FID,PMD1550 FROM QADATA WHERE substr(FID,1,11) = '" & PrevSpool & "' AND PMD1550 > 0", cn, adOpenKeyset, adLockReadOnly
'630               cntPMD = 0
'640               If RS.RecordCount > 0 Then
'650                   While RS.EOF = False
'660                       cntPMD = cntPMD + 1
'670                       RS.MoveNext
'680                   Wend
'690               End If
'700               RS.Close
        
              RS.Open "SELECT * FROM ZTPMD_COIL_ENTRY where substr(fid,1,11) = '" & Prevspool & "'", cn, adOpenDynamic, adLockOptimistic
              cntPMD = 0
              If RS.RecordCount > 0 Then
                      While RS.EOF = False
                          cntPMD = cntPMD + 1
                          RS.MoveNext
                      Wend
                 End If
              RS.Close

            If (ptSample - cntPMD) > 1 Then
                'msg = "PMD Sample of Spool " & PrevSpool & " not tested !" & vbCrLf
                msg = "Coil Entry of Spool " & Prevspool & " not Done  !" & vbCrLf
                msg = msg & IDStr & vbCrLf
                msg = msg & "Total Sample " & ptSample & " , Coil entry " & cntPMD
                MsgBox msg, vbCritical
                drawid.Text = ""
                Unload Me
            End If
        End If
    End If
    
    'here Added Code as lock for recommend by uvk and abhishek on 09032017
   ' If plant = "SHENDRA" Then
        Dim rscheck As New Recordset
            If rscheck.State = 1 Then
                rscheck.Close
                Set rscheck = Nothing
            End If
        
        rscheck.Open "Select PRODUCT_TYPE from Draw where FID='" & Mid(PTID.Text, 1, 11) & "'", cn, adOpenKeyset, adLockReadOnly
        If rscheck.EOF = False Then
         If IsNull(rscheck("PRODUCT_TYPE")) = False Then
          
                txtProducttype.Text = rscheck("PRODUCT_TYPE")
            Else
            
            MsgBox "Product Type Not found for this FID.Please Check Product Type"
            End If
           ' txtProducttype.Text = RSCHECK("PRODUCT_TYPE")
        End If
                If rscheck.State = 1 Then
                    rscheck.Close
                    Set rscheck = Nothing
                End If
     'End If
    
    ''''''''''''''''''
    
End Sub
Private Sub cmdexit_Click()
    Unload Me
End Sub
Private Sub cmdMELen_Click()
If Check1.value = 1 And DRReason.Text = "M" Then
    frmPTMultipleEndEntry.Show vbModal
Else
    MsgBox "Select multiple end drrejection check box", vbExclamation
End If
End Sub

Private Sub cmdOutScanPending_Click()
Dim RS As New ADODB.Recordset
lstOutScanPending.clear
        If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
RS.Open "Select BARCODEID,LENGTH,PDATE,POPR from PTBARCODE WHERE  PDATE>=SYSDATE-60 AND STATUS='O' AND STAGE IS NULL AND SUBSTR(FID,1,1) NOT IN ('M','R','D','A','N') and LENGTH >=4200 AND MCNO='" & ptno.Text & "' order by pdate", cn, adOpenKeyset, adLockReadOnly
Do While RS.EOF = False
  lstOutScanPending.AddItem RS(0) & " " & vbTab & RS(1) & vbTab & RS(2) & vbTab & RS(3)
  RS.MoveNext
Loop
        If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If

End Sub

Private Sub cmdSAP_Click()

Dim RS As New ADODB.Recordset
Dim rsdate As New ADODB.Recordset
   On Error GoTo cmdSAP_Click_Error:
'MsgBox "This Is Pravin SAP Click"
Process_Type = Mid(drawid.Text, 1, 1)

If Process_Type <> "M" And Process_Type <> "R" And Process_Type <> "D" And Process_Type <> "N" And Process_Type <> "A" And Process_Type <> "Q" Then
    MINLEN_NEW = 4200
Else
    MINLEN_NEW = Minlen
End If
If Val(ptlen.Text) < MINLEN_NEW Then
    If Check1.value = 1 Or Check2.value = 1 Then
        SAP_Batch_Status = "DRREJ"
        SAP_qty = Round(Val(drrej) / 1000, 3) + Round(Val(txtBrkCntLen.Text) / 1000, 3) + Round(Val(scrqty) / 1000, 3)     ''ADDED 21-3-14
        SAP_Header_qty = Round(Val(drrej) / 1000, 3) + Round(Val(txtBrkCntLen.Text) / 1000, 3) + Round(Val(scrqty) / 1000, 3) 'Inclusion of BrkCounter Lenght  ''ADDED 21-3-14

    Else
        SAP_Batch_Status = "PTSCRAP"
        SAP_qty = Val(scrqty) / 1000
        'SAP_Header_qty = Val(scrqty) / 1000
        SAP_Header_qty = Round(Val(scrqty) / 1000, 3) + Round(Val(txtBrkCntLen.Text) / 1000, 3) 'Inclusion of BrkCounter Lenght
        '--------Added on 20.03.2007
        If Val(ptlen.Text) >= Minlen And Val(ptlen.Text) <= MINLEN_NEW Then
            SAP_qty = Val(ptlen.Text) / 1000
            'SAP_Header_qty = Val(ptlen.Text) / 1000
            SAP_Header_qty = Round(Val(ptlen.Text) / 1000, 3) + Round(Val(txtBrkCntLen.Text) / 1000, 3) 'Inclusion of BrkCounter Lenght
        End If
        '--------Added on 20.03.2007 ----
        
    End If
Else
    If Check1.value = 1 Or Check2.value = 1 Then
        SAP_Batch_Status = "DRREJ"
        SAP_qty = Round(Val(drrej) / 1000, 3) + Round(Val(txtBrkCntLen.Text) / 1000, 3) + Round(Val(scrqty) / 1000, 3)     ''ADDED 21-3-14
        SAP_Header_qty = Val(drrej) / 1000 + Val(txtBrkCntLen.Text) / 1000 + Round(Val(scrqty) / 1000, 3) 'Inclusion of BrkCounter Lenght ''ADDED 21-3-14
    Else
        SAP_Batch_Status = "OK"
        SAP_qty = Val(ptlen.Text) / 1000
        'SAP_Header_qty = Val(ptlen.Text) / 1000
        SAP_Header_qty = Round(Val(ptlen.Text) / 1000, 3) + Round(Val(txtBrkCntLen.Text) / 1000, 3)  'Inclusion of BrkCounter Lenght
    End If
End If

If Val(ptlen.Text) >= MINLEN_NEW Then
    SAP_Batch_Status = "OK"
    SAP_qty = Val(ptlen.Text) / 1000
    'SAP_Header_qty = Val(ptlen.Text) / 1000
    SAP_Header_qty = Round(Val(ptlen.Text) / 1000, 3) + Round(Val(txtBrkCntLen.Text) / 1000, 3) + Round(Val(scrqty) / 1000, 3) 'Inclusion of BrkCounter Lenght      ''added scrqty by nirja 3314
End If


Production_Stage = "PT"
SAP_BatchID = UCase$(txtBarcodeID.Text)
SAP_BatchId_Converted = UCase$(txtBarcodeID.Text)
SAP_Batch_Mcno = ptno.Text
SAP_Header_BatchID = UCase$(drawid.Text)

If plant = "SHENDRA" Then
    Call SAP_Mat_Master_SEZ(Production_Stage, UCase$(Trim(drawid.Text)))
Else
    Call SAP_Mat_Master(Production_Stage, UCase$(Trim(drawid.Text)))
End If


plant = Get_Plant(drawid.Text)

If plant = "SHENDRA" Then
    frmSAPMatMove_Sez.Show vbModal
Else
    frmSAPMatMove.Show vbModal
End If

   
Exit Sub

cmdSAP_Click_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "cmdSAP_Click" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "cmdSAP_Click", Erl)
   err.clear
End Sub

Private Sub cmdShow_Click()
    If drawid.Text <> "" Then
        frmDisplayPTLog.Show
    End If
End Sub

Private Sub cmdUnloadME_Click()
Dim Ans As VbMsgBoxResult
Dim RS As New ADODB.Recordset
Dim Curdate
   On Error GoTo cmdUnloadME_Click_Error:
   
If drawid.Text <> "" Then
    Ans = MsgBox("Are you want to unload main spool for Multiple end", vbYesNoCancel, "Confirmation")
    If Ans = vbYes Then
        If Mid(drawid.Text, 1, 1) <> "R" And Mid(drawid.Text, 1, 1) <> "M" Then
            RS.Open "Select sysdate from dual", cn, adOpenKeyset, adLockReadOnly
            If RS.EOF = False Then
                Curdate = RS(0)
            End If
            RS.Close
            RS.Open "select * from mainspoolrejection WHERE SPOoLID='" & drawid.Text & "'", cn, adOpenDynamic, adLockOptimistic
            If RS.EOF = True Then
                RS.AddNew
                RS("SPOOLID") = UCase$(drawid.Text)
                RS("LENGTH") = Val(drqty.Text)
                RS("edate") = Curdate
                RS("Operator") = UCase$(opr.Text)
                RS("INSPECTED_By") = ""
                RS("Reason") = "MULTIPLE END"
                RS("Remark") = "MULTIPLE END"
                RS.update
            End If
            RS.Close
            RS.Open "Select * from pt_allocation where spool_id='" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
            If RS.EOF = False Then
                If IsNull(RS("UNLOAD")) = False Then
                    If RS("UNLOAD") = "T" Then
                        cn.Execute "Update pt_Allocation set mcno='0',ALLOCATE='S',UNLOAD='T',UNLOAD_REASON='MULTIPLE END',UNLOAD_TIME=sysdate,Unload_by='" & Mid(opr.Text, 1, 10) & "' where spool_id='" & drawid.Text & "'"
                        MsgBox "Spool is Re-Unloaded Sucessfully", vbInformation
                        'MsgBox "Spool is already unloaded previously,You can't unloded same spool twise", vbCritical
                        RS.Close
                        Exit Sub
                    Else
                        cn.Execute "Update pt_Allocation set mcno='0',ALLOCATE='H',UNLOAD='T',UNLOAD_REASON='MULTIPLE END',UNLOAD_TIME=sysdate,Unload_by='" & Mid(opr.Text, 1, 10) & "' where spool_id='" & drawid.Text & "'"
                        MsgBox "Spool is unloaded Sucessfully", vbInformation
                        drawid.Text = ""
                    End If
                End If
            End If
            RS.Close
        Else
            MsgBox "You cann't unload Mulitmode fibre here", vbExclamation
        End If
    End If
End If
   
   Exit Sub

cmdUnloadME_Click_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "cmdUnloadME_Click" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "cmdUnloadME_Click", Erl)
   err.clear
End Sub

Private Sub Command1_Click()
    ' DATA VALIDATION OVER
    ' NOW GENERATE NEW FIBRE ID
   On Error GoTo Command1_Click_Error:
   
   
    
      CNT = 0
If drawid.Text <> "" Then
      RS.Open "select * from MATSTOCK where PROCESS = 'PT' AND ID LIKE " & "'" & drawid.Text & "%' order by id", cn, adOpenKeyset, adLockReadOnly
        Do While RS.EOF = False
            CNT = CNT + 1
            ID = RS("id")
            RS.MoveNext
        Loop
      RS.Close

      RS.Open "select * from MATSTOCK where PROCESS = 'PTR' AND ID LIKE " & "'" & drawid.Text & "Y%' order by id", cn, adOpenKeyset, adLockReadOnly
        Do While RS.EOF = False
            CNT = CNT + 1
            ID = RS("id")
            RS.MoveNext
        Loop
      RS.Close
      RS.Open "select * from MATSTOCK where PROCESS = 'PTR' AND ID LIKE " & "'" & drawid.Text & "Z%' order by id", cn, adOpenKeyset, adLockReadOnly
        Do While RS.EOF = False
            CNT = CNT + 1
            ID = RS("id")
            RS.MoveNext
        Loop
      RS.Close

        If CNT >= 40 Then
            PTID = drawid & "Z" & Chr(CNT + 65 - 40)
        ElseIf CNT >= 24 Then
            PTID = drawid & "Y" & Chr(CNT + 65 - 24)
        Else
            PTID = drawid & Chr(CNT + 65)
        End If
        
    'here Added Code as lock for recommend by uvk and abhishek on 09032017
    If UCase(plant) = "SHENDRA" Or UCase(plant) = "WALUJ" Then
        Dim rscheck As New Recordset
                 If rscheck.State = 1 Then
                    rscheck.Close
                    Set rscheck = Nothing
                 End If
        rscheck.Open "Select PRODUCT_TYPE from Draw where FID='" & Mid(PTID.Text, 1, 11) & "'", cn, adOpenKeyset, adLockReadOnly
        If rscheck.EOF = False Then
         If IsNull(rscheck("PRODUCT_TYPE")) = False Then
          
                txtProducttype.Text = rscheck("PRODUCT_TYPE")
            Else
            
            MsgBox "Product Type Not found for this FID.Please Check Product Type"
            End If
        End If
                If rscheck.State = 1 Then
                    rscheck.Close
                    Set rscheck = Nothing
                End If
     End If
    
    ''''''''''''''''''
   
End If
   
   Exit Sub

Command1_Click_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Command1_Click" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Command1_Click", Erl)
   err.clear
End Sub



Private Sub Command23_Click()

End Sub

Private Sub Command2_Click()
    GenerateSampleID "PB830AI4060O", 25200
End Sub


Private Sub DRReason_LostFocus()
   On Error GoTo DRReason_LostFocus_Error:
   
    Tag = 0
    DRReason.Text = UCase$(DRReason.Text)
    If DRReason <> "" Then
          For i = 0 To DRReason.ListCount - 1
              If DRReason.Text = DRReason.List(i) Then
                  Tag = 1
                  Exit For
              End If
          Next i
          If Tag = 0 Then
              DRReason.Text = ""
              MsgBox "Draw Rejection not correct", vbExclamation
          End If
    End If
   
   Exit Sub

DRReason_LostFocus_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "DRReason_LostFocus" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "DRReason_LostFocus", Erl)
   err.clear
End Sub

Private Sub grdInit()
   On Error GoTo grdInit_Error:
   
grdFlaw.Rows = 14
grdFlaw.Cols = 7
grdFlaw.RowHeight(0) = 15
grdFlaw.Column(0).Width = 12
grdFlaw.Column(1).Width = 25
grdFlaw.Column(2).Width = 45
grdFlaw.Column(3).Width = 45
grdFlaw.Column(4).Width = 45
grdFlaw.Column(5).Width = 45
grdFlaw.Column(6).Width = 25
grdFlaw.Cell(0, 2).Text = "Reason"
grdFlaw.Cell(0, 3).Text = "From"
grdFlaw.Cell(0, 4).Text = "To"
grdFlaw.Cell(0, 5).Text = "Scrap"
If FSYS.FolderExists("C:\ICON") = False Then
    FSYS.CreateFolder ("C:\Icon")
    FSYS.CopyFolder "\\" & ORACLEPATH & "\production\icon", "C:\icon", True
Else
    If FSYS.FileExists("C:\icon\Unchecked.bmp") = False Or FSYS.FileExists("C:\icon\checked.bmp") = False Then
        FSYS.CopyFolder "\\" & ORACLEPATH & "\production\icon", "C:\icon", True
    End If
End If
grdFlaw.Images.Add "C:\icon\" & "UnChecked.bmp", "UnChecked"
grdFlaw.Images.Add "C:\icon\" & "Checked.bmp", "Checked"
grdFlaw.Images.Add "C:\icon\" & "CrossChecked.bmp", "CrossChecked"
For i = 1 To grdFlaw.Rows - 1
    grdFlaw.Cell(i, 1).SetImage "UnChecked"
Next i
For i = 1 To grdFlaw.Rows - 1
    grdFlaw.Cell(i, 6).SetImage "UnChecked"
Next i
With grdFlaw.Range(1, 1, grdFlaw.Rows - 1, 1)
    .Alignment = cellCenterCenter
    .BackColor = vbButtonFace
End With
With grdFlaw.Range(1, 6, grdFlaw.Rows - 1, 6)
    .Alignment = cellCenterCenter
    .BackColor = vbButtonFace
End With
'grdFlaw.AutoRedraw = True
grdFlaw.ReadOnly = True
 '===========================================================================
''===================================================For silic-bY Mohit-2013-07-19
If plant = "WALUJ" Then
    Grid3.Visible = True
    Grid3.Rows = 14
    Grid3.Cols = 5
    Grid3.RowHeight(0) = 15
    Grid3.Column(0).Width = 12
    Grid3.Column(1).Width = 45
    Grid3.Column(2).Width = 45
    Grid3.Column(3).Width = 45
    Grid3.Column(4).Width = 45
    'Grid3.Column(5).Width = 45
    'Grid3.Column(6).Width = 25
    Grid3.Cell(0, 1).Text = "Reason"
    Grid3.Cell(0, 2).Text = "From"
    Grid3.Cell(0, 3).Text = "To"
'Grid3.Cell(0, 5).Text = "Scrap"
'If FSYS.FolderExists("C:\ICON") = False Then
'    FSYS.CreateFolder ("C:\Icon")
'    FSYS.CopyFolder "\\" & ORACLEPATH & "\production\icon", "C:\icon", True
'Else
'    If FSYS.FileExists("C:\icon\Unchecked.bmp") = False Or FSYS.FileExists("C:\icon\checked.bmp") = False Then
'        FSYS.CopyFolder "\\" & ORACLEPATH & "\production\icon", "C:\icon", True
'    End If
'End If
'Grid3.Images.Add "C:\icon\" & "UnChecked.bmp", "UnChecked"
'Grid3.Images.Add "C:\icon\" & "Checked.bmp", "Checked"
'Grid3.Images.Add "C:\icon\" & "CrossChecked.bmp", "CrossChecked"
'For i = 1 To Grid3.Rows - 1
'    Grid3.Cell(i, 1).SetImage "UnChecked"
'Next i
'For i = 1 To Grid3.Rows - 1
'    Grid3.Cell(i, 6).SetImage "UnChecked"
'Next i
'With Grid3.Range(1, 1, Grid3.Rows - 1, 1)
'    .Alignment = cellCenterCenter
'    .BackColor = vbButtonFace
'End With
'With Grid3.Range(1, 6, Grid3.Rows - 1, 6)
'    .Alignment = cellCenterCenter
'    .BackColor = vbButtonFace
'End With
'grid3.AutoRedraw = True
Grid3.ReadOnly = True

''ADDED BY ABDUL ON 13.09.2013 FOR AIRLINE

    Grid4.Visible = True
    Grid4.Rows = 14
    Grid4.Cols = 5
    Grid4.RowHeight(0) = 15
    Grid4.Column(0).Width = 12
    Grid4.Column(1).Width = 45
    Grid4.Column(2).Width = 45
    Grid4.Column(3).Width = 45
    Grid4.Column(4).Width = 45
    Grid4.Cell(0, 1).Text = "Reason"
    Grid4.Cell(0, 2).Text = "From"
    Grid4.Cell(0, 3).Text = "To"
    Grid4.ReadOnly = True



'====================================================================================
End If
     
''for clad (BCSL REMARK) MARKING ON 05122015,FOR WALUJ ON 05.01.2016
'If plant = "SHENDRA" Then
    Grid5.Visible = True
    Grid5.Rows = 14
    Grid5.Cols = 5
    Grid5.RowHeight(0) = 15
    Grid5.Column(0).Width = 12
    Grid5.Column(1).Width = 45
    Grid5.Column(2).Width = 45
    Grid5.Column(3).Width = 45
    Grid5.Column(4).Width = 45
    Grid5.Cell(0, 1).Text = "Reason"
    Grid5.Cell(0, 2).Text = "From"
    Grid5.Cell(0, 3).Text = "To"
    Grid5.ReadOnly = True
'End If

   Exit Sub

grdInit_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "grdInit" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "grdInit", Erl)
   err.clear
End Sub

Private Sub Form_Load()
 'chkFlaw.value = 1
bobbinprefend = 0
Dim RS As New ADODB.Recordset
Dim FSYS As New Scripting.FileSystemObject
If Mid(GetIPAddresses, 8, 2) = 12 Or Mid(GetIPAddresses, 8, 2) = 13 Then
    plant = "SHENDRA"
Else
    plant = "WALUJ"
End If

  plant = GET_IP_PLANT
  
BOBBINLENGTH = 52
If rs1.State = 1 Then
rs1.Close
Set rs1 = Nothing
End If
rs1.Open "select * from TBL_RETRIVAL_TIMELOSS_STATUS where TYPE='BOBBINLENGTH'", cn, adOpenDynamic, adLockOptimistic

If rs1.EOF = False Then

If IsNull(rs1("DESCRIPTION")) = False Then BOBBINLENGTH = rs1("DESCRIPTION")
End If


If rs1.State = 1 Then
rs1.Close
Set rs1 = Nothing
End If





''Automatic Mc. No
On Error Resume Next
Call grdInit
Text1.Height = 6615
Text1.Left = 240
Text1.Top = 360
Text1.Width = 11895
txtCoatType.Text = ""

If plant = "" Then
    If Mid(GetIPAddresses, 8, 2) = 12 Or Mid(GetIPAddresses, 8, 2) = 13 Then
        plant = "SHENDRA"
    Else
        plant = "WALUJ"
    End If
    plant = GET_IP_PLANT
End If



With ptno
 If plant = "SHENDRA" Then
''    For i = 41 To 47
''        .AddItem i
''    Next i
''    For i = 52 To 58
''        .AddItem i
''    Next i
    .Enabled = False
    PT = Mid(GetCompName, 8, 2)
    If (Val(PT) >= 41 And Val(PT) <= 47) Or (Val(PT) >= 52 And Val(PT) <= 60) Then
        .AddItem Val(PT)
        .Text = Val(PT)
    Else
        MsgBox "PT entry has been allowed only PT Machine PC", vbCritical
    End If
Else
    CHK = FSYS.FileExists("c:\ptno.txt")
    If CHK = True Then
        Set INSTREAM = FSYS.OpenTextFile("c:\ptno.txt", ForReading, True)
        If err Then
            err.clear
            INSTREAM.Close
        Else
            PT = Val(INSTREAM.ReadLine)
            If err Then
                err.clear
            End If
            INSTREAM.Close
        End If
    Else
        MsgBox "File no c:\ptno.text doesn't exist", vbCritical
    End If
    
    If Val(PT) > 0 Then
        ptno.Text = Val(PT)
    Else
        MsgBox "Please check c:\ptno.text content"
    End If


    For i = 1 To 25
      .AddItem i
    Next i
      .AddItem 32
      .AddItem 33
      .AddItem 40
 End If
End With

 cboColor.Text = "BLUE"
 opr = uname


'''''code added by navnath divekar on dt 2018-05-25
'Set rs = Nothing
'rs.Open "Select sysdate from dual", cn, adOpenKeyset, adLockReadOnly
'lblDateDiff.Caption = Round(((Date + Time) - rs("sysdate")) * 60 * 24)
'rs.Close
'cboShiftIncharge.clear
'Set rs = Nothing
'rs.Open "select OPID,sysdate from OPID_NEW where dept = 'PT' and SUBDEPT = 'PT' and Status='A' and plant='" & plant & "'", cn, adOpenKeyset, adLockReadOnly
'    Do While rs.EOF = False
'        cboShiftIncharge.AddItem rs(0)
'        rs.MoveNext
'    Loop
'rs.Close
'If rs.State = 1 Then rs.Close
'Set rs = Nothing
'rs.Open "select * from FORM_ACCESS_MASTER where dept = 'PT' and FRM_NAME = 'PT AUTOMATION' and FRAME_NO='cmdMELen' and empid='" & uname & "' and plant='" & plant & "'", cn, adOpenKeyset, adLockReadOnly
'    If rs.RecordCount > 0 Then
'        cmdMELen.Enabled = True
'    End If
'rs.Close
'''''End of code added by navnath divekar on dt 2018-05-25








If plant = "WALUJ" Then
''    txtBobType.AddItem "EKDANT"
''    txtBobType.AddItem "GAMA"
''    txtBobType.AddItem "CRELLIN"
''    txtBobType.AddItem "NEWGAMA"
''    txtBobType.AddItem "MERCY50"
''    txtBobType.AddItem "EKDANT 25"
   ' txtBobType.AddItem "CRELLIN OLD 50"
''    txtBobType.AddItem "CRELLIN NEW 25"
    txtBobType.AddItem "CRELLIN NEW 50"
    txtBobType.AddItem "25 Km BOBBIN"
Else
    txtBobType.AddItem "EKDANT"
    txtBobType.AddItem "CRELLIN"
    txtBobType.AddItem "GAMA"
    txtBobType.AddItem "REUSE-GAMA"
    txtBobType.AddItem "REUSE-CRELLIN"
    txtBobType.AddItem "REUSE-EKDANT"
End If


'On Error GoTo CHK

If FSYS.FolderExists("C:\GFP\LOG") = True Then
    McPath = "C:\GFP\LOG\"
Else
    McPath = "C:\GFP\LOG\"
End If
 ' McPath = "\\10.101.12.198\GFP\LOG\"
             date2 = FormatDateTime(Date, vbShortDate)
             TIME2 = Format(Time, "HH:MM:SS")
 ''''''''''''''Commented by pravin 07-06-2018'''''''''''
'If FSYS.FileExists(McPath & medakfilename) Then
'    Set INSTREAM = FSYS.OpenTextFile(McPath & medakfilename, ForReading)
'        If INSTREAM.AtEndOfStream = False Then
'         Do While INSTREAM.AtEndOfStream = False
'           FStr = INSTREAM.ReadLine
'           Dat1 = CDate(Trim(Mid(FStr, 44, 10)))
'           Tim1 = Mid(FStr, 57, 12)
'           If (PrvDat1 > Dat1) Or (PrvDat1 = Dat1 And PrvTim1 > Tim1) Then
'             Call Enable_Me(Me, False)
'             txtCtrlMsg.Text = ">>Data of event table is no consistent" & vbCrLf
'             txtCtrlMsg.Text = txtCtrlMsg.Text & ">>Please delete event table immediately" & vbCrLf
'             txtCtrlMsg.Text = txtCtrlMsg.Text & ">>Restart PT Automation after deleting event table" & vbCrLf
'             txtCtrlMsg.Visible = True
'             txtCtrlMsg.Height = 2535
'             txtCtrlMsg.Left = 120
'             txtCtrlMsg.Top = 2760
'             txtCtrlMsg.Width = 11895
'             Exit Sub
'             '''Code added by navnath divekar on dt 2018-05-25
'               '' Exit Do
'             '''End of Code added by navnath divekar on dt 2018-05-25
'           Else
'            PrvDat1 = Dat1
'            PrvTim1 = Tim1
'           End If
'           If InStr(FStr, "FAULT MACHINE STOP BY") = 1 Then
'            date2 = CDate(Trim(Mid(FStr, 44, 10)))
'            TIME2 = Mid(FStr, 57, 12)
'            date3.Text = date2
'            time3.Text = TIME2
'           End If
'         Loop
'        End If
'    INSTREAM.Close
'Else
'   MsgBox "defaut.hor not found", vbInformation
'   'Call Enable_Me(Me, False)
'   'Exit Sub
'End If
'TagDrrej = 0
'checkbarcodeid = ""
'xx = 0
'YY = 0
'If FSYS.FileExists(McPath & "Sansnom") Then
'    Set INSTREAM = FSYS.OpenTextFile(McPath & "Sansnom", ForReading)
'        If INSTREAM.AtEndOfStream = False Then
'         Do While INSTREAM.AtEndOfStream = False
'           FStr = INSTREAM.ReadLine
'           If InStr(FStr, "1DECALAGE D'ORIGINE") = 1 Then
'                xx = InStr(FStr, "(")
'                YY = InStr(FStr, ")")
'              txtOriginGap.Text = Val(Trim(Mid(FStr, xx + 1, YY - xx)))
'              Exit Do
'           End If
'         Loop
'        End If
'    INSTREAM.Close
'End If
'
''''''''''''''Commented by pravin 07-06-2018'''''''''''
Set RS = Nothing
RS.Open "Select sysdate from dual", cn, adOpenKeyset, adLockReadOnly
lblDateDiff.Caption = Round(((Date + Time) - RS("sysdate")) * 60 * 24)
   
   FileName = Format(RS("sysdate"), "dd-MM-yyyy")
   FileName = FileName & ".txt"
   medakfilename = FileName
RS.Close
cboShiftIncharge.clear
Set RS = Nothing
RS.Open "select OPID,sysdate from OPID_NEW where dept = 'PT' and SUBDEPT = 'PT' and Status='A' and plant='" & plant & "'", cn, adOpenKeyset, adLockReadOnly
    Do While RS.EOF = False
        cboShiftIncharge.AddItem RS(0)
        RS.MoveNext
    Loop
RS.Close


If RS.State = 1 Then RS.Close
Set RS = Nothing
RS.Open "select * from FORM_ACCESS_MASTER where dept = 'PT' and FRM_NAME = 'PT AUTOMATION' and FRAME_NO='cmdMELen' and empid='" & uname & "' and plant='" & plant & "'", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        cmdMELen.Enabled = True
    End If
RS.Close




'stCompName = GetCompName
   
   'If UCase$(Mid(stCompName, 1, 2)) = "PT" Then
      If FSYS.FileExists(McPath & FileName) = True Then
          'If Pub_PTLoginRec = False Then
          MsgBox "File " & McPath & FileName & " access successfully...."
              Timer2.Enabled = True
              PTLoginRecFlag = True
              Pub_PTLoginRec = True
          'End If
      End If
   'End If


CHK:
    If err Then
        err.clear
        MsgBox "Defaut.hor not found", vbInformation, "No File"
        Exit Sub
    End If
End Sub
Private Sub PTAllocation()
   On Error GoTo PTAllocation_Error:
   
If ptno.Text <> "" Then
    If RS.State = 1 Then RS.Close
    RS.Open "Select * from PT_Allocation where mcno='" & Val(Me.ptno.Text) & "' ORDER BY PRIORITY", cn, adOpenKeyset, adLockReadOnly
    If RS.EOF = False Then
        drawid.Text = RS("Spool_ID")
        If IsNull(RS("PRIORITY")) = False Then
        If RS("PRIORITY") = "2" Then
            cn.Execute "Update PT_Allocation set PRIORITY='1' where mcno='" & Val(Me.ptno.Text) & "' and PRIORITY='2'"
        End If
        End If
    Else
        MsgBox "No Spool is allocated for this PTM/c" & vbCrLf & "Please allocate new spool from PT Allocation entry Screen", vbExclamation
        RS.Close
        Exit Sub
    End If
    RS.Close
    drawid_lostfocus
Else
    MsgBox "Mc no not found", vbExclamation
End If
   
   Exit Sub

PTAllocation_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "PTAllocation" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "PTAllocation", Erl)
   err.clear
End Sub
Private Sub drawid_dropdown()
   On Error GoTo drawid_dropdown_Error:
   
With drawid
      drawid.clear
      RS.Open "select * from matSTOCK where MATNO IN('330000183','330000155','330000126') AND PROCESS='DRAW' AND STOCKQTY > 0", cn, adOpenKeyset, adLockReadOnly
        Do While RS.EOF = False
            .AddItem RS.Fields("ID").value
            RS.MoveNext
        Loop
      RS.Close
End With
Exit Sub

drawid_dropdown_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "drawid_dropdown" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "drawid_dropdown", Erl)
   err.clear
End Sub
Private Sub drawid_lostfocus()
    Dim SDrDate As String
    Dim sdrdate1 As String
    Dim rschk As New Recordset
    'Print Set Length
    Dim rsms As New ADODB.Recordset
    Dim RS As New ADODB.Recordset
    Dim rslock As New ADODB.Recordset
    Dim RsChinaPt As New ADODB.Recordset
    Dim rsA As New ADODB.Recordset
    Dim rscoat As New ADODB.Recordset
    Dim rsCut As New ADODB.Recordset
    
   On Error GoTo drawid_lostfocus_Error:
   
    iPTContBrkCnt = 0 'Cont. Break Counter for PT SOP is initialized
    iPTContBrkFlag = 0 'Cont Break Flag initialized
    txtCoatType.Text = ""
    sDrNo11 = Val(Mid(drawid, 9, 2))
    txtDrawBarcodeId.Text = vbNullString
    rsdate.Open "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL", cn
        msg1 = rsdate(0)
    rsdate.Close
        
    CURTIME = TimeSerial(Hour(msg1), Minute(msg1), second(msg1))
    If CURTIME >= TimeSerial(7, 0, 0) And CURTIME < TimeSerial(15, 0, 0) Then
        SF = "A"
    ElseIf CURTIME >= TimeSerial(15, 0, 0) And CURTIME < TimeSerial(23, 0, 0) Then
        SF = "B"
    Else
        SF = "C"
    End If

    drawid = UCase$(Trim(drawid.Text))
    plant = Get_Plant(drawid.Text)
  

    Timer1.Enabled = False
    RS.Open "select * from matSTOCK where PROCESS='DRAW' and id =" & "'" & drawid.Text & "'" & " and stockqty>0 ", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        ptype.Text = RS(6)
        balqty = RS("stockqty") * 1000
        SDrDate = Format(RS("createdate"), "YYYY-MM-DD hh:nn:ss")
        sdrdate1 = Format(RS("createdate") - 1, "YYYY-MM-DD hh:nn:ss")
        MatNo = RS("MATNO")
    Else
        RS.Close
        MsgBox "ID not found", vbCritical, "error"
        cn.Execute "Delete from pt_allocation where SPOOL_ID='" & drawid.Text & "'"
        drawid.Text = ""
        Exit Sub
    End If
    RS.Close

    If Mid(drawid.Text, 1, 1) <> "D" And Mid(drawid.Text, 1, 1) <> "N" And Mid(drawid.Text, 1, 1) <> "A" And Mid(drawid.Text, 1, 1) <> "R" And Mid(drawid.Text, 1, 1) <> "M" And Mid(drawid.Text, 1, 1) <> "Q" Then
    If RS.State = 1 Then RS.Close
    RS.Open "SELECT * FROM MATSTOCK WHERE PROCESS = 'SLEV' AND ID='" & Mid(UCase$(drawid.Text), 1, 8) & "'  AND COLR='FELCON'", cn, adOpenKeyset, adLockReadOnly
    If RS.EOF = False Then
        If RS("Stockqty") = 0 Then
            If RS.State = 1 Then RS.Close
            RS.Open "select * from DRAW where Matno ='330000126' and FID LIKE " & "'" & Mid(UCase$(drawid.Text), 1, 8) & "%' order by pdate", cn, adOpenKeyset, adLockReadOnly
            If RS.EOF = False Then
                Drcnt = RS.RecordCount - 1
                If Drcnt = Val(Mid(UCase$(drawid.Text), 11, 1)) Then
                    If RS.State = 1 Then RS.Close
                    RS.Open "Select * from MSENTRY where msid='" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
                    If RS.EOF = True Then
                        MsgBox "Spool Entry Locked due Main Spool Entry of this spool is not Done"
                        drawid.Text = ""
                        If RS.State = 1 Then RS.Close
                        Exit Sub
                    End If
                    If RS.State = 1 Then RS.Close
                End If
            End If
            If RS.State = 1 Then RS.Close
        End If
    End If
    End If
    If RS.State = 1 Then RS.Close

    
    Label10.Caption = ""

    
    RS.Open "Select preid,INSTR_PT from sint_trial where preid='" & Mid(drawid, 1, 8) & "'", cn, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            If IsNull(RS("INSTR_PT")) = False Then Label10.Caption = "Instruction for PT : " & RS("INSTR_PT")
        End If
    RS.Close
          
    RS.Open "select * from matmove where PROCESS='DRAW' and id =" & "'" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
    If RS.EOF = False Then drqty = RS(3) * 1000 Else drqty = 0
    RS.Close
    Rem1.Text = ""

    
    txtBS.Caption = ""
    
    SSRC_Coat_Flag = False
    LUVANTIX_Coat_Flag = False
    txtTotPT.Text = Val(drqty.Text) - Val(balqty.Text)
    FlagScratch = False
    RS.Open "Select BOTSCRAP,prremarks,winding,appear,drlen,ATTN1310,ATTN1550,CUTOFF,mfd,clado,PT_INSTR,COATTYPE,SCR_OBS,DRAW_BARCODEID,MAINSPOOL,NEWMAINSPOOL_NO from DRAW  where matno='330000126' and FID='" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        If IsNull(RS("prremarks")) = False Then lblMSStatus.Caption = RS("prremarks")
        txtBS.Caption = (RS(0) * 1000)
         If RS("scr_obs") = "YES" Then
              If Val(txtTotPT.Text) = 0 Then
                 FlagScratch = True
                 Text1.Visible = True
                 Text1.BackColor = vbYellow
                 Text1.Text = "Scracthes Observed," & vbCrLf & "Give first 4200 metes sample to BSA."
                 MsgBox "Scracthes Observed", vbInformation
                 Text1.Text = ""
                 Text1.Visible = False
              End If
              Label10.Caption = "Scracthes Observed,Recheck Main Spool for Scractch.If OK then do PT else scrap 10 Km...."
         End If
        If IsNull(RS("COATTYPE")) = False Then
            txtCoatType.Text = RS("COATTYPE")
            If Mid(RS("COATTYPE"), 1, 4) = "SSRC" Then
                SSRC_Coat_Flag = True
            End If
            If InStr(1, UCase(RS("COATTYPE")), "LUVANTIX") > 0 Then
                LUVANTIX_Coat_Flag = True
            End If
            If IsNull(RS("DRAW_BARCODEID")) = False Then
                txtDrawBarcodeId.Text = RS("DRAW_BARCODEID")
            End If
        End If
        If IsNull(RS("NEWMAINSPOOL_NO")) = False Then txtSpoolNo.Text = RS("NEWMAINSPOOL_NO")
    End If
    RS.Close
   
    spreid = ""
    sDRNo = Val(Mid(drawid, 9, 2))

    
    TAGSpool = 25
    
    '===========GET THE SOC MACHINE NO.============================
    txtInfo = "Tower: " & Mid(drawid, 9, 2)
    RS.Open "SELECT LATHE FROM OVDD WHERE matno in ('330000152','330000183','330001941','330002291','330002401','330002402','330002481','330002601') and DPREID like '" & Mid(drawid, 1, 8) & "%' and status='OK'", cn, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            txtInfo = txtInfo & " Dep.Mc: " & RS("LATHE")
            SocDepMcno = RS("LATHE")
        End If
    RS.Close
    RS.Open "SELECT TWNO FROM OVDS WHERE matno in ('330000155','330000184','330001937','330002300','330002482','330002501','330002502','330002602') and SPREID like '" & Mid(drawid, 1, 8) & "%' and status='OK'", cn, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
            txtInfo = txtInfo & " Sin.Mc: " & RS("Twno")
        End If
    RS.Close
    RS.Open "Select CORE_SINT_MC from Core_sinter where preid like '" & Mid(drawid, 1, 7) & "%'", cn, adOpenKeyset, adLockReadOnly
    If RS.EOF = False Then
        If RS.RecordCount > 0 Then
            If IsNull(RS(0)) = False Then
                txtInfo = txtInfo & " CoreSinterMc: " & RS(0)
                CoreSinterMcno = RS(0)
            End If
        End If
    End If
    RS.Close
    
    RS.Open "Select DEP_MCNO from Coredep where preid like '" & Mid(drawid, 1, 7) & "%'", cn, adOpenKeyset, adLockReadOnly
    If RS.EOF = False Then
        If RS.RecordCount > 0 Then
            If IsNull(RS(0)) = False Then
                txtInfo = txtInfo & " CoreDepMc: " & RS(0)
                CoreDepMcno = RS(0)
            End If
        End If
    End If
    RS.Close
    
    
    
    
   
    'rs.Open "SELECT * FROM DRAWFLAW WHERE FID = '" & drawid.Text & "' AND POSITION < " & Val(balqty) / 1000 & " AND REMARK NOT IN ('PITCHCHANGE') ORDER BY POSITION DESC ", cn, adOpenKeyset, adLockReadOnly
    'Included PITCHCHANGE remark also , as there are cases in draw logging where
    'PICHCHANGE is there , but NO BFD or LUMP record.This waas causing some flaws (BFD/Lumps) not getting recorded if mc doesn't stop
    'Change on 02.03.2006
    
    NextFault = 1000000 'SUNIL 06032008 FOR ZTPMD INSTRUCTION
    
    RS.Open "SELECT * FROM DRAWFLAW WHERE FID = '" & drawid.Text & "' AND POSITION < " & Val(balqty) / 1000 & "  ORDER BY POSITION DESC ", cn, adOpenKeyset, adLockReadOnly
    TypeOfFlaw = ""
    If RS.RecordCount > 0 Then
        Do While RS.EOF = False
'            If rs("REMARK") <> "AIRLINE" Then      ''''''''''ADDED BY ABDUL ON 12.07.2013 AS PER DILIP/ASHISH
                Fault = Val(balqty) - RS(2) * 1000
                If Fault <= 5000 Then NextFault = Fault
                Rem1.Text = "NEXT FAULT AT " & Fault & " M"
                If Val(Fault) < 2000 Then
                    If RS("REMARK") = "LUMPS" Or Mid(RS("REMARK"), 1, 1) = "L" Then
                        TypeOfFlaw = "L"
                        TypeOfRemark = "LUMPS"
                    End If
                    If RS("REMARK") = "BFD" Or Mid(RS("REMARK"), 1, 1) = "B" Then
                        TypeOfFlaw = "B"
                        TypeOfRemark = "BFD"
                    End If
                    If RS("REMARK") = "AIRLINE" Or Mid(RS("REMARK"), 1, 1) = "A" Then
                        TypeOfFlaw = "A"
                        TypeOfRemark = "AIRLINE"
                    End If
                    If RS("REMARK") = "SCD" Then
                        TypeOfFlaw = "C"
                        TypeOfRemark = "SCD"
                    End If
                End If
                Exit Do
'            End If
            RS.MoveNext
        Loop
    End If
    RS.Close

    'ChkFlaw
    stpos = Val(balqty) / 1000 - 2
    EdPos = Val(balqty) / 1000 + 2
    RS.Open "select * from DRAWflaw where FID = '" & drawid.Text & "' and  position >'" & stpos & "' and position <'" & EdPos & "'", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While RS.EOF = False
'            If rs("REMARK") <> "AIRLINE" Then   '''''''''ADDED BY ABDUL ON 12.07.2013 AS PER DILIP/ASHISH
                If RS("REMARK") = "LUMPS" Or Mid(RS("REMARK"), 1, 1) = "L" Then
                    TypeOfFlaw = "L"
                    TypeOfRemark = "LUMPS"
                End If
                If RS("REMARK") = "BFD" Or Mid(RS("REMARK"), 1, 1) = "B" Then
                    TypeOfFlaw = "B"
                    TypeOfRemark = "BFD"
                End If
                If RS("REMARK") = "AIRLINE" Or Mid(RS("REMARK"), 1, 1) = "A" Then
                    TypeOfFlaw = "A"
                    TypeOfRemark = "AIRLINE"
                End If
                If RS("REMARK") = "SCD" Then
                    TypeOfFlaw = "C"
                    TypeOfRemark = "SCD"
                End If
                If RS("REMARK") = "PITCHCHANGE" Then
                    If TypeOfFlaw = "" Then
                        TypeOfFlaw = "P"
                        TypeOfRemark = "PITCHCHANGE"
                    End If
                End If
'            End If
            RS.MoveNext
        Loop
    End If
    RS.Close
    
    If TypeOfFlaw = "" Then
        TypeOfFlaw = "U" 'For PT Machine Fault
        TypeOfRemark = "UNKNOWN"
    End If

    If Val(Fault) > 55000 Then Rem1.Text = ""
    List1.clear
    
    st = "select * from drawflaw where fid='" & drawid.Text & "'"
    RS.Open st, cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While RS.EOF = False
'            If rs("REMARK") <> "AIRLINE" Then      ''''''''''''ADDED BY ABDUL ON 12.07.2013 AS PER DILIP/ASHISH
                If IsNull(RS("POSITION1")) = False Then
                    If (RS("position1").value) > 0 Then
                        List1.AddItem ((Val(drqty.Text) / 1000 - RS("position1").value) & " to " & (Val(drqty.Text) / 1000 - RS("position").value) & "," & RS("remark").value)
                    ElseIf (Val(drqty.Text) / 1000 - RS("position").value) > 0 Then
                        List1.AddItem ((Val(drqty.Text) / 1000 - RS("position").value) & ", " & RS("remark").value)
                    End If
                ElseIf (Val(drqty.Text) / 1000 - RS("position").value) > 0 Then
                    List1.AddItem ((Val(drqty.Text) / 1000 - RS("position").value) & ", " & RS("remark").value)
                End If
'            End If
            RS.MoveNext
        Loop
    End If
    RS.Close
    List2.clear
    
    Dim BAR()
    i = 0
    cnt1 = 0
    
    rsbar.Open "SELECT FID,BARCODEID FROM PTBARCODE WHERE FID like '" & drawid.Text & "%'", cn, adOpenKeyset, adLockReadOnly
    ReDim BAR(rsbar.RecordCount, 2)
    cnt1 = rsbar.RecordCount - 1
    Do While rsbar.EOF() = False
        BAR(i, 0) = rsbar(0)
        BAR(i, 1) = rsbar(1)
        i = i + 1
        rsbar.MoveNext
    Loop
    rsbar.Close
    
    Dim rsz As New ADODB.Recordset
    
    RS.Open "SELECT * FROM PT WHERE ptID like '" & drawid.Text & "%' order by proddate", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
      Do While RS.EOF = False
        If RS("scrap1") > 0 Then
            List2.AddItem RS("scrap1") & " SCRAP"
        End If
        If RS("OKLEN") > 0 Then
            iPTContBrkCnt = 0 'if OK len then reset PTContbrkcnt
            For i = 0 To cnt1
                If BAR(i, 0) = RS("PTID") Then
                    List2.AddItem RS("OKLEN") & "     " & RS("PTID") & "      " & BAR(i, 1)
                    rsz.Open "select * from pt_inst_ztpmd where barcodeid='" & BAR(i, 1) & "'", cn, adOpenKeyset, adLockReadOnly
                    If rsz.RecordCount > 0 Then
                        If IsNull(rsz("ZTPMDID")) = False Then List2.AddItem "                    " & rsz("ZTPMDID") & "       ZTPMDID"
                    End If
                    rsz.Close
                    Exit For
                End If
            Next i
        End If
        If RS("scrap2") > 0 Then List2.AddItem RS("scrap2") & " SCRAP"
        If RS("drrej") > 0 Then List2.AddItem RS("drrej") & " Drrej"
        If RS("NBREAK") > 0 Then
            'BrkCnt = BrkCnt + 1
            iPTContBrkCnt = iPTContBrkCnt + 1 'If Brk then increase PTcontbrkcnt counter
        End If
        If IsNull(RS("SAMPLE_NO")) = False Then
            If RS("SAMPLE_NO") <> "F" And RS("SAMPLE_NO") <> "L" Then LBLSAMPLENO.Caption = RS("SAMPLE_NO")       '15.12.2014 NEW SAMPLE,active on 30.12.2014
        End If
        RS.MoveNext
      Loop
    End If
    RS.Close
    
    
    If drqty > 0 Then ProgressBar1.value = (Val(drqty) - Val(balqty)) / drqty * 100
    
    
       
    ''**logic written by ankit on 05.02.2014
    If drawid.Text = "" Then Exit Sub
    '''
    
    '----------Get Position Factor First Time the Spool starts-----------
    Stag = 0
    If drqty = balqty Then
        GetPos
        Stag = 1
    Else
        RS.Open "Select min(pos) from pt where ptid like '" & drawid.Text & "%'", cn, adOpenKeyset, adLockReadOnly
        If IsNull(RS(0)) = False Then txtFactor1.Text = RS(0)
        RS.Close
    End If
    
    RS.Open "SELECT sum(nbreak) FROM PT WHERE  ptid like '" & drawid.Text & "%'", cn, adOpenKeyset, adLockReadOnly
    If IsNull(RS(0)) = False Then txtbreakidno.Text = RS(0)
    RS.Close
    
    Call PitchChange_Detail
    If Val(balqty.Text) = Val(drqty.Text) And chkFlaw.value = 1 Then
    Else
        Call PitchChange_NextPos
    End If
    cmdAllocatedID.Enabled = False
    Timer1.Enabled = True
   
    Set_Length_Automatically (Fault)
    
    
    RS.Open "SELECT * FROM  TL_ASSIGN WHERE AREA = 'PT' ORDER BY TLSUBCODE ", cn, adOpenKeyset, adLockReadOnly
      While RS.EOF = False
          txtTLReason.AddItem RS(1)
          RS.MoveNext
      Wend
    RS.Close
    
    Exit Sub
  
drawid_lostfocus_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "drawid_lostfocus" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "drawid_lostfocus", Erl)
   err.clear
    
End Sub
Private Sub PitchChange_Detail()
Dim RS As New ADODB.Recordset
Dim PrevEnd
   On Error GoTo PitchChange_Detail_Error:
   
grdFlaw.ReadOnly = False
If grdFlaw.Rows > 1 Then grdFlaw.Range(1, 2, grdFlaw.Rows - 1, grdFlaw.Cols - 2).ClearAll
For i = 1 To grdFlaw.Rows - 1
    grdFlaw.Cell(i, 1).SetImage "UnChecked"
Next i
For i = 1 To grdFlaw.Rows - 1
    grdFlaw.Cell(i, 6).SetImage "UnChecked"
Next i
i = 0 'Init
RS.Open "Select * from drawpitchchange where fid='" & drawid.Text & "' ORDER BY PITCH_START ASC", cn, adOpenKeyset, adLockReadOnly
While RS.EOF = False
    If RS("PITCH_START") - PrevEnd > 4.2 Or i = 0 Then
        i = i + 1
        If grdFlaw.Rows <= i + 1 Then grdFlaw.Rows = i + 1
        If RS("PITCH_CUT") = "Y" Then
            grdFlaw.Cell(i, 6).SetImage "Checked"
        Else
            grdFlaw.Cell(i, 6).SetImage "UnChecked"
        End If
        If RS("PITCH_STOP") = "Y" Then
            grdFlaw.Cell(i, 1).SetImage "Checked"
        Else
            grdFlaw.Cell(i, 1).SetImage "UnChecked"
        End If
        grdFlaw.Cell(i, 4).Text = drqty.Text / 1000 - RS("PITCH_START") + 0.25   '0.05 FROM 0.3 AS PER MISHRA/BHOUMIK APPROVAL 2008.11.27,0.05 to 0.250 as per Bhaumik 06.01.2009
        
        If Val(Mid(drawid.Text, 9, 2)) >= 20 And Val(Mid(drawid.Text, 9, 2)) < 40 Then 'For New Tower
            grdFlaw.Cell(i, 4).Text = drqty.Text / 1000 - RS("PITCH_START") + 0.35    '0.250 to 0.350 as Per Bhaumik/Yogesh 22042010
        ElseIf Val(Mid(drawid.Text, 9, 2)) >= 40 Then
            grdFlaw.Cell(i, 4).Text = drqty.Text / 1000 - RS("PITCH_START") + 0.4    '0.250 to 0.350 as Per Bhaumik/Yogesh 22042010,'0.35 to 0.4 as per milind sir on 16032015.
        End If

        grdFlaw.Cell(i, 0).Text = i
        If IsNull(RS("PITCH_REASON")) = False Then
            grdFlaw.Cell(i, 2).Text = RS("PITCH_REASON")
        Else
            grdFlaw.Cell(i, 2).Text = "PITCHCHANGE"
         End If
        If RS("PITCH_END") > 0 Then grdFlaw.Cell(i, 3).Text = drqty.Text / 1000 - RS("PITCH_END")
        PrevEnd = RS("PITCH_END")
    Else
        If RS("PITCH_END") > 0 Then grdFlaw.Cell(i, 3).Text = drqty.Text / 1000 - RS("PITCH_END")
        PrevEnd = RS("PITCH_END")
    End If
    RS.MoveNext
Wend
RS.Close
''===================================================
grdFlaw.Rows = i + 1 'sunil
For i = 1 To grdFlaw.Rows - 1
    If grdFlaw.Cell(i, 3).Text = "" Then grdFlaw.Cell(i, 3).Text = Val(grdFlaw.Cell(i, 4).Text) - 0.4
    If Val(grdFlaw.Cell(i, 3).Text) < 0 Then grdFlaw.Cell(i, 3).Text = 0
    grdFlaw.Cell(i, 5).Text = Round(Val(grdFlaw.Cell(i, 4).Text) - Val(grdFlaw.Cell(i, 3).Text), 2)
Next i
If Val(balqty.Text) = Val(drqty.Text) Then
    For i = 1 To grdFlaw.Rows - 1
        If grdFlaw.Cell(i, 3).Text < 4.2 Then
            Call PitchChange_Reason(i)
            grdFlaw.Cell(i, 1).SetImage ("Checked")
            CuttingLen = Val(grdFlaw.Cell(i, 5).Text) + Val(grdFlaw.Cell(i, 3).Text) 'Cutting len after drrej
            txtFlawMsg.Text = "CUT " & CuttingLen & " KM FOR DRAW REJECTION"
            txtFlawMsg.Visible = True
            txtCuttingLen.Text = CuttingLen
            cn.Execute "Update drawpitchchange set PITCH_STOP='Y' where fid='" & drawid.Text & "' and PITCH_START>= " & Val(drqty.Text) / 1000 - Val(grdFlaw.Cell(i, 4).Text) & " and PITCH_END<= " & Val(drqty.Text) / 1000 - Val(grdFlaw.Cell(i, 3).Text) & ""
            chkFlaw.value = 1
            txtFlawNo.Text = i
            Exit For
        End If
    Next i
End If
grdFlaw.ReadOnly = True
grdFlaw.Refresh

If RS.State = 1 Then RS.Close
RS.Open "Select * from PT_PITCH_STATUS where fid='" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then
    If IsNull(RS("status")) = False Then chkFlaw.value = RS("status")
    If IsNull(RS("GRDINDEX")) = False Then txtFlawNo.Text = RS("GRDINDEX")
End If
RS.Close

'''===================================================For silic-bY Mohit-2013-07-19 FOR BOTH PLANTS DONE

Grid3.ReadOnly = False
i = 0 'Init

RS.Open "Select * from draw_info where fid='" & drawid.Text & "' ORDER BY INFO_START ASC", cn, adOpenKeyset, adLockReadOnly
While RS.EOF = False
    i = i + 1
    If Grid3.Rows <= i + 1 Then Grid3.Rows = i + 1
    Grid3.Cell(i, 3).Text = Val(drqty.Text) / 1000 - RS("info_START")
    Grid3.Cell(i, 0).Text = i
    If IsNull(RS("PITCH_REASON")) = False Then
        Grid3.Cell(i, 1).Text = RS("PITCH_REASON")
    Else
        Grid3.Cell(i, 1).Text = "PITCHCHANGE"
    End If
    If RS("INFO_END") > 0 Then Grid3.Cell(i, 2).Text = Val(drqty.Text) / 1000 - RS("INFO_END")
    PrevEnd = RS("INFO_END")
    RS.MoveNext
Wend
RS.Close
Grid3.Rows = i + 1
Grid3.ReadOnly = True
Grid3.Refresh

''''''''''''''''''''''''''''''''''for clad (BCSL REMARK) MARKING ON 05122015,FOR WALUJ ON 05.01.2016
'If plant = "SHENDRA" Then
    Grid5.ReadOnly = False
    i = 0 'Init
    
    RS.Open "Select * from draw_CLAD where fid='" & drawid.Text & "' ORDER BY CLAD_START ASC", cn, adOpenKeyset, adLockReadOnly
    While RS.EOF = False
        i = i + 1
        If Grid5.Rows <= i + 1 Then Grid5.Rows = i + 1
        Grid5.Cell(i, 3).Text = Val(drqty.Text) / 1000 - RS("CLAD_START")
        Grid5.Cell(i, 0).Text = i
        If IsNull(RS("PITCH_REASON")) = False Then
            Grid5.Cell(i, 1).Text = RS("PITCH_REASON")
        Else
            Grid5.Cell(i, 1).Text = "PITCHCHANGE"
        End If
        If RS("CLAD_END") > 0 Then Grid5.Cell(i, 2).Text = Val(drqty.Text) / 1000 - RS("CLAD_END")
        PrevEnd = RS("CLAD_END")
        RS.MoveNext
    Wend
    RS.Close
    Grid5.Rows = i + 1
    Grid5.ReadOnly = True
    Grid5.Refresh
'End If
'''''''''''''''''''''''''''''''''''

If plant = "WALUJ" Or plant = "SHENDRA" Then
''ADDED BY ABDUL ON 13.09.2013 FOR AIRLINE,16.09.2013
    Grid4.ReadOnly = False
    i = 0 'Init
    
    RS.Open "Select * from draw_air where fid='" & drawid.Text & "' ORDER BY INFO_START ASC", cn, adOpenKeyset, adLockReadOnly
    While RS.EOF = False
        i = i + 1
        If Grid4.Rows <= i + 1 Then Grid4.Rows = i + 1
        Grid4.Cell(i, 3).Text = Val(drqty.Text) / 1000 - RS("info_START")
        Grid4.Cell(i, 0).Text = i
        If IsNull(RS("PITCH_REASON")) = False Then
            Grid4.Cell(i, 1).Text = RS("PITCH_REASON")
        Else
            Grid4.Cell(i, 1).Text = "PITCHCHANGE"
        End If
        If RS("INFO_END") > 0 Then Grid4.Cell(i, 2).Text = Val(drqty.Text) / 1000 - RS("INFO_END")
        PrevEnd = RS("INFO_END")
        RS.MoveNext
    Wend
    RS.Close
    Grid4.Rows = i + 1
    Grid4.ReadOnly = True
    Grid4.Refresh


End If
   
   Exit Sub

PitchChange_Detail_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "PitchChange_Detail" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "PitchChange_Detail", Erl)
   err.clear

End Sub
Private Sub PitchChange_Reason(ByVal i As Integer)
'        TypeOfFlaw = ""
   On Error GoTo PitchChange_Reason_Error:
   
        If grdFlaw.Cell(i, 2).Text = "LUMPS" Or Mid(grdFlaw.Cell(i, 2).Text, 1, 1) = "L" Then
            TypeOfFlaw = "L"
            TypeOfRemark = "LUMPS"
        ElseIf grdFlaw.Cell(i, 2).Text = "BFD" Or Mid(grdFlaw.Cell(i, 2).Text, 1, 1) = "B" Then
            TypeOfFlaw = "B"
            TypeOfRemark = "BFD"
        ElseIf grdFlaw.Cell(i, 2).Text = "AIRLINE" Or Mid(grdFlaw.Cell(i, 2).Text, 1, 1) = "A" Then
            TypeOfFlaw = "A"
            TypeOfRemark = "AIRLINE"
        ElseIf grdFlaw.Cell(i, 2).Text = "CD2" Or Mid(grdFlaw.Cell(i, 2).Text, 1, 1) = "C" Then
            TypeOfFlaw = "C"
            TypeOfRemark = "SCD"
        ElseIf grdFlaw.Cell(i, 2).Text = "FSU" Or Mid(grdFlaw.Cell(i, 2).Text, 1, 1) = "F" Then
            TypeOfFlaw = "F"
            TypeOfRemark = "FSU"
        ElseIf grdFlaw.Cell(i, 2).Text = "RUMPUP" Then
            TypeOfFlaw = "S"
            TypeOfRemark = "RUMPUP"
        End If
        frmptautomation.DRReason = TypeOfFlaw
        Remark.Text = TypeOfRemark
   
   Exit Sub

PitchChange_Reason_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "PitchChange_Reason" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "PitchChange_Reason", Erl)
   err.clear
End Sub
Private Sub PitchChange_NextPos()
    Dim NextFlawPos
   On Error GoTo PitchChange_NextPos_Error:
    NextFlawPosGlobal = ""
    txtFlawMsg.Text = ""
    txtCuttingLen.Text = ""
    For i = 1 To grdFlaw.Rows - 1
        If Val(grdFlaw.Cell(i, 3).Text) > Val(drqty.Text) / 1000 - Val(balqty.Text) / 1000 And grdFlaw.Cell(i, 1).ImageKey = "UnChecked" Then
            NextFlawPos = Round(Val(grdFlaw.Cell(i, 3).Text) - (Val(drqty.Text) / 1000 - Val(balqty.Text) / 1000), 3)
            If Val(NextFlawPos) < 60# Then txtFlawMsg.Text = "NEXT PITCH CHANGE  AT " & NextFlawPos & " KM" ''changed from 55 to 60
            NextFlawPosGlobal = Val(NextFlawPos)
            If NextFlawPos < 4.2 Then
                Call PitchChange_Reason(i)
            End If
        End If
    Next i
   
   Exit Sub

PitchChange_NextPos_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "PitchChange_NextPos" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "PitchChange_NextPos", Erl)
   err.clear
End Sub
'''code added by dinesh on ref bajirao sir on 25042018
Private Sub Ptflawallocation()
    Dim ptflawlen
 On Error GoTo Ptflawallocation_Error:
   Dim pos
   
   pos = Val(txtFactor1.Text)
   pos = Val(txtTotPT.Text) / 1000
    For i = 1 To grdFlaw.Rows - 1
       If pos < Val(grdFlaw.Cell(i, 3).Text) And grdFlaw.Cell(i, 1).ImageKey = "UnChecked" Then
           If Val(pos + Val(Text2.Text)) > Val(grdFlaw.Cell(i, 3).Text) Then
                   If Val(pos) >= Val(grdFlaw.Cell(i, 3).Text) - 0.03 Then
                    Text2.Text = Val(Val(grdFlaw.Cell(i, 5).Text) + 0.02)
                   Else
                    Text2.Text = Val(Val(grdFlaw.Cell(i, 3).Text) - 0.01) - Val(pos)
                   End If
                    
              Else
              
               If Mid(txtFlawMsg.Text, 1, 4) = "NEXT" Then
                    If NextFlawPosGlobal < 4.6 Then       '(4.2 + 0.5 +1.5)  < 4.6 no usaable length
                        SetLen = 4.6     ' scope of improvement
                    ElseIf NextFlawPosGlobal >= StandardLen Then    ' team will let us know if std len is 25.2 and next flaw is >50.4 still dey want to set 25.2 if not logic will change
                       If NextFlawPosGlobal >= 25.6 And NextFlawPosGlobal < 50 And StandardLen <= 26 Then
                        SetLen = (NextFlawPosGlobal)
                       Else
                        SetLen = StandardLen
                       End If
                        
                       
                    Else
                        SetLen = (NextFlawPosGlobal) '- 0.15
                    End If
                    
                    Text2.Text = SetLen
                    
                     If Val(pos + Val(Text2.Text)) > Val(grdFlaw.Cell(i, 3).Text) Then
                        Text2.Text = Val(Val(grdFlaw.Cell(i, 3).Text) - 0.01) - Val(pos)
                     End If
               Else
                    
                    If Val(pos) >= Val(grdFlaw.Cell(i, 3).Text) - 0.02 Then
                        Text2.Text = Val(Val(grdFlaw.Cell(i, 5).Text) + 0.02)
                    End If
                End If
                
                
           End If
       
       End If
               
      Next
 Exit Sub

Ptflawallocation_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Ptflawallocation" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Ptflawallocation", Erl)
   err.clear
End Sub

Private Sub PitchChange_Stop()
    Dim CuttingLen
    Dim pos
    Dim RS As New ADODB.Recordset
   On Error GoTo PitchChange_Stop_Error:
   
    If TagDrrej = 1 Then
        txtFlawMsg.Text = ""
        txtCuttingLen.Text = ""
        txtFlawMsg.Visible = True
        txtFlawMsg.Text = "CUT " & 0.5 & " KM FOR DRAW REJECTION" 'Default Instruction
'60            txtCuttingLen.Text = 0.3 'Default Cutting Len 'Changed from 0.5 km to 0.250 km as per mishra/Bhaumik 27.11.08
        txtCuttingLen.Text = 0.5 'Default Cutting Len 'Changed from 0.250 km to 0.5 km as per Bhaumik 06.01.2009 Sunil

        



        pos = Val(drqty.Text) / 1000 - Val(balqty.Text) / 1000
        For i = 1 To grdFlaw.Rows - 1
            If pos > Val(grdFlaw.Cell(i, 3).Text) - 3.5 And pos < Val(grdFlaw.Cell(i, 3).Text) + 3.5 And grdFlaw.Cell(i, 1).ImageKey = "UnChecked" Then
                Check1.value = 1
                Call PitchChange_Reason(i)
                grdFlaw.Cell(i, 1).SetImage ("Checked")
                CuttingLen = grdFlaw.Cell(i, 5).Text 'Cutting len after drrej
                If CuttingLen < 0.5 Then CuttingLen = 0.5 'Default Cutting Len
                
          
                
                txtFlawMsg.Text = "CUT " & CuttingLen & " KM FOR DRAW REJECTION"
                txtCuttingLen.Text = CuttingLen
                cn.Execute "Update drawpitchchange set PITCH_STOP='Y' where fid='" & drawid.Text & "' and PITCH_START>= " & Val(drqty.Text) / 1000 - Val(grdFlaw.Cell(i, 4).Text) & " and PITCH_END<= " & Val(drqty.Text) / 1000 - Val(grdFlaw.Cell(i, 3).Text) & ""
                cn.Execute "INSERT INTO PT_PITCH_STATUS VALUES ('" & drawid.Text & "',1," & i & ")"
                chkFlaw.value = 1
                txtFlawNo.Text = i
                Exit Sub
            End If
        Next i
    End If
   
   Exit Sub

PitchChange_Stop_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "PitchChange_Stop" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "PitchChange_Stop", Erl)
   err.clear
End Sub

Private Sub PitchChange_Cut()
Dim i As Integer
   On Error GoTo PitchChange_Cut_Error:
   
If chkFlaw.value = 1 Then
    If Val(drrej) + Val(scrqty) > 0 Then
        If Val(txtFlawNo.Text) > 0 Then
            If ((Val(drrej) + Val(scrqty)) / 1000) + 0.25 > Val(txtCuttingLen.Text) Then   '0.05 to 0.250 Sunil 06.01.2009 as per Bhaumik
                i = Val(txtFlawNo.Text)
                grdFlaw.Cell(i, 6).SetImage ("Checked")
                cn.Execute "Update drawpitchchange set PITCH_CUT='Y' where fid='" & drawid.Text & "' and PITCH_START>= " & Val(drqty.Text) / 1000 - Val(grdFlaw.Cell(i, 4).Text) & " and PITCH_END<= " & Val(drqty.Text) / 1000 - Val(grdFlaw.Cell(i, 3).Text) & ""
            Else
                txtCuttingLen.Text = Val(txtCuttingLen.Text) - Val(drrej) / 1000
                txtFlawMsg.Text = "CUT " & txtCuttingLen.Text & " KM FOR DRAW REJECTION"
                Exit Sub
            End If
        End If
    Else
        If Val(txtFlawNo.Text) > 0 Then
            i = Val(txtFlawNo.Text)
            grdFlaw.Cell(i, 6).SetImage ("CrossChecked")
        End If
    End If
    cn.Execute "DELETE FROM PT_PITCH_STATUS WHERE FID='" & drawid.Text & "'"
    txtFlawNo.Text = 0
    chkFlaw.value = 0
End If
   
   Exit Sub

PitchChange_Cut_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "PitchChange_Cut" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "PitchChange_Cut", Erl)
   err.clear
End Sub

Private Sub Slider1_Click()
    Slider1.MAX = drqty / 1000
    Slider1.value = (drqty - balqty) / 1000
End Sub
Private Sub Form_Unload(Cancel As Integer)
Pub_PTLoginRec = False
'    MDIProd.mnuPTAutomation.Checked = False
End Sub
Private Sub ChkPT50(ByVal Fault)
    Dim rsA As New ADODB.Recordset
    Dim DrawnKM, Optok
    Dim st As String
    Dim BrkCnt As Integer
    Dim PTSum As Double
   On Error GoTo ChkPT50_Error:
   
    PT50Flag = True
    If PT50Flag = True Then
        If Val(balqty.Text) < 50600 Then
            PT50Flag = False
            Exit Sub
        End If
    End If
    If PT50Flag = True Then
        If (Val(drqty) - Val(balqty.Text)) < 26000 Then
            PT50Flag = False
            Exit Sub
        End If
    End If
    If PT50Flag = True Then
        If Val(Fault) > 0 And Val(Fault) < 50600 Then
            PT50Flag = False
            Exit Sub
        End If
    End If
    If PT50Flag = True Then   'Breaks /100 for last 3 days for the tower
        st = "Select Round((sum(Nbreak)/sum(oklen+Scrap1+scrap2))*100,3) from pt where proddate BETWEEN sysdate-3  AND sysdate and substr(fid,1,1) in ('I','P')and substr(fid,9,2)='" & Mid(drawid.Text, 9, 2) & "'"
        rsA.Open st, cn, adOpenForwardOnly, adLockReadOnly
        If rsA.EOF = False Then
            If IsNull(rsA(0)) = False Then
                If rsA(0) >= 2 Then
                    PT50Flag = False
                End If
            End If
        End If
        rsA.Close
    End If
    If PT50Flag = True Then
        st = "Select SUM(DRLEN) from DRAW where pdate BETWEEN sysdate-7  AND sysdate and substr(fid,1,1) in ('I','P') and substr(fid,9,2)='" & Mid(drawid.Text, 9, 2) & "'"
        rsA.Open st, cn, adOpenForwardOnly, adLockReadOnly
        If rsA.EOF = False Then
            If IsNull(rsA(0)) = False Then
                DrawnKM = rsA(0)
            End If
        End If
        rsA.Close
        
        st = "Select SUM(optlen) from ptfibre where testdate BETWEEN sysdate-7  AND sysdate and substr(fid,1,1) in ('I','P')  and matno='330000127' and grade not in ('REW','FAIL','D','CATV','CATV-SL','MECH','AOTDR','BOTDR') and substr(fid,9,2)='" & Mid(drawid.Text, 9, 2) & "'"
        rsA.Open st, cn, adOpenForwardOnly, adLockReadOnly
        If rsA.EOF = False Then
            If IsNull(rsA(0)) = False Then
                Optok = rsA(0) / 1000
            End If
        End If
        rsA.Close
        If DrawnKM > 0 Then
        If Round(Optok / DrawnKM, 3) < 0.6 Then  'ADY for the tower  '0.8
            PT50Flag = False
        End If
        End If
    End If
    If PT50Flag = True Then
        'If there are two cont. break then switch to 24.4
        BrkCnt = 0
        st = "select OKLEN,SCRAP1,SCRAP2,NBREAK from (select * from pt where fid='" & drawid.Text & "' order by proddate desc) where rownum <3"
        rsA.Open st, cn, adOpenForwardOnly, adLockReadOnly
        If rsA.EOF = False Then
            If IsNull(rsA("NBREAK")) = False Then
                BrkCnt = BrkCnt + rsA("NBREAK")
            End If
            If IsNull(rsA("OKLEN")) = False Then
                PTSum = PTSum + rsA("OKLEN")
            End If
            rsA.MoveNext
        End If
        rsA.Close
        If BrkCnt >= 4 Then   'EARLIER 2
            PT50Flag = False
        End If
    End If
    If PT50Flag = True Then
        'If PT Scrap more then 10 km then switch to 24.4
        st = "select sum(SCRAP1+SCRAP2) from pt where ptid like '" & Mid(drawid.Text, 1, 11) & "%'"
        rsA.Open st, cn, adOpenForwardOnly, adLockReadOnly
        If rsA.EOF = False Then
            If IsNull(rsA(0)) = False Then
                If rsA(0) > 10 Then
                    PT50Flag = False
                End If
            End If
        End If
        rsA.Close
    End If
    
    'If there are two continous 24.4 then switch to 50.6
    If PT50Flag = False Then
        BrkCnt = 0
        PTSum = 0
        st = "select OKLEN,SCRAP1,SCRAP2,NBREAK from (select * from pt where fid='" & drawid.Text & "' order by proddate desc) where rownum <3"
        rsA.Open st, cn, adOpenForwardOnly, adLockReadOnly
        If rsA.EOF = False Then
            If IsNull(rsA("NBREAK")) = False Then
                BrkCnt = BrkCnt + rsA("NBREAK")
            End If
            If IsNull(rsA("OKLEN")) = False Then
                PTSum = PTSum + rsA("OKLEN")
            End If
            rsA.MoveNext
        End If
        rsA.Close
        If PTSum > 24.4 And BrkCnt = 0 Then  'even single bobbin of 24.4 04092008
            PT50Flag = True
        End If
    End If
   
   Exit Sub

ChkPT50_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "ChkPT50" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "ChkPT50", Erl)
   err.clear
End Sub



Private Sub ptno_LostFocus()
   On Error GoTo ptno_LostFocus_Error:
   
If RS.State = 1 Then RS.Close
RS.Open "SELECT SYSDATE FROM DUAL", cn, adOpenKeyset, adLockReadOnly
    todate1 = Format(RS(0), "DD-MM-YYYY")
RS.Close
'==============For Norms Alert
        FG1.clear
        FG1.TextMatrix(0, 0) = "Cons_code"
        FG1.TextMatrix(0, 1) = "Norms"
        FG1.TextMatrix(0, 2) = "Last Changed Date"
        FG1.TextMatrix(0, 3) = "Due Date"
        FG1.TextMatrix(0, 4) = "Status"
        'FG1.TextMatrix(0, 5) = "Days Balance"
'        fg1.TextMatrix(0, 6) = "Min_Lock"
'        fg1.TextMatrix(0, 7) = "Max_Lock"
'        fg1.TextMatrix(0, 8) = "LEADTIME"
'        fg1.TextMatrix(0, 9) = "Status"

        OUTPUT = 0
   
    RS.Open "Select UNIQUE(a.cons_code),b.norms, max(a.edate),max(a.edate)+B.NORMS,b.status from child_consfactor a,master_consfactor b where a.cons_code=b.cons_code and b.status in ('A') group by a.cons_code,b.norms,b.status  order by a.cons_code", cn, adOpenKeyset, adLockReadOnly
   
'   rs.Open "Select a.cons_code, max(a.edate), b.norms,b.unit, b.min_alert, b.max_alert, b.min_lock,b.max_lock,b.lead_time,b.status from child_consfactor a,master_consfactor b where a.cons_code=b.cons_code and b.status ='A' group by a.cons_code,b.norms,b.unit,b.min_alert,b.min_lock, b.max_alert,b.max_lock,b.status,b.lead_time order by a.cons_code", cn, adOpenKeyset, adLockReadOnly
'    rs.Open "select unique(name) ,status from prod.masters where status = 'A' order by name", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While RS.EOF = False
        Subarea = Split(RS(0), "/")(1)
        Machine = Split(RS(0), "/")(3)
        YY = Val(RS(3)) - Val(todate1)
            If Subarea = "PT" And Machine = ptno.Text Then
                FG1.AddItem RS(0) & vbTab & RS(1) & vbTab & RS(2) & vbTab & RS(3) & vbTab & RS(4) '& vbTab & Round(YY, 0)
                OUTPUT = 2
            End If
            RS.MoveNext
        Loop
    End If
    Subarea = "PT"
    Machine = ptno.Text
    RS.Close

    If OUTPUT = 2 Then
'290           convert_to_excel "C:\ProdReports\PT Machine Alert.xls", FG1, ""
'300           FileName = "C:\ProdReports\PT Machine Alert.xls"
'      '        SendMail "ofit@sterlite.com", "Dear All," & vbCrLf & "  Alert For Change/Measure " & Subarea & "-" & Machine & "." & vbCrLf & "Please Find the List of Consumables to be Changed." & vbCrLf, FileName, "AutoMail : Alert For Consumable in : " & Subarea & "-" & Machine
'310           SendMail1 "sarvanan.guru@sterlite.com;ajay.gupta@sterlite.com;prakash.chindake@sterlite.com", "Dear All," & vbCrLf & "  Alert For Change/Measure  " & Subarea & "-" & Machine & "." & vbCrLf & "Please Find the List of Items to be Changed/Measured." & vbCrLf, FileName, "AutoMail : Alert For Item Change / Measure in : " & Subarea & "-" & Machine, "OFIT@STERLITE.COM"
    End If
    
    
'===============For Mc Entry Lock
        FG2.clear
        FG2.TextMatrix(0, 0) = "Cons_code"
        FG2.TextMatrix(0, 1) = "Norms"
        FG2.TextMatrix(0, 2) = "Last Changed Date"
        FG2.TextMatrix(0, 3) = "Due Date"
        FG2.TextMatrix(0, 4) = "Status"
        FG2.TextMatrix(0, 5) = "Days Balance"

        OUTPUT = 0
    RS.Open "Select UNIQUE(a.cons_code),b.norms, max(a.edate),max(a.edate)+B.NORMS,b.status from child_consfactor a,master_consfactor b where a.cons_code=b.cons_code and b.status in ('L') group by a.cons_code,b.norms,b.status  order by a.cons_code", cn, adOpenKeyset, adLockReadOnly
    
'    RS.Open "Select a.cons_code, max(a.edate), b.norms,b.unit, b.min_alert, b.max_alert, b.min_lock,b.max_lock,b.lead_time,b.status from child_consfactor a,master_consfactor b where a.cons_code=b.cons_code and b.status ='L' group by a.cons_code,b.norms,b.unit,b.min_alert,b.min_lock, b.max_alert,b.max_lock,b.status,b.lead_time order by a.cons_code", cn, adOpenKeyset, adLockReadOnly
'    rs.Open "select unique(name) ,status from prod.masters where status = 'L' order by name", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While RS.EOF = False
        Subarea = Split(RS(0), "/")(1)
        Machine = Split(RS(0), "/")(3)
        YY = Val(RS(3)) - Val(todate1)
            If Subarea = "PT" And Machine = Val(ptno.Text) Then
                FG2.AddItem RS(0) & vbTab & RS(1) & vbTab & RS(2) & vbTab & RS(3) & vbTab & RS(4) & vbTab & Round(YY, 0)
                OUTPUT = 1
            End If
            RS.MoveNext
        Loop
    End If
    Subarea = "PT"
    Machine = Val(ptno.Text)
    RS.Close

    If OUTPUT = 1 Then
        Exit Sub
    End If
    
'=====End of alert & lock =============
   
   Exit Sub

ptno_LostFocus_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "ptno_LostFocus" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "ptno_LostFocus", Erl)
   err.clear

End Sub

Private Sub PTscrap_Click()
   On Error GoTo PTscrap_Click_Error:
   
If PTscrap.value = 1 Then
   Check1.value = 0
   DRReason.Text = ""
   If txtScReason.Text = "" Then
      MsgBox ("Enter PT Scrap Reason")
      PTscrap.value = 0
   Exit Sub
   End If
End If
   
   Exit Sub

PTscrap_Click_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "PTscrap_Click" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "PTscrap_Click", Erl)
   err.clear
End Sub
Private Sub Timer1_Timer()

datasaved = False
Dim RS As New ADODB.Recordset
Dim CurDiff
Dim Brkdate As Date, BrkTime
   On Error GoTo Timer1_Timer_Error:
   
If drawid.Text = "" Then Exit Sub
Call PitchChange_Detail
 
FOpenFlag = False
WLFlag = False
PTMode = ""
Timer1.Enabled = False
If chkME.value = 1 Then 'Multiple End e ntry
   Storedata
   chkME.value = 0
    
End If
RS.Open "Select sysdate from dual", cn, adOpenKeyset, adLockReadOnly
 CurDiff = Round(((Date + Time) - RS("sysdate")) * 60 * 24)
RS.Close
If Abs(CurDiff - Val(lblDateDiff.Caption)) > 5 Then
    txtCtrlMsg.Text = ">>Somebody had changed PC date" & vbCrLf
    txtCtrlMsg.Text = txtCtrlMsg.Text & ">>It caused inconsistency in event data" & vbCrLf
    txtCtrlMsg.Text = txtCtrlMsg.Text & ">>Please delete event table immediately" & vbCrLf
    txtCtrlMsg.Text = txtCtrlMsg.Text & ">>Restart PT Automation after deleting event table" & vbCrLf
    txtCtrlMsg.Visible = True
    txtCtrlMsg.Height = 2535
    txtCtrlMsg.Left = 120
    txtCtrlMsg.Top = 2760
    txtCtrlMsg.Width = 11895
    Call Enable_Me(Me, False)
    Timer1.Enabled = False
    PTLoginRecFlag = False
    Exit Sub
Else
    lblDateDiff.Caption = CurDiff
End If


If FSYS.FileExists(McPath & medakfilename) Then
    Set INSTREAM = FSYS.OpenTextFile(McPath & medakfilename, ForReading)
    LINECOUNTER = 0
    'Dim lines() As String
    
   ' lines() = IO.File.ReadAllLines(McPath & medakfilename)
    LENGHGET = False
    If INSTREAM.AtEndOfStream = False Then
        Do While INSTREAM.AtEndOfStream = False
            FStr = INSTREAM.ReadLine
           Dim medakstatus
           medakstatus = Split(FStr, ";")
            
            If InStr(FStr, "RUN") = 1 Then 'From PTMc Start on '26.08.2007
                McStdate = CDate(Trim(Mid(FStr, 44, 10)))
                McStTIME = Mid(FStr, 57, 12)
            End If
            If TIME1 = "10:21:30" Then
                 MsgBox "Hi"
            End If
'            If InStr(FStr, "STOP -> Stop Length reached") > 0 Then
'                DateDrrej = CDate(Trim(Mid(FStr, 44, 10)))
'                TimedrRej = Mid(FStr, 57, 12)
'                If (DateDrrej > date2) Or (DateDrrej = date2 And TimedrRej > TIME2) Then
'                    Remark.Text = "Automatic Draw Rejection"
'                    TagDrrej = 1
'                End If
'            End If
'            If InStr(FStr, "FAULT MINIMUM STRAIN ON PROOF TEST") = 1 Or InStr(FStr, "FAULT FIBER BREAK ON TAKE UP") = 1 Then
'                Brkdate = CDate(Trim(Mid(FStr, 44, 10)))
'                BrkTime = Mid(FStr, 57, 12)
'                If (Brkdate > date2) Or (Brkdate = date2 And BrkTime > TIME2) Then
'                    ActBrk = 1
'                End If
'            End If
Debug.Print FStr
''LINECOUNTER = LINECOUNTER + 1
 If LENGHGET = True Then
   If InStr(FStr, "BCODE") > 0 Then
        If InStr(FStr, "SCRAP") > 0 Then
            If nbreak = 1 Then
                 Remark.Text = "Fiber Break at Sensor"
            Else
                 Remark.Text = "Automatic Draw Rejection"
            End If
            TagDrrej = 1
        Else
            txtBarcodeID.Text = UCase(medakstatus(1))
        End If
        
   End If
 
 End If
 'CODE MOVED BY DINESH KANSE ON 12072018
                DATE1 = CDate(medakstatus(2)) ''' CDate(Trim(Mid(medakfilename, 1, Len(medakfilename) - 4)))
                TIME1 = medakstatus(3)
               
               
  If (DATE1 > date2) Or (DATE1 = date2 And TIME1 > TIME2) Then
              If InStr(FStr, "STOP -> Stop Length reached") > 0 Or InStr(FStr, "STOP -> request by remote") > 0 Or InStr(FStr, "STOP -> Fiber Break at TU Fiber Sensor") > 0 Then
                 If (InStr(FStr, "STOP -> Stop Length reached") > 0 Or InStr(FStr, "STOP -> request by remote") > 0 Or InStr(FStr, "STOP -> Fiber Break at TU Fiber Sensor") > 0) And TagDrrej = 0 And PTscrap = 0 Then
                      nbreak1 = 1 'IF Mc stop for Fault then it may be break or PT Machine Fault '2008-09-26
                      If InStr(FStr, "STOP -> Fiber Break at TU Fiber Sensor") > 0 Then  'IF Mc stop for Minimum strain then it count actual break '2008-09-26
                          nbreak = 1
                      Else
                          nbreak = 0
                      End If
                 Else
                      nbreak = 0
                      nbreak1 = 0
                 End If
                       
                    length = Trim(medakstatus(6))              'Changed from Length = Trim(Mid(FStr, 70, 10))
                    'length = Replace(length, " ", "0")
                    ptlen.Text = length
                    '===========================================================================================
                    'Introduce BrkCounter Length for len counter mismatching between draw & PT
                   ' txtBrkCntLen.Text = Val(Trim(Mid(FStr, 75, 5)))              'Introduce BrkCounter Length for len counter mismatching between draw & PT
                    
                     txtBrkCntLen.Text = 0
'                    If Trim(Mid(FStr, 80, 1)) = "0" Then
'                      If Val(Trim(Mid(FStr, 81, 3))) > 30 Then 'If tension is >30 there then PT else REW
'                        PTMode = "P"
'                      Else
'                        PTMode = "R"
'                      End If
'                    ElseIf Trim(Mid(FStr, 80, 1)) = "1" Then
'                      PTMode = "D"
'                    End If
                    If nbreak = 0 Then txtBrkCntLen.Text = 0                     'If there are no brk then brkcounterlen should be zero
                    If Val(txtBrkCntLen.Text) > 300 Then
                        txtBrkCntLen.Text = 180                     'If PT machine display wrong Counter length(Some case it show wrong value already inform to mait) then default counter length is 180 mts
                        If Val(ptlen.Text) < 1300 Or Val(ptlen.Text) > 23500 Then txtBrkCntLen.Text = 50
                    End If
                    If nbreak = 1 And txtBrkCntLen = 0 Then
                        txtBrkCntLen.Text = 250                     'If PT machine display no Counter length then default counter length is 180 mts
                        If Val(ptlen.Text) < 1300 Or Val(ptlen.Text) > 23500 Then txtBrkCntLen.Text = 50 'for less then 1300 and >23500 default counter len is 50
                    End If
                    '==========================================================================================
                    If Val(TagDrrej) = 1 And Val(ptlen) < 2200 Then Check1.value = 1
                     If ptlen > 10 Then
                        If ptlen > 4200 Then
                            LENGHGET = True
                            ''continue WHILE
                             'txtBarcodeID.Text = lines(LINECOUNTER + 1)
                        Else
                         
                         
                       ' Storedata
                        If datasaved = True Then
                         
                         ''Set_Length_Automatically (Fault)
                        End If
                        TagDrrej = 0
                        ActBrk = 0
                        nbreak = 0
                        nbreak1 = 0
                        PTMode = ""
                        End If
                    
                       ' MsgBox ptlen & "  " & nbreak '
                                  
                       
                    Else
                        date2 = DATE1
                        TIME2 = TIME1
                        date3.Text = DATE1
                        time3.Text = TIME1
                        ptlen.Text = ""
                        PTMode = ""
                    
                   Exit Do
                   
                End If
                
           End If
           If LENGHGET = True And txtBarcodeID.Text <> "" Then
               MsgBox "For Entry " & ptlen & "  " & nbreak  '
           ElseIf TagDrrej = 1 Then
               MsgBox "Scrap Entry " & ptlen & "  " & nbreak  '
           End If
           
                   
                     If (LENGHGET = True And txtBarcodeID.Text <> "") Or (TagDrrej = 1) Then
                        Storedata
                            If datasaved = True Then
                             LENGHGET = False
                             ''Set_Length_Automatically (Fault)
                            End If
                        TagDrrej = 0
                        ActBrk = 0
                        nbreak = 0
                        nbreak1 = 0
                        PTMode = ""
                    End If
           End If
        Loop
    End If
    INSTREAM.Close
End If
xx = 0
YY = 0
'If FSYS.FileExists(McPath & "Sansnom") Then
'    Set INSTREAM = FSYS.OpenTextFile(McPath & "Sansnom", ForReading)
'        If INSTREAM.AtEndOfStream = False Then
'         Do While INSTREAM.AtEndOfStream = False
'           FStr = INSTREAM.ReadLine
'           If InStr(FStr, "1DECALAGE D'ORIGINE") = 1 Then
'                xx = InStr(FStr, "(")
'                YY = InStr(FStr, ")")
'              txtOriginGap.Text = Val(Trim(Mid(FStr, xx + 1, YY - xx)))
'              Exit Do
'           End If
'         Loop
'        End If
'    INSTREAM.Close
'End If

If Val(balqty.Text) = 0 Then cmdexit_Click Else Timer1.Enabled = True
Exit Sub

Timer1_Timer_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Timer1_Timer" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Timer1_Timer", Erl)
   err.clear
End Sub
Private Sub Storedata()
Dim PT50 As Integer

    
    '-----------------------------Sunil for capturing time loss automatically from Machine 25052007-------
    Dim rsTLoss As New ADODB.Recordset
    Dim BobbinActTime, McSpeed, BobbinIdleTime
    Dim RS As New ADODB.Recordset
    finalstatus = False
   On Error GoTo Storedata_Error:
   
           
   
   ''''''
    ''''' added flag for pref stretching
     pref_stretch = False
     If RS.State = 1 Then
     RS.Close
     Set RS = Nothing
     End If
     RS.Open "select preid from stretch_temp where preid = '" & Mid(drawid.Text, 1, 8) & "'", cn, adOpenKeyset, adLockReadOnly
        If RS.RecordCount > 0 Then
         pref_stretch = True
        End If
     If RS.State = 1 Then
     RS.Close
     Set RS = Nothing
     End If
   
   txtBarcodeID.Text = UCase(Trim(txtBarcodeID.Text))
   
    d1 = DateSerial(Year(DATE1), Month(DATE1), Day(DATE1))
    T1 = TimeSerial(Val(Mid(TIME1, 1, 2)), Val(Mid(TIME1, 5, 2)), Val(Mid(TIME1, 10, 2)))
    d1t1 = d1 + T1
    
    D2 = DateSerial(Year(date2), Month(date2), Day(date2))
    T2 = TimeSerial(Val(Mid(TIME2, 1, 2)), Val(Mid(TIME2, 5, 2)), Val(Mid(TIME2, 10, 2)))
    d2t2 = D2 + T2
    
    'Machince Start Time
    D3 = DateSerial(Year(McStdate), Month(McStdate), Day(McStdate))
    T3 = TimeSerial(Val(Mid(McStTIME, 1, 2)), Val(Mid(McStTIME, 5, 2)), Val(Mid(McStTIME, 10, 2)))
    d3t3 = D3 + T3
    DRReason.Text = Trim(UCase$(DRReason.Text))
    'Actual Timeloss is taken from PT mc directly
    txtTimeLoss.Text = 0
    txtSpeedLoss.Text = 0
    If Check2.value <> 1 Then
        txtTimeLoss.Text = Round((d3t3 - d2t2) * 24 * 60, 3)
        If Val(txtTimeLoss.Text) > 10000 Or Val(txtTimeLoss.Text) < 0 Then
            txtTimeLoss.Text = 0
        End If
        BobbinActTime = Round((d1t1 - d3t3) * 24 * 60, 3)
        McSpeed = 2200
        If Val(ptno.Text) >= 20 Then McSpeed = 2500
        BobbinIdleTime = Round(Val(ptlen.Text) / McSpeed, 3)
        txtSpeedLoss.Text = BobbinActTime - BobbinIdleTime
        If Val(txtSpeedLoss.Text) > 10000 Or Val(txtSpeedLoss.Text) < 0 Then
            txtSpeedLoss.Text = 0
        End If
    End If
    
protype = ""
 
 Dim rsGrcomp As New ADODB.Recordset

 If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
rsGrcomp.Open "select product_type from draw where preid='" & Mid(drawid.Text, 1, 8) & "'", cn, adOpenKeyset, adLockReadOnly
If rsGrcomp.RecordCount > 0 Then
    If IsNull(rsGrcomp("PRODUCT_TYPE")) = False Then
        protype = rsGrcomp("PRODUCT_TYPE")
    End If
End If
    If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If

If protype = "" Then
If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
    rsGrcomp.Open "select PRODUCT_TYPE_E from coredep where preid = '" & Mid(drawid.Text, 1, 7) & "'  AND PRODUCT_TYPE_E IS NOT NULL ", cn, adOpenKeyset, adLockReadOnly
        If rsGrcomp.RecordCount > 0 Then
            protype = rsGrcomp(0)
        End If
    If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
End If
If protype = "" Then
    If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
  rsGrcomp.Open "select PRODUCT_TYPE from coredep where preid = '" & Mid(drawid.Text, 1, 7) & "'  AND PRODUCT_TYPE IS NOT NULL ", cn, adOpenKeyset, adLockReadOnly
      If rsGrcomp.RecordCount > 0 Then
          protype = rsGrcomp(0)
      End If
    If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
End If
If protype = "" Then
If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
  rsGrcomp.Open "select PROD_TYPE from RIC_PREFORM_ASSEMBLY where preid = '" & Mid(drawid.Text, 1, 8) & "'  AND PROD_TYPE IS NOT NULL ", cn, adOpenKeyset, adLockReadOnly
      If rsGrcomp.RecordCount > 0 Then
          protype = rsGrcomp(0)
      End If
 If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
End If
  
Dim RICPrefFlag As Boolean
RICPrefFlag = False
Dim RICPrefWt As Double
Dim RICRodCount As Integer
Dim RICRod1DrawnKm, RICRod2DrawnKm, RICRod3DrawnKm, RICRod4DrawnKm, RICRod5DrawnKm, RICRod6DrawnKm, RICRod7DrawnKm
Dim RIDRodID1, RIDRodID2, RIDRodID3, RIDRodID4, RIDRodID5, RIDRodID6, RIDRodID7
RICPrefFlag = False
RICRod1DrawnKm = 0
RICRod2DrawnKm = 0
RICRod3DrawnKm = 0
RICRod4DrawnKm = 0
RICRod5DrawnKm = 0
RICRod6DrawnKm = 0
RICRod7DrawnKm = 0
RIDRodID1 = ""
RIDRodID2 = ""
RIDRodID3 = ""
RIDRodID4 = ""
RIDRodID5 = ""
RIDRodID6 = ""
RIDRodID7 = ""
If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
rsGrcomp.Open "select * from RIC_PREFORM_ASSEMBLY where preid = '" & Mid(drawid.Text, 1, 8) & "'", cn, adOpenKeyset, adLockReadOnly
If rsGrcomp.RecordCount > 0 Then
    RICPrefFlag = True
    RICFibDrawnperKG = (rsGrcomp("PREFWT") * 37) / rsGrcomp("RODASSEMBLYWT")
    If IsNull(rsGrcomp("NO_ROD_USED")) = False Then RICRodCount = rsGrcomp("NO_ROD_USED")
    If IsNull(rsGrcomp("ROD1_WT")) = False Then
        RICRod1DrawnKm = RICFibDrawnperKG * rsGrcomp("ROD1_WT")
        RIDRodID1 = rsGrcomp("ROD1_ID")
    End If
    If IsNull(rsGrcomp("ROD2_WT")) = False Then
        RICRod2DrawnKm = RICRod1DrawnKm + (RICFibDrawnperKG * rsGrcomp("ROD2_WT"))
        RIDRodID2 = rsGrcomp("ROD2_ID")
    End If
    If IsNull(rsGrcomp("ROD3_WT")) = False Then
        RICRod3DrawnKm = RICRod2DrawnKm + (RICFibDrawnperKG * rsGrcomp("ROD3_WT"))
        RIDRodID3 = rsGrcomp("ROD3_ID")
    End If
    If IsNull(rsGrcomp("ROD4_WT")) = False Then
        RICRod4DrawnKm = RICRod3DrawnKm + (RICFibDrawnperKG * rsGrcomp("ROD4_WT"))
        RIDRodID4 = rsGrcomp("ROD4_ID")
    End If
    If IsNull(rsGrcomp("ROD5_WT")) = False Then
        RICRod5DrawnKm = RICRod4DrawnKm + (RICFibDrawnperKG * rsGrcomp("ROD5_WT"))
        RIDRodID5 = rsGrcomp("ROD5_ID")
    End If
    If IsNull(rsGrcomp("ROD6_WT")) = False Then
        RICRod6DrawnKm = RICRod5DrawnKm + (RICFibDrawnperKG * rsGrcomp("ROD6_WT"))
        RIDRodID6 = rsGrcomp("ROD6_ID")
    End If
    If IsNull(rsGrcomp("ROD7_WT")) = False Then
        RICRod7DrawnKm = RICRod6DrawnKm + (RICFibDrawnperKG * rsGrcomp("ROD7_WT"))
        RIDRodID7 = rsGrcomp("ROD7_ID")
    End If
End If
If rsGrcomp.State = 1 Then
     rsGrcomp.Close
     Set rsGrcomp = Nothing
     End If
    
    
    
 '--------Sunil- 17042009----------
  If Val(txtTimeLoss.Text) >= 1 And txtTLReason.Text = "" Then
      MsgBox "Please enter Time Loss Reason", vbCritical
      Exit Sub
  End If
  
  '----------Sunil 17042009-----------
    
    'Check for Shift Incharge Name
    If cboShiftIncharge.Text = "" Then
        MsgBox "Enter Shift Incharge Name", vbCritical
        Exit Sub
    End If
    
    
    If opr.Text = "" Then
        MsgBox "Please Enter Operator Name", vbCritical
        Exit Sub
    End If
'--------------End of Above Sunil for capturing time loss automatically from Machine 25052007-------
    If txtBobType.Text = "" Then
          MsgBox "Enter Bobbin Type", vbCritical
          Exit Sub
    End If
   
    If cboColor.Text = "" Then
        MsgBox "Please select Bobbin Color", vbCritical
        Exit Sub
    End If

    If txtVibration.Text = "" Then
          MsgBox "Enter Vibration Status", vbCritical
          Exit Sub
    End If
    
    
    If txtDancer.Text = "" Then
        MsgBox "Enter Dancer Vibration Status", vbCritical
        Exit Sub
    End If

  If txtBobNo.Text = "" Then
      MsgBox "Enter Bobbin No", vbCritical
      Exit Sub
  End If
    
    
    CURTIME = TimeSerial(Hour(Time), Minute(Time), second(Time))
    If CURTIME >= TimeSerial(7, 0, 0) And CURTIME < TimeSerial(15, 0, 0) Then
        SF = "A"
    ElseIf CURTIME >= TimeSerial(15, 0, 0) And CURTIME < TimeSerial(23, 0, 0) Then
        SF = "B"
    Else
        SF = "C"
    End If
    scrqty = 0
    drrej = 0
    Go_for_Sample = 0 'No Reliability Sample
    
    'Please Follow This Procedure in Every Sub
    '========================================

    Dim Xbool As Boolean
    Xbool = False
    SubName = "Storedata"
    '========================================
    If iPTContBrkFlag = 2 Then
        If (PTscrap.value = 1) And Val(ptlen.Text) > 4300 Then
            'MsgBox "The Scrap Length can not be greater than 10.0 kms", vbCritical
            MsgBox "The Scrap Length can not be greater than 4.2 kms", vbCritical
            Exit Sub
        End If
    ElseIf iPTContBrkFlag = 1 Then
        If (PTscrap.value = 1) And Val(ptlen.Text) > 4300 Then
            'MsgBox "The Scrap Length can not be greater than 5.0 kms", vbCritical
            MsgBox "The Scrap Length can not be greater than 4.2 kms", vbCritical
            Exit Sub
        End If
    Else
        If (PTscrap.value = 1) And Val(ptlen.Text) > 4225 Then ''length change from 2300 to 4225 as per discussion with pravin on 24712
            MsgBox "The Scrap Length can not be greater than 4.225 kms", vbCritical
            Exit Sub
        End If
    End If
    
    If (chkpmd.value = 1) And Val(ptlen.Text) > 250 Then
        MsgBox "The Scrap Length for ZTPMD can not be greater than 250 mts", vbCritical
        Exit Sub
    End If
    
    If Check1.value = 1 And Val(ptlen.Text) > 4200 Then     'Type Mismatche errror , ptlen.text changed to val(ptlen.text) -Sunil 16012007
        If chkME.value = 0 And Check2.value = 0 Then
            Ans11 = MsgBox("Draw Rejection Qty is " & ptlen.Text & vbCrLf & "If you want to send this quantity into Draw rejection Then Click OK" & vbCrLf & "OtherWise Click No button and then Assign BarcodeID for this QTY", vbYesNo, "Confirm")
            If Ans11 = vbNo Then
                Check1.value = 0
                Exit Sub
            End If
        End If
    End If
    If Check1.value = 1 And Val(ptlen.Text) > 40000 Then   'Length > 20 kms can not be sent for draw rejection Sunil /Deshmukh 05052009
          MsgBox "Draw Rejection > 40 kms not allowed ", vbCritical
          Exit Sub
    End If
    
    
    If chkME.value = 1 Then chkME.value = 0
     'Min. length for acceptance 4200 for SMF & 4400 for DOF
    'Minlen = 2000   'MINIMUM lENGTH Changed to 2100 as per inst from Arora Sir on 27.06.2005
    
    Minlen = 4200   'MINIMUM lENGTH Changed to 4200 as per inst from Bhaumik 06012008 SUNIL
    
    If Mid(drawid, 1, 1) = "D" Or Mid(drawid, 1, 1) = "N" Or Mid(drawid, 1, 1) = "A" Or Mid(drawid, 1, 1) = "Q" Then
        Minlen = 4250
    End If

''changed from 09.03.2015
''    If Mid(drawid, 1, 1) = "D" Or Mid(drawid, 1, 1) = "Q" Then
''        Minlen = 6350
''    End If
    
    If Mid(drawid, 1, 1) = "M" Or Mid(drawid, 1, 1) = "R" Then
        Minlen = 2200 'Change from 1100 to 2200 as per sathis mail 26.07.2006
    End If
    

    'Check for Mc No
    If Trim(ptno.Text) = "" Or Val(ptno.Text) <= 0 Or Len(ptno.Text) > 2 Then
        MsgBox "Enter PT Mc No.", vbCritical
        Exit Sub
    End If
    
    If Check2.value = 1 Then
        If DRReason.Text <> "S" Then
            MsgBox "Please Select 'S' for B/E Scrap", vbCritical
            Exit Sub
        End If
    End If
    'If Bottom end being scrapped check it against actual bot scrap kms given by draw
    If Check2.value = 1 And DRReason.Text = "S" Then
        If Val(ptlen.Text) > (Val(txtBS.Caption)) + Minlen Then   '4100 changed to 2100 SUN
            MsgBox "Bottom End Scrap in PT can not be greater than " & Val(txtBS.Caption) + Minlen
            Exit Sub
        End If
    End If

' check for proper barcode ID if OK PT Length

plant = Get_Plant(drawid.Text)

If plant = "WALUJ" Then
    If (chkpmd.value = 0 And PTscrap.value = 0 And Check1.value = 0 And Check2.value = 0) Then
        If Val(ptlen.Text) >= Minlen Then
            txtBarcodeID = UCase$(Trim(txtBarcodeID))
            CHKCHR1 = Val(Mid(txtBarcodeID, 1, 1))
            CHKCHR2 = Mid(txtBarcodeID, 2, 1)
                       
            'Ids starting with 8CXXXXX to be used with K type fibre only 120112 Sunil /Anand Pande
''            If Mid(drawid.Text, 1, 1) <> "K" And Mid(drawid.Text, 1, 1) <> "Y" Then
''                If Mid(txtBarcodeID.Text, 1, 2) = "8C" Then
''                    MsgBox "ENTER Correct Barcode Id", vbCritical, "ERROR"
''                    txtBarcodeID.SetFocus
''                    Exit Sub
''                End If
''            End If
                       
            If Len(txtBarcodeID.Text) <> 7 Then
                MsgBox "ENTER BARCODE ID", vbCritical, "ERROR"
                'txtBarcodeID.SetFocus
                Exit Sub
            End If
            
            If Mid(UCase(Trim(txtBarcodeID.Text)), 1, 1) = "Z" Then
                MsgBox "ENTER Correct Barcode Id", vbCritical, "ERROR"
                txtBarcodeID.SetFocus
                Exit Sub
            End If
            
            
            
            'Check for Validity of 1st and 2nd Character (upto M)
            'If (CHKCHR1 <> 4 Or CHKCHR1 <> 5) And (Val(Asc(CHKCHR2)) < 65 Or Asc(CHKCHR2) > 77) Then
            'If (CHKCHR1 <> 4 Or CHKCHR1 <> 5) And (Val(Asc(CHKCHR2)) < 65 Or Asc(CHKCHR2) > 89) Then
            If Val(CHKCHR1) < 4 Or Val(CHKCHR1) > 9 Then
                MsgBox "ENTER CORRECT BARCODE ID", vbCritical, "ERROR"
                txtBarcodeID.SetFocus
                Exit Sub
            End If
            
            
            For k = 1 To 7        ''ADDED BY ABDUL ON 15.01.2015 TO AVOID SPL CHAR IN BARCODE
                  newchr = Mid(txtBarcodeID.Text, k, 1)
                  If (Val(Asc(newchr)) >= 48 And Val(Asc(newchr)) <= 57) Or (Val(Asc(newchr)) >= 65 And Val(Asc(newchr)) <= 90) Then
                  Else
                      MsgBox "ENTER CORRECT BARCODE ID", vbCritical, "ERROR"
                      txtBarcodeID.SetFocus
                      Exit Sub
                  End If
            Next k
            
            
            If (Val(Asc(CHKCHR2)) < 65 Or Asc(CHKCHR2) > 90) Then
                MsgBox "ENTER CORRECT BARCODE ID", vbCritical, "ERROR"
                txtBarcodeID.SetFocus
                Exit Sub
            End If
            'For DOF,Core,LowOH and Multimode fibre validate third charater as process type ,for SMF normal
            '20080808SUNIL
            For P = 4 To 7
                CHR4 = (Mid(txtBarcodeID, P, 1))
                If (Val(Asc(CHR4))) < 48 Or (Val(Asc(CHR4))) > 57 Then
                    MsgBox "ENTER CORRECT BARCODE ID", vbCritical, "ERROR"
                    txtBarcodeID.SetFocus
                    Exit Sub
                End If
            Next P
            'Check third character in barcode id for correct process type
''''            If Mid(drawid.Text, 1, 1) = "N" Or Mid(drawid.Text, 1, 1) = "M" Or Mid(drawid.Text, 1, 1) = "R" Then
''''                CHR3 = Mid(txtBarcodeID, 3, 1)
''''                If CHR3 <> Mid(drawid.Text, 1, 1) Then
''''                    MsgBox ("Enter Correct Process in Barcode ID")
''''                    Exit Sub
''''                End If
''''            End If
            
''            If Mid(drawid.Text, 1, 1) = "D" Then        ''ADDED BY ABDUL ON 02.01.2015 AS PER MAYUR
''                CHR3 = (Mid(txtBarcodeID.Text, 3, 1))
''                If CHR3 <> Mid(drawid.Text, 1, 1) And CHR3 <> "A" Then
''                   MsgBox ("Enter Correct Process in Barcode ID")
''                   Exit Sub
''                End If
''            End If

            If Mid(drawid.Text, 1, 1) <> "D" And Mid(drawid.Text, 1, 1) <> "Q" And Mid(drawid.Text, 1, 1) <> "A" Then
                'CHR3 = (Mid(txtBarcodeID.Text, 3, 1))
                If Mid(txtBarcodeID.Text, 1, 1) = "7" Then
                   MsgBox "Enter Correct Process Type in BarcodeID", vbCritical
                   Exit Sub
                End If
            End If
            
            
            If Mid(drawid.Text, 1, 1) = "D" Or Mid(drawid.Text, 1, 1) = "Q" Or Mid(drawid.Text, 1, 1) = "A" Then
                'CHR3 = (Mid(txtBarcodeID.Text, 3, 1))
                If Mid(txtBarcodeID.Text, 1, 1) <> "7" Then
                   MsgBox "Enter Correct Process Type in BarcodeID", vbCritical
                   Exit Sub
                End If
            End If
            
            
            If Mid(drawid.Text, 1, 1) = "I" Or Mid(drawid.Text, 1, 1) = "P" Or Mid(drawid.Text, 1, 1) = "C" Or Mid(drawid.Text, 1, 1) = "H" Or Mid(drawid.Text, 1, 1) = "K" Or Mid(drawid.Text, 1, 1) = "E" Then     'For SMF do not accept third character as D,N,M,R  'Sunil 10102009,E TYPE ADDED ON 11.12.2014
             '   CHR3 = Mid(txtBarcodeID.Text, 3, 1)
                If Mid(txtBarcodeID.Text, 1, 1) = "7" Then
                    MsgBox ("Enter Correct Barcode ID")
                    Exit Sub
                End If
            End If
            
''''            If Mid(drawid.Text, 1, 1) = "A" Then
''''                CHR3 = Mid(txtBarcodeID, 3, 1)
''''                If CHR3 <> "N" And CHR3 <> "D" And CHR3 <> "A" Then
''''                    MsgBox ("Enter N OR D type Barcodeid for A type Process")
''''                    Exit Sub
''''                End If
''''            End If
                    
        ''***logic written for Q type''changed by abdul on 18.03.2014
'''        If Mid(drawid.Text, 1, 1) = "Q" Then
'''            If Mid(txtBarcodeID, 3, 1) <> "D" And Mid(txtBarcodeID, 3, 1) <> "Q" And Mid(txtBarcodeID, 3, 1) <> "A" Then
'''''                If Mid(drawid.Text, 1, 1) <> "D" And Mid(drawid.Text, 1, 1) <> "Q" Then
'''                   MsgBox ("Enter Correct Process in Barcode ID")
'''                   Exit Sub
'''''                End If
'''            End If
'''        End If
        
    '====================================for Single Mode fibre should not accept ID like N,M,R,L,D===============================================================
    '====================================Introduced on 21 Mar 2006 by Rajneesh Kr Singh
    '1780          If Mid(drawid.Text, 1, 1) <> "D" And Mid(drawid.Text, 1, 1) <> "N" And Mid(drawid.Text, 1, 1) <> "M" And Mid(drawid.Text, 1, 1) <> "R" And Mid(drawid.Text, 1, 1) <> "L" And Mid(drawid.Text, 1, 1) <> "A" Then
    '1790              CHR3 = (Mid(txtBarcodeID.Text, 3, 1))
    '1800              If (CHR3 <> "S") And (CHR3 <> "L") Then
    '1810                  If (Val(Asc(CHR3))) < 48 Or (Val(Asc(CHR3))) > 57 Then
    '                          'MsgBox "You Cann't Assign M/R/N/D/L/V... type sticker for Process" & vbCrLf & "Other process (I,S) ", vbCritical
    '1820                      MsgBox "You Cann't Assign M/R/N/D... type sticker for Process" & vbCrLf & "Other process (I,S) ", vbCritical
    '1830                      Exit Sub
    '1840                  End If
    '1850              End If
    '1860          End If
    '============================================================================================================================================
            'Check if barcode id already exists
            If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
            RS.Open "SELECT * FROM PTBARCODE WHERE BARCODEID = '" & txtBarcodeID & "'", cn, adOpenKeyset, adLockReadOnly
            If RS.RecordCount > 0 Then
                MsgBox "WRONG BAR CODE ID, ALREADY ASSIGNED", vbCritical, "ERROR"
                'txtBarcodeID.SetFocus
                If RS.State = 1 Then
                RS.Close
                Set RS = Nothing
                End If
                Exit Sub
            End If
           If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        End If
    End If
Else
'------------------Check for Shendra---------------------------------------------------------------
    If (chkpmd.value = 0 And PTscrap.value = 0 And Check1.value = 0 And Check2.value = 0) Then
        If Val(ptlen.Text) >= Minlen Then
            txtBarcodeID = UCase$(Trim(txtBarcodeID))
            CHKCHR1 = Val(Mid(txtBarcodeID, 1, 1))
            CHKCHR2 = Mid(txtBarcodeID, 2, 1)
                       
            If Len(txtBarcodeID.Text) <> 10 Then
                MsgBox "ENTER BARCODE ID", vbCritical, "ERROR"
                Exit Sub
            End If
            
            If Mid(UCase(Trim(txtBarcodeID.Text)), 1, 1) <> "S" Then
                MsgBox "ENTER Correct Barcode Id", vbCritical, "ERROR"
                txtBarcodeID.SetFocus
                Exit Sub
            End If
            
            For k = 1 To 10        ''ADDED BY ABDUL ON 15.01.2015 TO AVOID SPL CHAR IN BARCODE
                  newchr = Mid(txtBarcodeID.Text, k, 1)
                  If (Val(Asc(newchr)) >= 48 And Val(Asc(newchr)) <= 57) Or (Val(Asc(newchr)) >= 65 And Val(Asc(newchr)) <= 90) Then
                  Else
                      MsgBox "ENTER CORRECT BARCODE ID", vbCritical, "ERROR"
                      txtBarcodeID.SetFocus
                      Exit Sub
                  End If
            Next k
            
            If (Val(Asc(CHKCHR2)) < 65 Or Asc(CHKCHR2) > 90) Then
                MsgBox "ENTER CORRECT BARCODE ID", vbCritical, "ERROR"
                txtBarcodeID.SetFocus
                Exit Sub
            End If
                     
            'Check if barcode id already exists
            If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
            RS.Open "SELECT * FROM PTBARCODE WHERE BARCODEID = '" & txtBarcodeID & "'", cn, adOpenKeyset, adLockReadOnly
            If RS.RecordCount > 0 Then
                MsgBox "WRONG BAR CODE ID, ALREADY ASSIGNED", vbCritical, "ERROR"
                RS.Close
                Exit Sub
            End If
            If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        End If
    End If
End If
    
'----------------End of Above-----------------------------------------------------------------------------

'Dim currDate As Date
'Dim Drawdate As Date
'Dim rsdate As New ADODB.Recordset
'
'rsdate.Open "select SYSDATE FROM DUAL", cn
'     currDate = rsdate(0)
'rsdate.Close
'If Val(drqty.Text) > 200000 And (Mid(drawid.Text, 1, 1) = "I" Or Mid(drawid.Text, 1, 1) = "P") Then
'    RS.Open "SELECT PDATE FROM DRAW WHERE FID = '" & drawid.Text & "' and TWNO IN ('6','16','21') AND SUBSTR(FID,1,1) NOT IN ('D')", cn, adOpenKeyset, adLockReadOnly
'      If RS.RecordCount > 0 Then
'          Drawdate = RS(0)
'          If DateDiff("h", Drawdate, currDate) <= 12 Then
'              MsgBox "Difference between Drawdate and PT Date should be greater than 12 hours ", vbCritical, "error"
'              Exit Sub
'          End If
'      End If
'    RS.Close
'End If
'''''''''''''26/11/2012 AS PER MAYUR,ARUN
'If Mid(drawid, 1, 1) = "D" Then
'    RSDATE.Open "select SYSDATE FROM DUAL", cn
'         currDate = RSDATE(0)
'    RSDATE.Close
'    rs.Open "SELECT PDATE FROM DRAW WHERE FID = '" & drawid.Text & "'  , cn, adOpenKeyset, adLockReadOnly"
'      If rs.RecordCount > 0 Then
'          Drawdate = rs(0)
'          If DateDiff("h", Drawdate, currDate) <= 12 Then
'              MsgBox "Difference between Drawdate and PT Date should be greater than 12 hours ,Please unload the spool", vbCritical, "error"
'              Exit Sub
'          End If
'      End If
'    rs.Close
'End If


' VALIDATE  fibreid
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    RS.Open "select * from matSTOCK where PROCESS='DRAW'AND ID=" & "'" & drawid.Text & "'" & " and stockqty>0 ORDER BY ID", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount = 0 Then
        MsgBox "Fibre ID Not found", vbCritical, "ERROR"
       If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        Exit Sub
    End If
    ptype.Text = Mid(drawid.Text, 1, 1)
    balqty = RS("stockqty") * 1000
    datedr = RS(4)
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
 Dim scr1 As Integer, scr2 As Integer
    If Val(ptlen) < Minlen And Check1 = False Then
        scrqty = ptlen
        ptlen = 0
    End If
    
    If Check1 = 1 Then
       drrej = Val(ptlen.Text)
       ptlen = 0
       scrqty = 0
       nbreak = 0
       nbreak1 = 0
    End If
    If PTscrap = 1 And Val(ptlen) >= Minlen Then
       scrqty = Val(ptlen.Text)
       ptlen = 0
       drrej = 0
       nbreak = 0
       nbreak1 = 0
    End If
       
    If Abs(balqty - ptlen - scrqty - drrej) <= (Minlen - 300) Then scrqty = balqty - ptlen - drrej
    If (ptlen + scrqty + drrej) > (balqty + 200) Then
        MsgBox "Total PT quantity Greater than available quantity ", vbCritical, "ERROR"
        Exit Sub
    End If
    
    If drrej > 0 And Check2.value = 0 Then
        If chkME.value = 0 Then 'No Auto selection for ME automation
            DRReason.Text = TypeOfFlaw 'Manual Selection disabled
        End If
    End If
    'If TypeOfFlaw = "" And drrej > 0 Then
    '    If DRReason.Text <> "S" And DRReason.Text <> "M" Then
    '        MsgBox "Only B/E Scrap (Drreason-S), Multilple end (Drreason-M)  are allow manually", vbExclamation
    '        Exit Sub
    '    End If
    'End If
    
    'Introduce on 22 jan 2007 by Rajneesh
    If drrej > 0 Then
        If Trim(DRReason.Text) = "" Then
            MsgBox "Draw Rejection reason cann't be blank", vbExclamation
            Exit Sub
        End If
    End If
    
    ' DATA VALIDATION OVER
    ' NOW GENERATE NEW FIBRE ID
      CNT = 0
      
         
        ''**logic written by ankit on 05.02.2014
        If drawid.Text = "" Then Exit Sub
        '''
      If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
      RS.Open "select * from MATSTOCK where PROCESS = 'PT' AND ID LIKE " & "'" & drawid.Text & "%' order by id", cn, adOpenKeyset, adLockReadOnly
        Do While RS.EOF = False
            CNT = CNT + 1
            ID = RS("id")
            RS.MoveNext
        Loop
     If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If

        If CNT >= 25 Then
            PTID = drawid & "Z" & Chr(CNT + 65 - 25)
        Else
            PTID = drawid & Chr(CNT + 65)
        End If
        

Dim SumDate
            If rsdate.State = 1 Then
            rsdate.Close
            Set rsdate = Nothing
            End If
rsdate.Open "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL", cn, adOpenKeyset, adLockReadOnly
PRDate = rsdate(0)
SumDate = rsdate(0)
         If rsdate.State = 1 Then
            rsdate.Close
            Set rsdate = Nothing
            End If
PTLoginRecFlag = False


PT50 = 0
If Val(ptlen.Text) >= 48450 Then PT50 = 1
If Val(ptlen.Text) >= 24450 And Val(ptlen.Text) < 48450 Then
  If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    RS.Open "SELECT OKLEN,FLAG50 FROM PT WHERE FID = '" & drawid.Text & "' ORDER BY PRODDATE DESC", cn, adOpenDynamic, adLockOptimistic
    If RS.RecordCount > 0 Then
          If RS("oklen") >= 24.45 And RS("FLAG50") <> "Y" Then
                PT50 = 1
          Else
                PT50 = 0
          End If
    Else
          PT50 = 0
    End If
        If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
End If


 ''''''''''This is New Code Added By Pravin For Swap Barcode recommend by ABhishek On 07042017
 If plant = "SHENDRA" Then
   ' MsgBox "This Is Pravin Store"
   If ptlen.Text >= 4200 Then
        lockbarcedflag = True
        lockforbarcode
        If lockbarcedflag = False Then
            Exit Sub
        End If
    End If
Else
    If ptlen.Text >= 4200 Then
        lockbarcedflag = True
        lockforbarcodewaluj
        If lockbarcedflag = False Then
            Exit Sub
        End If
    End If
End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'====================Use in All Sub where transactions take palce to the database======
    cn.BeginTrans
    Xbool = True
'======================================================================================
    SAP_Init
    SAP_Date = PRDate 'For maintaining same date as production date
    SAP_Process_Time = 0 'For calculating sap_process_time (FAB Time)
    SAP_Labout = 0  'SAP Labout time

    
    cmdSAP_Click
    If SAP_Entry = False Then
        MsgBox "No PT entry without SAP Entry", vbCritical
        cn.RollbackTrans
        Exit Sub
    End If

'  FIBRE ID GENERATED
  'UPDATE DATA IN TABLE MATSTOCK FOR PT FIBRE
      If ptlen > 0 Then
      If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        RS.Open "select * from MATSTOCK", cn, adOpenDynamic, adLockOptimistic
        RS.AddNew
                RS(0).value = "330000127"
                If ptype.Text = "N" Or ptype.Text = "D" Or ptype.Text = "Q" Or ptype.Text = "A" Then RS(0).value = "330001939" 'Added on 30.11.2004
                If ptype.Text = "M" Then RS(0).value = "330000156"  'Added on 30.11.2004
                If ptype.Text = "R" Then RS(0).value = "330000184"
                RS(1).value = PTID
                RS(2).value = Round(Val(ptlen.Text) / 1000, 3)
                RS(3).value = ""
                RS(4).value = PRDate
                RS(5).value = "PT"
                RS(6).value = ptype.Text
                RS(7).value = "NATURAL"
                RS(9).value = Get_Plant(drawid.Text)
                
           RS.update
        If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    ' record missed flaws
        If Val(Mid(Rem1, 14, 10)) > 0 And (Val(ptlen.Text) - Val(Mid(Rem1, 14, 10))) > 1000 And Val(Mid(Rem1, 14, 10)) > 1000 Then
        ' changed from 1500m to 100 m on 18.08.04
        'Changed from 100 m to 1000 m on 16.03.2006 as per Bhaumik
        If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
            RS.Open "select * from ptflaw", cn, adOpenDynamic, adLockOptimistic
            RS.AddNew
               RS(0) = PTID
               RS(1) = Round(Val(ptlen.Text) / 1000, 3)
               RS(3) = PRDate
               RS(4) = Val(Mid(Rem1, 14, 10))
               RS(5) = Get_Plant(drawid.Text)
            RS.update
          If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        ElseIf TagDrrej = 0 And Val(Mid(Rem1, 14, 10)) > 0 And (Val(ptlen.Text) - Val(Mid(Rem1, 14, 10))) > 0 And Val(Mid(Rem1, 14, 10)) > 1000 Then
        ' Record if Machine stop other then Dr Cutting on 16.03.2006 as per Bhaumik
            If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
            RS.Open "select * from ptflaw", cn, adOpenDynamic, adLockOptimistic
            RS.AddNew
               RS(0) = PTID
               RS(1) = Round(Val(ptlen.Text) / 1000, 3)
               RS(3) = PRDate
               RS(4) = Val(Mid(Rem1, 14, 10))
               RS(5) = Get_Plant(drawid.Text)
            RS.update
           If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        End If
      End If
            If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        RS.Open "select * from MATSTOCK WHERE MATNO in( '330000126','330000183','330000155','330001937') AND ID = '" & drawid.Text & "' and process='DRAW'", cn, adOpenDynamic, adLockOptimistic
               RS(2).value = Round((Val(balqty) - scrqty - Val(ptlen) - drrej - Val(txtBrkCntLen.Text)) / 1000, 3)
        RS.update
       If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        If drqty > 0 Then ProgressBar1.value = (Val(drqty) - Val(balqty) + scrqty + Val(ptlen) + drrej + Val(txtBrkCntLen.Text)) / drqty * 100

    ' UPDATE DATA IN TABLE MATMOVE FOR DRAWN FIBRE
        If Val(ptlen) + scrqty > 0 Then
        If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        RS.Open "select * from MATMOVE ", cn, adOpenDynamic, adLockOptimistic 'for Ahmed 2013-02-05
        RS.AddNew
                RS(0).value = "330000126"
                If ptype.Text = "N" Or ptype.Text = "D" Or ptype.Text = "Q" Or ptype.Text = "A" Then RS(0).value = "330001937"
                If ptype.Text = "M" Then RS(0).value = "330000155"
                If ptype.Text = "R" Then RS(0).value = "330000183"
                RS(1).value = drawid.Text
                RS(2).value = PRDate
                'rs(3).Value = (Val(ptlen) + scrqty) / 1000
                RS(3).value = Round((Val(ptlen) + scrqty + Val(txtBrkCntLen.Text)) / 1000, 3) 'BrkCounter Length on 26.08.2007
                RS(4).value = "PT"
                RS(5).value = ptno.Text
                RS(6).value = Mid(opr.Text, 1, 4)
                RS(7).value = ""
                RS(8).value = ptype.Text
                RS(9).value = "261"
                RS(10).value = "I"
                RS(11).value = PRDate
                RS(12) = SF
                RS("BARCODEID") = UCase$(txtBarcodeID)
                RS("PLANT") = Get_Plant(drawid.Text)
           RS.update
       If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        End If
        If drrej > 0 Then
        If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        RS.Open "select * from MATMOVE ", cn, adOpenDynamic, adLockOptimistic 'for Ahmed 2013-02-05
        RS.AddNew
                RS(0).value = "330000126"
                If ptype.Text = "N" Or ptype.Text = "D" Or ptype.Text = "Q" Or ptype.Text = "A" Then RS(0).value = "330001937"
                If ptype.Text = "M" Then RS(0).value = "330000155"
                If ptype.Text = "R" Then RS(0).value = "330000183"
                RS(1).value = drawid.Text
                RS(2).value = PRDate
                RS(3).value = Round(drrej / 1000, 3)
                RS(4).value = "PT"
                RS(5).value = ptno.Text
                RS(6).value = Mid(opr.Text, 1, 4)
                RS(7).value = ""
                RS(8).value = ptype.Text
                RS(9).value = "551"
                RS(10).value = "I"
                RS(11).value = PRDate
                RS(12) = SF
                RS("BARCODEID") = UCase$(txtBarcodeID)
                RS("PLANT") = Get_Plant(drawid.Text)
           RS.update
      If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        End If
' UPDATE DATA IN TABLE MATMOVE FOR PT FIBRE
     If ptlen > 0 Then
     If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        RS.Open "select * from MATMOVE ", cn, adOpenDynamic, adLockOptimistic  'for Ahmed 2013-02-05
        RS.AddNew
                RS(0).value = "330000127"
               If ptype.Text = "N" Or ptype.Text = "D" Or ptype.Text = "Q" Or ptype.Text = "A" Then RS(0).value = "330001939"
               If ptype.Text = "M" Then RS(0).value = "330000156"
               If ptype.Text = "R" Then RS(0).value = "330000184"
                RS(1).value = PTID.Text
                RS(2).value = PRDate
                RS(3).value = Round(Val(ptlen) / 1000, 3)
                RS(4).value = "PT"
                RS(5).value = ptno.Text
                RS(6).value = Mid(opr.Text, 1, 4)
                RS(7).value = ""
                RS(8).value = ptype.Text
                RS(9).value = "101"
                RS(10).value = "R"
                RS(11).value = PRDate
                RS(12) = SF
                RS("BARCODEID") = UCase$(txtBarcodeID)
                RS("PLANT") = Get_Plant(drawid.Text)
           RS.update
       If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
     End If
    'Get PT Breaks
    If nbreak > 0 Then
        'rs.Open "SELECT sum(nbreak) FROM PT WHERE matno='330000126' and substr(ptid,1,11) = '" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
        'RS.Open "SELECT sum(nbreak) FROM PT WHERE  substr(ptid,1,11) = '" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
        RS.Open "SELECT sum(nbreak) FROM PT WHERE  ptid like '" & drawid.Text & "%'", cn, adOpenKeyset, adLockReadOnly
        If IsNull(RS(0)) = False Then BrkCnt = RS(0) Else BrkCnt = 0
        BREAKID = (Mid(drawid.Text, 11, 1) & (BrkCnt + 1))
        If IsNull(BrkCnt) = False Then txtbreakidno.Text = BrkCnt + 1
        RS.Close
    End If
     Dim rsdep As New ADODB.Recordset
    ' UPDATE DATA IN TABLE PT FOR PT FIBRE
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    RS.Open "select * from PT ", cn, adOpenDynamic, adLockOptimistic
    RS.AddNew
    RS(0).value = "330000126"
    If Mid(drawid.Text, 1, 1) = "N" Or Mid(drawid.Text, 1, 1) = "D" Or Mid(drawid.Text, 1, 1) = "Q" Or Mid(drawid.Text, 1, 1) = "A" Then RS(0).value = "330001937"
    If Mid(drawid.Text, 1, 1) = "M" Then RS(0).value = "330000155"
    If Mid(drawid.Text, 1, 1) = "R" Then RS(0).value = "330000183"
    RS(1).value = drawid.Text
    RS(2).value = PRDate
    RS(3).value = Round(Val(ptlen) / 1000, 3)
    RS(4).value = Round(scrqty / 1000, 3)
    
    'new code added 2013-06-27
    totscrap = RS(4).value
    
    RS(5).value = Round(Val(txtBrkCntLen.Text) / 1000, 3) 'Including Brk Counter length from 26.08.2007 by Rajneesh

    RS(6).value = "PT"
    RS(7).value = PTID.Text
    RS(8).value = Val(ptno.Text)
    If nbreak > 0 Then
        RS(9).value = BREAKID
    End If
    RS(10).value = Round(Val(nbreak), 0)
    RS(11).value = ptype.Text
    RS(12).value = Round(drrej / 1000, 3)
    RS(13).value = 0
    
    If pref_stretch = True Then RS("pref_stretch") = "YES" Else RS("pref_stretch") = "NO"   ''added for pref_stretching by nirja on 07-08-2014
    'new code added 2013-06-27
    totrej = RS(12).value - RS(13).value 'rs(4).value
    
    If totscrap > 0 Or totrej > 0 Then
        finalstatus = True
    End If
    
    If drrej > 0 Then
      RS("remark") = Remark
      RS(15) = DRReason 'FlawType
    End If
    If chkpmd.value = 1 Then
      RS("sample") = "ZTPMD"
    End If
    If chkdocsample.value = 1 Then   ''added by abdul for Reliability sample on 07.12.2013,27.01.2014
        RS("sample") = "DOC SAMPLE"
    End If
    If PTscrap.value = 1 Then
      RS("sample") = Trim(txtScReason.Text)
    End If
    RS("OPR") = Mid(opr.Text, 1, 4)
    RS("SHIFT_INCHARGE") = Mid(cboShiftIncharge.Text, 1, 10)
    If Stag = 1 Then
        RS("FACTOR") = Round(Factor, 3)
        RS("POS") = Round(Factor - (Val(ptlen.Text) + scrqty + drrej + Val(txtBrkCntLen.Text)) / 1000, 3)
    Else
        RS("Factor") = Round(Val(txtFactor1), 3)
        RS("POS") = Round(Val(txtFactor1) - (Val(ptlen.Text) + scrqty + drrej + Val(txtBrkCntLen.Text)) / 1000, 3)
    End If
    
    'code added by dinesh kanse ref factor issue of 0 regards mbend and mohhamad sir
'    If Val(RS("FACTOR")) <= 0 Then
'            MsgBox "Factore value should not saved zero", vbCritical, "Factor Zero"
'            Exit Sub
'     End If
    
    RS("SCRTYPE") = txtScr.Text
    RS("TIMELOSS") = Round(Val(txtTimeLoss), 3)
    RS("logdate") = d1t1  'Machine Logging Date
    RS("SPEEDLOSS") = Round(Val(txtSpeedLoss.Text), 3)
    RS("SPOS") = Round((Val(drqty.Text) - Val(balqty.Text)) / 1000, 3)
    RS("EPOS") = Round(RS("SPOS") + (Val(ptlen.Text) + scrqty + drrej + Val(txtBrkCntLen.Text)) / 1000, 3)
    
  'Checking First 25-50/Last 25-50 Km of Main Spool         '15.12.2014 NEW SAMPLE,active on 30.12.2014
    
    If RS("OKLEN") > 0 Then
        If RS("SPOS") < 50# Then
            LBLSAMPLENO = 0
            RS("SAMPLE_NO") = "F"
            'Top 50 km
        ElseIf Val(drqty.Text) / 1000 - RS("EPOS") < 50# Then
             LBLSAMPLENO = 0
             RS("SAMPLE_NO") = "L"
            'Bottom 50Km
        Else
             LBLSAMPLENO.Caption = Val(LBLSAMPLENO.Caption) + 1
             RS("SAMPLE_NO") = Val(LBLSAMPLENO.Caption)
            'mark as 1
        End If
    End If
    
    
    '''FOR BOTH PLANTS
    For i = 1 To Grid3.Rows - 1
          If (Val(Grid3.Cell(i, 2).Text) > RS("SPOS")) And (Val(Grid3.Cell(i, 2).Text) < RS("EPOS")) Then
                RS("draw_info") = "T"
          End If
          If (Val(Grid3.Cell(i, 3).Text) > RS("SPOS")) And (Val(Grid3.Cell(i, 3).Text) < RS("EPOS")) Then
                RS("draw_info") = "T"
          End If
    Next i
    
'    If plant = "SHENDRA" Then       ''FOR CLAD (BCSL) MARKING ON 05122015,FOR WALUJ ON 05.01.2016
        For i = 1 To Grid5.Rows - 1
              If (Val(Grid5.Cell(i, 2).Text) > RS("SPOS")) And (Val(Grid5.Cell(i, 2).Text) < RS("EPOS")) Then
                    RS("DRAW_CLAD") = "Q"
              End If
              If (Val(Grid5.Cell(i, 3).Text) > RS("SPOS")) And (Val(Grid5.Cell(i, 3).Text) < RS("EPOS")) Then
                    RS("DRAW_CLAD") = "Q"
              End If
        Next i
'    End If
    
    If UCase(plant) = "WALUJ" Or UCase(plant) = "SHENDRA" Then
        For i = 1 To Grid4.Rows - 1  ''ADDED BY ABDUL ON 13.09.2013 FOR AIRLINE
             If (Val(Grid4.Cell(i, 2).Text) > RS("SPOS")) And (Val(Grid4.Cell(i, 2).Text) < RS("EPOS")) Then
                   RS("draw_air") = "A"
             End If
            
            If (Val(Grid4.Cell(i, 3).Text) > RS("SPOS")) And (Val(Grid4.Cell(i, 3).Text) < RS("EPOS")) Then
                   RS("draw_air") = "A"
             End If
        Next i
    End If
    
    If TagDrrej = 1 Then RS("SLReason") = "D"
    If nbreak1 = 1 Or nbreak = 1 Then
       If ActBrk = 1 Then
          RS("SLReason") = "B"
       Else
          RS("SLReason") = "M"
       End If
    End If
    RS("PTMODE") = Mid(PTMode, 1, 1)
    RS("VIBRATION") = UCase(txtVibration.Text)
    RS("dancer_vibration") = UCase(txtDancer.Text)
    RS("PLANT") = Get_Plant(drawid.Text)
    RS("LOGNAME") = uname
    RS("coattype") = txtCoatType.Text

    If protype <> "" Then
        RS("PRODUCT_TYPE") = protype
    End If
    
    If PT50 = 1 Then RS("FLAG50") = "Y" Else RS("FLAG50") = "N"
        rsdep.Open "SELECT SPINDLE FROM OVDD WHERE DPREID like '" & Mid(drawid.Text, 1, 8) & "%'", cn, adOpenKeyset, adLockReadOnly
          If rsdep.RecordCount > 0 Then
              RS("SPINDLE") = rsdep("SPINDLE")
          End If
        rsdep.Close
        
    If chkMERej.value = 1 Then            ''added by abdul for mutlipel end rejection reason on 04.04.2015
        RS("MULTIEND_REASON") = cmbmutiend.Text
    End If
    
    
    RS.update
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    
    'UPDATE DATE IN TABLE PTBARCODE if OK PT Len
    If Val(ptlen.Text) >= Minlen Then
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
       RS.Open "SELECT * FROM PTBARCODE", cn, adOpenDynamic, adLockOptimistic
       RS.AddNew
          RS(0) = UCase$(txtBarcodeID)
          RS(1) = PTID.Text
          RS(2) = Round(Val(ptlen.Text), 3)
          RS(3) = Val(ptno.Text)
          RS(4) = PRDate
          RS("RELLEN") = Round((Val(drqty.Text) - Val(balqty.Text) + ptlen) / 1000, 3) 'Added on 28.05.05 for Along the Length Analysis
          RS("DIRECTION") = "E"
'          If Plant = "SHENDRA" Then
            RS("STATUS") = "D"
'          Else
'            rs("STATUS") = "O"
'          End If
          
          RS("POPR") = UCase(Mid(opr.Text, 1, 4))
          RS("PROCESS") = "PT"  'SUNIL 21012008
          RS("PLANT") = Get_Plant(drawid.Text)
          If Stag = 1 Then
              RS("PREFPOS") = Round(Factor - (Val(ptlen.Text) + scrqty + drrej + Val(txtBrkCntLen.Text)) / 1000, 3)
          Else
              RS("PREFPOS") = Round(Val(txtFactor1) - (Val(ptlen.Text) + scrqty + drrej + Val(txtBrkCntLen.Text)) / 1000, 3)
          End If
          
       
       
       
       RS.update
     If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    End If
    chkpmd = 0
    Stag = 0

' For Robo Target Location is 'C'--OTDR TESTING
''''    If Val(ptlen.Text) >= Minlen Then
''''       If rs.State = 1 Then rs.Close
''''       rs.Open "SELECT * FROM OBJTARGET", cn, adOpenDynamic, adLockOptimistic
''''       rs.AddNew
''''          rs("OBJNO") = UCase$(txtBarcodeID)
''''          rs("TARGET") = "C"
''''          rs("ORDNO") = 0
''''          rs("POSNR") = 0
''''          rs("SORTING") = 0
''''          If Mid(drawid.Text, 1, 1) = "M" Or Mid(drawid.Text, 1, 1) = "R" Then rs("TARGET") = "K"
''''
''''          'If Val(Mid(drawid.Text, 9, 2)) = 10 Or Val(Mid(drawid.Text, 9, 2)) = 11 Or Val(Mid(drawid.Text, 9, 2)) = 12 Or Val(Mid(drawid.Text, 9, 2)) = 13 Or Val(Mid(drawid.Text, 9, 2)) = 14 Or Val(Mid(drawid.Text, 9, 2)) = 15 Then
''''          If Val(Mid(drawid.Text, 9, 2)) = 10 Or Val(Mid(drawid.Text, 9, 2)) = 15 Then
''''                rs("TARGET") = "K"
''''          End If
''''          rs("NEWSIGN") = "Y"
''''       rs.Update
''''       If rs.State = 1 Then rs.Close
''''    End If
'==============For PT25.2 Tracking============================
    If Val(ptlen.Text) >= 24450 Then
        If Mid(drawid.Text, 1, 1) <> "M" And Mid(drawid.Text, 1, 1) <> "R" And Mid(drawid.Text, 1, 1) <> "D" And Mid(drawid.Text, 1, 1) <> "Q" And Mid(drawid.Text, 1, 1) <> "N" And Mid(drawid.Text, 1, 1) <> "A" Then
           If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
           RS.Open "SELECT * FROM PT25TRACKING", cn, adOpenDynamic, adLockOptimistic
           RS.AddNew
              RS("FID") = UCase$(PTID.Text)
              RS("LENGTH") = Round(Val(ptlen.Text), 3)
              RS("STATUS") = "O"
              RS("PDATE") = PRDate
              RS("PLANT") = Get_Plant(drawid.Text)
           RS.update
           If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        End If
    End If
'==============For PT50.4 Tracking============================
    If plant = "SHENDRA" Then
        If Val(ptlen.Text) >= 48850 Then
             If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
                 RS.Open "SELECT * FROM PT50TRACKING", cn, adOpenDynamic, adLockOptimistic
                 RS.AddNew
                 RS("FID") = UCase$(PTID.Text)
                 RS("LENGTH") = Val(ptlen.Text)
                 RS("STATUS") = "O"
                 RS("PDATE") = PRDate
                 RS("PLANT") = Get_Plant(drawid.Text)
                 RS.update
                If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        End If
    End If
    


    Dim A As Double
          
    If ptlen > 0 Then
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        RS.Open "select * from ptcolor WHERE FID = '" & PTID.Text & "'", cn, adOpenDynamic, adLockOptimistic
        If RS.EOF = True Then
            RS.AddNew
        End If
        RS(0) = PTID.Text
        RS(1) = cboColor.Text
        RS(2) = PRDate
        RS("BOBBIN_TYPE") = UCase(txtBobType.Text)
        RS("DEV_SPOOL") = chkDeviationSpol.value
        RS("BOB_NO") = Val(txtBobNo.Text)
        RS("PLANT") = plant
        RS("PTLEN") = Round(Val(ptlen) / 1000, 3)
        RS.update
        If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    End If
    
    Dim PT_25
    PT_25 = 0
    If Val(ptlen.Text) / 1000 >= 25.35 Then
        If Val(ptlen.Text) / 1000 >= 58.9 Then
            PT_25 = 58.8
        ElseIf Val(ptlen.Text) / 1000 >= 50.55 Then
            PT_25 = 50.4
        Else
            PT_25 = 25.2
        End If
    End If
    'For Realtime production Report
  
   If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    RS.Open "Select * from PRODREPORT_DRAW where pdate='" & Format(PRDate, "yyyy-mm-dd") & "' and process='" & Mid(drawid.Text, 1, 1) & "' and dtno=" & Val(Mid(drawid.Text, 9, 2)) & "", cn, adOpenDynamic, adLockPessimistic
    If RS.EOF = True Then
        RS.AddNew
        RS("PDATE") = Format(PRDate, "YYYY-MM-DD")
        RS("PROCESS") = Mid(drawid.Text, 1, 1)
        RS("DTNO") = Val(Mid(drawid.Text, 9, 2))
        For i = 3 To RS.Fields.Count - 1
            RS(i) = 0  'Init all Param
        Next i
        RS("TOT_PT") = (Val(ptlen.Text) + Val(scrqty) + Val(txtBrkCntLen.Text)) / 1000
        RS("PT_OK") = Val(ptlen.Text) / 1000
        RS("DRAW_REJ") = drrej / 1000
        RS("PT_25") = PT_25
        RS("PT_BRK") = Val(nbreak)
        RS.update
       If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    Else
        If RS.State = 1 Then RS.Close
        StrUpdate = "TOT_PT=TOT_PT+" & (Val(ptlen.Text) + Val(scrqty) + Val(txtBrkCntLen.Text)) / 1000
        StrUpdate = StrUpdate & ",PT_OK=PT_OK+" & Val(ptlen.Text) / 1000
        StrUpdate = StrUpdate & ",DRAW_REJ=DRAW_REJ+" & drrej / 1000
        StrUpdate = StrUpdate & ",PT_25=PT_25+" & PT_25
        StrUpdate = StrUpdate & ",PT_BRK=PT_BRK+" & Val(nbreak)
        cn.Execute "Update PRODREPORT_DRAW set " & StrUpdate & " where pdate='" & Format(PRDate, "yyyy-mm-dd") & "' and process='" & Mid(drawid.Text, 1, 1) & "' and dtno=" & Val(Mid(drawid.Text, 9, 2)) & ""
    End If
'    For Realtime production Report
         If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    RS.Open "Select * from PRODREPORT where pdate='" & Format(PRDate, "yyyy-mm-dd") & "' and process='" & Mid(drawid.Text, 1, 1) & "'", cn, adOpenDynamic, adLockOptimistic
    If RS.EOF = True Then
        RS.AddNew
        RS("PDATE") = Format(PRDate, "YYYY-MM-DD")
        RS("PROCESS") = Mid(drawid.Text, 1, 1)
        For i = 2 To RS.Fields.Count - 1
            RS(i) = 0  'Init all Param
        Next i
        RS("TOT_PT") = (Val(ptlen.Text) + Val(scrqty) + Val(txtBrkCntLen.Text)) / 1000
        RS("PT_OK") = Val(ptlen.Text) / 1000
        RS("DRAW_REJ") = drrej / 1000
        RS("PT_25") = PT_25
        RS("PT_BRK") = Val(nbreak)
        RS.update
         If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    Else
       If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        StrUpdate = "TOT_PT=TOT_PT+" & (Val(ptlen.Text) + Val(scrqty) + Val(txtBrkCntLen.Text)) / 1000
        StrUpdate = StrUpdate & ",PT_OK=PT_OK+" & Val(ptlen.Text) / 1000
        StrUpdate = StrUpdate & ",DRAW_REJ=DRAW_REJ+" & drrej / 1000
        StrUpdate = StrUpdate & ",PT_25=PT_25+" & PT_25
        StrUpdate = StrUpdate & ",PT_BRK=PT_BRK+" & Val(nbreak)
        cn.Execute "Update PRODREPORT set " & StrUpdate & " where pdate='" & Format(PRDate, "yyyy-mm-dd") & "' and process='" & Mid(drawid.Text, 1, 1) & "'"
    End If
'    If rs.State = 1 Then rs.Close
'====================Use in All Sub where transactions take palce to the database
    cn.CommitTrans
    Xbool = False
    PTLoginRecFlag = True
'==================================================================================================

  Dim RICPreID, PrefPosEnd, PrefPosStart
  If RICPrefFlag = True Then
    If Stag = 1 Then
        PrefPosEnd = Round(Factor, 3)
        PrefPosStart = Round(Factor - (Val(ptlen.Text) + scrqty + drrej + Val(txtBrkCntLen.Text)) / 1000, 3)
    Else
        PrefPosEnd = Round(Val(txtFactor1), 3)
        PrefPosStart = Round(Val(txtFactor1) - (Val(ptlen.Text) + scrqty + drrej + Val(txtBrkCntLen.Text)) / 1000, 3)
    End If
    RICPreID = ""
    If Val(RICRod1DrawnKm) > 0 And ((Val(PrefPosStart) + Val(PrefPosEnd)) / 2) <= Val(RICRod1DrawnKm) Then
        RICPreID = RIDRodID1
    ElseIf Val(RICRod2DrawnKm) > 0 And ((Val(PrefPosStart) + Val(PrefPosEnd)) / 2) <= Val(RICRod2DrawnKm) Then
        RICPreID = RIDRodID2
    ElseIf Val(RICRod3DrawnKm) > 0 And ((Val(PrefPosStart) + Val(PrefPosEnd)) / 2) <= Val(RICRod3DrawnKm) Then
        RICPreID = RIDRodID3
    ElseIf Val(RICRod4DrawnKm) > 0 And ((Val(PrefPosStart) + Val(PrefPosEnd)) / 2) <= Val(RICRod4DrawnKm) Then
        RICPreID = RIDRodID4
    ElseIf Val(RICRod5DrawnKm) > 0 And ((Val(PrefPosStart) + Val(PrefPosEnd)) / 2) <= Val(RICRod5DrawnKm) Then
        RICPreID = RIDRodID5
    ElseIf Val(RICRod6DrawnKm) > 0 And ((Val(PrefPosStart) + Val(PrefPosEnd)) / 2) <= Val(RICRod6DrawnKm) Then
        RICPreID = RIDRodID6
    ElseIf Val(RICRod7DrawnKm) > 0 And ((Val(PrefPosStart) + Val(PrefPosEnd)) / 2) <= Val(RICRod7DrawnKm) Then
        RICPreID = RIDRodID7
    End If
    If RICPreID <> "" Then
        If Val(ptlen.Text) >= Minlen Then
            cn.Execute "Update ptbarcode set RICPreID='" & RICPreID & "' where barcodeid='" & txtBarcodeID.Text & "'"
            cn.Execute "Update pt set RICPreID='" & RICPreID & "' where PTID='" & PTID.Text & "'"           ''added on 14.09.2016
        End If
    End If
  End If


  If Val(txtTimeLoss.Text) >= 1 Then    'From 6 to 1 Minute 22012011 Deshmukh/Parag
        If rsTLoss.State = 1 Then
            rsTLoss.Close
            Set rsTLoss = Nothing
            End If
      rsTLoss.Open "select * from PTIME_LOSS", cn, adOpenDynamic, adLockOptimistic
      rsTLoss.AddNew
      rsTLoss("LDATE") = PRDate
      rsTLoss("FIBREID") = PTID.Text
      rsTLoss("AREA") = "PT"
      rsTLoss("LOSSTYPE") = UCase(txtTLReason.Text)
      rsTLoss("TIMELOSS") = Val(txtTimeLoss.Text)
      'rsTLoss("REMARK") = UCase(Trim(Mid(txtRemarks.Text, 25)))
      'rsTLoss("TLSUBREASON") = UCase(Trim(txtTLSubReason.Text))
      rsTLoss("OPR") = Mid(opr.Text, 1, 3)
      rsTLoss("MCNO") = Val(ptno.Text)
      rsTLoss.update
      If rsTLoss.State = 1 Then
            rsTLoss.Close
            Set rsTLoss = Nothing
            End If
  End If
'----------sunil 17042009

    If scrqty > 0 Then List2.AddItem scrqty / 1000 & " SCRAP"
    If ptlen > 0 Then List2.AddItem ptlen / 1000 & "   " & PTID.Text & "   " & txtBarcodeID.Text
    
    Dim rsz As New ADODB.Recordset
    If rsz.State = 1 Then
            rsz.Close
            Set rsz = Nothing
            End If
    rsz.Open "select * from pt_inst_ztpmd where barcodeid='" & UCase(txtBarcodeID.Text) & "'", cn, adOpenKeyset, adLockReadOnly
    If rsz.RecordCount > 0 Then
        If IsNull(rsz("ZTPMDID")) = False Then List2.AddItem "                    " & rsz("ZTPMDID") & "       ZTPMDID"
    End If
   If rsz.State = 1 Then
            rsz.Close
            Set rsz = Nothing
            End If
    
    If drrej > 0 Then List2.AddItem drrej / 1000 & " DRRej"
    If Val(txtBrkCntLen.Text) > 0 Then List2.AddItem Val(txtBrkCntLen.Text) / 1000 & " BRKLEN"
    
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    RS.Open "select * from matSTOCK where PROCESS='DRAW' and id ='" & drawid.Text & "' AND STOCKQTY > 0", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        ptype.Text = RS(6)
        balqty.Text = RS("stockqty") * 1000
    Else
        balqty = 0
        cn.Execute "Delete from pt_allocation where SPOOL_ID='" & drawid.Text & "'"
    End If
   If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    '------------------Give Msg for Star bobbins- 08092011-----------------------------
    If PTID.Text <> "" Then
        GenerateSampleID PTID.Text, Val(ptlen.Text)
    End If
    '------------------------------------------------------------------------------
        
    '------------------------------implemented for mbend sampling ids on 24.04.2014,given exe on 26.04.2014,02.05.2014
    If Mid(PTID.Text, 1, 1) = "E" Then
        If PTID.Text <> "" And Val(ptlen) > 0 Then
            If Stag = 1 Then
                GenerateSampleID_mbend PTID.Text, Round(Factor, 3)
            Else
                GenerateSampleID_mbend PTID.Text, Round(Val(txtFactor1.Text), 3)
            End If
        End If
    End If
    '------------------------------------------------------------------------------
    date2 = DATE1
    TIME2 = TIME1
    date3.Text = DATE1
    time3.Text = TIME1
    
    Dt = Date
    If nbreak > 0 Then
        MsgBox "BREAKID :" & BREAKID
    End If
    FlagPMD = False
    If HoldForInst = True Then
        Text1.Visible = True
        Text1.BackColor = &HFFC0C0
        Text1.Text = "Please take 220 mts sample for ZTPMD."
        MsgBox "", vbInformation
        Text1.Text = ""
        Text1.Visible = False
        HoldForInst = False
        FlagPMD = True
    End If
    If Trim(txtBarcodeID.Text) <> "" Then
        ZTPMD_New  'New Sampling Instruction for ZTPMD 28022008 by Sunil as per request from Maitryee/Nilesh
    End If
    iPTContBrkFlag = 0
    If nbreak > 0 And Val(ptlen) = 0 Then
        iPTContBrkCnt = iPTContBrkCnt + 1
    End If
    If Val(ptlen) / 1000 > 2.2 Then
        iPTContBrkCnt = 0
        iPTContBrkFlag = 0
    End If
    If Val(Mid(drawid.Text, 9, 2)) = 9 Or Val(Mid(drawid.Text, 9, 2)) = 6 Or Val(Mid(drawid.Text, 9, 2)) = 7 Or Val(Mid(drawid.Text, 9, 2)) = 8 Then
        MsgBox "DT 6,7,8,9 Fibre ,Send for Tight Wound PMD Measurement"
    End If
    
    txtScReason.ListIndex = -1
    Rem1.Text = ""
    txtBrkCntLen.Text = "" 'Refreshing BrkCntLen
    Remark.Text = ""
    'txtBarcodeID.Text = ""
    DRReason.Text = ""
    Screen.MousePointer = vbNormal
    If rsdate.State = 1 Then
            rsdate.Close
            Set rsdate = Nothing
            End If
    rsdate.Open "SELECT TO_CHAR(SYSDATE,'YYYY-MM-DD HH24:MI:SS') FROM DUAL", cn, adOpenKeyset, adLockReadOnly
        NDate = rsdate(0)
        PRDate = rsdate(0)
    rsdate.Close
     
    '==========Same Logic which is used in LostFocus Event 02.03.2006
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    RS.Open "SELECT * FROM DRAWFLAW WHERE FID = '" & drawid.Text & "' AND POSITION < " & Val(balqty) / 1000 & "  ORDER BY POSITION DESC ", cn, adOpenKeyset, adLockReadOnly
    TypeOfFlaw = ""
    NextFault = 1000000 'SUNIL 06032008
    If RS.RecordCount > 0 Then
        Do While RS.EOF = False
'            If rs("REMARK") <> "AIRLINE" Then          ''''''''''ADDED BY ABDUL ON 12.07.2013 AS PER DILIP/ASHISH
                Fault = Val(balqty) - RS(2) * 1000
                If Fault <= 5000 Then NextFault = Fault  'SUNIL 06032008
                Rem1.Text = "NEXT FAULT AT " & Fault & " M"
                If Fault > 1000 Or TagDrrej = 0 Then
                    If Fault < 2000 Then
                        If RS("REMARK") = "LUMPS" Or Mid(RS("REMARK"), 1, 1) = "L" Then
                            TypeOfFlaw = "L"
                            TypeOfRemark = "LUMPS"
                        End If
                        If RS("REMARK") = "AIRLINE" Or Mid(RS("REMARK"), 1, 1) = "A" Then
                            TypeOfFlaw = "A"
                            TypeOfRemark = "AIRLINE"
                        End If
                        If RS("REMARK") = "BFD" Or Mid(RS("REMARK"), 1, 1) = "B" Then
                            TypeOfFlaw = "B"
                            TypeOfRemark = "BFD"
                        End If
                        If RS("REMARK") = "SCD" Then
                            TypeOfFlaw = "C"
                            TypeOfRemark = "SCD"
                        End If
                    End If
                    Exit Do
                End If
'            End If
            RS.MoveNext
        Loop
    End If
   If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    
    'ChkFlaw
    stpos = Val(balqty) / 1000 - 2
    EdPos = Val(balqty) / 1000 + 2
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    RS.Open "select * from DRAWflaw where FID = '" & drawid.Text & "' and  position >'" & stpos & "' and position <'" & EdPos & "'", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        Do While RS.EOF = False
'            If rs("REMARK") <> "AIRLINE" Then          '''''''''''ADDED BY ABDUL ON 12.07.2013 AS PER DILIP/ASHISH
                If RS("REMARK") = "LUMPS" Or Mid(RS("REMARK"), 1, 1) = "L" Then
                    TypeOfFlaw = "L"
                    TypeOfRemark = "LUMPS"
                End If
                If RS("REMARK") = "BFD" Or Mid(RS("REMARK"), 1, 1) = "B" Then
                    TypeOfFlaw = "B"
                    TypeOfRemark = "BFD"
                End If
                If RS("REMARK") = "AIRLINE" Or Mid(RS("REMARK"), 1, 1) = "A" Then
                    TypeOfFlaw = "A"
                    TypeOfRemark = "AIRLINE"
                End If
                If RS("REMARK") = "SCD" Then
                    TypeOfFlaw = "C"
                    TypeOfRemark = "SCD"
                End If
                If RS("REMARK") = "PITCHCHANGE" Then
                    If TypeOfFlaw = "" Then
                        TypeOfFlaw = "P"
                        TypeOfRemark = "PITCHCHANGE"
                    End If
                End If
'            End If
            RS.MoveNext
        Loop
    End If
    If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
    If TypeOfFlaw = "" Then
        TypeOfFlaw = "U" 'For PT Machine Fault
        TypeOfRemark = "UNKNOWN"
    End If
    '============================================
    'Rajneesh 22 Feb-2008
    Call PitchChange_Cut          'For Cross Checking whether previous Pitchchange is cut or not
    If chkFlaw.value = 0 Then
        Call PitchChange_NextPos    'For Next Pitchchange postion Position
    End If
    Call PitchChange_Stop        'For confirming whether M/c is stop for pitchchange or not and also giving cutting instruction
    If TagDrrej = 1 Or chkFlaw.value = 1 Then
        txtFlawMsg.Visible = True
    Else
      '''code added by dinesh on ref 29012018 BAJIRAO SIR
       If UCase(plant) = "WALUJ" Then
         txtFlawMsg.Visible = False
       Else
         txtFlawMsg.Visible = True
       End If

        
    End If
    txtTotPT.Text = Val(drqty.Text) - Val(balqty.Text)
    '=============================================
    
    
    If Fault > 55000 Then Rem1.Text = ""
    PTscrap.value = 0
    Check1.value = 0
    chkMERej.value = 0
    ChkScratch.value = 0
    chkpmd.value = 0
    chkdocsample.value = 0      ''added by abdul for Reliability sample on 27.01.2014
    DRReason.Text = ""
    If Val(balqty) = 0 Then
        drawid = ""
        txtDrawBarcodeId.Text = vbNullString
    End If
    
    datasaved = True
    ''**logic written by ankit on 05.02.2014
    If drawid.Text = "" Then Exit Sub
    '''
    If RS11.State = 1 Then
            RS11.Close
            Set RS11 = Nothing
            End If
    RS11.Open "Select Min(pos) from pt where ptid like '" & drawid.Text & "%'", cn, adOpenKeyset, adLockReadOnly
        If IsNull(RS11(0)) = False Then txtFactor1.Text = RS11(0)
    If RS11.State = 1 Then
            RS11.Close
            Set RS11 = Nothing
            End If
    '-----------------------------------------------------------------
    
   
    '----------------------------Message for Cutting ,PTSOP,By Rajneesh 111007 20 pm-----------
    iPTContBrkFlag = 0
    If nbreak > 0 Then
        Call PTSOPNew
    End If
            
    Set_Length_Automatically (Fault)
    '----------------------------22052008--Reliability Sample
'    If Go_for_Sample = 1 Then
'        MsgBox "Please Take 200 mts Sample for Reliability", vbInformation
'    End If
'----------------------------End of Above-------------------------------------------------
    txtBarcodeID.Text = "" 'SUNIL24052008
    ptlen.Text = ""
    scrqty = 0
    drrej = 0
    txtTimeLoss.Text = ""
    txtSpeedLoss.Text = ""
    txtBobType.ListIndex = -1
    If Mid(drawid.Text, 1, 1) <> "M" And Mid(drawid.Text, 1, 1) <> "R" Then
        If Val(Val(balqty.Text) / 1000) > 0 And Val(Val(balqty.Text) / 1000) <= 50 Then
            'MsgBox "", vbInformation
            If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
            RS.Open "Select * from mcstatus where area='PT' and mcno='" & Val(ptno.Text) & "'", cn, adOpenKeyset, adLockReadOnly
            If RS.EOF = False Then
                If RS("Status") = "PT" Then
                    Call AllocateNextID
                End If
            End If
            If RS.State = 1 Then
            RS.Close
            Set RS = Nothing
            End If
        End If
    End If
    Call cmdOutScanPending_Click
    txtTLReason.ListIndex = -1
   Exit Sub

Storedata_Error:
   If Xbool = True Then cn.RollbackTrans
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & .Source & "' occurred in " & "Storedata" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description & err.Source, "frmptautomation", "Storedata", Erl)
   err.clear

End Sub
Private Sub AllocateNextID()
Dim RS As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
   On Error GoTo AllocateNextID_Error:
   
RS.Open "Select Spool_id,length from pt_allocation where mcno='" & Val(ptno.Text) & "' and priority='2'", cn, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then
    If IsNull(RS(0)) = False Then txtNextAllocateID.Text = RS(0)
    If IsNull(RS(1)) = False Then txtnextAllocateidLen.Text = RS(1)
    
    If plant = "WALUJ" Then
        Fill_ZTPMD_IDS (txtNextAllocateID.Text)
    End If
    
    rs1.Open "SELECT DRAW_BARCODEID FROM DRAW WHERE FID = '" & RS(0) & "'", cn, adOpenKeyset, adLockReadOnly
      If rs1.RecordCount > 0 Then
          txtNextAllocateID.Text = rs1("DRAW_BARCODEID")
      End If
    rs1.Close
Else
    If RS.State = 1 Then RS.Close
    Load frmPTAllocation
    frmPTAllocation.Hide
    frmPTAllocation.blnAutoAllocation = True
    frmPTAllocation.txtPTNO.Text = ptno.Text
    frmPTAllocation.Cal_AllocateID
    txtNextAllocateID.Text = frmPTAllocation.txtAllocateID.Text
    txtnextAllocateidLen.Text = frmPTAllocation.txtAllocateLen.Text
    Unload frmPTAllocation
End If
If RS.State = 1 Then RS.Close
   If txtNextAllocateID.Text <> "" Then
      MsgBox "New Spool ID allocated to you ", vbInformation
   End If
   Exit Sub

AllocateNextID_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "AllocateNextID" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "AllocateNextID", Erl)
   err.clear
End Sub
Private Sub Timer2_Timer()
'On Error GoTo Timer2_Timer_Error:
On Error Resume Next
Timer2.Enabled = False
If PTLoginRecFlag = True And Pub_PTLoginRec = True Then
    Call EventRead
End If
Timer2.Enabled = True


'
'    If Trim(checkbarcodeid) <> "" Then
'labelch:
'        checkp = InputBox("Enter Previous Barcodeid", "barcodeid")
'        If UCase(Trim(checkp)) <> UCase(Trim(checkbarcodeid)) Then
'            GoTo labelch
'        End If
'    End If
'    checkbarcodeid = ""

   Exit Sub

Timer2_Timer_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Timer2_Timer" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Timer2_Timer", Erl)
   err.clear
End Sub
Private Sub EventRead()
Dim RS As New ADODB.Recordset
Dim ErFsys As New Scripting.FileSystemObject
Dim ErInStream As Scripting.TextStream
Dim ErFStr As String, ErDescr As String
Dim ErDate1, ErTime1
Dim ErLength
Dim Erptlen As Double, ErBrkLen As Double
Dim ErD1, ErT1
Dim stLastLineDesc As String
Dim stFirstLineDesc As String
Dim iFlagFstr As Integer
Dim iFlagEventDelete As Integer
Dim i As Integer
Dim dtORAdate As Date
Dim dtpCDate As Date
Dim dtPRDate As Date
Dim lngLastLineNo As Long
Dim dblDateDiff As Long
Dim lngLineNumber As Long
Dim lngLastTimeDiff As Long
Dim CNT
Dim PTMcno As Integer
Dim stCompName As String
'Dim EventReadFlag As Boolean
  On Error GoTo EventRead_Error:
   
'stCompName = GetCompName
'If UCase$(Mid(stCompName, 1, 2)) = "PT" Then
'PTMcno = Round(Val(Mid(stCompName, 3, 2)))
'Else
'    PTLoginRecFlag = False
'    Exit Sub
'End If
PTMcno = Val(ptno.Text)
If PTMcno = 0 Then
    PTLoginRecFlag = False
    Exit Sub
End If
If PTMcno < 8 Then
    PTLoginRecFlag = False
    Exit Sub
End If

CNT = 0
dtORAdate = Date + Time
dtpCDate = Date + Time
RS.Open "Select LastLineno,LastLineDesc,FirstLineDesc,LastTimeDiff,sysdate from PTLOGINSTATUS where ptno='" & Val(PTMcno) & "'", cn, adOpenKeyset, adLockReadOnly
If RS.EOF = False Then
    lngLastLineNo = RS("LastLineno")
    If IsNull(RS("LastLineDesc")) = False Then stLastLineDesc = RS("LastLineDesc")
    If IsNull(RS("FirstLineDesc")) = False Then stFirstLineDesc = RS("FirstLineDesc")
    lngLastTimeDiff = RS("LastTimeDiff")
    dtORAdate = RS("sysdate")
    dtpCDate = Date + Time
Else
    PTLoginRecFlag = False
End If
RS.Close

iFlagEventDelete = 0
lngLineNumber = 0
dblDateDiff = (dtORAdate - dtpCDate) * 86400

Dim iNewEntryFlag As Integer
iNewEntryFlag = 0 'Init
If ErFsys.FileExists(McPath & medakfilename) = True Then
    Set ErInStream = ErFsys.OpenTextFile(McPath & medakfilename, ForReading)
    Do While ErInStream.AtEndOfStream = False
        lngLineNumber = ErInStream.Line
        ErFStr = ErInStream.ReadLine
        If lngLineNumber = 1 Then
            If ErFStr = stFirstLineDesc Then 'If First Event is same as previous read then event table is not deleted
                If lngLastLineNo > 1 Then
                    For i = 2 To lngLastLineNo
                        lngLineNumber = i
                        ErInStream.SkipLine
                    Next i
                    If ErInStream.AtEndOfStream = False Then
                        lngLineNumber = ErInStream.Line
                        ErFStr = ErInStream.ReadLine
                    Else
                        Exit Do
                    End If
                End If
            Else
                stFirstLineDesc = ErFStr
            End If
        End If
        iNewEntryFlag = 1
        iFlagFstr = 0
        If InStr(ErFStr, "MACHINE Start") = 1 Then
            iFlagFstr = 1
        ElseIf InStr(ErFStr, "FAULT DRAW CUTTING INSTRUCTION") = 1 Then
            iFlagFstr = 1
        ElseIf InStr(ErFStr, "FAULT MACHINE STOP BY") = 1 Then
            iFlagFstr = 1
        ElseIf InStr(ErFStr, "FAULT MINIMUM STRAIN ON PROOF TEST") = 1 Then
            iFlagFstr = 1
        ElseIf InStr(ErFStr, "COVER OPEN") = 1 Then
            iFlagFstr = 1
        ElseIf InStr(ErFStr, "FAULT THREADING NOT GOOD TO REW. MODE") = 1 Then
            iFlagFstr = 1
        End If
        If iFlagFstr = 1 Then
            ErDate1 = CDate(Trim(Mid(ErFStr, 44, 10)))
            ErTime1 = Mid(ErFStr, 57, 12)
            ErD1 = DateSerial(Year(ErDate1), Month(ErDate1), Day(ErDate1))
            ErT1 = TimeSerial(Val(Mid(ErTime1, 1, 2)), Val(Mid(ErTime1, 5, 2)), Val(Mid(ErTime1, 10, 2)))
            ErLength = Trim(Mid(ErFStr, 70, 5))
            ErLength = Replace(ErLength, " ", "0")
            ErDescr = Trim(Mid(ErFStr, 1, 42))
            dtPRDate = ErD1 + ErT1
            dtPRDate = dtPRDate + (dblDateDiff / 86400)
            Erptlen = Val(ErLength)
            ErBrkLen = Val(Trim(Mid(ErFStr, 75, 5)))
            cn.Execute ("INSERT INTO PTLOGIN (PTNO, DESCR, PDATE, VAL1, VAL2, OPR ) VALUES  (" & Val(PTMcno) & ",'" & ErDescr & "',to_date('" & Format(dtPRDate, "YYYY-MM-DD HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss')," & Erptlen & "," & ErBrkLen & ",'" & opr.Text & "')")
            CNT = CNT + 1
            If CNT > 5 Then
                Exit Do
            End If
            DoEvents
        End If
    Loop
    ErInStream.Close
    If iNewEntryFlag = 1 Then
        lngLastLineNo = lngLineNumber
        stLastLineDesc = ErFStr
        lngLastTimeDiff = dblDateDiff
        If RS.State = 1 Then RS.Close
        RS.Open "Select * from PTLOGINSTATUS where ptno='" & Val(PTMcno) & "'", cn, adOpenDynamic, adLockOptimistic
        If RS.EOF = True Then
            RS.AddNew
            RS("ptno") = Val(PTMcno)
            RS("LastLineno") = lngLastLineNo
            RS("LastLineDesc") = stLastLineDesc
            RS("FirstLineDesc") = stFirstLineDesc
            RS("LastTimeDiff") = lngLastTimeDiff
            RS.update
        Else
            RS("LastLineno") = lngLastLineNo
            RS("LastLineDesc") = stLastLineDesc
            RS("FirstLineDesc") = stFirstLineDesc
            RS("LastTimeDiff") = lngLastTimeDiff
            RS.update
        End If
        If RS.State = 1 Then RS.Close
    End If
End If
   
   Exit Sub

EventRead_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "EventRead" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "EventRead", Erl)
   err.clear
End Sub

Private Sub Timer3_Timer()
   'On Error GoTo Timer3_Timer_Error:
    If drawid.Text <> "" Then
        plant = Get_Plant(drawid.Text)
    Else
        If Mid(GetIPAddresses, 8, 2) = 12 Or Mid(GetIPAddresses, 8, 2) = 13 Then
            plant = "SHENDRA"
        Else
            plant = "WALUJ"
        End If
        plant = GET_IP_PLANT
   End If
   
   If plant = "WALUJ" Then
        If srnogotf = 1 And (Len(txtBarcodeID.Text) < 7 And Len(txtBarcodeID.Text) > 0) Then
            MsgBox "Enter barcodeid from scanner only", vbCritical
            txtBarcodeID.Text = ""
        End If
        If srnogotf = 1 And (Len(txtBarcodeID.Text) > 7) Then
            MsgBox "Enter barcodeid from scanner only", vbCritical
            txtBarcodeID.Text = ""
        End If
        If srnogotf = 1 And (Len(txtPTOut.Text) < 7 And Len(txtPTOut.Text) > 0) Then
            MsgBox "Enter barcodeid from scanner only", vbCritical
            txtPTOut.Text = ""
        End If
        If srnogotf = 1 And (Len(txtPTOut.Text) > 7) Then
            MsgBox "Enter barcodeid from scanner only", vbCritical
            txtPTOut.Text = ""
        End If
        If txtPTOut.SelLength > 0 Then
            txtPTOut.Text = ""
        End If
    Else
        If srnogotf = 1 And (Len(txtBarcodeID.Text) < 10 And Len(txtBarcodeID.Text) > 0) Then
            MsgBox "Enter barcodeid from scanner only", vbCritical
            txtBarcodeID.Text = ""
        End If
        If srnogotf = 1 And (Len(txtBarcodeID.Text) > 10) Then
            MsgBox "Enter barcodeid from scanner only", vbCritical
            txtBarcodeID.Text = ""
        End If
        If srnogotf = 1 And (Len(txtPTOut.Text) < 10 And Len(txtPTOut.Text) > 0) Then
            MsgBox "Enter barcodeid from scanner only", vbCritical
            txtPTOut.Text = ""
        End If
        If srnogotf = 1 And (Len(txtPTOut.Text) > 10) Then
            MsgBox "Enter barcodeid from scanner only", vbCritical
            txtPTOut.Text = ""
        End If
        If txtPTOut.SelLength > 0 Then
            txtPTOut.Text = ""
        End If
    End If

Exit Sub

Timer3_Timer_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Timer3_Timer" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Timer3_Timer", Erl)
   err.clear
End Sub

Private Sub txtBarcodeID_Change()
    lock_txtBarcodeID
End Sub

Private Sub txtBarcodeID_Click()
    lock_txtBarcodeID
End Sub

Private Sub txtBarcodeID_GotFocus()
    lock_txtBarcodeID
End Sub

Private Sub txtBarcodeID_KeyDown(KeyCode As Integer, shift As Integer)
    lock_txtBarcodeID
End Sub

Private Sub lockforbarcodewaluj()

 If txtProducttype.Text = "E1" Or txtProducttype.Text = "L2" Or txtProducttype.Text = "B" Then
        If Mid(txtBarcodeID.Text, 1, 1) <> "2" And Mid(txtBarcodeID.Text, 1, 1) <> "6" Then
            MsgBox "Please Select 2 or 6 Intitial Digit barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "E2" Or txtProducttype.Text = "E3" Then
         If Mid(txtBarcodeID.Text, 1, 1) <> "5" Then
            MsgBox "Please Select 5 Digit Intital Digit barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "A" Or txtProducttype.Text = "Q" Then
         If Mid(txtBarcodeID.Text, 1, 1) <> "7" Then
            MsgBox "Please Select 7 Initial Digit barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "L1" Then
        If Mid(txtBarcodeID.Text, 1, 1) <> "4" Then
            MsgBox "Please Select 4 Intital Digit barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "" Then
            MsgBox "Please Click On CHECKPTID Button..", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        
    Else
        
        
    End If
   

End Sub

Private Sub lockforbarcode()
    If txtProducttype.Text = "E1" Then
        If Mid(txtBarcodeID.Text, 1, 3) <> "SAE" Then
            MsgBox "Please Select SAE barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "E2" Then
         If Mid(txtBarcodeID.Text, 1, 3) <> "SAR" Then
            MsgBox "Please Select SAR barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "B" Then
         If Mid(txtBarcodeID.Text, 1, 3) <> "SAE" Then
            MsgBox "Please Select SAE barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "L1" Then
        If Mid(txtBarcodeID.Text, 1, 3) <> "SAL" Then
            MsgBox "Please Select SAL barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "SRB" Then
        If Mid(txtBarcodeID.Text, 1, 3) <> "SAS" Then
            MsgBox "Please Select SAS barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "L3" Then
        If Mid(txtBarcodeID.Text, 1, 3) <> "SAL" Then
            MsgBox "Please Select SAL barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
    ElseIf txtProducttype.Text = "S" Then
        If Mid(txtBarcodeID.Text, 1, 3) <> "SA0" Then
            MsgBox "Please Select SA0 barcode id", vbCritical
            txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        End If
     ElseIf txtProducttype.Text = "" Then
            MsgBox "Please Click On CHECKPTID Button..", vbCritical
             txtBarcodeID.Text = ""
            txtBarcodeID.SetFocus
            lockbarcedflag = False
            Exit Sub
        
    Else
        
        
    End If
    
End Sub

Private Sub txtBarcodeid_KeyPress(KeyAscii As Integer)
    lock_txtBarcodeID
End Sub

Private Sub lock_txtBarcodeID()
'       Clipboard.SetText ("")
 '      srnogotf = 1
End Sub

Private Sub txtBarcodeid_LostFocus()
'    If (Mid(txtBarcodeID.Text, 1, 1)) <> "7" And (Mid(txtBarcodeID.Text, 1, 1)) <> "8" And (Mid(txtBarcodeID.Text, 1, 1)) <> "9" Then
'        MsgBox "Enter Valid Barcodeid", vbCritical
'        txtBarcodeID.Text = ""
'        Exit Sub
'    End If
    srnogotf = 0
    If Trim(txtBarcodeID.Text) <> "" Then
        checkbarcodeid = UCase(Trim(txtBarcodeID.Text))
       
    End If
End Sub

Private Sub txtBarcodeID_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    lock_txtBarcodeID
End Sub
Function GetPos()
On Error GoTo GetPos_Error:
thisspool = 0
ptlen1 = 0
drkms = 0
Factor = 0
Dim DATA(13, 2)
If RS.State = 1 Then
RS.Close
Set RS = Nothing
End If
RS.Open "Select FID,DRLEN from draw where matno IN('330000183','330000155','330000126') and fid like '" & Mid(drawid.Text, 1, 8) & "%' ORDER BY SUBSTR(FID,11,1)", cn, adOpenKeyset, adLockReadOnly
If RS.RecordCount > 0 Then
   ' spoolno = Val(Mid(drawidd.Text, 11, 1))
'------------------- Code Added by dinesh ref mohhamad sir pos issue
   If UCase(plant) = "SHENDRA" Then
        Dim retval
          retval = Mid(drawid.Text, 11, 1)
          retval = UCase(retval)
      
      If retval = "A" Or retval = "B" Or retval = "C" Or retval = "D" Or retval = "E" Or retval = "F" Then
              retval = UCase(retval)
              spoolno = Asc(retval)
              spoolno = Val(spoolno - 55)
         
      Else
          spoolno = Val(Mid(drawid.Text, 11, 1))
      End If
    Else
            spoolno = Val(Mid(drawid.Text, 11, 1))
    
    End If
        
   
    For i = 0 To RS.RecordCount - 1
        DATA(i, 0) = RS(0)
        DATA(i, 1) = RS(1)
        If i < spoolno Then
            drkms = drkms + Val(DATA(i, 1)) 'Draw Kms Before this Piece
        End If
        RS.MoveNext
    Next i
End If
thisspool = Val(DATA(spoolno, 1))
If RS.State = 1 Then
RS.Close
Set RS = Nothing
End If

ptlen1 = 0
Factor = (thisspool - ptlen1 + drkms)
   
Exit Function

GetPos_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "GetPos" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "GetPos", Erl)
   err.clear
End Function
Sub ChkSpoolRejection(nid As String)
    Dim rslogic As New ADODB.Recordset
   On Error GoTo ChkSpoolRejection_Error:
   
    flag_rejection = False
    rslogic.Open "select * from MAINSPOOLREJECTION where spoolid = '" & nid & "'", cn, adOpenKeyset, adLockReadOnly
    If rslogic.EOF = False Then
        flag_rejection = True
    End If
    rslogic.Close
   
   Exit Sub

ChkSpoolRejection_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "ChkSpoolRejection" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "ChkSpoolRejection", Erl)
   err.clear
End Sub

Private Sub PTSOPNew()
'============================11-Oct-2007====================================================
' New SOP as per Maitreyee/Bruno
' 1. Start
' 2. 3 Brks < 2.2 Kms -----Scrap 5 Kms
' 3. Restart
' 4. 1 Brks < 2.2 Kms -----Scrap 10 Kms
' 5. Restart
' 6. 1 Brks < 2.2 Kms -----Scrap 10 Kms
' 7. Continue Step 5,6 until good length > 2.2 Kms
' 8. if any length > 2.2 then restart from Step 1
'===========================================================================================
'Change in SOP as per request by Samarth and approve by G.C.Mishra/Ajay Gupta
   On Error GoTo PTSOPNew_Error:
   
If iPTContBrkCnt > 3 Then
    iPTContBrkFlag = 2
    'MsgBox "1 Cont. Brk(After 3 Brks) Scrap 10 KM", vbExclamation
    MsgBox "1 Cont. Brk(After 3 Brks) Scrap 4.2 KM", vbExclamation
ElseIf iPTContBrkCnt = 3 Then
    iPTContBrkFlag = 1
    'MsgBox "3 Brks ,Scrap 5 KM", vbExclamation
    MsgBox "3 Brks ,Scrap 4.2 KM", vbExclamation
End If
   
   Exit Sub

PTSOPNew_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "PTSOPNew" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "PTSOPNew", Erl)
   err.clear
End Sub





Private Sub txtBobType_Click()
    '---------------NO Crellin/YUQIU Spool for towers 9,10,11,12,14,15,16,17
    
'    If txtBobType.Text = "CRELLIN" Or txtBobType.Text = "YUQIU" Then '' lock is disabled by ankit on 08.11.2013 as per discussion with pravin
'        If Val(Mid(drawid.Text, 9, 2)) = 9 Or Val(Mid(drawid.Text, 9, 2)) = 10 Or Val(Mid(drawid.Text, 9, 2)) = 11 Or Val(Mid(drawid.Text, 9, 2)) = 12 Or Val(Mid(drawid.Text, 9, 2)) = 14 Or Val(Mid(drawid.Text, 9, 2)) = 15 Or Val(Mid(drawid.Text, 9, 2)) = 16 Or Val(Mid(drawid.Text, 9, 2)) = 17 Then
'            MsgBox "Crellin or YUQIU spool not allowed for tower 9,10,11,12,14,15,16,17"
'            txtBobType.ListIndex = -1
'        End If
'    End If
End Sub

Private Sub txtMultimodeID_LostFocus()
   On Error GoTo txtMultimodeID_LostFocus_Error:
   
    If txtMultimodeID.Text <> "" Then
        txtMultimodeID.Text = UCase$(Trim(txtMultimodeID.Text))
        If Mid(txtMultimodeID.Text, 1, 1) = "M" Or Mid(txtMultimodeID.Text, 1, 1) = "R" Then
            drawid.Text = UCase$(Trim(txtMultimodeID.Text))
            drawid_lostfocus
            cmdUnloadME.Enabled = False
        Else
            MsgBox "Only Multimode id are allow here", vbExclamation
            txtMultimodeID.Text = ""
        End If
    End If
   
   Exit Sub

txtMultimodeID_LostFocus_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "txtMultimodeID_LostFocus" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "txtMultimodeID_LostFocus", Erl)
   err.clear
End Sub
Private Sub Set_Horus_Parameters(ByVal SetLen As Double)

''    #Region ;**** Directives created by AutoIt3Wrapper_GUI ****
''#AutoIt3Wrapper_UseAnsi=y
''#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
    Dim oAuto As New AutoItX3
   On Error GoTo Set_Horus_Parameters_Error:
   
    Set oAuto = CreateObject("AUTOITX3.CONTROL")
    'Call oAuto.BlockInput(1)
    Call oAuto.Opt("WinWaitDelay", 100)
    Call oAuto.Opt("WinTitleMatchMode", 4)
    Call oAuto.Opt("WinDetectHiddenText", 1)
    Call oAuto.Opt("MouseCoordMode", 0)
    Call oAuto.WinWait("HORUS 4.0 Tower", "")
    If Not oAuto.WinActive("HORUS 4.0 Tower", "") Then Call oAuto.WinActivate("HORUS 4.0 Tower", "")
    Call oAuto.WinWaitActive("HORUS 4.0 Tower", "")
'    Call oAuto.MouseMove(28, 29)
'    Call oAuto.MouseDown("left")
'    Call oAuto.MouseUp("left")
'    Call oAuto.MouseMove(41, 67)
'    Call oAuto.MouseDown("left")
'    Call oAuto.MouseUp("left")
    Call oAuto.Send("{ALTDOWN}s{ALTUP}c")
    '''code added by dinesh on ref bajirao sir on 26042018
    If UCase(plant) = "SHENDRA" Then
       Call oAuto.MouseMove(515, 210)
    Else
       Call oAuto.MouseMove(515, 198)
    End If
    
   
    Call oAuto.MouseDown("left")
    Call oAuto.MouseUp("left")
    
'    Call oAuto.Send(SetLen)
'    Call oAuto.Send(vbCrLf)
'
'    Call oAuto.Send("{ESC}")


    'Call oAuto.MouseMove(227, 96)
    'Call oAuto.MouseDown("left")
    'Call oAuto.MouseUp("left")
    'Call oAuto.BlockInput(0)
   
   Exit Sub

Set_Horus_Parameters_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Set_Horus_Parameters" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Set_Horus_Parameters", Erl)
   err.clear
End Sub
Private Sub ZTPMD_New()
    Dim PTQTY As Double
    Dim SampleTaken As Integer
    Dim rsnew As New ADODB.Recordset
    SampleTaken = 0
   On Error GoTo ZTPMD_New_Error:
   PTQTY = Val(drqty.Text) - Val(balqty.Text)
   
'   If Mid(drawid.Text, 1, 1) = "D" Then   'After every Bobbin ZTPMD Sample Saurav's mail 17112009
'        RSNEW.Open "select * from pt_inst_ztpmd where SUBSTR(FID,1,11) = '" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
'        If RSNEW.RecordCount = 0 Then SampleCnt = 1 Else SampleCnt = RSNEW.RecordCount
'        SampleCnt = SampleCnt + 1
'        RSNEW.Close
'        fm_len = 0  'Not Relavant
'        to_len = 500000 'Not Relavant
'        Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, SampleCnt
'        S1 = 1
'        MsgBox fm_len & "-" & to_len
'   End If
    
   
   If Mid(drawid.Text, 1, 1) = "D" Or Mid(drawid.Text, 1, 1) = "Q" Or Mid(drawid.Text, 1, 1) = "A" Then
        If Val(balqty.Text) <= 50000 Then
            fm_len = (Val(drqty.Text) - 50000)
            to_len = Val(drqty.Text)
            Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "BOT"
            S10 = 1
            MsgBox fm_len & "-" & to_len
        End If
        
        If S10 = 0 Then  'If Bot end samle has been taken then do not go into this loop.
            If Val(PTQTY) <= 50000 And Val(S1) = 0 Then
                fm_len = 0
                to_len = 50000
                Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "1"
                S1 = 1
                MsgBox fm_len & "-" & to_len
            ElseIf Val(PTQTY) <= 175000 And Val(S2) = 0 Then
                fm_len = 50000
                to_len = 175000
                Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "2"
                S2 = 1
                MsgBox fm_len & "-" & to_len
            ElseIf Val(PTQTY) <= 300000 And Val(S3) = 0 Then
                fm_len = 150000
                to_len = 200000
                Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "3"
                S3 = 1
                MsgBox fm_len & "-" & to_len
            ElseIf Val(PTQTY) <= 425000 And Val(S4) = 0 Then
                fm_len = 200000
                to_len = 250000
                Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "4"
                S4 = 1
                MsgBox fm_len & "-" & to_len
            ElseIf Val(PTQTY) <= 550000 And Val(S5) = 0 Then
                fm_len = 250000
                to_len = 300000
                Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "5"
                S5 = 1
                MsgBox fm_len & "-" & to_len
            End If
        End If
   End If
   
   
    
    
        
   If Mid(drawid.Text, 1, 1) <> "D" And Mid(drawid.Text, 1, 1) <> "Q" And Mid(drawid.Text, 1, 1) <> "A" Then
'        If Val(PTQTY) <= 26000 Then     'For Top 25 kms ,Sample is necesssary 15062009 for all pieces / DISABLED SUNIL/PAVAN 09102009 NOT REQUIRED NOW
'            Piece = Mid(PTID.Text, 12, 1)
'            fm_len = 0
'            to_len = 25000
'            Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, Piece
'            SampleTaken = 1
'        End If
'        If SampleTaken = 0 Then  'For Bot 25 kms ,Sample is necesssary 15062009 for all pieces
'            If Val(balqty) <= 26000 Then
'                Piece = Mid(PTID.Text, 12, 1)
'                fm_len = (Val(drqty.Text) - 25000)
'                to_len = Val(drqty.Text)
'                Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, Piece
'                SampleTaken = 1
'            End If
'        End If
        
      
        If SampleTaken = 0 Then
            If Val(drqty.Text) <= 100000 Then
                If Mid(PTID.Text, 12, 1) = "A" Then
                    fm_len = 0
                    to_len = 100000
                    Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "TOP"
                End If
            ElseIf Val(drqty.Text) >= 50000 And Val(drqty.Text) <= 100000 Then
                    fm_len = 0
                    to_len = 100000
                    Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "TOP1"   '2 Samples in 100 kms Sunil 30032010
            ElseIf Val(drqty.Text) < 150000 Then
                If Mid(PTID.Text, 12, 1) = "A" Then
                        fm_len = 0
                        to_len = 100000
                        Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "TOP"
                ElseIf Val(balqty.Text) <= 50000 Then
                    'MsgBox "Last 50"
                    fm_len = (Val(drqty.Text) - 50000)
                    
                    to_len = Val(drqty.Text)
                    Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "BOT"
                End If
            Else
                If Mid(PTID.Text, 12, 1) = "A" Then
                    If lblMSStatus.Caption = "PREFORM_END" Then
                        fm_len = 0
                        to_len = 50000
                    Else
                        fm_len = 0
                        to_len = (Val(drqty.Text) - 50000)
                        If Val(drqty.Text) <= 50000 Then to_len = Val(drqty.Text)
                    End If
                    Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "TOP"
                ElseIf Val(balqty.Text) <= 50000 Then
                    'MsgBox "Last 50"
                    fm_len = (Val(drqty.Text) - 50000)
                    to_len = Val(drqty.Text)
                    Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "BOT"
                Else
                    If lblMSStatus.Caption = "PREFORM_END" Then
                        If TagDrrej = 0 Then    'If Last Entry is of Draw Rejection
                            If Abs(Val(balqty.Text) - Val(drqty.Text) / 2) <= 25000 Then
                                'MsgBox "Middle"
                                fm_len = 50000
                                to_len = (Val(drqty.Text) - 50000)
                                Insert_into_ZTPMD_Table txtBarcodeID.Text, PTID.Text, fm_len, to_len, "MID"
                            End If
                        End If
                    End If
                End If
            End If
        End If  'SampleTaken
    End If
    
    Exit Sub

ZTPMD_New_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "ZTPMD_New" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "ZTPMD_New", Erl)
   err.clear
End Sub
Private Sub Insert_into_ZTPMD_Table(Barcodeid As String, FID As String, ByVal fm_len As Double, ByVal to_len As Double, ByVal Sample As String)
    Dim rsnew As New ADODB.Recordset
    Dim rszz As New ADODB.Recordset
    Dim rschid As New ADODB.Recordset
   On Error GoTo Insert_into_ZTPMD_Table_Error:
   
    FlagPMD = False
    Dim ztpmdId
    'Timer2.Enabled = False
    rsnew.Open "select * from pt_inst_ztpmd where SUBSTR(FID,1,11) = '" & Mid(PTID.Text, 1, 11) & "' and sample = '" & Sample & "'", cn, adOpenDynamic, adLockOptimistic
    If rsnew.EOF = True Then
        If TagDrrej = 0 Then
            Text1.Visible = True
            Text1.BackColor = &HFFC0C0
            Text1.Text = "Please take 220 mts sample for ZTPMD."
            MsgBox "", vbInformation
            Text1.Text = ""
            Text1.Visible = False
            FlagPMD = True
        Else
            HoldForInst = True
        End If
        
ask_ztpmd_Id:
        ztpmdId = ""
        frmInputBoxZTPMD.Label2.Visible = False
        frmInputBoxZTPMD.txtrelId.Visible = False
        frmInputBoxZTPMD.Label1.Visible = True
        frmInputBoxZTPMD.txtZtpmdId.Visible = True
        frmInputBoxZTPMD.Show vbModal
        ztpmdId = Trim(UCase(frmInputBoxZTPMD.txtZtpmdId.Text))
        frmInputBoxZTPMD.txtZtpmdId.Text = ""
        frmInputBoxZTPMD.Hide
        
        If Trim(ztpmdId) = "" Or ztpmdId = 0 Then
            GoTo ask_ztpmd_Id
        End If
        
        rschid.Open "select * from pt_inst_ztpmd where ztpmdid='" & Trim(ztpmdId) & "'", cn, adOpenKeyset, adLockReadOnly
        If rschid.RecordCount > 0 Then
            MsgBox "This ZTPMD ID Allready exists, Enter another ztpmd id", vbCritical
            rschid.Close
            GoTo ask_ztpmd_Id
        End If
        rschid.Close
        
        rsnew.AddNew
        rsnew("barcodeid") = UCase(txtBarcodeID.Text)
        rsnew("FID") = PTID.Text
        rsnew("FM_LEN") = Round(Val(fm_len), 3)
        rsnew("TO_LEN") = Round(Val(to_len), 3)
        rsnew("PDATE") = getcurdate
        rsnew("SAMPLE") = Sample
        rsnew("ZTPMDID") = UCase(ztpmdId)
        rsnew("PLANT") = Get_Plant(PTID.Text)
        rsnew.update
        
        rszz.Open "select * from pt_inst_ztpmd where barcodeid='" & UCase(txtBarcodeID.Text) & "'", cn, adOpenKeyset, adLockReadOnly
        If rszz.RecordCount > 0 Then
            If IsNull(rszz("ZTPMDID")) = False Then List2.AddItem "                    " & rszz("ZTPMDID") & "       ZTPMDID"
        End If
        rszz.Close
        
    End If
    rsnew.Close
    'Timer2.Enabled = True
   
   Exit Sub

Insert_into_ZTPMD_Table_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Insert_into_ZTPMD_Table" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Insert_into_ZTPMD_Table", Erl)
   err.clear
End Sub
Function Check_Sample_For_Reliability(Barcodeid As String, FID As String) As Boolean
    Dim RsRel As New ADODB.Recordset
    Dim RsX As New ADODB.Recordset
    Dim rschid As New ADODB.Recordset
    Dim rszz As New ADODB.Recordset
    Dim relId
   On Error GoTo Check_Sample_For_Reliability_Error:
    Check_Sample_For_Reliability = False
    Dim draw_date
    Dim draw_date1
    'Sunil 13112009 as per Sachin
    'Give Instruction only for bobbins where drawdate is Today's Date
    '''code modified by dinesh kanse on ref sachin sir on 29012018
    RsRel.Open "select pdate,twno from draw where fid='" & drawid.Text & "'", cn, adOpenKeyset, adLockReadOnly
    If RsRel.RecordCount > 0 Then
        draw_date = Format(RsRel(0), "yyyy-mm-dd")
        draw_date1 = Format(RsRel(0), "yyyy-mm-dd")
    End If
    RsRel.Close
    
    Cur_date = Format(getcurdate, "yyyy-mm-dd")
    
    
    ''''CODE ADDED BY DINESH ON REF Sachin Girawale ON 19012018

    If plant = "WALUJ" Then
      If draw_date = Cur_date Then
      
      Else
        If RsRel.State = 1 Then
         RsRel.Close
         Set RsRel = Nothing
        End If
            RsRel.Open "Select pdate from pt_in where fid='" & drawid & "'", cn, adOpenKeyset, adLockReadOnly
             If RsRel.RecordCount > 0 Then
                 draw_date = Format(RsRel(0), "yyyy-mm-dd")
             End If
              If RsRel.State = 1 Then
                 RsRel.Close
                 Set RsRel = Nothing
                End If
      End If
    End If
    '''''''''''''''''''''''''''''''''''''''added for reliability sample on 07.12.2013 by abdul,27.01.2014,05022016 for trail
''    If plant = "SHENDRA" Then
''        If draw_date = Cur_date Then
''            RsRel.Open "Select * from pt_rel_sample where to_char(edate,'yyyy-mm-dd') = to_char(sysdate,'yyyy-mm-dd')and twno = '" & Val(Mid(drawid.Text, 9, 2)) & "'", cn, adOpenKeyset, adLockReadOnly
''            If RsRel.RecordCount <= 0 Then
''                Check_Sample_For_Reliability = True
''                RsX.Open "Select * from pt_rel_sample", cn, adOpenDynamic, adLockOptimistic
''                RsX.AddNew
''                RsX("barcodeid") = txtBarcodeID.Text
''                RsX("fid") = PTID.Text
''                RsX("twno") = Val(Mid(drawid.Text, 9, 2))
''                RsX("edate") = getcurdate
''                RsX("PTMC") = Val(ptno.Text)
''                RsX("PLANT") = Get_Plant(PTID.Text)
''                RsX.update
''                RsX.Close
''            End If
''            RsRel.Close
''        End If
''    Else
        If draw_date = Cur_date Then
        If UCase(plant) = "WALUJ" Then
         RsRel.Open "Select * from pt_rel_sample where to_char(drawdate,'yyyy-mm-dd') ='" & Format(draw_date1, "yyyy-mm-dd") & "' and twno = '" & Val(Mid(drawid.Text, 9, 2)) & "'", cn, adOpenKeyset, adLockReadOnly
        Else
        
         RsRel.Open "Select * from pt_rel_sample where to_char(edate,'yyyy-mm-dd') = to_char(sysdate,'yyyy-mm-dd')and twno = '" & Val(Mid(drawid.Text, 9, 2)) & "'", cn, adOpenKeyset, adLockReadOnly
        End If
      
      
            If RsRel.RecordCount <= 0 Then
    
ask_Reliability_Id:
                relId = ""
                frmInputBoxZTPMD.Label1.Visible = False
                frmInputBoxZTPMD.txtZtpmdId.Visible = False
                frmInputBoxZTPMD.Label2.Visible = True
                frmInputBoxZTPMD.txtrelId.Visible = True
                frmInputBoxZTPMD.Show vbModal
                relId = Trim(UCase(frmInputBoxZTPMD.txtrelId.Text))
                frmInputBoxZTPMD.txtrelId.Text = ""
                frmInputBoxZTPMD.Hide
    
                If Trim(relId) = "" Or relId = 0 Then
                    GoTo ask_Reliability_Id
                End If
    
                rschid.Open "select * from pt_rel_sample where relid='" & Trim(relId) & "'", cn, adOpenKeyset, adLockReadOnly
                If rschid.RecordCount > 0 Then
                    MsgBox "This Reliability ID Allready exists, Enter another Reliability id", vbCritical
                    rschid.Close
                    GoTo ask_Reliability_Id
                End If
                rschid.Close
                Check_Sample_For_Reliability = True
    
                    RsX.Open "Select * from pt_rel_sample", cn, adOpenDynamic, adLockOptimistic
                    RsX.AddNew
                    RsX("barcodeid") = txtBarcodeID.Text
                    RsX("fid") = PTID.Text
                    RsX("twno") = Val(Mid(drawid.Text, 9, 2))
                    RsX("edate") = getcurdate
                    RsX("PTMC") = Val(ptno.Text)
                    RsX("RELID") = UCase(relId)
                    RsX("PLANT") = Get_Plant(PTID.Text)
                    RsX("drawdate") = draw_date1
                    RsX.update
                        rszz.Open "select * from pt_rel_sample where barcodeid='" & UCase(txtBarcodeID.Text) & "'", cn, adOpenKeyset, adLockReadOnly
                        If rszz.RecordCount > 0 Then
                            If IsNull(rszz("RELID")) = False Then List2.AddItem "                    " & rszz("RELID") & "       DOCID"
                        End If
                        rszz.Close
                    RsX.Close
            End If
            RsRel.Close
        End If
''    End If
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
   Exit Function

Check_Sample_For_Reliability_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Check_Sample_For_Reliability" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Check_Sample_For_Reliability", Erl)
   err.clear
End Function

Private Sub Set_Standard_Length(Fault As Double)
   On Error GoTo Set_Standard_Length_Error:
  
    Dim Bendlite As Boolean
    Dim rs1 As New ADODB.Recordset
    
'    check_date = GetCurDate
'    rs.Open "select * from pt_standardLen where to_char(EDATE,'YYYY-MM-DD')='" & Format(check_date, "YYYY-MM-DD") & "'", cn, adOpenKeyset, adLockReadOnly
'    If rs.RecordCount > 0 Then
'        If IsNull(rs("LENGTH")) = False Then
'            StandardLen = rs("LENGTH")
'        Else
'            StandardLen = 25.4
'            'StandardLen = 24.65  '25092010 Sunil
'        End If
'    Else
'        StandardLen = 25.4
'        'StandardLen = 24.65   '25092010 Sunil
'    End If
'    rs.Close

    twno = Val(Mid(drawid.Text, 9, 2))
    
    StandardLen = 25.4
    
    If Check_for_preform_position(drawid.Text) = True Then
        StandardLen = 24.65
    End If
        
    
   
'Disbaled Sunil/Milind 28052010 till further instruction
'23042011 -Disabled as per Milind's Mail till further instruction
'DISABLED MILIND/SUNIL 25052011
   
    If Val(ptno.Text) = "8" Or ptno.Text = "14" Or ptno.Text = "15" Or ptno.Text = "16" Or ptno.Text = "17" Or ptno.Text = "18" Then
        If Check_for_preform_position_50(drawid.Text) = True Then
            'StandardLen = 50.75
            rs1.Open "select * from cutting_logic where  twno='" & twno & "' ", cn, adOpenKeyset, adLockReadOnly
                If rs1.RecordCount > 0 Then
                    newLength = rs1("length")
                    If newLength = 48.8 Then StandardLen = 49.15
                    If newLength = 50.4 Then StandardLen = 50.75
                End If
            rs1.Close
        End If
    End If
   
 Exit Sub

Set_Standard_Length_Error:
   With err
  sErrorDescription = "Error '" & .Number & " " & .Description & "' occurred in " & "Set_Standard_Length" & IIf(Erl <> 0, "at line " & CStr(Erl) & ".", ".")
   End With
   Call Err_Handler(True, err.Number, err.Description, "frmptautomation", "Set_Standard_Length", Erl)
   err.clear
End Sub

Private Sub txtPTOut_Change()
    Clipboard.SetText ("")
    srnogotf = 1
End Sub

Private Sub txtPTOut_Click()
    Clipboard.SetText ("")
    srnogotf = 1

End Sub

Private Sub txtPTOut_GotFocus()
    Clipboard.SetText ("")
    srnogotf = 1

End Sub

Private Sub txtPTOut_KeyDown(KeyCode As Integer, shift As Integer)
    Clipboard.SetText ("")
    srnogotf = 1

End Sub

Private Sub txtPTOut_KeyPress(KeyAscii As Integer)
    Clipboard.SetText ("")
    srnogotf = 1

    If KeyAscii = 13 Then
      Dim RS As New ADODB.Recordset
      Dim Entry_date As Date
      Entry_date = getcurdate
      txtPTOut.Text = UCase$(Trim(txtPTOut.Text))
      RS.Open "Select BARCODEID,LENGTH,PDATE,fid from PTBARCODE where barcodeid='" & UCase(txtPTOut.Text) & "' and STATUS='O' AND  LENGTH >=4025 ", cn, adOpenKeyset, adLockReadOnly
      If RS.RecordCount <= 0 Then
          MsgBox "Wrong id", vbCritical
          txtPTOut.Text = ""
          RS.Close
          Exit Sub
      Else
        MC = Val(Mid(RS("fid"), 9, 2))
''''        If (Entry_date - rs("pdate")) * 24 * 60 < 1.5 Then '''lock by ankit on 08.06.12
''''            MsgBox "Do out entries after 2 min of PT entry", vbCritical
''''            txtPTOut.Text = ""
''''            rs.Close
''''            Exit Sub
''''        End If
      End If
      RS.Close
    
'      '------------------------------Do not allow for PT Out in case TWPMD is not done-------------------Sunil /Pavan 13102009
'      If MC = 6 Or MC = 7 Or MC = 8 Or MC = 9 Then
'        rs.Open "Select * from twpmd where barcodeid = '" & txtPTOut.Text & "'", cn, adOpenKeyset, adLockReadOnly
'        If rs.RecordCount = 0 Then
'              MsgBox "TWPMD Pending", vbCritical
'              rs.Close
'              Exit Sub
'        End If
'        rs.Close
'      End If
'      '----------------------------------end of above-----------------------------------------------------
    
      CURTIME = TimeSerial(Hour(Entry_date), Minute(Entry_date), second(Entry_date))
      If CURTIME >= TimeSerial(7, 0, 0) And CURTIME < TimeSerial(15, 0, 0) Then
          SF = "A"
      ElseIf CURTIME >= TimeSerial(15, 0, 0) And CURTIME < TimeSerial(23, 0, 0) Then
          SF = "B"
      Else
          SF = "C"
      End If
    
      RS.Open "select * from BobbinInOutRw where barcodeid='" & UCase(txtPTOut.Text) & "'", cn, adOpenDynamic, adLockOptimistic
      If RS.EOF = True Then
          RS.AddNew
          RS("rwoutdate") = Entry_date
          RS("rwoutuser") = UCase(opr.Text)
          RS("rwoutshift") = UCase(SF)
          RS("barcodeid") = UCase(txtPTOut.Text)
          RS("rwstatus") = "OUT"
          RS.update
      Else
          RS("rwoutdate") = Entry_date
          RS("rwoutuser") = UCase(opr.Text)
          RS("rwoutshift") = UCase(SF)
          RS("barcodeid") = UCase(txtPTOut.Text)
          RS("rwstatus") = "OUT"
          RS.update
      End If
      RS.Close
    
      RS.Open "select * from ptbarcode where  barcodeid='" & UCase(txtPTOut.Text) & "'", cn, adOpenDynamic, adLockOptimistic
      If RS.EOF = False Then
          RS("stage") = "OUTRW"
          RS.update
      End If
      RS.Close
      Call cmdOutScanPending_Click
      txtPTOut.Text = ""
    End If
End Sub

Private Sub txtPTOut_LostFocus()
    Clipboard.SetText ("")
    srnogotf = 1

End Sub

Private Sub txtPTOut_MouseDown(Button As Integer, shift As Integer, X As Single, Y As Single)
    Clipboard.SetText ("")
    srnogotf = 1

End Sub

Function Check_Bendlite(preid) As Boolean
    
    Dim RsDel As New ADODB.Recordset
    Dim RsBend As New ADODB.Recordset
    Check_Bendlite = True
    '---------Check DELP---------------------
    RsDel.Open "Select Max(DELP) from P104 where PREID ='" & preid & "'", cn, adOpenKeyset
    If RsDel.EOF = False Then
        If IsNull(RsDel(0)) = False Then
            If RsDel(0) = 0 Then
                Check_Bendlite = False
            ElseIf RsDel(0) < 4.81 Or RsDel(0) > 5.2 Then     'earlier If less than 4.9 or > 5.2 then 'Not Bendlite (BZ/BZSL)
                Check_Bendlite = False
            End If
        Else
            Check_Bendlite = False
        End If
    Else
        Check_Bendlite = False
    End If
    RsDel.Close
    '------------Check Centre Dip--
    RsBend.Open "Select Max(CENTREDIP) from P104 where CENTREDIP > 0 and  substr(PREID,1,7)='" & Mid(preid, 1, 7) & "' and  substr(PREID,1,1) not in ('D','N','L')", cn, adOpenKeyset
    If RsBend.EOF = False Then
        If IsNull(RsBend(0)) = False Then
            If RsBend(0) <= 0 Or RsBend(0) >= 50 Then
                Check_Bendlite = False
            End If
        Else
            Check_Bendlite = False
        End If
    Else
        Check_Bendlite = False
    End If
    RsBend.Close
    'Check_Bendlite = data
    Exit Function
End Function
Private Sub Fill_ZTPMD_IDS(ID As String)
    Dim rsdate As New ADODB.Recordset
    Dim RS As New ADODB.Recordset
    If rsdate.State = 1 Then rsdate.Close
    rsdate.Open "select SYSDATE-5 FROM DUAL", cn
    fromdate = Format(rsdate(0), "yyyy-mm-dd 00:00:00")
    rsdate.Close
    TowerNo = Val(Mid(ID, 9, 2))
    Grid1.ReadOnly = True
    RS.Open "Select * from draw where pdate  >= '" & fromdate & "' and twno = '" & TowerNo & "' order by pdate desc", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        For i = 0 To RS.RecordCount - 1
              If RS("fid") = ID Then
                  RS.MoveNext
                  If RS.EOF = False Then RS.MoveNext '3rd no spool
                  If RS.EOF = False Then RS.MoveNext '3rd no spool
                  If RS.EOF = False Then Prevspool = RS("FID")
                  RS.Close
                  Exit For
              Else
                  RS.MoveNext
              End If
        Next i
    Else
        RS.Close
        Exit Sub
    End If
      
    If Prevspool = "" Then
      Exit Sub
    End If
        
  'List of ZTPMD IDs
  IDStr = ""
  RS.Open "select * from pt_inst_ztpmd where substr(fid,1,11) = '" & Prevspool & "'", cn, adOpenKeyset, adLockReadOnly
  If RS.RecordCount > 0 Then
      Grid1.Rows = RS.RecordCount + 1
      Grid1.Cols = 2
      Grid1.Column(0).Width = 0
      While RS.EOF = False
          Grid1.Cell(i, 1).Text = RS("ZTPMDID")
          i = i + 1
          RS.MoveNext
      Wend
  End If
  RS.Close
End Sub
Function Check_for_preform_position(FID As String) As Boolean
    Dim RS As New ADODB.Recordset
    Check_for_preform_position = False
    Dim rsDelta As New ADODB.Recordset
    
    Dim rs1 As New ADODB.Recordset
    
    twno = Val(Mid(drawid.Text, 9, 2))   '10,20,21,22 added after Milind's mail 21122010
    
    rs1.Open "select * from cutting_logic where  twno= '" & twno & "'", cn, adOpenKeyset, adLockReadOnly
      If rs1.RecordCount > 0 Then
        If rs1("LENGTH") = 24.4 Then
'            rs.Open "select min(pos) from pt where fid like  '" & FID & "%'", cn, adOpenKeyset, adLockReadOnly
'            If IsNull(rs(0)) = False Then
'                ppos = rs(0)
'            Else
'                ppos = 0
'            End If
'            rs.Close
'
'            rs.Open "select sum(drlen) from draw where fid like  '" & Mid(FID, 1, 8) & "%'", cn, adOpenKeyset, adLockReadOnly
'            If IsNull(rs(0)) = False Then
'                Totdrlen = rs(0)
'            End If
'            rs.Close
            
    '        If ppos > 50 And ppos <= (Totdrlen - 50) Then
    '           Check_for_preform_position = True
    '        End If
             Check_for_preform_position = True
    '         rs.Close
        End If
        
      End If
    rs1.Close
End Function
Function Check_for_preform_position_50(FID As String) As Boolean
    Dim RS As New ADODB.Recordset
    Check_for_preform_position_50 = False
    Dim rsDelta As New ADODB.Recordset
    Dim rs1 As New ADODB.Recordset
    
    twno = Val(Mid(drawid.Text, 9, 2))
     
    rs1.Open "select * from cutting_logic where  twno='" & twno & "' ", cn, adOpenKeyset, adLockReadOnly
    If rs1.RecordCount > 0 Then
        If rs1("LENGTH") = 50.4 Or rs1("LENGTH") = 48.8 Or rs1("LENGTH") = 54.6 Or rs1("LENGTH") = 58.8 Then
            RS.Open "select min(pos) from pt where fid like  '" & FID & "%'", cn, adOpenKeyset, adLockReadOnly
            If IsNull(RS(0)) = False Then
                ppos = RS(0)
            Else
                ppos = 0
            End If
            RS.Close
            
            RS.Open "select sum(drlen) from draw where fid like  '" & Mid(FID, 1, 8) & "%'", cn, adOpenKeyset, adLockReadOnly
            If IsNull(RS(0)) = False Then
                Totdrlen = RS(0)
            End If
            RS.Close
            
            rsDelta.Open "select delp from p104 where preid = '" & Mid(FID, 1, 8) & "'", cn, adOpenKeyset, adLockReadOnly
            If rsDelta.RecordCount > 0 Then
                DELP = rsDelta(0)
            End If
            rsDelta.Close
            If ppos > 50 And ppos <= (Totdrlen - 50) Then
                'If delp > 4.85 Then
                    Check_for_preform_position_50 = True
                'End If
            End If
        End If
    End If
    rs1.Close
    
End Function
Private Sub GenerateSampleID(ByVal ID As String, ByVal bobbin_length As Double)
    If ID = "" Then
        Exit Sub
    End If
    BFDFlag = False
    FPcFlag = False
    LPcFlag = False
    F25Flag = False
    L25Flag = False
    F50Flag = False
    L50Flag = False
    proceed = 0
    Dim PT()
    Dim ids()
    Dim rsPar As New ADODB.Recordset
    Dim rsRej As New ADODB.Recordset
    ParentId = ""
    
    Dim rsdraw1 As New ADODB.Recordset
    Dim rsmat As New ADODB.Recordset
    
    Dim LP As Boolean
    LP = False
    proceed = 1
    
    '========================For Checking First and Last Piece============================
    rsmat.Open "Select stockqty from matstock where id='" & Mid(ID, 1, 11) & "' AND matno in '330000126'", cn, adOpenKeyset, adLockReadOnly
    If rsmat.RecordCount > 0 Then
        If rsmat(0) = 0 Then
            LP = True
        Else
            LP = False
        End If
    End If
    rsmat.Close
   
    RS.Open "Select max(ptid),min(ptid) from pt where  ptid like '" & Mid(ID, 1, 11) & "%'  AND OKLEN >0", cn, adOpenKeyset, adLockReadOnly
    If RS.RecordCount > 0 Then
        If IsNull(RS(1)) = False Then
            If Mid(ID, 1, 12) = Mid(RS(1), 1, 12) Then
                FPcFlag = True
                L25Flag = True
                L50Flag = True
                FLTarget = True
                proceed = 0
                RS.Close
                Exit Sub
            End If
        End If
        If LP = True Then
            If IsNull(RS(0)) = False Then
                If Len(RS(0)) = 12 Then
                    If Mid(ID, 1, 12) = Mid(RS(0), 1, 12) Then
                        LPcFlag = True
                        F25Flag = True
                        F50Flag = True
                        FLTarget = True
                        RS.Close
                        proceed = 0
                        Exit Sub
                    End If
                Else
                    If Mid(ID, 1, 13) = Mid(RS(0), 1, 13) Then
                        LPcFlag = True
                        F25Flag = True
                        F50Flag = True
                        FLTarget = True
                        RS.Close
                        proceed = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If
    RS.Close
    
    
    Dim rsdraw As New ADODB.Recordset
    rsdraw.Open "Select drlen from draw where fid='" & Mid(ID, 1, 11) & "' and matno in ('330000126')", cn, adOpenKeyset, adLockReadOnly
    If rsdraw.RecordCount > 0 Then
        drlen = rsdraw(0)
    End If
    rsdraw.Close
    
    '---------------------New Logic for Spool Start and End Length above logic to  be deleted later Sunil
      rsdraw.Open "select epos,oklen,spos  from PT where  ptid like '" & Mid(ID, 1, 12) & "%' and oklen > 0", cn, adOpenKeyset, adLockReadOnly
      If rsdraw.RecordCount > 0 Then
          SpStLen = drlen - rsdraw(0)
          SpEdLen = SpStLen + rsdraw(1)
          ptlen1 = SpEdLen
      End If
      rsdraw.Close
    
    '---------------------
    'Checking First 25-50/Last 25-50 Km of Main Spool
    If SpStLen < 50# Then
        L50Flag = True
        If SpStLen < 25# Then
             L25Flag = True
        End If
    End If
    If Val(drlen - SpEdLen) < 50# Then
        F50Flag = True
        If Val(drlen - SpEdLen) < 25# Then
            F25Flag = True
        End If
    End If
    
    If F25Flag = True Or L25Flag = True Then
        FLTarget = True
        Exit Sub
    End If
    
    '=========================For Checking BFD cutting ===========================
    'Adjacent Fibre Should not have BFD i.e Fibre Start Length-1 to Fibre End Length +1 Kms
    'Changed By Sunil 01.03.2006 .This Logic will Check within and outside (+/- 1 Kms) Length.
    Dim rsdr As New ADODB.Recordset
    BFDFlag = False
    rsdr.Open "Select remark from drawflaw where fid='" & Mid(ID, 1, 11) & "' and position >'" & (SpStLen - 1) & "' and position <'" & (SpEdLen + 1) & "'", cn, adOpenKeyset, adLockReadOnly
    If rsdr.RecordCount > 0 Then
        Do While rsdr.EOF = False
            If rsdr(0) = "BFD" Or rsdr(0) = "B" Or rsdr(0) = "PITCHCHANGE" Then
                BFDFlag = True
                Exit Do
            End If
            rsdr.MoveNext
        Loop
    End If
    rsdr.Close
    Dim Range As String
    Range = ""
    If ptlen1 >= 0 And ptlen1 < 100 Then
        mn = 0
        mxm = 100
        Range = "A"
    End If
    If ptlen1 >= 100 And ptlen1 < 200 Then
        mn = 100
        mxm = 200
        Range = "B"
    End If
    If ptlen1 >= 200 And ptlen1 < 300 Then
        mn = 200
        mxm = 300
        Range = "C"
    End If
    If ptlen1 >= 300 And ptlen1 < 400 Then
        mn = 300
        mxm = 400
        Range = "D"
    End If
    If ptlen1 >= 400 And ptlen1 < 500 Then
        mn = 400
        mxm = 500
        Range = "E"
    End If
    If ptlen1 >= 500 And ptlen1 < 600 Then
        mn = 500
        mxm = 600
        Range = "F"
    End If
    If ptlen1 >= 600 And ptlen1 < 700 Then
        mn = 600
        mxm = 700
        Range = "G"
    End If
    If ptlen1 >= 700 And ptlen1 < 800 Then
        mn = 700
        mxm = 800
        Range = "H"
    End If
    
    If ptlen1 >= 800 And ptlen1 < 900 Then
        mn = 800
        mxm = 900
        Range = "I"
    End If
    
    If ptlen1 >= 900 And ptlen1 < 1000 Then
        mn = 900
        mxm = 1000
        Range = "J"
    End If
    
    If ptlen1 >= 1000 And ptlen1 < 1100 Then
        mn = 1000
        mxm = 1100
        Range = "K"
    End If
    
    '========================Introduced on 23 MArch 2006 for giving Message for OTDR person for Spectral TEsting
    Dim FlagSamplingID As Boolean
    FlagSamplingID = False
    
    Dim RsSampling As New ADODB.Recordset
    Dim SampleID
    
    SampleRange = Round((mn + mxm) / 2, 3)
    If ptlen1 >= (SampleRange - 30) And ptlen1 <= (SampleRange + 30) Then  'Sample has to be from middle of the range 20072009 Bhaumik's Mail ,CHANGED TO 30 FROM 20 24072009
        If Range <> "" Then
            If BFDFlag = False And Val(bobbin_length) >= 4400 Then
                RsSampling.Open "Select * from SamplingID where Range='" & Range & "' and substr(FID,1,11)='" & Mid(ID, 1, 11) & "'", cn, adOpenDynamic, adLockOptimistic
                If RsSampling.EOF = True Then
                    RsSampling.AddNew
                    RsSampling("FID") = ID
                    RsSampling("Range") = Range
                    RsSampling("EDATE") = getcurdate
                    RsSampling("GENFROM") = "PT"
                    RsSampling("Plant") = plant
                    RsSampling.update
                    FlagSamplingID = True
                    MsgBox "This is Star Bobbin ! Please Mark it as Star for Faster Processing", vbInformation
                Else
                    FlagSamplingID = False
                    SampleID = RsSampling("FID")
                End If
                RsSampling.Close
            End If
        End If
    End If
End Sub
Private Sub GenerateSampleID_mbend(ByVal ID As String, ByVal FactorVal As Double)
''Dim rsmb As New ADODB.Recordset
''Dim mbrange
''
''If FactorVal >= 0 And FactorVal < 200 Then
''    mbrange = "A"
''ElseIf FactorVal >= 200 And FactorVal <= 400 Then
''    mbrange = "B"
''ElseIf FactorVal > 400 And FactorVal <= 600 Then
''    mbrange = "C"
''ElseIf FactorVal > 600 And FactorVal <= 800 Then
''    mbrange = "D"
''ElseIf FactorVal > 800 And FactorVal <= 1000 Then
''    mbrange = "E"
''ElseIf FactorVal > 1000 And FactorVal <= 1200 Then
''    mbrange = "F"
''ElseIf FactorVal > 1200 Then
''    mbrange = "G"
''End If
''
''
''If rsmb.State = 1 Then rsmb.Close
''If mbrange <> "" Then
''    rsmb.Open "select * from mbend_samplingid where range='" & mbrange & "' and substr(fid,1,8)='" & Mid(PTID, 1, 8) & "'", cn, adOpenDynamic, adLockOptimistic
''    If rsmb.EOF = True Then
''        rsmb.AddNew
''        rsmb("FID") = ID
''        rsmb("RANGE") = mbrange
''        rsmb("edate") = getcurdate
''        rsmb("genfrom") = "PT"
''        rsmb("plant") = plant
''        rsmb.Update
''        'MsgBox "This is MBEND Bobbin ! Please Mark it as MBend for MBend Testing", vbInformation
''    End If
''    rsmb.Close
''End If


End Sub
Private Sub Command3_Click()
'''Dim rsmb As New ADODB.Recordset
'''
'''ID = Text4.Text
'''PTID = Text4.Text
'''factor = "720"
'''GenerateSampleID_mbend PTID.Text, factor
'''
'''If rs.State = 1 Then rs.Close
'''''rs.Open "Select * from pt where fid like '" & Mid(ID, 1, 8) & "%' oklen >0", cn, adOpenKeyset, adLockReadOnly
'''rs.Open "Select * from pt where PTID like '" & Mid(ID, 1, 8) & "%' and oklen > 0 order by ptid", cn, adOpenKeyset, adLockReadOnly
'''If rs.RecordCount > 0 Then
'''    While rs.EOF = False
'''    If Mid(rs("ptid"), 11, 2) = "0A" Then
'''        mbrange = "A"
'''        mbid = rs("PTID")
'''    ElseIf rs("FACTOR") >= 150 And rs("FACTOR") <= 400 Then
'''        mbrange = "B"
'''    ElseIf rs("FACTOR") > 400 And rs("FACTOR") <= 600 Then
'''        mbrange = "C"
'''    ElseIf rs("factor") > 600 Then
'''        mbrange = "D"
'''        mbid = rs("PTID")
'''    End If
'''    rs.MoveNext
'''    Wend
'''End If
'''rs.Close
'''
'''If mbrange <> "" Then
'''rsmb.Open "select * from mbend_samplingid where range='" & mbrange & "' and substr(fid,1,8)='" & Mid(PTID, 1, 8) & "'", cn, adOpenDynamic, adLockOptimistic
'''If rsmb.EOF = True Then
'''    rsmb.AddNew
'''    rsmb("FID") = PTID.Text
'''    rsmb("RANGE") = mbrange
'''    rsmb("edate") = getcurdate
'''    rsmb("genfrom") = "PT"
'''    rsmb("plant") = plant
'''    rsmb.update
'''Else
'''    If mbrange = "D" Then
'''        rsmb("FID") = ID
'''        rsmb("edate") = getcurdate
'''        rsmb.update
'''    End If
'''End If
'''rsmb.Close
'''End If
'''
End Sub



