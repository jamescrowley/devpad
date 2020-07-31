VERSION 5.00
Object = "{5C0E11AE-2C8C-4C35-BC7A-D9B469D5DE4D}#6.1#0"; "VBWTRE~1.OCX"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   4395
   ClientLeft      =   30
   ClientTop       =   240
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOptions_new.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "Options"
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   12
      Left            =   2715
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   51
      Top             =   660
      Width           =   4935
      Begin VB.CheckBox chkOption 
         Caption         =   "1351"
         Height          =   255
         Index           =   7
         Left            =   0
         TabIndex        =   146
         Tag             =   "RestoreWorkspace"
         Top             =   510
         Width           =   3435
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "1155"
         Height          =   252
         Index           =   6
         Left            =   0
         TabIndex        =   61
         Tag             =   "OneProgramInstance"
         Top             =   270
         Width           =   3585
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "1259"
         Height          =   345
         Left            =   3645
         TabIndex        =   55
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtItem 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   53
         Tag             =   "VBPath"
         Top             =   1230
         Width           =   2505
      End
      Begin VB.Label lblLabel 
         Caption         =   "1157"
         Height          =   255
         Index           =   24
         Left            =   0
         TabIndex        =   54
         Top             =   1230
         Width           =   975
      End
      Begin VB.Label lblLabel 
         Caption         =   "1224"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   25
         Left            =   0
         TabIndex        =   52
         Top             =   0
         Width           =   2265
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   11
      Left            =   3240
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   135
      Top             =   435
      Width           =   4935
      Begin VB.TextBox txtItem 
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   139
         Tag             =   "ScrollDelay"
         Text            =   "180"
         Top             =   285
         Width           =   1320
      End
      Begin VB.ComboBox cboWrap 
         Height          =   315
         ItemData        =   "frmOptions_new.frx":000C
         Left            =   1065
         List            =   "frmOptions_new.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   138
         Top             =   1635
         Width           =   1335
      End
      Begin VB.ComboBox cboFonts 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1065
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   137
         Top             =   900
         Width           =   2205
      End
      Begin VB.ComboBox cboFontSize 
         Height          =   315
         Left            =   1065
         Style           =   2  'Dropdown List
         TabIndex        =   136
         Tag             =   "NORES"
         Top             =   1260
         Width           =   1335
      End
      Begin VB.Label lblLabel 
         Caption         =   "1247"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   29
         Left            =   0
         TabIndex        =   145
         Top             =   0
         Width           =   2265
      End
      Begin VB.Label lblLabel 
         Caption         =   "1277"
         Height          =   270
         Index           =   31
         Left            =   0
         TabIndex        =   144
         Top             =   285
         Width           =   1020
      End
      Begin VB.Label lblLabel 
         Caption         =   "1136"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   143
         Top             =   1650
         Width           =   1020
      End
      Begin VB.Label lblLabel 
         Caption         =   "1208"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   0
         TabIndex        =   142
         Top             =   645
         Width           =   2265
      End
      Begin VB.Label lblLabel 
         Caption         =   "1130"
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   141
         Top             =   915
         Width           =   810
      End
      Begin VB.Label lblLabel 
         Caption         =   "1131"
         Height          =   255
         Index           =   6
         Left            =   0
         TabIndex        =   140
         Top             =   1275
         Width           =   810
      End
   End
   Begin VB.PictureBox picScreen 
      BorderStyle     =   0  'None
      Height          =   3045
      Index           =   10
      Left            =   2685
      ScaleHeight     =   3045
      ScaleWidth      =   4740
      TabIndex        =   126
      Top             =   540
      Width           =   4740
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   17
         Left            =   1890
         TabIndex        =   98
         Text            =   "'"
         Top             =   525
         Width           =   270
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   6
         Left            =   1890
         TabIndex        =   99
         Text            =   "'"
         Top             =   840
         Width           =   270
      End
      Begin VB.TextBox txtLang 
         Height          =   300
         Index           =   14
         Left            =   1515
         TabIndex        =   103
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtLang 
         Height          =   300
         Index           =   13
         Left            =   1515
         TabIndex        =   102
         Top             =   1710
         Width           =   1095
      End
      Begin VB.CheckBox chkLang 
         Caption         =   "1330"
         Height          =   240
         Index           =   5
         Left            =   0
         TabIndex        =   101
         Top             =   1425
         Width           =   2940
      End
      Begin VB.CheckBox chkLang 
         Caption         =   "1329"
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   100
         Top             =   1185
         Width           =   3315
      End
      Begin VB.CheckBox chkLang 
         Caption         =   "1335"
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   97
         Top             =   270
         Width           =   2145
      End
      Begin VB.Label lblLabel 
         Caption         =   "1328"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   57
         Left            =   0
         TabIndex        =   155
         Top             =   0
         Width           =   2205
      End
      Begin VB.Label lblLabel 
         Caption         =   "1369"
         Height          =   225
         Index           =   56
         Left            =   0
         TabIndex        =   154
         Top             =   540
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1370"
         Height          =   225
         Index           =   32
         Left            =   0
         TabIndex        =   153
         Top             =   855
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1348"
         Height          =   240
         Index           =   51
         Left            =   0
         TabIndex        =   128
         Top             =   2070
         Width           =   1380
      End
      Begin VB.Label lblLabel 
         Caption         =   "1331"
         Height          =   240
         Index           =   50
         Left            =   0
         TabIndex        =   127
         Top             =   1740
         Width           =   1380
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   8
      Left            =   2730
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   29
      Top             =   585
      Width           =   4935
      Begin VB.CheckBox chkLang 
         Caption         =   "1327"
         Height          =   240
         Index           =   2
         Left            =   0
         TabIndex        =   96
         Top             =   2250
         Width           =   2145
      End
      Begin VB.CheckBox chkLang 
         Caption         =   "1326"
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   95
         Top             =   2010
         Width           =   2145
      End
      Begin VB.ComboBox cboLang 
         Height          =   315
         Index           =   0
         ItemData        =   "frmOptions_new.frx":002C
         Left            =   1170
         List            =   "frmOptions_new.frx":003C
         Style           =   2  'Dropdown List
         TabIndex        =   92
         Tag             =   "Options\FileType"
         Top             =   1020
         Width           =   1860
      End
      Begin VB.CheckBox chkLang 
         Caption         =   "1309"
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   94
         Top             =   1770
         Width           =   2145
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   2
         Left            =   1185
         TabIndex        =   90
         Top             =   690
         Width           =   1830
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   0
         Left            =   1185
         TabIndex        =   88
         Top             =   360
         Width           =   1830
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   1
         Left            =   1185
         Locked          =   -1  'True
         TabIndex        =   86
         Top             =   15
         Width           =   1830
      End
      Begin VB.Label lblLabel 
         Caption         =   "1349"
         Height          =   255
         Index           =   52
         Left            =   0
         TabIndex        =   129
         Top             =   1035
         Width           =   1050
      End
      Begin VB.Label lblLabel 
         Caption         =   "1047"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   49
         Left            =   0
         TabIndex        =   125
         Top             =   1425
         Width           =   2370
      End
      Begin VB.Label lblLabel 
         Caption         =   "1181"
         Height          =   255
         Index           =   26
         Left            =   0
         TabIndex        =   64
         Top             =   690
         Width           =   1050
      End
      Begin VB.Label lblLabel 
         Caption         =   "1308"
         Height          =   255
         Index           =   16
         Left            =   0
         TabIndex        =   63
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label lblLabel 
         Caption         =   "1070"
         Height          =   255
         Index           =   15
         Left            =   0
         TabIndex        =   62
         Top             =   15
         Width           =   1050
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3015
      Index           =   9
      Left            =   2700
      ScaleHeight     =   3015
      ScaleWidth      =   4935
      TabIndex        =   108
      Top             =   540
      Width           =   4935
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   16
         Left            =   4260
         TabIndex        =   78
         Top             =   660
         Width           =   360
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   15
         Left            =   1890
         TabIndex        =   77
         Text            =   "'"
         Top             =   660
         Width           =   270
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   12
         Left            =   1890
         TabIndex        =   83
         Top             =   1905
         Width           =   360
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   11
         Left            =   1890
         TabIndex        =   84
         Top             =   2220
         Width           =   360
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   10
         Left            =   4260
         TabIndex        =   82
         Top             =   1305
         Width           =   360
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   9
         Left            =   4260
         TabIndex        =   80
         Text            =   "'"
         Top             =   975
         Width           =   270
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   8
         Left            =   4260
         TabIndex        =   76
         Text            =   "'"
         Top             =   330
         Width           =   270
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   7
         Left            =   1890
         TabIndex        =   81
         Top             =   1290
         Width           =   360
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   5
         Left            =   1890
         TabIndex        =   79
         Text            =   """"
         Top             =   975
         Width           =   270
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   4
         Left            =   1890
         TabIndex        =   75
         Text            =   "'"
         Top             =   345
         Width           =   270
      End
      Begin VB.Label lblLabel 
         Caption         =   "1368"
         Height          =   225
         Index           =   55
         Left            =   2340
         TabIndex        =   152
         Top             =   675
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1367"
         Height          =   225
         Index           =   8
         Left            =   0
         TabIndex        =   151
         Top             =   675
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1337"
         Height          =   225
         Index           =   48
         Left            =   0
         TabIndex        =   124
         Top             =   1920
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1338"
         Height          =   225
         Index           =   47
         Left            =   0
         TabIndex        =   123
         Top             =   2235
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1336"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   46
         Left            =   0
         TabIndex        =   122
         Top             =   1605
         Width           =   2370
      End
      Begin VB.Label lblLabel 
         Caption         =   "1342"
         Height          =   225
         Index           =   45
         Left            =   2340
         TabIndex        =   121
         Top             =   1320
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1345"
         Height          =   255
         Index           =   44
         Left            =   2340
         TabIndex        =   120
         Top             =   990
         Width           =   1170
      End
      Begin VB.Label lblLabel 
         Caption         =   "1345"
         Height          =   255
         Index           =   43
         Left            =   2340
         TabIndex        =   119
         Top             =   360
         Width           =   1170
      End
      Begin VB.Label lblLabel 
         Caption         =   "1341"
         Height          =   225
         Index           =   33
         Left            =   0
         TabIndex        =   118
         Top             =   1305
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1334"
         Height          =   225
         Index           =   30
         Left            =   0
         TabIndex        =   117
         Top             =   990
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1333"
         Height          =   225
         Index           =   222
         Left            =   0
         TabIndex        =   107
         Top             =   360
         Width           =   1740
      End
      Begin VB.Label lblLabel 
         Caption         =   "1340"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   223
         Left            =   0
         TabIndex        =   106
         Top             =   0
         Width           =   2205
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2550
      Index           =   15
      Left            =   2925
      ScaleHeight     =   2550
      ScaleWidth      =   4635
      TabIndex        =   149
      Top             =   765
      Width           =   4635
      Begin VB.CommandButton cmdNewLanguage 
         Caption         =   "1245"
         Height          =   345
         Left            =   45
         TabIndex        =   150
         Top             =   45
         Width           =   1215
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   1
      Left            =   2775
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   5
      Top             =   690
      Width           =   4935
      Begin VB.TextBox txtItem 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   14
         Tag             =   "ProjectAuthor"
         Top             =   630
         Width           =   2190
      End
      Begin VB.TextBox txtItem 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   13
         Tag             =   "ProjectName"
         Top             =   270
         Width           =   2190
      End
      Begin VB.Label lblLabel 
         Caption         =   "1209"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   7
         Left            =   0
         TabIndex        =   19
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1139"
         Height          =   255
         Index           =   19
         Left            =   0
         TabIndex        =   16
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label lblLabel 
         Caption         =   "1140"
         Height          =   255
         Index           =   20
         Left            =   0
         TabIndex        =   15
         Top             =   630
         Width           =   1080
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   5
      Left            =   2715
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   9
      Top             =   600
      Width           =   4935
      Begin VB.TextBox txtItem 
         Height          =   285
         Index           =   2
         Left            =   1185
         TabIndex        =   49
         Tag             =   "ServerFiles"
         Text            =   ".asp .cfm .cfml .ihtml .js .jsp .php"
         Top             =   1170
         Width           =   2250
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "1218"
         Height          =   270
         Index           =   1
         Left            =   0
         TabIndex        =   44
         Tag             =   "PreviewUseTempFile"
         Top             =   270
         Width           =   2895
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "1219"
         Height          =   270
         Index           =   0
         Left            =   0
         TabIndex        =   43
         Tag             =   "PreviewPromptToSave"
         Top             =   570
         Width           =   2895
      End
      Begin VB.TextBox txtItem 
         Height          =   285
         Index           =   1
         Left            =   1185
         TabIndex        =   42
         Tag             =   "ServerLocalPath"
         Text            =   "C:\inetpub\wwwroot"
         Top             =   1530
         Width           =   2250
      End
      Begin VB.TextBox txtItem 
         Height          =   285
         Index           =   0
         Left            =   1185
         TabIndex        =   41
         Tag             =   "Server"
         Text            =   "http://localhost"
         Top             =   1890
         Width           =   2250
      End
      Begin VB.CommandButton cmdBrowseLocal 
         Caption         =   "1259"
         Height          =   345
         Left            =   3510
         TabIndex        =   40
         Top             =   1485
         Width           =   1215
      End
      Begin VB.Label lblLabel 
         Caption         =   "1221"
         Height          =   255
         Index           =   23
         Left            =   0
         TabIndex        =   50
         Top             =   1185
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Caption         =   "1220"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   22
         Left            =   0
         TabIndex        =   48
         Top             =   930
         Width           =   2100
      End
      Begin VB.Label lblLabel 
         Caption         =   "1217"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   21
         Left            =   0
         TabIndex        =   47
         Top             =   0
         Width           =   2100
      End
      Begin VB.Label lblLabel 
         Caption         =   "1222"
         Height          =   255
         Index           =   18
         Left            =   0
         TabIndex        =   46
         Top             =   1545
         Width           =   1065
      End
      Begin VB.Label lblLabel 
         Caption         =   "1223"
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   45
         Top             =   1905
         Width           =   1095
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   4
      Left            =   3030
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   8
      Top             =   660
      Width           =   4935
      Begin VB.ComboBox cboFavourites 
         Height          =   315
         Left            =   15
         TabIndex        =   58
         Text            =   "Combo1"
         Top             =   1200
         Width           =   2130
      End
      Begin VB.CommandButton cmdDelFav 
         Caption         =   "1021"
         Height          =   345
         Left            =   3435
         TabIndex        =   35
         Top             =   1185
         Width           =   1215
      End
      Begin VB.CommandButton cmdAddFav 
         Caption         =   "1020"
         Height          =   345
         Left            =   2175
         TabIndex        =   34
         Top             =   1185
         Width           =   1215
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "1214"
         Height          =   270
         Index           =   3
         Left            =   0
         TabIndex        =   32
         Tag             =   "OpenRememberLocation"
         Top             =   570
         Width           =   2895
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "1213"
         Height          =   270
         Index           =   2
         Left            =   0
         TabIndex        =   31
         Tag             =   "OpenShowNewTab"
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label lblLabel 
         Caption         =   "1215"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   14
         Left            =   0
         TabIndex        =   33
         Top             =   930
         Width           =   2100
      End
      Begin VB.Label lblLabel 
         Caption         =   "1212"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   2100
      End
   End
   Begin vbwTreeView.TreeView tvwOpt 
      Height          =   3795
      Left            =   75
      TabIndex        =   3
      Top             =   510
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   6694
      Lines           =   0   'False
      LabelEditing    =   0   'False
      PlusMinus       =   0   'False
      RootLines       =   0   'False
      ToolTips        =   0   'False
      BorderStyle     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxScrollTime   =   0
      DisableCustomDraw=   -1  'True
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "1003"
      Enabled         =   0   'False
      Height          =   345
      Left            =   5310
      TabIndex        =   2
      Top             =   3930
      Width           =   1200
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00808080&
      Cancel          =   -1  'True
      Caption         =   "1002"
      Height          =   345
      Left            =   4005
      TabIndex        =   1
      Top             =   3930
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "1001"
      Default         =   -1  'True
      Height          =   345
      Left            =   2700
      TabIndex        =   0
      Top             =   3930
      Width           =   1215
   End
   Begin DevPad.ctlFrame ctlFrame1 
      Height          =   360
      Left            =   75
      TabIndex        =   59
      Top             =   90
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   635
      Begin VB.Label lblHeader1 
         BackStyle       =   0  'Transparent
         Caption         =   "1047"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   90
         TabIndex        =   60
         Top             =   75
         Width           =   2355
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   2
      Left            =   4005
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   6
      Top             =   600
      Width           =   4935
      Begin VB.TextBox txtConstantValue 
         Height          =   285
         Left            =   1215
         TabIndex        =   37
         Top             =   1965
         Width           =   2265
      End
      Begin VB.TextBox txtConstant 
         Height          =   285
         Left            =   1215
         TabIndex        =   36
         Top             =   1605
         Width           =   2250
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   345
         Left            =   2505
         TabIndex        =   22
         Top             =   525
         Width           =   1215
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   345
         Left            =   2505
         TabIndex        =   21
         Top             =   930
         Width           =   1215
      End
      Begin VB.ListBox lstConstants 
         Height          =   1035
         Left            =   0
         TabIndex        =   17
         Top             =   525
         Width           =   2415
      End
      Begin VB.Label lblLabel 
         Caption         =   "1075"
         Height          =   255
         Index           =   4
         Left            =   30
         TabIndex        =   39
         Top             =   1980
         Width           =   1065
      End
      Begin VB.Label lblLabel 
         Caption         =   "1130"
         Height          =   255
         Index           =   2
         Left            =   30
         TabIndex        =   38
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label lblLabel 
         Caption         =   "1130"
         Height          =   255
         Index           =   9
         Left            =   0
         TabIndex        =   20
         Top             =   270
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1205"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   2175
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   3
      Left            =   2715
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   7
      Top             =   570
      Width           =   4935
      Begin VB.CommandButton cmdDeleteTemplate 
         Caption         =   "1021"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3660
         TabIndex        =   57
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdNewTemplate 
         Caption         =   "102"
         Enabled         =   0   'False
         Height          =   345
         Left            =   3660
         TabIndex        =   56
         Top             =   225
         Width           =   1215
      End
      Begin VB.ComboBox cboTemplateLang 
         Height          =   315
         ItemData        =   "frmOptions_new.frx":0058
         Left            =   1065
         List            =   "frmOptions_new.frx":005A
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   615
         Width           =   2205
      End
      Begin VB.ComboBox cboTemplate 
         Height          =   315
         ItemData        =   "frmOptions_new.frx":005C
         Left            =   1065
         List            =   "frmOptions_new.frx":005E
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   255
         Width           =   2205
      End
      Begin VB.TextBox txtContents 
         Height          =   2070
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   23
         Top             =   1005
         Width           =   4890
      End
      Begin VB.Label lblLabel 
         Caption         =   "1073"
         Height          =   225
         Index           =   11
         Left            =   0
         TabIndex        =   28
         Top             =   630
         Width           =   1140
      End
      Begin VB.Label lblLabel 
         Caption         =   "1130"
         Height          =   225
         Index           =   12
         Left            =   0
         TabIndex        =   26
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label lblLabel 
         Caption         =   "1158"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   0
         TabIndex        =   25
         Top             =   0
         Width           =   1875
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2550
      Index           =   14
      Left            =   3705
      ScaleHeight     =   2550
      ScaleWidth      =   4635
      TabIndex        =   147
      Top             =   990
      Width           =   4635
      Begin VB.CheckBox chkOption 
         Caption         =   "1352"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   148
         Tag             =   "ApplyIndentToCodeLib"
         Top             =   15
         Width           =   4380
      End
   End
   Begin DevPad.vbwColourPicker cpkPicker 
      Height          =   1965
      Left            =   4065
      TabIndex        =   109
      Top             =   1005
      Visible         =   0   'False
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   3466
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3120
      Index           =   6
      Left            =   3300
      ScaleHeight     =   3120
      ScaleWidth      =   4935
      TabIndex        =   10
      Top             =   570
      Width           =   4935
      Begin VB.TextBox txtWordSeperators 
         Height          =   285
         Left            =   1635
         TabIndex        =   105
         Top             =   2190
         Width           =   2475
      End
      Begin VB.TextBox txtLang 
         Height          =   285
         Index           =   3
         Left            =   1635
         TabIndex        =   104
         Top             =   1875
         Width           =   2475
      End
      Begin VB.TextBox txtKeywords 
         Height          =   1440
         Left            =   1635
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   89
         Text            =   "frmOptions_new.frx":0060
         Top             =   360
         Width           =   1620
      End
      Begin VB.ComboBox cboKeyword 
         Height          =   315
         ItemData        =   "frmOptions_new.frx":0066
         Left            =   1620
         List            =   "frmOptions_new.frx":0073
         Style           =   2  'Dropdown List
         TabIndex        =   87
         Top             =   0
         Width           =   1650
      End
      Begin VB.Label lblLabel 
         Caption         =   "1314"
         Height          =   255
         Index           =   54
         Left            =   0
         TabIndex        =   131
         Top             =   345
         Width           =   1560
      End
      Begin VB.Label lblLabel 
         Caption         =   "1349"
         Height          =   255
         Index           =   53
         Left            =   0
         TabIndex        =   130
         Top             =   30
         Width           =   1560
      End
      Begin VB.Label lblLabel 
         Caption         =   "1324"
         Height          =   255
         Index           =   42
         Left            =   0
         TabIndex        =   93
         Top             =   2205
         Width           =   1530
      End
      Begin VB.Label lblLabel 
         Caption         =   "1325"
         Height          =   255
         Index           =   28
         Left            =   0
         TabIndex        =   91
         Top             =   1905
         Width           =   1560
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3180
      Index           =   7
      Left            =   2715
      ScaleHeight     =   3180
      ScaleWidth      =   4935
      TabIndex        =   11
      Top             =   600
      Width           =   4935
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   8
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   74
         Top             =   2505
         Width           =   510
      End
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   7
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   72
         Top             =   2190
         Width           =   510
      End
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   6
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   70
         Top             =   1875
         Width           =   510
      End
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   5
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   68
         Top             =   1575
         Width           =   510
      End
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   4
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   67
         Top             =   1260
         Width           =   510
      End
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   3
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   66
         Top             =   945
         Width           =   510
      End
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   2
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   65
         Top             =   630
         Width           =   510
      End
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   1
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   110
         Top             =   330
         Width           =   510
      End
      Begin VB.PictureBox picColour 
         Height          =   270
         Index           =   0
         Left            =   1545
         ScaleHeight     =   210
         ScaleWidth      =   450
         TabIndex        =   111
         Top             =   15
         Width           =   510
      End
      Begin VB.Label lblLabel 
         Caption         =   "1322"
         Height          =   255
         Index           =   41
         Left            =   0
         TabIndex        =   85
         Top             =   2535
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1321"
         Height          =   255
         Index           =   40
         Left            =   0
         TabIndex        =   73
         Top             =   2220
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1320"
         Height          =   255
         Index           =   39
         Left            =   0
         TabIndex        =   71
         Top             =   1905
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1319"
         Height          =   255
         Index           =   38
         Left            =   0
         TabIndex        =   69
         Top             =   1605
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1318"
         Height          =   255
         Index           =   37
         Left            =   0
         TabIndex        =   112
         Top             =   1290
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1315"
         Height          =   255
         Index           =   36
         Left            =   0
         TabIndex        =   113
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1316"
         Height          =   255
         Index           =   35
         Left            =   0
         TabIndex        =   114
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1317"
         Height          =   255
         Index           =   34
         Left            =   0
         TabIndex        =   115
         Top             =   660
         Width           =   1455
      End
      Begin VB.Label lblLabel 
         Caption         =   "1314"
         Height          =   255
         Index           =   27
         Left            =   0
         TabIndex        =   116
         Top             =   345
         Width           =   1455
      End
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2550
      Index           =   13
      Left            =   2940
      ScaleHeight     =   2550
      ScaleWidth      =   4425
      TabIndex        =   132
      Top             =   945
      Width           =   4425
      Begin VB.CheckBox chkOption 
         Caption         =   "1137"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   134
         Tag             =   "ProjectDockable"
         Top             =   270
         Width           =   3135
      End
      Begin VB.CheckBox chkOption 
         Caption         =   "1132"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   133
         Tag             =   "ProjectReloadLast"
         Top             =   15
         Width           =   3135
      End
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   "1240"
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   7260
      MouseIcon       =   "frmOptions_new.frx":0089
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   3990
      Width           =   360
   End
   Begin VB.Label lblHeader2 
      BackStyle       =   0  'Transparent
      Caption         =   "1047"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   2730
      TabIndex        =   4
      Top             =   150
      Width           =   4335
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   2715
      X2              =   7635
      Y1              =   3810
      Y2              =   3810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   2715
      X2              =   7620
      Y1              =   3825
      Y2              =   3825
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00808080&
      FillStyle       =   0  'Solid
      Height          =   360
      Left            =   2670
      Top             =   90
      Width           =   5070
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
' Developers Pad
' Version 1, BETA 2
' http://www.developerspad.com/
'
' © 1999-2000 VB Web Development
' You may not redistribute this source code,
' or distribute re-compiled versions of
' Developers Pad
'
Private cFavourites         As clsFavourites
Private cFlatCombo()        As clsFlatCombo
Private cFlatOpt()          As clsFlatOpt
Private bLangDirty()        As Boolean
Private sLastTemplate       As String
Private bTemplateChanged    As Boolean
Private bIgnore             As Boolean
Private sConstantsHeader    As String
Private lLastConstant       As Long
Private lLastItem           As Long
Private lCurrentColour      As Long
Private sCurLang            As String

'*** Button Events ***

'A different option has been selected...
'Enable Apply button
Private Sub cboFavourites_Click()
    cmdApply.Enabled = True
End Sub
Private Sub cboFonts_Click()
    cmdApply.Enabled = True
End Sub
Private Sub cboFontSize_Click()
    cmdApply.Enabled = True
End Sub

Private Sub cboKeyword_Click()

    If cboKeyword.Tag = -1 Or sCurLang = "" Then Exit Sub 'ignore...
    
    With cGlobalEditor.SyntaxFile(sCurLang).vSyntaxInfo
        'save the last data...
        pSaveKeyword
        Select Case cboKeyword.ListIndex
        Case 0 'keywords
            txtKeywords.Text = Left$(.sKeywords, .lSecondKeywordStart)
        Case 1 'keywords alt
            If Len(.sKeywords) > 0 Then
                txtKeywords.Text = Right$(.sKeywords, Len(.sKeywords) - .lSecondKeywordStart)
            Else
                txtKeywords.Text = ""
            End If
        Case 2 'procedures
            txtKeywords.Text = .sProcedures
        End Select
        
        txtKeywords.Text = Replace(txtKeywords.Text, "*", vbCrLf)
        txtKeywords.Text = StripChar(vbCrLf, txtKeywords.Text)
    End With
    
End Sub
Private Sub pSaveKeyword(Optional bOverride As Boolean = False)
Static lLastItem As Long
Dim cData As Syntax_Item
Dim sNew As String
    On Error Resume Next
    If txtKeywords.Tag = "1" And (lLastItem <> cboKeyword.ListIndex Or bOverride) Then
        'get a local copy
        cData = cGlobalEditor.SyntaxFile(sCurLang)
        With cData.vSyntaxInfo
            Select Case lLastItem
            Case 0 'keywords
                '
                sNew = "*" & Replace(txtKeywords.Text, vbCrLf, "*")
                If Len(.sKeywords) <> 0 Then
                    .sKeywords = Right$(.sKeywords, Len(.sKeywords) - .lSecondKeywordStart)
                    .sKeywords = sNew & .sKeywords
                    .lSecondKeywordStart = Len(sNew)
                Else
                    .sKeywords = sNew
                End If
            Case 1 'keywords alt
                sNew = "*" & Replace(txtKeywords.Text, vbCrLf, "*") & "*"
                .sKeywords = Left$(.sKeywords, .lSecondKeywordStart)
                .sKeywords = .sKeywords & sNew
            Case 2 'procedures...
                .sProcedures = "*" & Replace(txtKeywords.Text, vbCrLf, "*") & "*"
            End Select
        End With
        txtKeywords.Tag = ""
        lLastItem = cboKeyword.ListIndex
        'save...
        cGlobalEditor.SyntaxFile(sCurLang) = cData
    End If
End Sub
Private Sub cboLang_Click(Index As Integer)
    If Index = 0 Then chkLang(0).Enabled = (cboLang(0).ListIndex <> 3)
End Sub



Private Sub cmdBrowse_Click()
    Dim sPath As String
    'browse for VB exe...
    If cDialog.ShowOpenSaveDialog(False, "Find VB Executable", "Applications (*.exe)|*.exe", txtItem(5).Text, Me) = True Then
        txtItem(5).Text = CmDlg.FileName
    End If
End Sub

Private Sub cmdNewLanguage_Click()
    Dim sName As String
    Dim lIndex As Long
    '1366: Please enter a name for the new Language definition
    sName = cDialog.InputBox(LoadResString(1366), "New Language")
    If sName <> "" Then
        lIndex = cGlobalEditor.NewLanguage(sName)
        'resize dirty array
        ReDim Preserve bLangDirty(1 To cGlobalEditor.SyntaxFilesCount)
        pAddLanguage (lIndex)
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'save the option position
    SaveSetting REG_KEY, "Settings", "LastOptionSection", tvwOpt.ItemKey(tvwOpt.Selected)
    'restore controls...
    RestoreControls cFlatCombo, cFlatOpt
    sCurLang = ""
End Sub

Private Sub picColour_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    With cpkPicker
        If .Visible = True Then
            .Visible = False
        Else
            .ZOrder
            .Colour = picColour(Index).BackColor
            .Left = 2775 + picColour(Index).Left
            .Top = 630 + picColour(Index).Top + picColour(Index).Height
            If .Top + .Height > ScaleHeight Then
                .Top = 630 + picColour(Index).Top - .Height
                If .Top < 0 Then
                    .Top = 0
                    .Left = 2775 + picColour(Index).Left + picColour(Index).Width
                End If
            End If
            lCurrentColour = Index
            cmdCancel.Cancel = False
            cmdOK.Default = False
            .Visible = True
        End If
    End With
End Sub


Private Sub tvwOpt_ItemExpandingCancel(hItem As Long, ExpandType As vbwTreeView.ExpandTypeConstants, Cancel As Boolean)
    'If tvwOpt.ItemData(hItem) = 0 Then
    
        'is a folder...update it's image
    tvwOpt.ItemImage(hItem) = IIf(ExpandType = Expand, IndexForKey("FOLDEROPEN"), IndexForKey("FOLDERCLOSED"))
    lLastItem = tvwOpt.ItemChild(hItem)
    If hItem = tvwOpt.Selected Then
        tvwOpt.ItemImage(lLastItem) = IndexForKey("CUR_OPTION")
    End If
    If tvwOpt.ItemData(hItem) > 999 And ExpandType = Expand Then
        pSaveLangData
        pLoadSyntaxInfo tvwOpt.ItemText(hItem), 0
        'sCurLang = tvwOpt.ItemText(hItem) ' - 999
    End If
End Sub

Private Sub txtItem_Change(Index As Integer)
    cmdApply.Enabled = True
End Sub
Private Sub chkOption_Click(Index As Integer)
    cmdApply.Enabled = True
End Sub

Private Sub cmdCancel_Click()
    'close the options dialog... and don't
    'save changes
    Unload Me
End Sub

Private Sub cmdAdd_Click()
    'add a new constant (for use in insert window)
    lstConstants.AddItem txtConstant.Text & "=" & txtConstantValue.Text
End Sub
Private Sub cmdDelete_Click()
    'remove the selected constant
On Error Resume Next
    lstConstants.RemoveItem (lstConstants.ListIndex)
End Sub
Private Sub cmdAddFav_Click()
    'add a new favourite path for use in
    'open dialog
    Dim cBrowse As clsBrowseForFolder
    Dim sPath As String
    Set cBrowse = New clsBrowseForFolder
    'Get the folder
    sPath = cBrowse.BrowseForFolder()
    If sPath <> "" Then
        cFavourites.AddFavourite sPath, cboFavourites
        'changed...
        cmdApply.Enabled = True
    End If
End Sub
Private Sub cmdDelFav_Click()
    If cboFavourites.ListIndex <> -1 Then
        'delete the selected favourite
        cboFavourites.RemoveItem cboFavourites.ListIndex
        'changed...
        cmdApply.Enabled = True
    End If
End Sub
Private Sub cmdBrowseLocal_Click()
    'select a local path...
    'open dialog
    Dim cBrowse As clsBrowseForFolder
    Dim sPath As String
    Set cBrowse = New clsBrowseForFolder
    'Get the folder
    sPath = cBrowse.BrowseForFolder()
    If sPath <> "" Then
        txtItem(1).Text = sPath
        'changed...
        cmdApply.Enabled = True
    End If
End Sub
Private Sub cmdOK_Click()
    'changes made...save them
    If cmdApply.Enabled = True Then cmdApply_Click
    'close dialog
    Unload Me
End Sub
Private Sub cmdApply_Click()
On Error GoTo ErrHandler
    Dim i As Integer
    Dim sText As String
    'apply the changes
    
    'save the templates
    If SaveTemplate(cboTemplate.Text) = False Then Exit Sub
    SaveConstants
    'save the favourites
    cFavourites.SaveFavourites cboFavourites
    'save the 'easy' settings... we just use the info
    'stored in the option tags
    For i = 0 To chkOption.Count - 1
        SaveSetting REG_KEY, "Settings", chkOption(i).Tag, chkOption(i).Value
    Next
    'save the textbox values
    For i = 0 To txtItem.Count - 1
        sText = txtItem(i).Text
        If Right$(sText, 1) = "\" Or Right$(sText, 1) = "/" Then sText = Left$(sText, Len(sText) - 1)
        SaveSetting REG_KEY, "Settings", txtItem(i).Tag, txtItem(i).Text
    Next
    
    'save fonts
    If cboFonts.ListCount <> 0 Then
        'update local variables
        With vDefault
            .sFont = cboFonts.Text
            .nFontSize = cboFontSize.Text
            .nWordWrap = cboWrap.ListIndex
        End With
        'save to registry
        SaveSetting REG_KEY, "Settings", "FontName", cboFonts.Text
        SaveSetting REG_KEY, "Settings", "FontSize", cboFontSize.Text
        SaveSetting REG_KEY, "Settings", "WordWrap", cboWrap.ListIndex
    End If
    'save the current language data to current info
    pSaveLangData
    
    For i = 1 To cGlobalEditor.SyntaxFilesCount
        If bLangDirty(i) Then
            'write dirty data to file
            cGlobalEditor.SaveSyntaxInfo i
            'not dirty...
            bLangDirty(i) = False
        End If
    Next
    'we have saved everything
    cmdApply.Enabled = False
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Options.Apply"
End Sub

'*** Form Events ***

Private Sub Form_Load()
    On Error GoTo ErrHandler

    Dim i As Integer
    Dim lFolderIndex As Long
    Dim lBlankIndex  As Long
    Dim lOptIndex    As Long
    Dim sName        As String

    'Position containers
    For i = 1 To picScreen.Count
        With picScreen(i)
            .Move 2775, 630, 4935, 3120
            .Visible = False
        End With
    Next
    
    'Attach thin-3d style combo boxes and checkboxes
    MakeControlsFlat Controls, cFlatCombo, cFlatOpt

    'Init
    lLastConstant = -1
    
    With tvwOpt
        'Build the treeview
        .ShowSelected = True
        .SingleExpand = True
        'optimize
        .DisableCustomDraw = True
        .NoDragDrop = True
        'assign the image list
        .hImageList = frmMainForm.vbalMain.hIml 'imlIcons.hIml
        lFolderIndex = IndexForKey("FOLDERCLOSED")
        lBlankIndex = IndexForKey("BLANK")
        lOptIndex = IndexForKey("CUR_OPTION")
        'add all the categories
        .Add -1, LastChild, "Environment", LoadResString(1246), lFolderIndex
            .Add "Environment", LastChild, "Env_General", LoadResString(1247), lBlankIndex, IndexForKey("CUR_OPTION")
            .Add "Environment", LastChild, "Env_Documents", LoadResString(1249), lBlankIndex, lOptIndex
            .Add "Environment", LastChild, "Env_Editor", LoadResString(1257), lBlankIndex, lOptIndex
        .Add -1, LastChild, "Templates", LoadResString(1250), lFolderIndex
            .Add "Templates", LastChild, "Tmp_General", LoadResString(1247), lBlankIndex, lOptIndex
        .Add -1, LastChild, "IEPreview", LoadResString(1350), lFolderIndex
            .Add "IEPreview", LastChild, "IEP_General", LoadResString(1247), lBlankIndex, lOptIndex
        .Add -1, LastChild, "Projects", LoadResString(1248), lFolderIndex
            .Add "Projects", LastChild, "Prj_General", LoadResString(1247), lBlankIndex, lOptIndex
            .Add "Projects", LastChild, "Prj_Defaults", LoadResString(1307), lBlankIndex, lOptIndex
        .Add -1, LastChild, "CodeLibrary", LoadResString(1251), lFolderIndex
            .Add "CodeLibrary", LastChild, "CLb_General", LoadResString(1247), lBlankIndex, lOptIndex
            .Add "CodeLibrary", LastChild, "CLb_Constants", LoadResString(1205), lBlankIndex, lOptIndex
        .Add -1, LastChild, "Languages", LoadResString(1254), lFolderIndex
            
            'add all the installed languages...
            ReDim bLangDirty(1 To cGlobalEditor.SyntaxFilesCount)
            For i = 1 To cGlobalEditor.SyntaxFilesCount
                
                pAddLanguage (i)
            Next
            .Add "Languages", FirstChild, "Lng_General", LoadResString(1247), lBlankIndex, lOptIndex
        'assign the pictureboxes to the categories
        '(in picScreen control array)
        'Environment
        .ItemData("Env_General") = 12
        .ItemData("Env_Documents") = 4
        .ItemData("Env_Editor") = 11
        .ItemData("Lng_General") = 15
        'Templates
        .ItemData("Tmp_General") = 3 '5
        .ItemData("IEP_General") = 5
        'Project
        .ItemData("Prj_General") = 13
        .ItemData("Prj_Defaults") = 1
        'CodeLibrary
        .ItemData("CLb_General") = 14
        .ItemData("CLb_Constants") = 2
        'select the last option
        sName = GetSetting(REG_KEY, "Settings", "LastOptionSection", "Environment")
        If .IsValidNewKey(sName) = False Then .Selected = .ItemHandle(sName)
    End With
    'make colour boxes flat
    For i = 0 To picColour.Count - 1
        SetThin3DBorder picColour(i).hWnd
    Next
    'Load the resource stings
    LoadResStrings Controls
    'get the VB path
    GetVBPath
    'load the settings for options and
    'text boxes, using the info set in their tags
    For i = 0 To chkOption.Count - 1
        chkOption(i).Value = GetSetting(REG_KEY, "Settings", chkOption(i).Tag, 1)
    Next
    For i = 0 To txtItem.Count - 1
        txtItem(i).Text = GetSetting(REG_KEY, "Settings", txtItem(i).Tag, txtItem(i).Text)
    Next
    'select the correct wrap item
    cboWrap.ListIndex = vDefault.nWordWrap
    'load the folder favourites for use in open dialogs
    Set cFavourites = New clsFavourites
    cFavourites.LoadFavourites cboFavourites, Me
    'load the constants for use in Insert window
    LoadConstants
    'no changes have been made yet...
    cmdApply.Enabled = False
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "Options.Load"
End Sub
Private Sub pAddLanguage(Index As Long)
    Dim sName        As String
    Dim lFolderIndex As Long
    Dim lBlankIndex  As Long
    Dim lOptIndex    As Long
    Dim hItem        As Long
    
    lFolderIndex = IndexForKey("FOLDERCLOSED")
    lBlankIndex = IndexForKey("BLANK")
    lOptIndex = IndexForKey("CUR_OPTION")
    With tvwOpt
        sName = cGlobalEditor.SyntaxFile(Index).sName
        'load the data if we need to...
        If cGlobalEditor.SyntaxFile(Index).bLoaded = False Then cGlobalEditor.LoadSyntaxInfo sName
        hItem = .Add("Languages", AlphabeticalChild, "Lng_File_" & sName, sName, lFolderIndex)
        .ItemData(hItem) = 999 + Index
        'add the standard sub-items
        hItem = .Add("Lng_File_" & sName, LastChild, "Lng_File_" & sName & "_General", LoadResString(1247), lBlankIndex, lOptIndex)
        .ItemData(hItem) = 8
        'only add this item if it is code
        If cGlobalEditor.SyntaxFile(Index).vSyntaxInfo.bCode Then
            
            'don't add this item if it is just HTML
            If cGlobalEditor.SyntaxFile(Index).vSyntaxInfo.bHTML = False Or cGlobalEditor.SyntaxFile(Index).vSyntaxInfo.bHTMLExtension = True Then
                hItem = .Add("Lng_File_" & sName, LastChild, "Lng_File_" & sName & "_Keywords", LoadResString(1255), lBlankIndex, lOptIndex)
                .ItemData(hItem) = 6
            End If
            hItem = .Add("Lng_File_" & sName, LastChild, "Lng_File_" & sName & "_Colours", LoadResString(1256), lBlankIndex, lOptIndex)
            .ItemData(hItem) = 7
            hItem = .Add("Lng_File_" & sName, LastChild, "Lng_File_" & sName & "_Options", LoadResString(1346), lBlankIndex, lOptIndex)
            .ItemData(hItem) = 9
            hItem = .Add("Lng_File_" & sName, LastChild, "Lng_File_" & sName & "_Indent", LoadResString(1347), lBlankIndex, lOptIndex)
            .ItemData(hItem) = 10
        End If
    End With
End Sub
'Private Function plGetIconIndex(ByVal sKey As String) As Long
'    'Return the index of an item in the ImageList
'    If sKey = "-1" Then
'        plGetIconIndex = -1
'    Else
'        On Error Resume Next
'        plGetIconIndex = imlIcons.ItemIndex(sKey)
'    End If
'    If plGetIconIndex = 0 Then plGetIconIndex = -1
'End Function

'*** Constants Code ***
Private Sub SaveConstants()
    Dim iFileNum    As Integer
    Dim i           As Integer
    'save changes to the current item being edited
    Call lstConstants_Click
    'get a free filenum, and open the constants file
    iFileNum = FreeFile
    On Error Resume Next
    Open App.Path & "\constants.txt" For Output As iFileNum
    If Err = 0 Then
        'output the header
        Print #iFileNum, sConstantsHeader
        'output all the constants
        For i = 0 To lstConstants.ListCount - 1
            Print #iFileNum, lstConstants.List(i)
        Next i
        'close
        Close iFileNum
    End If
End Sub
Private Sub LoadConstants()
    'load the constants from the file...
    Dim iFileNum    As Integer
    Dim sConstant   As String
    'clear header
    sConstantsHeader = ""
    'get a free file number, and open the file
    iFileNum = FreeFile
    On Error Resume Next
    Open App.Path & "\constants.txt" For Input As iFileNum
    If Err = 0 Then
        Do While Not EOF(iFileNum)
            'read line by line...
            Line Input #iFileNum, sConstant
            If Left$(sConstant, 1) = ";" Then
                'it is a comment... append to header
                sConstantsHeader = sConstantsHeader & sConstant & vbCrLf
            ElseIf sConstant <> "" Then
                'add the item to the list
                lstConstants.AddItem sConstant
            End If
        Loop
        'close the file
        Close iFileNum
    End If
End Sub

'*** Fonts ***
Private Sub ListFonts()
    Dim i       As Integer
    Dim cCursor As clsCursor
    'setup cursor
    Set cCursor = New clsCursor
    cCursor.SetCursor vbHourglass
    For i = 0 To Screen.FontCount - 1 ' Determine number of fonts.
        cboFonts.AddItem Screen.Fonts(i)  ' Put each font into list box.
    Next i
    'set default font
    cboFonts.Text = vDefault.sFont
    'load font sizes
    For i = 8 To 40 Step 2
        cboFontSize.AddItem i
    Next i
    'set default font-size
    cboFontSize.Text = vDefault.nFontSize
End Sub

'*** Templates ***
Private Sub cboTemplate_Click()
Dim vTemplate   As TemplateInfo
Dim i           As Long
    'if we need to ignore this event, abort
    If bIgnore Then Exit Sub
    'save the last template...
    If SaveTemplate(sLastTemplate) = False Then
        'don't trigger this code again!
        bIgnore = True
        'if the save is cancelled, revert to last item
        cboTemplate.Text = sLastTemplate
        bIgnore = False
        Exit Sub
    End If
    If cboTemplate.Text <> "" Then
        'if there is an item selected...
        bIgnore = True
        'load the file
        txtContents.Text = LoadTextFile(App.Path & "\_templates\" & cboTemplate.Text & ".txt")
        'get the info for that template
        vTemplate = GetTemplateInfo(cboTemplate.Text & ".txt")
        'select the correct item from the syntax combo
        For i = 1 To cGlobalEditor.SyntaxFilesCount
            If cGlobalEditor.SyntaxFile(i).sFile = vTemplate.sSyntax Then
                cboTemplateLang.Text = cGlobalEditor.SyntaxFile(i).sName
                Exit For
            End If
        Next
        'no more ignoring!
        bIgnore = False
        'save the last template
        sLastTemplate = cboTemplate.Text
    End If
    'template hasn't changed
    bTemplateChanged = False
End Sub
Private Sub txtContents_Change()
    'ignore flag set, abort
    If bIgnore Then Exit Sub
    'template has changed
    bTemplateChanged = True
    'something has changed
    cmdApply.Enabled = True
End Sub
Private Function SaveTemplate(sTemplate As String) As Boolean
On Error GoTo ErrHandler
    Dim iFileNum As Integer
    'if there is no template specified, abort
    If sTemplate = Empty Then
        SaveTemplate = True
        Exit Function
    End If
    If bTemplateChanged Then
        'if the template has changed
        '"Save Changes To " & sTemplate & " Template?"
        Select Case cDialog.ShowYesNo(LoadResString(1159) & sTemplate & LoadResString(1160), True)
        Case No
            'successful, but no save
            SaveTemplate = True
        Case Cancelled
            'not successful
            SaveTemplate = False
        Case Yes
            ' save the file
            iFileNum = FreeFile
            Open App.Path & "\_templates\" & sTemplate & ".txt" For Output As iFileNum
            If txtContents.Text <> "" Then
                'output the content of the textbox
                Print #iFileNum, txtContents
            End If
            Close #iFileNum
            'successful
            SaveTemplate = True
        End Select
    Else
        SaveTemplate = True
    End If
    If SaveTemplate = True Then bTemplateChanged = False 'template has not changed..
    Exit Function
ErrHandler:
    cDialog.ErrHandler Err, Error, "Options.SaveTemplate"
    SaveTemplate = False
End Function

Private Function ListTemplates()
Dim i As Long
    'lists all available the templates
    'clear the combo
    cboTemplate.Clear
    'add all the items to the combo
    AddAllFilesInDir "_templates", cboTemplate
    'clear the languages
    cboTemplateLang.Clear
    'load them...
    For i = 1 To cGlobalEditor.SyntaxFilesCount
        cboTemplateLang.AddItem cGlobalEditor.SyntaxFile(i).sName
    Next
    'select first item, if possible
    If cboTemplate.ListCount <> 0 Then cboTemplate.ListIndex = 0
End Function

Private Sub lblHelp_Click()
    'display help...
    cDialog.ShowHelpTopic 3, hWnd
End Sub

'*** Constants ***

Private Sub lstConstants_Click()
    Dim lNamePos As Integer

    If lLastConstant <> -1 Then
        'save the last constant
        If lstConstants.List(lLastConstant) <> txtConstant.Text & "=" & txtConstantValue.Text Then
            'it has changed...
            'save its name and value
            lstConstants.List(lLastConstant) = txtConstant.Text & "=" & txtConstantValue.Text
        End If
    End If
    If lstConstants.Text <> "" Then
        'an item is selected
        bIgnore = True
        'get the = pos
        lNamePos = InStr(1, lstConstants.Text, "=")
        'display the name in txtConstant
        txtConstant.Text = Left$(lstConstants.Text, lNamePos - 1)
        'and its value in txtConstantValue
        txtConstantValue.Text = Right$(lstConstants.Text, Len(lstConstants.Text) - lNamePos)
        bIgnore = False
    End If
    'save the last constant
    lLastConstant = lstConstants.ListIndex
End Sub
Private Sub txtConstant_Change()
    'call constant change proc
    ConstantChange
End Sub
Private Sub txtConstantValue_Change()
    'call constant change proc
    ConstantChange
End Sub
Private Sub ConstantChange()
    'ignore flag is set... abort
    If bIgnore Then Exit Sub
    'something has changed
    cmdApply.Enabled = True
    'enable the add button if there is a name
    'and value set for the constant
    cmdAdd.Enabled = Not (txtConstant.Text = "" And txtConstantValue.Text = "")
End Sub

'*** TreeView ***
Private Sub tvwOpt_SelChanged()
On Error GoTo ErrHandler
    Dim bItem       As Boolean
    Dim lParent     As Long
    Dim lItem       As Long
    Dim lTab        As Long
    pSaveLangData
    'get the currently selected item, and its parent
    lItem = tvwOpt.Selected
    'is it an item?
    bItem = (tvwOpt.ItemChild(lItem) = 0)
    lParent = tvwOpt.ItemParent(lItem)
    If bItem = False Then
        'is a folder...
        lParent = lItem
        'use it's first child as the item
        lItem = tvwOpt.ItemChild(lItem)
    End If
'    If tvwOpt.ItemImage(lParent) = IndexForKey("FOLDEROPEN") Then
'        lSelected = tvwOpt.ItemChild(lSelected)
'    End If
    'If bItem Then
        'if it has a parent, display its parent info first (ie General)
        lblHeader1.Caption = tvwOpt.ItemText(lParent) '& "- " & tvwOpt.ItemText(tvwOpt.Selected)
        'then display the selected sub-items text
        lblHeader2.Caption = tvwOpt.ItemText(lItem)
        If lLastItem <> lItem And lLastItem <> 0 Then
            tvwOpt.ItemImage(lLastItem) = IndexForKey("BLANK")
            tvwOpt.ItemSelectedImage(lLastItem) = IndexForKey("CUR_OPTION")
            lLastItem = 0 'IIf(lParent = 0, 0, lSelected)
        End If
'    Else
'        'item selected is a parent... display its text
'        lblHeader1.Caption = tvwOpt.ItemText(lSelected)
'        'and then the text of its first child
'        lblHeader2.Caption = tvwOpt.ItemText(tvwOpt.ItemChild(lSelected))
'    End If
    
    'get the picbox we are supposed to display
   ' If lParent <> 0 Then
        'item has a parent.. the info we want is in the
        'selected item
        lTab = tvwOpt.ItemData(lItem)
'    Else
'        'item is a parent, the info we want is in its first child
'        lTab = tvwOpt.ItemData(tvwOpt.ItemChild(lSelected))
'    End If

    If lTab > 0 Then
        If lTab > 999 Then
            'lTab = lTab - 999
            'pLoadLanguageData(lTab)
        Else
            Select Case lTab
            Case 11 'fonts
                'load the fonts if we haven't already
                If cboFonts.ListCount = 0 Then ListFonts
            Case 3 'templates
                'load the templates, if we haven't already
                If cboTemplate.ListCount = 0 Then ListTemplates
            Case 6, 7, 8, 9, 10 'language options
                pLoadSyntaxInfo tvwOpt.ItemText(lParent), lTab
                
            End Select
            DoEvents
            'make the correct picbox visible and topmost
            picScreen(lTab).ZOrder
            picScreen(lTab).Visible = True
        End If
    End If
    Exit Sub
ErrHandler:
    cDialog.ErrHandler Err, Error, "DevPad.Options:tvwOpt_SelChanged"
End Sub

Private Sub pLoadSyntaxInfo(sName As String, lTab As Long)
    Dim lPos As Long
    Dim i As Long
    
    If sCurLang = sName Then Exit Sub
    sCurLang = sName
    bIgnore = True
    'Select Case lTab
   ' Case 8 'general
        With cGlobalEditor.SyntaxFile(sName)
            txtLang(1).Text = .sName 'Name
            txtLang(0).Text = .sFilterName 'filter name
            txtLang(2).Text = .sFilter 'filter
        End With
        With cGlobalEditor.SyntaxFile(sName).vSyntaxInfo
            chkLang(0).Value = Abs(.bCode)
            If .bRTF Then
                cboLang(0).ListIndex = 3
            ElseIf .bHTMLExtension Then
                cboLang(0).ListIndex = 2
            ElseIf .bHTML Then
                cboLang(0).ListIndex = 1
            Else
                cboLang(0).ListIndex = 0
            End If
            chkLang(1).Value = Abs(.vCaseSensitive = vbTextCompare)
            chkLang(2).Value = Abs(.bAutoCase)
            'bug with flat opt class!
            chkLang(1).Refresh
            chkLang(2).Refresh
            chkLang(3).Refresh
            'only allow colourizing if it isn't RTF
            chkLang(0).Enabled = (cboLang(0).ListIndex <> 3)
            'only allow ignore case if it is code
            chkLang(1).Enabled = (chkLang(0).Value = 1)
            'only allow auto case if we are not ignoring case!
            chkLang(2).Enabled = (chkLang(1).Value = 1)
        End With
   ' Case 7 'colours
        With cGlobalEditor.SyntaxFile(sName).vSyntaxInfo
            picColour(0).BackColor = .vClr_Text
            picColour(1).BackColor = .vClr_Keyword
            picColour(2).BackColor = .vClr_Keyword2
            picColour(3).BackColor = .vClr_Comment
            picColour(4).BackColor = .vClr_Operator
            picColour(5).BackColor = .vClr_HTMLComment
            picColour(6).BackColor = .vClr_HTMLTag
            picColour(7).BackColor = .vClr_HTMLScript
            picColour(8).BackColor = .vClr_HTMLExTag
            For i = 1 To 8
                If (Not (.bHTML Or .bHTMLExtension) And i > 4) Or (.bHTML And Not .bHTMLExtension And i < 4) Or .bRTF Or Not .bCode Then
                    picColour(i).BackColor = vbButtonFace
                    picColour(i).Enabled = False
                Else
                    picColour(i).Enabled = True
                End If
            Next
        End With
   ' Case 6
        cboKeyword.Tag = -1
        cboKeyword.ListIndex = 0
        cboKeyword.Tag = 0
        cboKeyword_Click
        With cGlobalEditor.SyntaxFile(sName).vSyntaxInfo
            txtLang(3).Text = .sOperators
            txtWordSeperators.Text = cGlobalEditor.ParseHiddenChars(Replace(Right$(.sSeps, Len(.sSeps) - Len(.sOperators)), "*", ""))
        End With
   ' Case 9 'options
        With cGlobalEditor.SyntaxFile(sName).vSyntaxInfo
            txtLang(4).Text = .sSingleComment
            txtLang(15).Text = .sSingleCommentAlt
            
            txtLang(16).Text = .sStringEncoded
            
            txtLang(5).Text = .sStrings

            txtLang(7).Text = .sMultiCommentStart
            txtLang(10).Text = .sMultiCommentEnd
            txtLang(9).Text = .sFalseQuote
            txtLang(8).Text = .sSingleCommentEsc
            txtLang(11).Text = .sHTMLExtensionStart
            txtLang(12).Text = .sHTMLExtensionEnd
            txtLang(11).Enabled = .bHTMLExtension
            txtLang(12).Enabled = .bHTMLExtension
            'server-side script labels
            lblLabel(47).Enabled = .bHTMLExtension
            lblLabel(48).Enabled = .bHTMLExtension
        End With
   ' Case 10 'indent
        With cGlobalEditor.SyntaxFile(sName).vSyntaxInfo
            chkLang(5).Value = Abs(.bTabIndent)
            chkLang(4).Value = Abs(.bDelIndent)
            chkLang(3).Value = Abs(.bAutoIndent)
            txtLang(17).Text = .sAutoIndentChar
            txtLang(6).Text = .sAutoOutdentChar

            txtLang(13).Text = cGlobalEditor.ParseHiddenChars(.sIndent)
            txtLang(14).Text = cGlobalEditor.ParseHiddenChars(.sHTMLIndent)
            txtLang(14).Enabled = .bHTMLExtension
            lblLabel(51).Enabled = .bHTMLExtension
        End With
   ' End Select
    bIgnore = False

End Sub
Private Sub txtItem_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 6 Then KeyAscii = NumericOnly(KeyAscii)
End Sub

Private Sub cpkPicker_CancelPick()
    pHidePicker
End Sub
Private Sub cpkPicker_ColourChanged(NewColour As stdole.OLE_COLOR)
    pHidePicker
    picColour(lCurrentColour).BackColor = cpkPicker.Colour
    pMakeLangDirty
End Sub

Private Sub pHidePicker()
    cpkPicker.Visible = False
    cmdCancel.Cancel = True
    cmdOK.Default = True
End Sub
Private Sub pSaveLangData()
Dim cData As Syntax_Item
    If sCurLang = "" Then Exit Sub
    If bLangDirty(cGlobalEditor.SyntaxFileIndex(sCurLang)) = False Then Exit Sub
    pSaveKeyword True
    cData = cGlobalEditor.SyntaxFile(sCurLang)
    'saves the lang data to the info array
    With cData
        .sName = txtLang(1).Text 'Name N/P
        .sFilterName = txtLang(0).Text 'filter name
        .sFilter = txtLang(2).Text 'filter
    End With
    With cData.vSyntaxInfo
        .bCode = chkLang(0).Value
        .bRTF = (cboLang(0).ListIndex = 3)
        .bHTMLExtension = (cboLang(0).ListIndex = 2)
        .bHTML = (cboLang(0).ListIndex = 1 Or .bHTMLExtension)
        .vCaseSensitive = IIf(chkLang(1).Value, vbTextCompare, vbBinaryCompare)
        .bAutoCase = chkLang(2).Value
        'only allow colourizing if it isn't RTF
    'colours
        .vClr_Text = picColour(0).BackColor
        .vClr_Keyword = picColour(1).BackColor
        .vClr_Keyword2 = picColour(2).BackColor
        .vClr_Comment = picColour(3).BackColor
        .vClr_Operator = picColour(4).BackColor
        .vClr_HTMLComment = picColour(5).BackColor
        .vClr_HTMLTag = picColour(6).BackColor
        .vClr_HTMLScript = picColour(7).BackColor
        .vClr_HTMLExTag = picColour(8).BackColor
    'keywords
        .sOperators = txtLang(3).Text
        .sSeps = .sOperators & cGlobalEditor.GetHiddenChars(txtWordSeperators.Text)
    'options
        .sSingleComment = txtLang(4).Text
        .sSingleCommentAlt = txtLang(15).Text
        
        .sStrings = txtLang(5).Text
        .sStringEncoded = txtLang(16).Text
        
        .sMultiCommentStart = txtLang(7).Text
        .sMultiCommentEnd = txtLang(10).Text
        .sFalseQuote = txtLang(9).Text
        .sSingleCommentEsc = txtLang(8).Text
        .sHTMLExtensionStart = txtLang(11).Text
        .sHTMLExtensionEnd = txtLang(12).Text
        .bHTMLExtension = txtLang(11).Enabled
        .bHTMLExtension = txtLang(12).Enabled
    'indent
        .bTabIndent = chkLang(5).Value
        .bDelIndent = chkLang(4).Value
        .bAutoIndent = chkLang(3).Value
        .sAutoIndentChar = txtLang(17).Text
        .sAutoOutdentChar = txtLang(6).Text

        .sIndent = cGlobalEditor.GetHiddenChars(txtLang(13).Text)
        .sHTMLIndent = cGlobalEditor.GetHiddenChars(txtLang(14).Text)
    End With
    'done!
    cGlobalEditor.SyntaxFile(sCurLang) = cData
End Sub
Private Sub pMakeLangDirty()
    If sCurLang <> "" And bIgnore = False Then
        bLangDirty(cGlobalEditor.SyntaxFileIndex(sCurLang)) = True
        cmdApply.Enabled = True
    End If
End Sub

Private Sub txtKeywords_Change()
    pMakeLangDirty
    txtKeywords.Tag = "1"
End Sub

Private Sub txtLang_Change(Index As Integer)
    pMakeLangDirty
End Sub
Private Sub chkLang_Click(Index As Integer)
    pMakeLangDirty
    chkLang(1).Enabled = (chkLang(0).Value = 1)
    chkLang(2).Enabled = (chkLang(1).Value = 1)
End Sub

Private Sub txtLang_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If Index = 5 Or Index = 6 Then
        'only one char each...
        If Len(txtLang(Index).Text) = 1 And (KeyCode <> vbKeyDelete And KeyCode <> vbKeyBack) Then KeyCode = 0
    End If
End Sub

Private Sub txtWordSeperators_Change()
    pMakeLangDirty
End Sub
