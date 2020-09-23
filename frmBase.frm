VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Begin VB.Form frmBase 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Liquid Skyes "
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "MS Dialog Light"
      Size            =   6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H000000FF&
   Icon            =   "frmBase.frx":0000
   KeyPreview      =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   11520
   ScaleMode       =   0  'User
   ScaleWidth      =   15360
   Begin VB.Frame fraProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   11535
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   3930
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00B9CBAF&
         Caption         =   "Set By"
         Height          =   225
         Index           =   0
         Left            =   1980
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   127
         Top             =   1830
         UseMaskColor    =   -1  'True
         Width           =   945
      End
      Begin VB.CheckBox chkPause 
         BackColor       =   &H00000000&
         Caption         =   "Pause"
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   2040
         TabIndex        =   221
         Top             =   1560
         Width           =   855
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D696A2&
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   33
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   312
         Top             =   8520
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00D696A2&
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   33
         Left            =   1005
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   311
         Top             =   8520
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00B9CBAF&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   1920
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00B9CBAF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   11
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   105
         Top             =   1920
         Width           =   300
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Index           =   28
         Left            =   -30
         Locked          =   -1  'True
         TabIndex        =   200
         Text            =   " Time  / 10^6"
         Top             =   1800
         Width           =   2055
      End
      Begin VB.CheckBox chkAvalue 
         BackColor       =   &H00000080&
         Caption         =   "Click To Manual Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   330
         Left            =   30
         TabIndex        =   212
         Top             =   1560
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FF00&
         Height          =   225
         Index           =   15
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   187
         Text            =   "M *"
         Top             =   6240
         Width           =   495
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Index           =   29
         Left            =   420
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   348
         Text            =   "fps"
         Top             =   630
         Width           =   495
      End
      Begin VB.CheckBox chkAgr 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00473842&
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   1
         Left            =   2040
         TabIndex        =   347
         Top             =   7800
         Width           =   185
      End
      Begin VB.CheckBox chkAgr 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00473842&
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   0
         Left            =   2280
         TabIndex        =   346
         Top             =   7800
         Width           =   185
      End
      Begin VB.CheckBox chkPant 
         BackColor       =   &H000000C0&
         Caption         =   "Throb Mode"
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Left            =   2640
         TabIndex        =   343
         Top             =   7800
         Width           =   1260
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   2430
         MaskColor       =   &H000000FF&
         TabIndex        =   342
         Top             =   3000
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   5
         Left            =   3585
         MaskColor       =   &H000000FF&
         TabIndex        =   341
         Top             =   3000
         Width           =   345
      End
      Begin VB.TextBox txtMaxShot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   3330
         MousePointer    =   1  'Arrow
         TabIndex        =   310
         Text            =   "100"
         Top             =   10350
         Width           =   500
      End
      Begin VB.CheckBox chkBreakShot 
         BackColor       =   &H00000080&
         Caption         =   "Break On"
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Index           =   1
         Left            =   1950
         TabIndex        =   309
         ToolTipText     =   "Auto Shot"
         Top             =   10320
         Width           =   1965
      End
      Begin VB.CheckBox chkAutoShot 
         BackColor       =   &H00000080&
         Caption         =   "&Auto Shot"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   30
         TabIndex        =   242
         ToolTipText     =   "Auto Shot"
         Top             =   10320
         Width           =   3765
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H0000FF00&
         Caption         =   "&Restart"
         Height          =   345
         Index           =   10
         Left            =   2010
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   165
         Top             =   9000
         UseMaskColor    =   -1  'True
         Width           =   600
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00000000&
         Caption         =   "Auto Clear"
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Index           =   4
         Left            =   1350
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   133
         Top             =   9000
         Value           =   1  'Checked
         Width           =   705
      End
      Begin VB.CommandButton CmdDefault 
         BackColor       =   &H0080FFFF&
         Caption         =   "Set Defaults"
         Height          =   345
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   339
         Top             =   9240
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin VB.TextBox txtRST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   1980
         TabIndex        =   195
         Text            =   "0"
         Top             =   2040
         Width           =   945
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00BCB89A&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   2640
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   299
         Top             =   5295
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   3630
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   298
         Top             =   5295
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   7
         Left            =   2940
         TabIndex        =   203
         Text            =   "255"
         Top             =   5295
         Width           =   660
      End
      Begin VB.CheckBox chkAlphaEnable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00473842&
         ForeColor       =   &H00E0E0E0&
         Height          =   185
         Left            =   3765
         TabIndex        =   206
         Top             =   5070
         Width           =   180
      End
      Begin VB.CheckBox chkAlpha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00473842&
         ForeColor       =   &H00E0E0E0&
         Height          =   185
         Left            =   2640
         TabIndex        =   205
         Top             =   5070
         Width           =   180
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FFFF&
         Height          =   480
         Index           =   23
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   338
         Text            =   "Alpha"
         Top             =   5040
         Width           =   1350
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00B69ACD&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   32
         Left            =   1245
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   327
         Top             =   5025
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00B69ACD&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   28
         Left            =   1245
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   329
         Top             =   4440
         Width           =   300
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00004040&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   30
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   204
         Text            =   "Draw Mode"
         Top             =   8280
         Width           =   1260
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E0E0E0&
         Height          =   270
         ItemData        =   "frmBase.frx":08CA
         Left            =   2640
         List            =   "frmBase.frx":0904
         TabIndex        =   124
         Top             =   8520
         Width           =   1260
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   2430
         MaskColor       =   &H000000FF&
         TabIndex        =   337
         Top             =   3240
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   2815
         MaskColor       =   &H000000FF&
         TabIndex        =   336
         Top             =   3240
         Value           =   1  'Checked
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "3"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   3
         Left            =   3200
         MaskColor       =   &H000000FF&
         TabIndex        =   335
         Top             =   3240
         Width           =   345
      End
      Begin VB.CheckBox chkCol 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "4"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   4
         Left            =   3585
         MaskColor       =   &H000000FF&
         TabIndex        =   334
         Top             =   3240
         Width           =   345
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   27
         Left            =   3630
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   253
         Top             =   7417
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   27
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   252
         Top             =   7417
         Width           =   300
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F8AFA9&
         Height          =   150
         Index           =   19
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   333
         Text            =   "View Size /"
         Top             =   8280
         Width           =   1065
      End
      Begin VB.CommandButton cmdChandSaveFolder 
         BackColor       =   &H00B9CBAF&
         Caption         =   "Change Save Folder"
         Height          =   345
         Left            =   1980
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   246
         Top             =   9720
         UseMaskColor    =   -1  'True
         Width           =   1935
      End
      Begin VB.CheckBox chkAutoMax 
         BackColor       =   &H00F275B0&
         Caption         =   "Set Max With Time"
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   1560
         MaskColor       =   &H000000FF&
         TabIndex        =   213
         Top             =   3960
         Width           =   705
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00B69ACD&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   21
         Left            =   1245
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   174
         Top             =   3675
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   20
         Left            =   1245
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   170
         Top             =   3195
         Width           =   300
      End
      Begin VB.TextBox txtRST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         ForeColor       =   &H001B171C&
         Height          =   225
         Index           =   6
         Left            =   1560
         TabIndex        =   197
         Text            =   "10"
         Top             =   3240
         Width           =   705
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   28
         Left            =   330
         TabIndex        =   332
         Text            =   "0"
         Top             =   4440
         Width           =   900
      End
      Begin VB.CheckBox chkLastP 
         BackColor       =   &H002D061B&
         Caption         =   "Last Points"
         ForeColor       =   &H00F1E4F3&
         Height          =   195
         Left            =   30
         MaskColor       =   &H000000FF&
         TabIndex        =   331
         Top             =   4800
         Width           =   1545
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Index           =   5
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   330
         Text            =   "Start Points"
         Top             =   4200
         Width           =   1550
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         HideSelection   =   0   'False
         Index           =   32
         Left            =   330
         TabIndex        =   328
         Text            =   "100"
         Top             =   5025
         Width           =   900
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   20
         Left            =   330
         TabIndex        =   184
         Text            =   "1"
         Top             =   3240
         Width           =   900
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   21
         Left            =   330
         TabIndex        =   182
         Text            =   "10000"
         Top             =   3720
         Width           =   900
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   20
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   172
         Top             =   3195
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   24
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   2535
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   193
         Top             =   2745
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   26
         Left            =   1680
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   225
         Top             =   2310
         Width           =   300
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Set By"
         Height          =   225
         Index           =   1
         Left            =   1980
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   126
         Top             =   2400
         UseMaskColor    =   -1  'True
         Width           =   945
      End
      Begin VB.TextBox txtRST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00B9CBAF&
         ForeColor       =   &H001B171C&
         Height          =   225
         Index           =   7
         Left            =   1560
         TabIndex        =   198
         Text            =   "25796"
         Top             =   3720
         Width           =   705
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00A5CFD8&
         Height          =   270
         ItemData        =   "frmBase.frx":09C5
         Left            =   2640
         List            =   "frmBase.frx":09C7
         Style           =   2  'Dropdown List
         TabIndex        =   326
         Top             =   8880
         Width           =   1260
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00F8AFA9&
         Caption         =   "Set By"
         Height          =   225
         Index           =   6
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   171
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00A2C7E1&
         Caption         =   "Set By"
         Height          =   260
         Index           =   7
         Left            =   1560
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   175
         Top             =   3480
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00B69ACD&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   32
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   277
         Top             =   5025
         Width           =   300
      End
      Begin VB.TextBox txtFrm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arabic Transparent"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   60
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   295
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   630
         Width           =   360
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F1E4F3&
         DragMode        =   1  'Automatic
         ForeColor       =   &H001B171C&
         Height          =   825
         Left            =   30
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   132
         Text            =   "frmBase.frx":09C9
         Top             =   600
         Width           =   2895
      End
      Begin VB.CheckBox chkAutoFix 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00473842&
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   360
         TabIndex        =   324
         Top             =   8550
         Width           =   185
      End
      Begin VB.CommandButton cmdDriectory 
         BackColor       =   &H00C9B6D1&
         Caption         =   "Add Driectory"
         Height          =   285
         Left            =   2730
         TabIndex        =   248
         Top             =   10080
         Width           =   1185
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         HideSelection   =   0   'False
         Index           =   33
         Left            =   480
         TabIndex        =   313
         Text            =   "1"
         Top             =   8550
         Width           =   540
      End
      Begin VB.TextBox txtMaxShot 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   3330
         MousePointer    =   1  'Arrow
         TabIndex        =   234
         Text            =   "100"
         Top             =   10650
         Width           =   500
      End
      Begin VB.CheckBox chkBreakShot 
         BackColor       =   &H00000080&
         Caption         =   "New Dir On "
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Index           =   0
         Left            =   1950
         TabIndex        =   308
         ToolTipText     =   "Auto Shot"
         Top             =   10620
         Width           =   1965
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   24
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   209
         Top             =   2535
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   125
         Top             =   2775
         Width           =   300
      End
      Begin VB.CommandButton cmdMaxPoints 
         BackColor       =   &H00C9A4DF&
         Caption         =   "Set Max = 148900"
         Height          =   255
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   176
         Top             =   3960
         Width           =   1520
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dr 3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B171C&
         Height          =   210
         Index           =   5
         Left            =   2430
         TabIndex        =   278
         Top             =   4680
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Draw String"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B171C&
         Height          =   210
         Index           =   4
         Left            =   2430
         TabIndex        =   276
         Top             =   3480
         Width           =   1500
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dr 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B171C&
         Height          =   210
         Index           =   3
         Left            =   2430
         TabIndex        =   275
         Top             =   4440
         Value           =   1  'Checked
         Width           =   1500
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dr 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B171C&
         Height          =   210
         Index           =   1
         Left            =   2430
         TabIndex        =   274
         Top             =   4200
         Width           =   1500
      End
      Begin VB.TextBox txtPath 
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE4E0&
         DragMode        =   1  'Automatic
         Height          =   225
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   267
         Top             =   10080
         Width           =   1815
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00B69ACD&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   28
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   266
         Top             =   4440
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   22
         Left            =   3000
         TabIndex        =   262
         Text            =   "640"
         Top             =   2070
         Width           =   450
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   22
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   261
         Top             =   2280
         Width           =   450
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         Caption         =   "è"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   22
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   260
         Top             =   1800
         Width           =   450
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   31
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   259
         Top             =   5617
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   31
         Left            =   3630
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   258
         Top             =   5617
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   30
         Left            =   3630
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   257
         Top             =   6217
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   30
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   256
         Top             =   6217
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   29
         Left            =   3630
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   255
         Top             =   6817
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   29
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   254
         Top             =   6817
         Width           =   300
      End
      Begin VB.TextBox txtRST 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   1980
         TabIndex        =   196
         Text            =   "1"
         Top             =   2640
         Width           =   945
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00B69ACD&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   21
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   173
         Top             =   3675
         Width           =   300
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Draw String"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B171C&
         Height          =   210
         Index           =   2
         Left            =   2430
         TabIndex        =   250
         Top             =   3960
         Width           =   1500
      End
      Begin VB.TextBox txtShotCount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00512D4B&
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   30
         TabIndex        =   235
         Text            =   "0"
         Top             =   10080
         Width           =   825
      End
      Begin VB.CheckBox chkShotAll 
         BackColor       =   &H00000080&
         Caption         =   "&Shot All Frame"
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   30
         TabIndex        =   249
         ToolTipText     =   "Auto Shot"
         Top             =   10620
         Width           =   3885
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00F275B0&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   8
         Left            =   3615
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   240
         Top             =   10920
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00F275B0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   8
         Left            =   2070
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   237
         Top             =   10920
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00F275B0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   236
         Top             =   10920
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00F275B0&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   12
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   10
         Left            =   1605
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   239
         Top             =   10920
         Width           =   300
      End
      Begin VB.CommandButton cmdSF 
         BackColor       =   &H0000FF00&
         Caption         =   "Shot a Pic"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   247
         Top             =   9720
         UseMaskColor    =   -1  'True
         Width           =   1905
      End
      Begin VB.TextBox txtQua 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0C2DA&
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   285
         TabIndex        =   245
         Text            =   "35"
         ToolTipText     =   "JPG Quality"
         Top             =   11130
         Width           =   1335
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         BorderStyle     =   0  'None
         ForeColor       =   &H001E1E1E&
         Height          =   285
         Index           =   3
         Left            =   285
         Locked          =   -1  'True
         TabIndex        =   244
         Text            =   "Photo Quality"
         Top             =   10920
         Width           =   1335
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0C2DA&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   8
         Left            =   2325
         MaxLength       =   4
         OLEDragMode     =   1  'Automatic
         TabIndex        =   243
         Text            =   "2"
         ToolTipText     =   "Auto Shot Interval"
         Top             =   11130
         Width           =   1335
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         BorderStyle     =   0  'None
         ForeColor       =   &H001E1E1E&
         Height          =   285
         Left            =   2325
         Locked          =   -1  'True
         TabIndex        =   241
         Text            =   "Interval"
         ToolTipText     =   "JPG Quality"
         Top             =   10920
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00B4E0CE&
         ForeColor       =   &H001E1E1E&
         Height          =   225
         Left            =   3330
         TabIndex        =   238
         Text            =   "Sec"
         ToolTipText     =   "JPG Quality"
         Top             =   10995
         Width           =   345
      End
      Begin VB.TextBox txtLQT2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F1E4F3&
         DragMode        =   1  'Automatic
         ForeColor       =   &H001B171C&
         Height          =   825
         Left            =   2920
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   131
         Text            =   "frmBase.frx":0A31
         Top             =   600
         Width           =   1010
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   23
         Left            =   3480
         MultiLine       =   -1  'True
         TabIndex        =   233
         Text            =   "frmBase.frx":0A41
         Top             =   2070
         Width           =   450
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   23
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   232
         Top             =   1800
         Width           =   450
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "ê"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   23
         Left            =   3480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   231
         Top             =   2280
         Width           =   450
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H0000FFFF&
         Height          =   165
         Index           =   25
         Left            =   3328
         TabIndex        =   230
         Text            =   "1"
         Top             =   2610
         Width           =   275
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "®"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   25
         Left            =   3000
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   229
         Top             =   2550
         Width           =   330
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FFFF&
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   25
         Left            =   3600
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   228
         Top             =   2550
         Width           =   330
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   26
         Left            =   30
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   2295
         Width           =   300
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404040&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Index           =   7
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   223
         Text            =   "Colors"
         Top             =   3000
         Width           =   1500
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00B9CBAF&
         Caption         =   "571"
         Height          =   225
         Index           =   2
         Left            =   345
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   5400
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00B9CBAF&
         Caption         =   "0.1"
         Height          =   225
         Index           =   5
         Left            =   345
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   169
         Top             =   7200
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00B9CBAF&
         Caption         =   "1"
         Height          =   225
         Index           =   4
         Left            =   345
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   6600
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.CommandButton cmd0 
         BackColor       =   &H00B9CBAF&
         Caption         =   "491"
         Height          =   225
         Index           =   3
         Left            =   345
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   167
         Top             =   6000
         UseMaskColor    =   -1  'True
         Width           =   705
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE4E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   16
         Left            =   840
         TabIndex        =   191
         Text            =   "1"
         Top             =   5640
         Width           =   2775
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FF00&
         Height          =   225
         Index           =   14
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   189
         Text            =   "E *"
         Top             =   5640
         Width           =   495
      End
      Begin VB.CommandButton cmdInvertPage 
         BackColor       =   &H00D696A2&
         Caption         =   "Inverse Screen"
         Height          =   345
         Left            =   1350
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   211
         Top             =   8640
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin VB.CommandButton cmdCls 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Clear Screen"
         Height          =   345
         Left            =   1350
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   210
         Top             =   8280
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE4E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   17
         Left            =   840
         TabIndex        =   190
         Text            =   "1"
         Top             =   6240
         Width           =   2775
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE4E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   19
         Left            =   840
         TabIndex        =   188
         Text            =   "0.1"
         Top             =   7440
         Width           =   2775
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00CFE4E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   18
         Left            =   840
         TabIndex        =   186
         Text            =   "1"
         Top             =   6840
         Width           =   2775
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   11
         Left            =   330
         Locked          =   -1  'True
         TabIndex        =   179
         Text            =   "1"
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   330
         TabIndex        =   192
         Text            =   "1"
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FF00&
         Height          =   270
         Index           =   16
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   185
         Text            =   "CC *"
         Top             =   7440
         Width           =   495
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F1E4F3&
         Height          =   260
         Index           =   1
         Left            =   330
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   194
         Text            =   "Time +=  Velocity"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Index           =   22
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   181
         Text            =   "Max Points"
         Top             =   3480
         Width           =   1550
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00F1E4F3&
         Height          =   195
         Index           =   21
         Left            =   30
         Locked          =   -1  'True
         TabIndex        =   183
         Text            =   "Aggregation"
         Top             =   3000
         Width           =   1550
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FF00&
         Height          =   225
         Index           =   11
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   180
         Text            =   "C *"
         Top             =   6840
         Width           =   495
      End
      Begin VB.CheckBox chkTimeEnable 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Draw String"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H001B171C&
         Height          =   210
         Index           =   0
         Left            =   2430
         TabIndex        =   177
         Top             =   3720
         Width           =   1500
      End
      Begin VB.PictureBox picBProcs 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   30
         ScaleHeight     =   225
         ScaleWidth      =   3945
         TabIndex        =   134
         Top             =   720
         Width           =   3975
      End
      Begin VB.CommandButton cmdCtrl 
         BackColor       =   &H00E4E0E0&
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1800
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   340
         ToolTipText     =   "Logs"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdOpenTelo 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   15.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2532
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   294
         ToolTipText     =   "Camera"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdLogs 
         BackColor       =   &H00E4E0E0&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   14.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2166
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   293
         ToolTipText     =   "Logs"
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdMiniMize 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   2898
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   214
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdMax 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   3249
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   202
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdExit 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "r"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   3600
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   215
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CommandButton cmdMini 
         Appearance      =   0  'Flat
         BackColor       =   &H00E4E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Webdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   3240
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   201
         Top             =   30
         UseMaskColor    =   -1  'True
         Visible         =   0   'False
         Width           =   330
      End
      Begin VB.Label cmdAbout 
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "About"
         ForeColor       =   &H00CFE4E0&
         Height          =   150
         Left            =   3360
         TabIndex        =   325
         Top             =   360
         Width           =   540
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "Time Eangine"
         ForeColor       =   &H00CFE4E0&
         Height          =   255
         Left            =   60
         TabIndex        =   220
         ToolTipText     =   "Version 5.1 Build 564"
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label lblLQSky 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         Caption         =   "Liquid Sky"
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   21.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   735
         Left            =   60
         TabIndex        =   135
         ToolTipText     =   "Powered By Kaveh Abdollahi"
         Top             =   -120
         Width           =   3975
      End
   End
   Begin VB.Frame fraFullScr 
      Appearance      =   0  'Flat
      BackColor       =   &H001B171C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      TabIndex        =   300
      Top             =   11325
      Width           =   15360
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   6
         Left            =   11520
         TabIndex        =   307
         Top             =   15
         Width           =   735
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   5
         Left            =   10560
         TabIndex        =   306
         Top             =   15
         Width           =   735
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   4
         Left            =   9600
         TabIndex        =   305
         Top             =   15
         Width           =   735
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   3
         Left            =   8640
         TabIndex        =   304
         Top             =   15
         Width           =   735
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   2
         Left            =   7200
         TabIndex        =   303
         Top             =   15
         Width           =   1215
      End
      Begin VB.Label lblFullscr 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A5CFD8&
         Height          =   180
         Index           =   1
         Left            =   6120
         TabIndex        =   302
         Top             =   15
         Width           =   855
      End
      Begin VB.Label lblFullscr 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "Liquid Sky 7.7.639"
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00CFE4E0&
         Height          =   180
         Index           =   0
         Left            =   13440
         TabIndex        =   301
         Top             =   15
         Width           =   1575
      End
   End
   Begin VB.Frame fraLogs 
      Appearance      =   0  'Flat
      BackColor       =   &H00473842&
      BorderStyle     =   0  'None
      Caption         =   "( X ) len of data "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4200
      Left            =   3990
      TabIndex        =   25
      Top             =   0
      Visible         =   0   'False
      Width           =   3930
      Begin VB.CommandButton cmdGetLog 
         BackColor       =   &H00C9B6D1&
         Caption         =   "Get Log"
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   345
         Top             =   3360
         Width           =   1575
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H002EDEC8&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   9
         Left            =   255
         Style           =   1  'Graphical
         TabIndex        =   297
         Top             =   15
         Width           =   255
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H002EDEC8&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   296
         Top             =   15
         Width           =   255
      End
      Begin VB.CheckBox chkALog 
         BackColor       =   &H00000000&
         Caption         =   "Auto"
         ForeColor       =   &H00E0E0E0&
         Height          =   315
         Left            =   1560
         MaskColor       =   &H000000FF&
         TabIndex        =   292
         Top             =   3360
         Width           =   705
      End
      Begin VB.CommandButton cmdLogClr 
         BackColor       =   &H000000FF&
         Caption         =   "Clear List"
         Height          =   270
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   281
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdCopy2Excel 
         Caption         =   "Send To Excel"
         Height          =   270
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   280
         Top             =   3720
         Width           =   1215
      End
      Begin VB.ListBox lstLogs 
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         ForeColor       =   &H00CFE4E0&
         Height          =   3030
         Left            =   0
         TabIndex        =   279
         Top             =   330
         Width           =   3975
      End
      Begin VB.CommandButton cmdHideLogs 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hide Logs"
         Height          =   255
         Left            =   3000
         TabIndex        =   28
         Top             =   45
         Width           =   855
      End
      Begin VB.PictureBox picBLogs 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   15
         ScaleHeight     =   3675
         ScaleWidth      =   3885
         TabIndex        =   26
         Top             =   360
         Width           =   3915
      End
      Begin VB.Label lblLogs 
         Alignment       =   2  'Center
         BackColor       =   &H002D061B&
         Caption         =   "Logs"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   15
         TabIndex        =   27
         Top             =   0
         Width           =   3930
      End
   End
   Begin VB.Frame fraControls 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "( X ) len of data "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   9120
      Left            =   11520
      TabIndex        =   29
      Top             =   0
      Visible         =   0   'False
      Width           =   3480
      Begin VB.CommandButton Command2 
         BackColor       =   &H00DDD3E0&
         Caption         =   "WebCam"
         Height          =   375
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   344
         Top             =   480
         Width           =   1305
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0099E1B3&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   360
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   269
         Top             =   5955
         Width           =   1065
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D696A2&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   9.75
            Charset         =   1
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   360
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   268
         Top             =   6390
         Width           =   1065
      End
      Begin VB.TextBox txtFpS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Cordia New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   390
         Left            =   360
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   273
         Text            =   "30"
         Top             =   6120
         Width           =   825
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   10
         Left            =   1095
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   272
         Top             =   6150
         Width           =   330
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0070616C&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   8
         Left            =   360
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   271
         Text            =   "Speed Set on             FPS"
         Top             =   5760
         Width           =   1905
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   9
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   270
         Text            =   "100"
         Top             =   5760
         Width           =   435
      End
      Begin VB.CommandButton cmdNormalSize 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "#"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   120
         MaskColor       =   &H00404040&
         Style           =   1  'Graphical
         TabIndex        =   222
         Top             =   960
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00665766&
         Caption         =   "Draw"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   0
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   130
         Top             =   3960
         Width           =   915
      End
      Begin VB.CheckBox chktest 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00473842&
         Caption         =   "High Light"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   2520
         TabIndex        =   82
         Top             =   8760
         Width           =   210
      End
      Begin VB.CheckBox chkScript 
         BackColor       =   &H00000000&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   270
         Left            =   180
         MousePointer    =   1  'Arrow
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   8460
         Width           =   780
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         DragMode        =   1  'Automatic
         ForeColor       =   &H001B171C&
         Height          =   225
         Index           =   12
         Left            =   720
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   98
         Top             =   4185
         Width           =   480
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         DragMode        =   1  'Automatic
         ForeColor       =   &H001B171C&
         Height          =   225
         Index           =   13
         Left            =   720
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   97
         Top             =   4425
         Width           =   480
      End
      Begin VB.CheckBox chkP4Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   2
         Left            =   3240
         MaskColor       =   &H000000FF&
         TabIndex        =   95
         Top             =   4440
         Width           =   195
      End
      Begin VB.CheckBox chkP4Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   1
         Left            =   2820
         MaskColor       =   &H000000FF&
         TabIndex        =   94
         Top             =   4440
         Width           =   435
      End
      Begin VB.CheckBox chkP4Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   0
         Left            =   2520
         MaskColor       =   &H000000FF&
         TabIndex        =   91
         Top             =   4440
         Width           =   315
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00665766&
         Caption         =   "Clr Draw"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   7
         Left            =   1200
         MaskColor       =   &H000000FF&
         TabIndex        =   89
         Top             =   4440
         Width           =   1275
      End
      Begin VB.CheckBox chkP3Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   2
         Left            =   3240
         MaskColor       =   &H000000FF&
         TabIndex        =   93
         Top             =   4200
         Width           =   195
      End
      Begin VB.CheckBox chkP3Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   1
         Left            =   2820
         MaskColor       =   &H000000FF&
         TabIndex        =   92
         Top             =   4200
         Width           =   435
      End
      Begin VB.CheckBox chkP3Opt 
         BackColor       =   &H00473842&
         CausesValidation=   0   'False
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   0
         Left            =   2520
         MaskColor       =   &H000000FF&
         TabIndex        =   76
         Top             =   4200
         Width           =   315
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00473842&
         Caption         =   "P3"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   1
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   96
         Top             =   4200
         Width           =   1035
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00473842&
         Caption         =   "Clr Draw"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   5
         Left            =   1200
         MaskColor       =   &H000000FF&
         TabIndex        =   80
         Top             =   4200
         Width           =   1275
      End
      Begin VB.CheckBox ChkDraw 
         BackColor       =   &H00665766&
         Caption         =   "P4"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Index           =   6
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   90
         Top             =   4440
         Width           =   1035
      End
      Begin VB.TextBox txtLScr 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         ForeColor       =   &H00FFFFC0&
         Height          =   165
         Left            =   447
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "00.00"
         Top             =   2040
         Width           =   425
      End
      Begin VB.TextBox txtBR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   3090
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   86
         Text            =   "2"
         Top             =   1635
         Width           =   200
      End
      Begin VB.TextBox txtBLR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2257
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   85
         Text            =   "2"
         Top             =   1635
         Width           =   240
      End
      Begin VB.TextBox txtBL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   1470
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   84
         Text            =   "2"
         Top             =   1635
         Width           =   200
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test"
         Height          =   255
         Left            =   2505
         TabIndex        =   83
         Top             =   8760
         Width           =   855
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   7575
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   7380
         Width           =   300
      End
      Begin VB.CheckBox chkClrAlter 
         BackColor       =   &H00800080&
         Caption         =   "Alternate Clear"
         ForeColor       =   &H00FFFFFF&
         Height          =   200
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   58
         Top             =   5160
         Width           =   1395
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2110
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   2513
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   6
         Left            =   480
         MaxLength       =   5
         MousePointer    =   1  'Arrow
         TabIndex        =   40
         Text            =   "1"
         Top             =   2230
         Width           =   675
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   1140
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   2230
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   195
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   2230
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   2520
         MaxLength       =   5
         MousePointer    =   1  'Arrow
         TabIndex        =   52
         Text            =   "15"
         Top             =   3360
         Width           =   555
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FAF1F3&
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   2400
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   39
         Text            =   "1"
         Top             =   2520
         Width           =   675
      End
      Begin VB.CheckBox chkABalance 
         BackColor       =   &H00000000&
         Caption         =   "Check1"
         Height          =   210
         Left            =   2040
         TabIndex        =   36
         Top             =   3352
         Width           =   210
      End
      Begin VB.CheckBox chkAHeight 
         BackColor       =   &H00473842&
         Caption         =   "Auto Balance Height"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Left            =   180
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   3357
         Width           =   2055
      End
      Begin VB.TextBox txtBL2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   1710
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   74
         Text            =   "2"
         Top             =   1485
         Width           =   200
      End
      Begin VB.TextBox txtBLR2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2257
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   73
         Text            =   "2"
         Top             =   1485
         Width           =   240
      End
      Begin VB.TextBox txtBLR3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2257
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   72
         Text            =   "2"
         Top             =   1335
         Width           =   240
      End
      Begin VB.TextBox txtBand 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   1950
         MousePointer    =   1  'Arrow
         TabIndex        =   71
         Text            =   "2"
         Top             =   495
         Width           =   200
      End
      Begin VB.TextBox txtBand3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   1950
         MousePointer    =   1  'Arrow
         TabIndex        =   70
         Text            =   "2"
         Top             =   825
         Width           =   200
      End
      Begin VB.TextBox txtBandAvg2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   69
         Text            =   "2"
         Top             =   660
         Width           =   195
      End
      Begin VB.TextBox txtBandAvg1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   68
         Text            =   "2"
         Top             =   495
         Width           =   195
      End
      Begin VB.TextBox txtBandAvg3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   67
         Text            =   "2"
         Top             =   825
         Width           =   195
      End
      Begin VB.TextBox txtBand2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   168
         Left            =   1950
         MousePointer    =   1  'Arrow
         TabIndex        =   66
         Text            =   "2"
         Top             =   660
         Width           =   200
      End
      Begin VB.TextBox txtBL3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   1950
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   65
         Text            =   "2"
         Top             =   1335
         Width           =   200
      End
      Begin VB.TextBox txtBR3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2595
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   64
         Text            =   "2"
         Top             =   1335
         Width           =   200
      End
      Begin VB.TextBox txtBR2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   150
         Left            =   2835
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   63
         Text            =   "2"
         Top             =   1485
         Width           =   200
      End
      Begin VB.TextBox Text31 
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2790
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   61
         Text            =   "Midl"
         Top             =   660
         Width           =   300
      End
      Begin VB.TextBox Text30 
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2940
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   60
         Text            =   "Bass"
         Top             =   480
         Width           =   360
      End
      Begin VB.TextBox Text29 
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   2595
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   59
         Text            =   "Treble"
         Top             =   840
         Width           =   435
      End
      Begin VB.TextBox txtProcess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H001E1E1E&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0FFFF&
         Height          =   168
         Index           =   9
         Left            =   1395
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   48
         Text            =   "00.00"
         Top             =   390
         Width           =   410
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   56
         Text            =   "2"
         Top             =   7380
         Width           =   795
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00F8AFA9&
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   2280
         MousePointer    =   1  'Arrow
         TabIndex        =   35
         Text            =   "4"
         Top             =   7575
         Width           =   795
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   2400
         MaxLength       =   4
         MousePointer    =   1  'Arrow
         TabIndex        =   38
         Text            =   "256"
         Top             =   2760
         Width           =   675
      End
      Begin VB.TextBox txtProcess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0FFFF&
         Height          =   168
         Index           =   7
         Left            =   2472
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   32
         Text            =   "00.00"
         Top             =   7800
         Width           =   410
      End
      Begin VB.CheckBox chkAdjFreq 
         BackColor       =   &H00473842&
         Caption         =   "Adjustment (X Level)"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   0
         Left            =   180
         TabIndex        =   75
         Top             =   7365
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.CheckBox chkAdjFreq 
         BackColor       =   &H00473842&
         Caption         =   "Adjustment (Z Level)"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   1
         Left            =   180
         TabIndex        =   62
         Top             =   7590
         Value           =   1  'Checked
         Width           =   1875
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   25
         Left            =   195
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   57
         Text            =   "Sec of Last Freq  Match In Screen "
         Top             =   2025
         Width           =   3150
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   3360
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00808080&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   2280
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3360
         Width           =   300
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   24
         Left            =   180
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   50
         Text            =   "Base Heigth of Scope  "
         Top             =   2745
         Width           =   1920
      End
      Begin VB.CheckBox chkInc 
         BackColor       =   &H00665766&
         Caption         =   "Increase Scope Height"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Left            =   180
         MaskColor       =   &H000000FF&
         TabIndex        =   49
         Top             =   2985
         Value           =   1  'Checked
         Width           =   3195
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   2
         Left            =   180
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   47
         Text            =   "Scopes Count  "
         Top             =   2505
         Width           =   1920
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   2760
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   2115
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   2760
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   2513
         Width           =   300
      End
      Begin VB.CheckBox chkTransparent 
         BackColor       =   &H00473842&
         Caption         =   "Transparent All Panels"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   1560
         TabIndex        =   37
         Top             =   5160
         Width           =   1935
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   7575
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   3060
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   7380
         Width           =   300
      End
      Begin VB.PictureBox PicFFT 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H001E1E1E&
         DrawWidth       =   2
         FillColor       =   &H00FF0000&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00665766&
         Height          =   1470
         Left            =   1350
         Negotiate       =   -1  'True
         ScaleHeight     =   117.517
         ScaleMode       =   0  'User
         ScaleWidth      =   40.306
         TabIndex        =   77
         Top             =   390
         Width           =   2010
      End
      Begin VB.PictureBox picBCtrl 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         ForeColor       =   &H80000008&
         Height          =   8760
         Left            =   0
         ScaleHeight     =   8730
         ScaleWidth      =   3450
         TabIndex        =   78
         Top             =   360
         Width           =   3475
         Begin VB.CheckBox chkATime 
            BackColor       =   &H00800000&
            Caption         =   "LQ Time 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   291
            Top             =   6240
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.CommandButton cmdLarger 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "+1"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   290
            Top             =   6240
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdSmaler 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "-1"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   15
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   289
            Top             =   6240
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.CommandButton cmdSmaler 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   288
            Top             =   6240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton cmdLarger 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "100"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   13
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   287
            Top             =   6240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton cmdLarger 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "1000"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   286
            Top             =   6240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.CommandButton cmdSmaler 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "1000"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   8.25
               Charset         =   1
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   14
            Left            =   1800
            MaskColor       =   &H00FFFFFF&
            Style           =   1  'Graphical
            TabIndex        =   285
            Top             =   6240
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.TextBox txtspm 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00F8AFA9&
            ForeColor       =   &H00000000&
            Height          =   225
            Index           =   13
            Left            =   1800
            MaxLength       =   6
            MousePointer    =   1  'Arrow
            TabIndex        =   284
            Text            =   "1"
            Top             =   6000
            Visible         =   0   'False
            Width           =   1395
         End
         Begin VB.TextBox txtLQT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H0070616C&
            DragMode        =   1  'Automatic
            ForeColor       =   &H00E0E0E0&
            Height          =   705
            Left            =   1800
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   283
            Text            =   "frmBase.frx":0A45
            Top             =   6240
            Visible         =   0   'False
            Width           =   825
         End
         Begin VB.CheckBox chkPR 
            BackColor       =   &H00004040&
            Caption         =   "View All Available Threads"
            ForeColor       =   &H00FFFFFF&
            Height          =   345
            Index           =   2
            Left            =   1800
            MaskColor       =   &H000000FF&
            TabIndex        =   282
            Top             =   6240
            Visible         =   0   'False
            Width           =   1635
         End
      End
      Begin VB.Label lblControls 
         Alignment       =   2  'Center
         BackColor       =   &H00665766&
         Caption         =   "   Controls"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   0
         TabIndex        =   79
         Top             =   30
         Width           =   3450
      End
      Begin VB.Line Line34 
         BorderColor     =   &H00808080&
         X1              =   1200
         X2              =   3180
         Y1              =   2340
         Y2              =   2340
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   2
         X1              =   15
         X2              =   3465
         Y1              =   2640
         Y2              =   2640
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   4
         X1              =   15
         X2              =   3470
         Y1              =   2760
         Y2              =   2760
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   5
         X1              =   360
         X2              =   3000
         Y1              =   5520
         Y2              =   5520
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   7
         X1              =   360
         X2              =   3000
         Y1              =   4200
         Y2              =   4200
      End
   End
   Begin VB.ComboBox DevicesBox 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmBase.frx":0A55
      Left            =   8160
      List            =   "frmBase.frx":0A5C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   11160
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Frame fraAbout 
      Appearance      =   0  'Flat
      BackColor       =   &H001B171C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   7920
      TabIndex        =   216
      Top             =   3713
      Visible         =   0   'False
      Width           =   3975
      Begin VB.CommandButton cmdCloseAbout 
         Appearance      =   0  'Flat
         BackColor       =   &H00F1EDED&
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   305
         Left            =   3600
         MaskColor       =   &H000040C0&
         Style           =   1  'Graphical
         TabIndex        =   217
         TabStop         =   0   'False
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   330
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyrighjt© 2010  Kaveh Abdollahi"
         ForeColor       =   &H00CFE4E0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   219
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         Caption         =   "                Liquid Sky"
         BeginProperty Font 
            Name            =   "AngsanaUPC"
            Size            =   21.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   855
         Left            =   0
         TabIndex        =   218
         ToolTipText     =   "Powerd By Kaveh Abdollahi"
         Top             =   -120
         Width           =   3975
      End
   End
   Begin VB.Frame fraTelo 
      Appearance      =   0  'Flat
      BackColor       =   &H002D061B&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00512D4B&
      Height          =   5130
      Left            =   3990
      TabIndex        =   226
      Top             =   0
      Visible         =   0   'False
      Width           =   3930
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "®"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   36
         Left            =   15
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   322
         Top             =   4320
         Width           =   330
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   34
         Left            =   1245
         MousePointer    =   1  'Arrow
         TabIndex        =   315
         Text            =   "512"
         Top             =   4320
         Width           =   570
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H0082ADD5&
         Caption         =   "è"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   34
         Left            =   1815
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   317
         Top             =   4320
         Width           =   330
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H0082ADD5&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   34
         Left            =   915
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   316
         Top             =   4320
         Width           =   330
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   8.25
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   36
         Left            =   15
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   323
         Top             =   4830
         Width           =   330
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H002D061B&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   36
         Left            =   15
         MousePointer    =   1  'Arrow
         TabIndex        =   321
         Text            =   "64"
         Top             =   4590
         Width           =   330
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "ê"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   35
         Left            =   345
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   320
         Top             =   4830
         Width           =   570
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00DDD3E0&
         Caption         =   "é"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   12
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   35
         Left            =   345
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   319
         Top             =   4320
         Width           =   570
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   35
         Left            =   345
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   318
         Text            =   "frmBase.frx":0A6C
         Top             =   4590
         Width           =   570
      End
      Begin VB.PictureBox picTele 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   3840
         Left            =   75
         ScaleHeight     =   3900.952
         ScaleMode       =   0  'User
         ScaleWidth      =   3900.952
         TabIndex        =   265
         Top             =   420
         Width           =   3840
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H002EDEC8&
         Caption         =   "ç"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   15
         Style           =   1  'Graphical
         TabIndex        =   264
         Top             =   15
         Width           =   255
      End
      Begin VB.CommandButton cmdNav 
         Appearance      =   0  'Flat
         BackColor       =   &H002EDEC8&
         Caption         =   "è"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         HelpContextID   =   3
         Index           =   6
         Left            =   270
         Style           =   1  'Graphical
         TabIndex        =   263
         Top             =   15
         Width           =   255
      End
      Begin VB.CheckBox chkBox 
         BackColor       =   &H00784138&
         Caption         =   "°"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   27.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   251
         Top             =   4320
         Width           =   510
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0099E1B3&
         Caption         =   "Camera 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   15
         TabIndex        =   227
         ToolTipText     =   "Powerd By Kaveh Abdollahi"
         Top             =   15
         Width           =   3930
      End
   End
   Begin VB.Frame fraBlur 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   5295
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3495
      Begin VB.PictureBox picBBlur 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         ForeColor       =   &H80000008&
         Height          =   4260
         Left            =   0
         ScaleHeight     =   4230
         ScaleWidth      =   3465
         TabIndex        =   14
         Top             =   480
         Width           =   3500
      End
      Begin VB.CheckBox chkBlur 
         BackColor       =   &H0099E1B3&
         Caption         =   "Motion"
         ForeColor       =   &H001E1E1E&
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   207
         Top             =   1920
         Width           =   825
      End
      Begin VB.CheckBox chkBlur 
         BackColor       =   &H00000000&
         Caption         =   "Blur 1"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   0
         Left            =   2640
         TabIndex        =   178
         Top             =   1680
         Width           =   795
      End
      Begin VB.CheckBox chkBlur 
         BackColor       =   &H00000000&
         Caption         =   "Type 4"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   3
         Left            =   2640
         TabIndex        =   129
         Top             =   1440
         Width           =   795
      End
      Begin VB.CheckBox chkBlur 
         BackColor       =   &H00000000&
         Caption         =   "Type 3"
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   2
         Left            =   2640
         TabIndex        =   128
         Top             =   1200
         Width           =   795
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   27
         Left            =   2130
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   122
         Text            =   "Cpu 0"
         Top             =   3720
         Width           =   720
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Index           =   26
         Left            =   2835
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   121
         Text            =   "Cpu 1"
         Top             =   3720
         Width           =   720
      End
      Begin VB.TextBox txtProcess1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2835
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   117
         Text            =   "00"
         Top             =   3915
         Width           =   720
      End
      Begin VB.TextBox txtProcess0 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Left            =   2130
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   116
         Text            =   "00"
         Top             =   3915
         Width           =   720
      End
      Begin VB.TextBox txtProcessSum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   2160
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   123
         Text            =   "00"
         Top             =   3495
         Width           =   1425
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   0
         Left            =   2160
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   119
         Text            =   "WIN Cpu Usage"
         Top             =   3270
         Width           =   1425
      End
      Begin VB.ListBox lstPsent 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   2880
         ItemData        =   "frmBase.frx":0A70
         Left            =   2055
         List            =   "frmBase.frx":0A77
         TabIndex        =   107
         Top             =   360
         Width           =   550
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   150
         Index           =   17
         Left            =   1695
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   114
         Text            =   "ms"
         Top             =   3315
         Width           =   250
      End
      Begin VB.TextBox txtEFRM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005B425A&
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1170
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   112
         Text            =   "10.00"
         Top             =   3240
         Width           =   540
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   150
         Index           =   9
         Left            =   1695
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   111
         Text            =   "ms"
         Top             =   3525
         Width           =   250
      End
      Begin VB.TextBox txtDoEvSleep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005B425A&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1170
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   110
         Text            =   "10.00"
         Top             =   3465
         Width           =   540
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00322732&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   18
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   115
         Text            =   "                Sum : "
         Top             =   3240
         Width           =   2370
      End
      Begin VB.TextBox TextLabel 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   6
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   118
         Text            =   "Others"
         Top             =   3465
         Width           =   2490
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   4
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   120
         Text            =   "Timers"
         Top             =   3720
         Width           =   2475
      End
      Begin VB.CheckBox chkSortP 
         BackColor       =   &H00665766&
         Caption         =   "Sort Lists"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   0
         TabIndex        =   113
         Top             =   4320
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.ListBox lstFunctions 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   2880
         ItemData        =   "frmBase.frx":0A85
         Left            =   120
         List            =   "frmBase.frx":0A8C
         TabIndex        =   109
         Top             =   360
         Width           =   1425
      End
      Begin VB.ListBox lstProcess 
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   2880
         ItemData        =   "frmBase.frx":0A9E
         Left            =   1530
         List            =   "frmBase.frx":0AA5
         TabIndex        =   108
         Top             =   360
         Width           =   540
      End
      Begin VB.TextBox txtProcess 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400040&
         BorderStyle     =   0  'None
         ForeColor       =   &H00C0FFFF&
         Height          =   168
         Index           =   0
         Left            =   3000
         Locked          =   -1  'True
         MousePointer    =   14  'Arrow and Question
         TabIndex        =   15
         Text            =   "00.00"
         Top             =   360
         Width           =   525
      End
      Begin VB.TextBox txtTimeP2Sky 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H005B425A&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2640
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   103
         Text            =   "10.00"
         Top             =   4080
         Width           =   540
      End
      Begin VB.Timer Timer_Sky 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   1920
         Top             =   4200
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   12
         Left            =   2685
         MousePointer    =   1  'Arrow
         TabIndex        =   102
         Text            =   "1"
         Top             =   4320
         Width           =   435
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5B797&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   2400
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   101
         Top             =   4320
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00866CBB&
         Caption         =   "+"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   8.25
            Charset         =   1
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   3105
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   4320
         Width           =   300
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   6
         X1              =   120
         X2              =   2760
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Label lblBlur 
         Alignment       =   2  'Center
         BackColor       =   &H002D061B&
         Caption         =   "Blur"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   15
         TabIndex        =   11
         Top             =   30
         Width           =   3465
      End
      Begin VB.Line Line18 
         BorderColor     =   &H00808080&
         X1              =   0
         X2              =   2400
         Y1              =   6725
         Y2              =   6725
      End
      Begin VB.Label lblFileName 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "------------"
         DragMode        =   1  'Automatic
         BeginProperty Font 
            Name            =   "MS Dialog Light"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   165
         Left            =   240
         MousePointer    =   1  'Arrow
         TabIndex        =   10
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame fraColors 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   8180
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   3350
      Begin VB.CheckBox chkCM 
         BackColor       =   &H00512D4B&
         Caption         =   " C = M"
         Enabled         =   0   'False
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Left            =   720
         TabIndex        =   359
         Top             =   1440
         Width           =   1140
      End
      Begin VB.CheckBox chkLock 
         BackColor       =   &H00512D4B&
         Caption         =   " M = E"
         Enabled         =   0   'False
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Left            =   720
         MaskColor       =   &H000000FF&
         TabIndex        =   358
         Top             =   1200
         Width           =   1140
      End
      Begin VB.CheckBox chkAGrow 
         BackColor       =   &H00512D4B&
         Caption         =   "Grow"
         Enabled         =   0   'False
         ForeColor       =   &H00F1E4F3&
         Height          =   225
         Left            =   720
         TabIndex        =   357
         Top             =   1680
         Width           =   1140
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   17
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   356
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   355
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   16
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   354
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   19
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   353
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   352
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdSmaler 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   18
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   351
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   17
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   350
         Top             =   1080
         Width           =   300
      End
      Begin VB.CommandButton cmdLarger 
         Appearance      =   0  'Flat
         BackColor       =   &H00C7B6AB&
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   19
         Left            =   2040
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   349
         Top             =   1080
         Width           =   300
      End
      Begin VB.TextBox txtspm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00B9D3B8&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         HideSelection   =   0   'False
         Index           =   10
         Left            =   1080
         MousePointer    =   1  'Arrow
         TabIndex        =   314
         Text            =   "1"
         Top             =   1920
         Width           =   495
      End
      Begin VB.CheckBox chkInverse 
         BackColor       =   &H0000FFFF&
         Caption         =   "Coloring"
         ForeColor       =   &H001E1E1E&
         Height          =   225
         Left            =   1680
         MaskColor       =   &H000000FF&
         TabIndex        =   199
         Top             =   3720
         Width           =   1005
      End
      Begin VB.CheckBox chkFallCol 
         BackColor       =   &H00000000&
         Caption         =   "Fall Colors"
         ForeColor       =   &H00E0E0E0&
         Height          =   200
         Left            =   240
         MaskColor       =   &H000000FF&
         TabIndex        =   104
         Top             =   3780
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.CommandButton cmdSnCGr 
         BackColor       =   &H00FF8080&
         Caption         =   "B"
         Height          =   250
         Index           =   2
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3195
         Width           =   680
      End
      Begin VB.CommandButton cmdSnCGr 
         BackColor       =   &H0080FF80&
         Caption         =   "G"
         Height          =   250
         Index           =   1
         Left            =   1275
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3195
         Width           =   680
      End
      Begin VB.CommandButton cmdSnCGr 
         BackColor       =   &H000000FF&
         Caption         =   "R"
         Height          =   250
         Index           =   0
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   3195
         Width           =   680
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   0
         Left            =   600
         MousePointer    =   1  'Arrow
         TabIndex        =   21
         Text            =   "2"
         Top             =   3480
         Width           =   675
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   2
         Left            =   1920
         MousePointer    =   1  'Arrow
         TabIndex        =   23
         Text            =   "2"
         Top             =   3480
         Width           =   675
      End
      Begin VB.TextBox txtRGB 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FF0000&
         DragMode        =   1  'Automatic
         ForeColor       =   &H00C0FFFF&
         Height          =   195
         Index           =   1
         Left            =   1260
         MousePointer    =   1  'Arrow
         TabIndex        =   22
         Text            =   "2"
         Top             =   3480
         Width           =   675
      End
      Begin VB.CommandButton cmdRGB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "BRG"
         Height          =   250
         Index           =   2
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   450
         Width           =   495
      End
      Begin VB.CommandButton cmdRGB 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Set RGB With 0"
         Height          =   250
         Index           =   10
         Left            =   600
         TabIndex        =   16
         Top             =   450
         Width           =   1335
      End
      Begin VB.TextBox txtMinC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "0"
         Top             =   2880
         Width           =   680
      End
      Begin VB.TextBox txtMinC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   915
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "0"
         Top             =   2880
         Width           =   680
      End
      Begin VB.TextBox txtMinC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "0"
         Top             =   2880
         Width           =   680
      End
      Begin VB.TextBox txtMaxC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   2
         Left            =   1590
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "255"
         Top             =   2685
         Width           =   680
      End
      Begin VB.TextBox txtMaxC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00008000&
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   1
         Left            =   915
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "255"
         Top             =   2685
         Width           =   680
      End
      Begin VB.TextBox txtMaxC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   240
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "255"
         Top             =   2685
         Width           =   680
      End
      Begin VB.TextBox TextLabel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00665766&
         BorderStyle     =   0  'None
         DragMode        =   1  'Automatic
         ForeColor       =   &H00E0E0E0&
         Height          =   195
         Index           =   20
         Left            =   235
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         TabIndex        =   12
         Text            =   "RGB Limiter"
         Top             =   2490
         Width           =   2040
      End
      Begin VB.PictureBox picBCol 
         Appearance      =   0  'Flat
         BackColor       =   &H00473842&
         ForeColor       =   &H80000008&
         Height          =   3730
         Left            =   0
         ScaleHeight     =   3705
         ScaleWidth      =   3315
         TabIndex        =   13
         Top             =   360
         Width           =   3345
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00C0C0C0&
            Height          =   250
            Index           =   4
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   164
            Top             =   1680
            Width           =   680
         End
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00473842&
            Height          =   250
            Index           =   3
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   163
            Top             =   1410
            Width           =   680
         End
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00665766&
            Height          =   250
            Index           =   2
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   162
            Top             =   1140
            Width           =   680
         End
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00FFFFFF&
            Height          =   250
            Index           =   1
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   161
            Top             =   870
            Width           =   680
         End
         Begin VB.CommandButton cmdBackCol 
            BackColor       =   &H00000000&
            Height          =   250
            Index           =   0
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   160
            Top             =   600
            Width           =   680
         End
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   3330
         X2              =   3330
         Y1              =   -4545
         Y2              =   -475
      End
      Begin VB.Label lblColorSet 
         Alignment       =   2  'Center
         BackColor       =   &H00322732&
         Caption         =   "Colors"
         BeginProperty Font 
            Name            =   "Modern"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   2
         Top             =   15
         Width           =   3330
      End
      Begin VB.Line Line46 
         BorderColor     =   &H00C0C0C0&
         Index           =   3
         X1              =   338
         X2              =   2978
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   4005
      Locked          =   -1  'True
      TabIndex        =   159
      Text            =   "ETC"
      Top             =   11160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arabic Transparent"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   195
      Left            =   4005
      Locked          =   -1  'True
      TabIndex        =   158
      Text            =   "Blur"
      Top             =   11160
      Visible         =   0   'False
      Width           =   450
   End
   Begin VB.TextBox txtDraw 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Dialog"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Left            =   4215
      Locked          =   -1  'True
      TabIndex        =   157
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Dialog"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   8
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   156
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   6
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   155
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   5
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   154
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   4
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   153
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   3
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   152
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   1
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   151
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Index           =   2
      Left            =   4035
      Locked          =   -1  'True
      TabIndex        =   150
      Text            =   "00.00"
      Top             =   11175
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtEtcSumD 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3990
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   149
      Top             =   11160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtWaveSumD 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   3990
      Locked          =   -1  'True
      MaxLength       =   4
      TabIndex        =   148
      Top             =   11160
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox txtETCT 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Dialog"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   168
      Left            =   4680
      Locked          =   -1  'True
      TabIndex        =   147
      Text            =   "00.00"
      Top             =   11280
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   10
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   146
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   11
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   145
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   12
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   144
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   13
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   143
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   14
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   142
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   15
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   141
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   16
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   140
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   17
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   139
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   18
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   138
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   19
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   137
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin VB.TextBox txtProcess 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00400040&
      BorderStyle     =   0  'None
      ForeColor       =   &H00C0FFFF&
      Height          =   168
      Index           =   20
      Left            =   4755
      Locked          =   -1  'True
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   136
      Text            =   "00.00"
      ToolTipText     =   "FFT Calculate Timer "
      Top             =   11280
      Visible         =   0   'False
      Width           =   410
   End
   Begin MSScriptControlCtl.ScriptControl KScript 
      Left            =   6000
      Top             =   10920
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Timer Timer_AutoSave 
      Interval        =   100
      Left            =   6960
      Top             =   11040
   End
   Begin VB.Timer Timer_AutoLog 
      Interval        =   200
      Left            =   6600
      Top             =   11040
   End
   Begin VB.Timer Timer_Process 
      Interval        =   100
      Left            =   7680
      Top             =   11040
   End
   Begin VB.PictureBox picViewEE 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1.72800e5
      Left            =   0
      ScaleHeight     =   11520
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   87
      Top             =   0
      Width           =   2.30400e5
      Begin VB.Timer Timer_Seconds 
         Interval        =   1000
         Left            =   7320
         Top             =   11040
      End
   End
   Begin VB.PictureBox picBuffEE 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   11520
      ScaleMode       =   0  'User
      ScaleWidth      =   15360
      TabIndex        =   88
      Top             =   0
      Visible         =   0   'False
      Width           =   15360
   End
   Begin VB.PictureBox picBuffEE2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      CausesValidation=   0   'False
      ClipControls    =   0   'False
      DrawStyle       =   2  'Dot
      Enabled         =   0   'False
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H00FFFFFF&
      Height          =   11520
      Left            =   0
      ScaleHeight     =   11520
      ScaleMode       =   0  'User
      ScaleWidth      =   11520
      TabIndex        =   81
      Top             =   0
      Visible         =   0   'False
      Width           =   15360
      Begin VB.Timer Timer_AHeight 
         Interval        =   5
         Left            =   7320
         Top             =   11040
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00A5CFD8&
         Height          =   855
         Left            =   0
         Top             =   0
         Width           =   975
      End
      Begin VB.Line Line3 
         X1              =   90
         X2              =   1260
         Y1              =   1680
         Y2              =   1680
      End
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      X1              =   -120
      X2              =   -120
      Y1              =   -2640
      Y2              =   -840
   End
End
Attribute VB_Name = "frmBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Wave As WaveHdr, te As Long, te2 As Long


Private Sub chkAutoMax_Click()
    If chkAutoMax Then
        cmdMaxPoints.Enabled = False
        txtspm(21).Enabled = False
        cmdLarger(21).Enabled = False
        cmdSmaler(21).Enabled = False
        txtRST(7).Enabled = False
        cmd0(7).Enabled = False
      Else
        cmdMaxPoints.Enabled = True
        txtspm(21).Enabled = True
        cmdLarger(21).Enabled = True
        cmdSmaler(21).Enabled = True
        txtRST(7).Enabled = True
        cmd0(7).Enabled = True
     End If
End Sub

Private Sub chkAutoShot_Click()
    If chkAutoShot Then
        Timer_AutoSave.Enabled = True
    Else
        Timer_AutoSave.Enabled = False
    End If
End Sub


Private Sub chkAvalue_Click()
    If chkAvalue Then
        chkAvalue.Caption = "Click To Manual Time"
    Else
        chkAvalue.Caption = "Click To Automatic Time"
    End If
End Sub

Private Sub chkClrAlter_Click()
    minY = 0: maxY = 768
    kaAltrCls
End Sub

Private Sub chkCM_Click()
    If chkCM Then txtspm(18) = txtspm(17)
End Sub

Private Sub chkCol_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim A As Byte
    If chkCol(Index) = 1 Then
        For A = 0 To chkCol.count - 1
           chkCol(A) = 0
        Next A
     chkCol(Index) = 1
    End If
End Sub

Private Sub ChkDraw_Click(Index As Integer)
Dim X As Integer
    If Index = 0 Or Index = 1 Or Index = 6 And ChkDraw(Index) = 0 Then ' cmdCls_Click
         picTmp.ForeColor = vbBlack
        For X = 1 To 20
            Polyline picTmp.hdc, PtL(0, X), 512
        Next X
    End If
End Sub

Private Sub chkDrawCntr_Click(Index As Integer)
    
    Select Case Index
        
        Case 0:
        
        Case 1:
        
        Case 3:
        
        Case 4:
    
    End Select

End Sub

Private Sub chkInc_Click()
    If chkInc Then
      chkInc.ForeColor = &HFFFF&
      chkInc.BackColor = &HFF&
     Else
      chkInc.ForeColor = &HE0E0E0
      chkInc.BackColor = &H0&
    End If
End Sub
'

Private Sub chkLock_Click()
    If chkLock Then txtspm(16) = 1 / txtspm(17)
End Sub

Private Sub chkP3Opt_Click(Index As Integer)
If Index = 2 Then
      chkP3Opt(0) = chkP3Opt(0) * -1 + 1
      chkP3Opt(1) = chkP3Opt(1) * -1 + 1
End If
End Sub

Private Sub chkP4Opt_Click(Index As Integer)
If Index = 2 Then
      chkP4Opt(0) = chkP4Opt(0) * -1 + 1
      chkP4Opt(1) = chkP4Opt(1) * -1 + 1
End If
End Sub

Private Sub chkTimeEnable_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 4 Then chkTimeEnable(0).Value = 0: chkTimeEnable(2).Value = 0
    If Index = 2 Then chkTimeEnable(0).Value = 0: chkTimeEnable(4).Value = 0
    If Index = 0 Then chkTimeEnable(2).Value = 0: chkTimeEnable(4).Value = 0
'    DoEvents

End Sub

Private Sub chkTransparent_Click()
 
 If chkTransparent Then Exit Sub
    
      picBCtrl.Cls
      picBCol.Cls
      picBBlur.Cls
      picBProcs.Cls
      picBLogs.Cls
 
End Sub


Private Sub cmd0_Click(Index As Integer)
    
    If Index = 0 Then LQT2 = txtRST(0): txtspm(11) = txtRST(0)
    If Index = 1 Then txtspm(2) = txtRST(1)
    If Index = 2 Then txtspm(16) = 571 ' txtRST(2)
    If Index = 3 Then txtspm(17) = 491 ' txtRST(3)
    If Index = 4 Then txtspm(18) = 1 ' txtRST(4)
    If Index = 5 Then txtspm(19) = 0.1 ' txtRST(5)
    If Index = 6 Then txtspm(20) = txtRST(6)
    If Index = 7 Then txtspm(21) = txtRST(7)
    
    If Index = 10 Then
        KaCls
        txtspm(13) = 0: LQT = 0
        txtspm(11) = txtRST(0): LQT2 = txtRST(0)
    End If

End Sub

Private Sub cmdAbout_Click()
   If fraAbout.Visible = False Then
       fraAbout.Visible = True
    Else
       fraAbout.Visible = False
   End If
End Sub

Private Sub cmdBackCol_Click(Index As Integer)
    picTmp.BackColor = cmdBackCol(Index).BackColor
End Sub

Private Sub cmdChandSaveFolder_Click()
Dim S As String
    If sPath <> "" Then
        S = BrowseForFolder(bPath, Me.hWnd, "Select Folder For Save Image")  'App.Path & "\"
    Else
        S = BrowseForFolder("c:\", Me.hWnd, "Select Folder For Save Image") 'App.Path & "\"
    End If
    If S <> "" Then sPath = S: bPath = S
txtPath = sPath
End Sub

Private Sub cmdCloseAbout_Click()
    fraAbout.Visible = False
End Sub

Private Sub cmdCls_Click()
    KaCls
End Sub

Private Sub cmdCtrl_Click()
    If fraControls.Visible = False Then
        fraControls.Visible = True
    Else
        fraControls.Visible = False
    End If
End Sub

Private Sub CmdDefault_Click()
    
    chkAvalue.Value = 1 ' Auto Time
    LQT2 = 0            ' Time
    txtspm(2) = 1       ' V
    txtspm(20) = 1      ' Aggr
    txtspm(21) = 3000   ' Max Points
    txtspm(28) = 0      ' Start Point
    txtspm(32) = 0      ' Last Points
    txtspm(16) = 1      ' E
    txtspm(17) = 1      ' M
    txtspm(18) = 1      ' C
    txtspm(19) = 0.1      ' Cc
    txtspm(22) = 640
    txtspm(23) = 384
    txtspm(25) = 1
    txtspm(33) = 1
    txtspm(7) = 255
    chkTimeEnable(0).Value = 0:    chkTimeEnable(1).Value = 0:    chkTimeEnable(2).Value = 0
    chkTimeEnable(3).Value = 0:    chkTimeEnable(4).Value = 1:    chkTimeEnable(5).Value = 1
    chkCol(1).Value = 0: chkCol(2).Value = 1: chkCol(3).Value = 0: chkCol(4).Value = 0
    chkAutoMax.Value = 0
    ChkDraw(4).Value = 1
    chkAutoFix.Value = 0
    chkAlpha.Value = 0
    chkAlphaEnable.Value = 0
    Combo1.ListIndex = 14   'Marge Pen



    DoEvents
End Sub

Private Sub cmdDriectory_Click()
On Error Resume Next
    MkDir bPath & "\" & Trim(txtShotCount.Text)
    sPath = bPath & "\" & Trim(txtShotCount.Text)
    txtPath = sPath
End Sub

Private Sub cmdExit_Click()
    Call Form_Unload(0)
End Sub

Private Sub cmdGetLog_Click()
Dim X As Long, Y
'    lstLogs.Visible = False
'    For x = 1 To 30000
'        lstLogs.AddItem x & "," & Primes(x) & "," & PrK(2, x) & "," & PrK(3, x) ' & ","
'    Next x
'    lstLogs.Visible = True
    Loger
End Sub

Private Sub cmdHideLogs_Click()
    cmdLogs_Click
End Sub

Private Sub cmdBlureOpen_Click()
Dim CommonDialog1 As OSDialog
Set CommonDialog1 = New OSDialog

' Examples:-
  Dim title$, Filt$, InDir$, FileSpec$, CurrPath$
  Dim FIndex As Long

'  LOAD egs
   title$ = "Blur Files"
   Filt$ = "Blur Files (*.BLR)|*.BLR|All files (*.*)|*.*" '"Pics bmp,jpg,gif,ico,cur,wmf,emf|*.bmp;*.jpg;*.gif;*.ico;*.cur;*.wmf;*.emf"
'   Filt$ = "Open vbp (*.vbp)|*.vbp|All files (*.*)|*.*"
   FileSpec$ = ""
   InDir$ = CurrPath$ 'Pathspec$
'   Set CommonDialog1 = New OSDialog

   CommonDialog1.ShowOpen FileSpec$, title$, Filt$, InDir$, "", Me.hWnd, FIndex
'   FIndex = 1 bmp
'   FIndex = 2 jpg
'   etc

   Set CommonDialog1 = Nothing

'  SAVE eg
'   Title$ = "Save Mask as 2-color bmp"
'   Filt$ = "Save bmp|*.bmp"
'   InDir$ = CurrPath$ 'Pathspec$
'   FileSpec$=""
'   Set CommonDialog1 = New OSDialog
'   CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hwnd
'   Set CommonDialog1 = Nothing
'
'   Len(FileSpec$)=0 for cancel
'''''''''''''''''''''''''''''''''''''''''
  Dim intf, X As Integer
  Dim S As String
'
'    CommonDialog1.Filter = "Blur Files (*.BLR)|*.BLR"
'    CommonDialog1.CancelError = True
'    On Error GoTo ErrHandler
'    CommonDialog1.ShowOpen
'
'    If CommonDialog1.FileName <> "" Then
'        MousePointer = 11
'        '        LoadNewFile (CommonDialog1.FileName)
'        intf = FreeFile
'        If Trim(FileSpec$) = "" Then Exit Sub
'        Open FileSpec$ For Input As #intf
'        lstFa.Clear
'        lstFaName.Clear
'        While Not EOF(intf)
'            Input #intf, S
'            lstFa.AddItem Trim(S)
'            Input #intf, S
'            lstFaName.AddItem Trim(S)
'        Wend
'        Close #intf
'        S = ""
'        For x = Len(CommonDialog1.FileName) To 1 Step -1
'            If Mid(CommonDialog1.FileName, x, 1) = "\" Then Exit For
'            DoEvents
'        Next x
'        lblFileName.Caption = Mid(CommonDialog1.FileName, x + 1, Len(CommonDialog1.FileName) - x - 4)
'        picBuff.Cls
'        MousePointer = 0
'        Exit Sub '---> Bottom
'    End If
'ErrHandler:
'
'Exit Sub

End Sub

Public Sub cmdInvertPage_Click()
    KaInvert 0, 0, 1024, 768
    reAl = True
End Sub

Public Sub cmdLarger_Click(Index As Integer)
On Error Resume Next
Dim ct As Integer
    If Index = 16 Then txtspm(16) = txtspm(16) * txtspm(16): txtspm(16).Refresh
    If Index = 17 Then txtspm(17) = txtspm(17) * txtspm(17): txtspm(17).Refresh
    If Index = 18 Then txtspm(18) = txtspm(18) * txtspm(18): txtspm(18).Refresh
    If Index = 19 Then txtspm(19) = txtspm(19) * txtspm(19): txtspm(19).Refresh
    
    If Index = 32 Then txtspm(32) = txtspm(32) + 1
    If Index = 33 Then txtspm(33) = txtspm(33) + 0.1
    
    If Index = 34 Then txtspm(34) = txtspm(34) + 8
    If Index = 35 Then txtspm(35) = txtspm(35) + 8
    If Index = 36 Then txtspm(36) = txtspm(36) + 4
    
    If Index = 31 Then txtspm(16) = txtspm(16) * 2
    If Index = 30 Then txtspm(17) = txtspm(17) * 2
    If Index = 29 Then txtspm(18) = txtspm(18) * 2
    If Index = 27 Then txtspm(19) = txtspm(19) * 2
    
    If Index = 28 Then txtspm(28) = txtspm(28) + 1
    
    If Index = 20 Then txtspm(20) = txtspm(20) + 0.1: txtspm(20).Refresh
    If Index = 21 Then txtspm(21) = txtspm(21) + 1: txtspm(21).Refresh
    If Index = 22 Then txtspm(22) = txtspm(22) + 16: txtspm(22).Refresh
    If Index = 23 Then txtspm(23) = txtspm(23) + 16: txtspm(23).Refresh
    If Index = 25 Then txtspm(25) = txtspm(25) + 0.1: txtspm(25).Refresh
    
    If Index = 0 And Val(txtspm(0).Text) < 768 Then txtspm(0).Text = Val(txtspm(0).Text) + 32
    If Index = 1 And txtspm(1) < 10 Then txtspm(1).Text = Val(txtspm(1).Text) + 1
    If Index = 2 Then txtspm(2) = txtspm(2) + 0.0001
    If Index = 24 Then txtspm(2) = txtspm(2) + 0.01
    If Index = 26 Then txtspm(2) = txtspm(2) + 0.1
    If Index = 3 Then txtspm(3) = txtspm(3) + 1
    If Index = 4 Then txtspm(4) = txtspm(4) + 1
    If Index = 5 And Val(txtspm(5).Text) < 50 Then txtspm(5).Text = Val(txtspm(5).Text) + 0.15 '* ((11 - txtspm(5)) \ 10 + 1)
    If Index = 6 And txtspm(6) < 4 Then txtspm(6) = txtspm(6) + 0.1:    txtspm(6).Refresh
    If Index = 7 And txtspm(7) < 255 Then txtspm(7).Text = Val(txtspm(7).Text) + 1
    If Index = 8 And txtspm(8) < 100 Then txtspm(8).Text = Val(txtspm(8).Text) + 0.1
    If Index = 9 Then txtspm(9) = txtspm(9) + 1
    If Index = 10 And txtQua < 100 Then txtQua = txtQua + 5
    If Index = 11 Then LQT2 = LQT2 + txtspm(2): txtspm(11) = LQT2
    
    If Index = 12 And txtspm(12) < 16 Then txtspm(12) = txtspm(12) + 0.1
    If Index = 15 And LQT < 148931 Then LQT = LQT + 1:  txtspm(13) = LQT: txtspm(13).Refresh
    If Index = 14 And LQT < 148931 - 1000 Then LQT = LQT + 1000: txtspm(13) = LQT: txtspm(13).Refresh
    If Index = 13 And LQT < 148931 - 100 Then
       LQT = LQT + 100
       txtspm(13) = LQT
       txtspm(13).Refresh: frmBase.txtLQT.Refresh
       DoEvents
   End If

End Sub

Private Sub cmdLarger_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    idxL = Index
    DoClickL = True
    DoL = 0
End Sub

Private Sub cmdLarger_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoClickL = False
End Sub

Private Sub cmdLoadSig_Click()

Dim X As Long, Y As Long, intf As Integer, By As Byte, BBy As Byte
    intf = FreeFile
    Open txtPath & "\" & txtMaxShot(0) & ".kpi" For Random As #intf Len = 1
    BBy = 2
    For X = 1 To FileLen(txtPath & "\" & txtMaxShot(0) & ".kpi")
        Get #intf, X, By
        PrK(2, X) = By
        If By > BBy Then BBy = By
        PrK(3, X) = BBy
    Next X

    Close #intf
End Sub

Private Sub cmdLogClr_Click()
    lstLogs.Clear
End Sub

Private Sub cmdLogs_Click()
    If fraLogs.Visible = True Then
            fraLogs.Visible = False
    Else
        fraLogs.Top = 30
        fraLogs.Visible = True
        fraLogs.ZOrder 0
    End If
End Sub

Private Sub cmdMaxPoints_Click()
'    txtspm(20) = 10
    txtspm(21) = 148900
End Sub

Private Sub cmdMiniMize_Click()

    frmBase.WindowState = 1

End Sub

Private Sub cmdMax_Click()
    cmdMax.Visible = False
    cmdMini.Visible = True
    fraProcess.Visible = False
'    fraBlur.Visible = False
'    fraControls.Visible = False
'    fraColors.Visible = False
    fraLogs.Visible = False
'    fraFullScr.Visible = True
    txtspm(22) = 512
End Sub

Private Sub cmdMini_Click()
    cmdMax.Visible = True
    cmdMini.Visible = False
    fraProcess.Visible = True
'    fraBlur.Visible = True
'    fraControls.Visible = True
'    fraColors.Visible = True
'    fraFullScr.Visible = False
    txtspm(22) = 640
End Sub

Private Sub cmdNav_Click(Index As Integer)
'    If Index = 0 Then txtNavigate(1) = txtNavigate(1) + 4
'    If Index = 1 Then txtNavigate(1) = txtNavigate(1) - 4
'
'    If Index = 2 Then txtNavigate(3) = txtNavigate(3) - 8
'    If Index = 3 Then txtNavigate(2) = txtNavigate(2) + 8
'    If Index = 4 Then txtNavigate(3) = txtNavigate(3) + 8
'    If Index = 5 Then txtNavigate(2) = txtNavigate(2) - 8
    If Index = 6 Then fraTelo.Left = 3990: fraTelo.Top = 0
    If Index = 7 Then fraTelo.Left = 0: fraTelo.Top = 620
    
End Sub


Private Sub cmdNormalSize_Click()
If cmdNormalSize.Tag <> "0" Then
    frmBase.Width = frmBase.Width - 3000
    frmBase.Height = frmBase.Height - 2000
    
    picViewEE.Width = frmBase.Width
    picViewEE.Height = frmBase.Height
    picViewEE.Left = 0
    picViewEE.Top = 0
    
    fraControls.Left = fraControls.Left - 3000
    fraColors.Left = fraColors.Left - 2000
    fraBlur.Left = fraBlur.Left - 1000
    fraControls.Top = picViewEE.Top
    fraColors.Top = picViewEE.Top
    fraBlur.Top = picViewEE.Top
    fraProcess.Top = picViewEE.Top
    
'    txtFrm.Left = txtFrm.Left - 3000 - 30
'    txtFrm.Top = txtFrm.Top + 30
    cmdNormalSize.Tag = "0"
Else
    frmBase.Width = frmBase.Width + 3000
    frmBase.Height = frmBase.Height + 2000
    picViewEE.Width = frmBase.Width
    picViewEE.Height = frmBase.Height
    picViewEE.Left = 0
    picViewEE.Top = 0
    
    fraControls.Left = fraControls.Left + 3000
    fraColors.Left = fraColors.Left + 2000
    fraBlur.Left = fraBlur.Left + 1000
    fraControls.Top = picViewEE.Top
    fraColors.Top = picViewEE.Top
    fraBlur.Top = picViewEE.Top
    fraProcess.Top = picViewEE.Top
    
'    txtFrm.Left = txtFrm.Left + 3000
'    txtFrm.Top = picViewEE.Top
    cmdNormalSize.Tag = "1"
End If
 
End Sub


Private Sub cmdOpenTelo_Click()
    Nvg(1) = 286: Nvg(2) = 15: Nvg(3) = 738
    If fraTelo.Visible = False Then
        fraTelo.Visible = True
        fraTelo.ZOrder 0
     Else
        fraTelo.Visible = False
        chkBox.Value = 0
    End If
    
End Sub
Private Sub cmdRGB_Click(Index As Integer)
    chRGB (Index)
End Sub

Private Sub cmdRGB_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraColors.ZOrder 0
End Sub

Private Sub cmdRColSt_Click()
Dim A As Byte
    For A = 0 To chkCol.count - 1
       chkCol(A) = 0
    Next A
    Colv_R = Primes(LQT) Mod 256
    Colv_G = Primes(LQT - 1) Mod 256
    Colv_B = Primes(LQT - 2) Mod 256
    frmBase.chkCol(Rnd * 5) = 1
    
    A = 0
    Do While cS(0) = 0 Or cS(1) = 0 Or cS(2) = 0
        cS(0) = Rnd * 1: cS(1) = Rnd * 1: cS(2) = Rnd * 1
        A = A + 1
        If A > 200 Then cS(0) = 1: cS(1) = 1: cS(2) = 1: Exit Do
    Loop

End Sub

Private Sub cmdRstCol_Click()

End Sub

Private Sub cmdSaveSig_Click()
Dim X As Long, Y As Long, intf As Integer, By As Byte
    intf = FreeFile
    Open txtPath & "\" & txtMaxShot(0) & ".kpi" For Random As #intf Len = 1
     
'    Keyp.H1 = "LQ_SKY_v7.7! 2  "
'    Keyp.H2 = 1
'    Keyp.H3 = 64000
    For X = 1 To 100000
        By = PrK(2, X)
        Put #intf, X, By
    Next X
       
    Close #intf
    
End Sub

Public Sub cmdSmaler_Click(Index As Integer)
Dim cd As Integer
On Error Resume Next
    If Index = 16 Then txtspm(16) = Sqr(txtspm(16)): txtspm(16).Refresh
    If Index = 17 Then txtspm(17) = Sqr(txtspm(17)): txtspm(17).Refresh
    If Index = 18 Then txtspm(18) = Sqr(txtspm(18)): txtspm(18).Refresh
    If Index = 19 Then txtspm(19) = Sqr(txtspm(19)): txtspm(19).Refresh
    
    If Index = 33 And txtspm(33) > 0.1 Then txtspm(33) = txtspm(33) - 0.1
    If Index = 32 And txtspm(32) > 0 Then txtspm(32) = txtspm(32) - 1
    
    If Index = 34 Then txtspm(34) = txtspm(34) - 8
    If Index = 35 Then txtspm(35) = txtspm(35) - 8
    If Index = 36 Then txtspm(36) = txtspm(36) - 4
    
    If Index = 31 Then txtspm(16) = txtspm(16) / 2
    If Index = 30 Then txtspm(17) = txtspm(17) / 2
    If Index = 29 Then txtspm(18) = txtspm(18) / 2
    If Index = 27 Then txtspm(19) = txtspm(19) / 2
    
    If Index = 28 And txtspm(28) > 0 Then txtspm(28) = txtspm(28) - 1
    
    If Index = 20 And txtspm(20) > 0.1 Then txtspm(20) = txtspm(20) - 0.1: txtspm(20).Refresh
    If Index = 21 Then txtspm(21) = txtspm(21) - 1: txtspm(21).Refresh
    If Index = 22 Then txtspm(22) = txtspm(22) - 16: txtspm(22).Refresh
    If Index = 23 Then txtspm(23) = txtspm(23) - 16: txtspm(23).Refresh
    If Index = 25 Then txtspm(25) = txtspm(25) - 0.1: txtspm(25).Refresh
    
    If Index = 0 And Val(txtspm(0).Text) > 32 Then txtspm(0).Text = Val(txtspm(0).Text) - 32: txtspm(0).Refresh
    If Index = 1 And Val(txtspm(1).Text) > 1 Then txtspm(1).Text = Val(txtspm(1).Text) - 1
    If Index = 2 Then txtspm(2) = txtspm(2) - 0.0001
    If Index = 24 Then txtspm(2) = txtspm(2) - 0.01
    If Index = 26 Then txtspm(2) = txtspm(2) - 0.1
    If Index = 3 And txtspm(3) > 1 Then txtspm(3).Text = Val(txtspm(3).Text) - 1
    If Index = 4 And txtspm(4) > 1 Then txtspm(4).Text = Val(txtspm(4).Text) - 1
    If Index = 5 And txtspm(5) > 0.15 Then txtspm(5).Text = Val(txtspm(5).Text) - 0.15 '* (txtspm(5) \ 10 + 1)
    If Index = 6 And txtspm(6) > 0 Then txtspm(6) = txtspm(6) - 0.1: txtspm(6).Refresh
    If Index = 7 And txtspm(7) Then txtspm(Index).Text = Val(txtspm(Index).Text) - 1
    If Index = 8 And txtspm(8) > 0.1 Then txtspm(8).Text = Val(txtspm(8).Text) - 0.1
    If Index = 9 And txtspm(9) > 1 Then txtspm(9) = txtspm(9) - 1
    If Index = 10 And txtQua > 5 Then txtQua = txtQua - 5
    If Index = 11 Then LQT2 = LQT2 - txtspm(2): txtspm(11) = LQT2
    
    If Index = 12 And txtspm(12) > 3.1 Then txtspm(12) = txtspm(12) - 0.1
    If Index = 15 And LQT > 1 Then LQT = LQT - 1:  txtspm(13) = LQT: txtspm(13).Refresh
    If Index = 14 And LQT > 1001 Then LQT = LQT - 1000: txtspm(13) = LQT: txtspm(13).Refresh
    If Index = 13 And LQT > 101 Then
        LQT = LQT - 100:  txtspm(13) = LQT
        txtspm(13).Refresh: frmBase.txtLQT.Refresh
        DoEvents
    End If
End Sub

Private Sub cmdSmaler_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    idxS = Index
    DoClickS = True
    DoS = 0
End Sub

Private Sub cmdSmaler_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoClickS = False
End Sub

Public Sub cmdSF_Click()
Dim S As String, B As Boolean, X As Integer
Dim s2 As String, s3 As String
    cmdSF.BackColor = vbWhite
    cmdSF.Enabled = False
    txtShotCount = Val(txtShotCount) + 1
    
    s2 = txtspm(19) & "-" & txtspm(18) & "-" & txtspm(17) & _
                "-" & txtspm(16) & "-" & Int(LQT2)  'txtspm(17) & "-" & txtspm(18) & "-" & txtspm(19) & "-" & txtspm(7)
        
    If txtShotCount Mod txtMaxShot(0).Text = 0 And chkBreakShot(0) Then
        cmdDriectory_Click
    End If
    If Val(txtShotCount.Text) >= Val(txtMaxShot(1).Text) And chkBreakShot(1) Then
        Exit Sub
    End If
    
    SaveCount = SaveCount + 1
    
    If Trim(sPath) = "" Then sPath = BrowseForFolder("e:\", Me.hWnd, "Select Folder For Save Image")  'App.Path & "\"
    If Len(Trim(sPath)) < 2 Then Exit Sub
    
    S = sPath & "\HPP-" & CStr(SaveCount) & "-" & s2 & ".jpg"
    
    chkPause.Value = 1
    
    SaveJpeg S, txtQua, picView
    
    chkPause.Value = 0
    
    SaveSetting "KV_M_B", "kvvisulation", "SaveCount", SaveCount
    SaveSetting "KV_M_B", "kvvisulation", "sPath", sPath
    
    
    cmdSF.Enabled = True
    cmdSF.BackColor = &HFF00&

End Sub

Private Sub Combo1_Click()
    Combo1_Validate True
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
    If Combo1.ListIndex > 0 Then picBuff.DrawMode = Combo1.ListIndex + 1
End Sub
Private Sub Combo2_Click()
    Combo2_Validate True
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
    If Combo2.ListIndex > 0 Then
        
    End If
End Sub

Public Sub cmdCopy2Excel_Click()
Dim X As Integer, z1 As Double, Z2 As Double                '''' only for test KRandom()
  ReDim ExelArray(0 To lstLogs.ListCount, 0 To 6)
    
    For X = 0 To lstLogs.ListCount - 1
        ExelArray(X, 0) = lstLogs.List(X)
    Next X


  ExcelSaveArray
End Sub


Public Sub cmdPrime_Click()
Dim a1 As Double, B As Double, co(0 To 255) As Long, coc(0 To 9) As Long, cocS(0 To 255) As Long, tm As Integer
Dim Tm2 As Long

    For a1 = 0 To 255
        co(a1) = RGB(a1 ^ 2 \ (4 * txtspm(10)), a1 ^ 2 \ (6 * txtspm(10)), a1 ^ 2 \ (3 * txtspm(10)))
    Next a1
    
    For B = 1 To 768 Step 3
    For a1 = 1 To 1024 Step 2
      
      tm = PrK(1, (a1 + B * 148.932 - 148.932)) Mod 10
      Tm2 = PrK(2, (a1 + B * 148.932 - 148.932))
'      SetPixel frmBase.PicPrime.hdc, a1, B, co(Tm2)
      
      coc(tm) = coc(tm) + 1
      tm = PrK(2, (a1 + B * 148.932 - 148.932))
      cocS(tm) = cocS(tm) + 1
    
    Next a1
    Next B
    
    
End Sub

Private Sub Command2_Click()
'    picImg.ZOrder 1
End Sub

Private Sub Form_Activate()
'    Sound.make330Hz
'    Init_DX7 (frmBase.Hwnd)
'
'    DSBWRITE 0, myByte(): DSBWRITE 1, myByte() 'Play it
''    HScroll1_Scroll    'Set directSound Frequency
''    HScroll2_Scroll    'Set directSound Frequency
''    VScroll1_Change    'Set directSound Volume
        

  Static waveFormat As WaveFormatEx

    With waveFormat
        .FormatTag = WAVE_FORMAT_PCM
        .Channels = 2
        .SamplesPerSec = 11025 '11khz
        .BitsPerSample = 8
        .BlockAlign = (.Channels * .BitsPerSample) \ 8
        .AvgBytesPerSec = .BlockAlign * .SamplesPerSec
        .ExtraDataSize = 0
    End With
    waveInOpen DevHandle, DevicesBox.ListIndex, VarPtr(waveFormat), 0, 0, 0
    If DevHandle = 0 Then
        Call MsgBox("Wave input device didn't open!", vbExclamation, "Ack!") ' 
        Exit Sub
    End If
    Call waveInStart(DevHandle)
    Inited = True
    DoEv = True
    
    
    BaseSub

End Sub

Private Sub Form_DblClick()
    If fraProcess.Visible = True Then
        cmdMax.Visible = False
        cmdMini.Visible = True
        
        fraProcess.Visible = False
'        fraFullScr.Visible = True
        txtspm(22) = 512
      Else
        cmdMax.Visible = True
        cmdMini.Visible = False
    
        fraProcess.Visible = True
'        fraFullScr.Visible = False
        txtspm(22) = 640
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyS Then cmdSF_Click
    If KeyCode = vbKeyC Then KaCls
    If KeyCode = vbKeyR Then cmd0_Click 10
    
    If KeyCode = vbKeyA Then
        If chkAutoShot.Value = 1 Then
            chkAutoShot.Value = 0
        Else
            chkAutoShot.Value = 1
        End If
        chkAutoShot_Click
    End If
    
'    If KeyCode = 107 Then cmdLarger_Click 2
'    If KeyCode = 109 Then cmdSmaler_Click 2
'    If KeyCode = vbKeyRight Then cmdLarger_Click 2: txtspm(2).Refresh
'    If KeyCode = vbKeyLeft Then cmdSmaler_Click 2: txtspm(2).Refresh
'    If KeyCode = vbKeyUp Then cmdLarger_Click 11: txtspm(11).Refresh
'    If KeyCode = vbKeyDown Then cmdSmaler_Click 11: txtspm(11).Refresh
    If KeyCode = 27 And cmdMini.Tag = "1" Then Unload Me
    If KeyCode = 27 Then cmdMini.Tag = "1": cmdMini_Click

End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'DoEvents
    PointerX = X
    PointerY = Y
End Sub


Private Sub Form_Resize()
    picViewEE.Width = frmBase.Width
    picViewEE.Height = frmBase.Height
'    txtFrm.Left = frmBase.Width - 360
    fraControls.Left = frmBase.Width - fraControls.Width - 360
    fraColors.Left = fraControls.Left - fraColors.Width
    fraBlur.Left = fraColors.Left - fraBlur.Width
End Sub

Private Sub Label12_Click()
    fraTelo.Visible = False
    fraTelo.ZOrder 0
    chkBox.Value = 0
End Sub

Private Sub lblClick_Click()
    lblLQSky_Click
End Sub

Private Sub lblControls_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraControls.ZOrder 0
End Sub

Public Sub txtFocus_GotFocus()
    DoEvents
End Sub

Private Sub txtFormula_Click()
    chkScript.Value = 0
End Sub

Private Sub Form_Load()
    
    Set picView = picViewEE
    Set picBuff = picBuffEE
    Set picBuffSe = picBuffEE
    Set picBuffSe2 = picBuffEE
    Set picTmp = picBuffEE2
    picView.Refresh
    picView.Visible = True
    picView.ZOrder 1
    
    initSetData
    lblFullscr(0).Caption = "Liquid Sky " & App.Major & "." & App.Minor & "." & App.Revision
    SsPtr = 0
    txP = 1: tyP = 1
    Set Clk = New cCpuClk                                     'Create the CpuClk instance
    DoEvents
    Call QueryPerformanceFrequency(cCycles) ' 
    
    stFirst = 1
    dRFlag = 1
    
    fraProcess.Height = 350
    fraBlur.Height = 350
    fraControls.Height = 350
    picBProcs.Height = 10215
    
    FlgBlur = 1
    Set_Process
    
    xColStp = 1
    yCol = 5
    xCol = 5
    
    LoadREG
    
    cmdChandSaveFolder_Click

'    On Error Resume Next
'      Debug.Print 1 / 0
'      If Err Then
'          MsgBox " . If Compile The Code Before Run . Its Runing About 2 Times Farster!!!", , " LQ_SKYS Present  ..."
'      End If

End Sub

Private Sub initSetData()
Dim caps As WAVEINCAPS, Which As Long
Dim X As Integer
     
    MVolu = 1
    Fst = True
    DoEv = 11
   
    ABass = 1
    AMidl = 1
    ATreb = 1
    AFreq = 1
    ABass2 = 1
    AMidl2 = 1
    ATreb2 = 1
    AFreq2 = 1
    Randomize (Timer)
    RV = CDbl(Rnd(2) * 255)
    GV = CDbl(Rnd(3) * 255)
    BV = CDbl(Rnd(5) * 255)
    Randomize (Timer)
    RN = CDbl(Rnd(2) * 255)
    GN = CDbl(Rnd(3) * 255)
    BN = CDbl(Rnd(5) * 255)
    
    PiTAdd1 = 8.72664625997165E-03
    PiTAdd2 = 1.74532925199433E-02
    
    For X = 0 To 2
        MaxC(X) = Val(txtMaxC(X).Text)
        MinC(X) = Val(txtMinC(X).Text)
    Next X
    
    Ox = 0
    Oy = 128
    Ox2 = 256
    Oy2 = 128
    BlurNum = 0
    
    tx = 1
    ty = 1

    DevicesBox.Clear
    For Which = 0 To waveInGetNumDevs - 1
        Call waveInGetDevCaps(Which, VarPtr(caps), Len(caps)) ' 
        If caps.Formats And WAVE_FORMAT_1S08 Then 'Now is 1S08 -- Check for devices that can do stereo 8-bit 11kHz
            Call DevicesBox.AddItem(StrConv(caps.ProductName, vbUnicode), Which) ' 
        End If
    Next ' Repeat For-Variable: WHICH
    If DevicesBox.ListCount = 0 Then
        MsgBox "You have no audio input devices!", vbCritical, "Ack!"
        End ' There are better ways to terminate
    End If
    DevicesBox.ListIndex = 0

    ColSt(0) = &H0&     'Black 0+0+0
    ColSt(1) = &HFF&    'Red
    ColSt(2) = &HFF00&  'Green
    ColSt(3) = &HFF0000 'Blue
    ColSt(4) = &HFFFF00 'Cyan B+G
    ColSt(5) = &HFF00FF 'Maginta B+R
    ColSt(6) = &HFFFF&  'Yelow G+R
    ColSt(7) = &H7F7F7F 'Gray  127+127+127
    ColSt(8) = &HFFFFFF 'White 255+255+255
    ColSt(9) = ColSt(8) Xor ColSt(2)
    ColSt(10) = ColSt(5) Xor ColSt(4)

End Sub


Private Sub Form_Unload(Cancel As Integer)

    
    If DevHandle <> 0 Then
    Call waveInReset(DevHandle) ' 
    Call waveInClose(DevHandle) ' 
    DoEvents
    DevHandle = 0
    
    End If
    SaveREG

    End
End Sub

Private Sub fraBlur_Click()
    fraBlur.ZOrder (0)
End Sub

Private Sub fraBlur_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEv = True
End Sub

Private Sub fraBlur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraBlur.ZOrder (0)
End Sub

Private Sub fraBlur_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEv = False
End Sub

Private Sub fraColors_Click()
    fraColors.ZOrder (0)
End Sub

Private Sub fraColors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEv = True
End Sub

Private Sub fraColors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEv = False
End Sub

Private Sub fraControls_Click()
    fraControls.ZOrder (0)
End Sub


Private Sub fraControls_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEv = True
End Sub

Private Sub fraControls_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEv = False
End Sub

Private Sub fraProcess_Click()
    fraProcess.ZOrder (0)
End Sub


Private Sub fraProcess_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEv = True
End Sub

Private Sub fraProcess_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    DoEv = False
End Sub

Private Sub lblBlur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraBlur.ZOrder 0
End Sub

Private Sub lblColorSet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraColors.ZOrder 0
End Sub

Private Sub lblLogs_Click()
'    fraLogs.Visible = False
    If fraLogs.Height > 400 Then
        fraLogs.Height = 350
      Else
        fraLogs.Height = 4200
        fraLogs.ZOrder (0)
    End If
End Sub

Private Sub lblColorSet_Click()
    If fraColors.Height > 400 Then
        fraColors.Height = 350
      Else
        fraColors.Height = 4095
        fraColors.ZOrder (0)
    End If
End Sub

Private Sub lblControls_Click()
    If fraControls.Height > 400 Then
        fraControls.Height = 350
      Else
        fraControls.Height = 9120
        fraControls.ZOrder (0)
    End If
End Sub

Private Sub lblBlur_Click()
    If fraBlur.Height > 400 Then
        fraBlur.Height = 350
      Else
        fraBlur.Height = 4620
        fraBlur.ZOrder (0)
    End If
End Sub

Private Sub lblLogs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraLogs.ZOrder 0
End Sub

Private Sub lblLQSky_Click()
    If fraProcess.Height > 635 Then
        fraProcess.Height = 600
      Else
        fraProcess.Height = 11535
        fraProcess.ZOrder (0)
    End If
End Sub

Private Sub lblVol_Click()
    lblControls_Click
End Sub

Private Sub cmdSnCGr_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If cS(Index) > 0 Then
        cS(Index) = -1
    ElseIf cS(Index) < 0 Then
        cS(Index) = 0
    Else
        cS(Index) = 1
    End If
End Sub

Private Sub List1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraLogs.ZOrder 0
End Sub

Private Sub List2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraLogs.ZOrder 0
End Sub

Private Sub lstFunctions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    fraProcess.ZOrder 0
End Sub


Private Sub lstProcess_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    fraProcess.ZOrder 0
End Sub

Private Sub lstPsent_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    fraProcess.ZOrder 0
End Sub

Private Sub picBBlur_Click()
    fraBlur.ZOrder 0
End Sub

Private Sub picBBlur_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraBlur.ZOrder (0)
End Sub

Private Sub picBCol_Click()
    fraColors.ZOrder 0
End Sub

Private Sub picBCol_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraColors.ZOrder 0
End Sub

Private Sub picBCtrl_Click()
    fraControls.ZOrder 0
End Sub

Private Sub picBCtrl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraControls.ZOrder 0
End Sub

Private Sub picBLogs_Click()
    fraLogs.ZOrder 0
End Sub

Private Sub picBLogs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    fraLogs.ZOrder 0
End Sub

Private Sub picBProcs_Click()
    fraProcess.ZOrder 0
End Sub

Private Sub picBProcs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    fraProcess.ZOrder 0
End Sub

Private Sub picViewEE_DblClick()
    Call Form_DblClick
End Sub

Private Sub Timer_AHeight_Timer()
    Timer_AHeight.Tag = Val(Timer_AHeight.Tag) + 1
    If Val(Timer_AHeight.Tag) > 1 Then
        Timer_AHeight.Enabled = False
        chkAHeight.Value = 0
        Timer_AHeight.Tag = "0"
    End If
    If (Abs(maxY - minY) < (txtspm(0) * 1.1)) And (Abs(maxY - minY) > (txtspm(0) * 0.9)) Then
        Timer_AHeight.Enabled = False
        chkAHeight.Value = 0
        Timer_AHeight.Tag = "0"
    End If
End Sub


Private Sub Timer_AutoLog_Timer()
'    If frmBase.chkALog.Value Then Loger
End Sub

Private Sub Timer_Process_Timer()
    Dim hIcon As Long, pdhStatus As Long
    Dim dbl As Double

        If AvgUsageCount = 0 Then
            AvgCpu0 = 0
            AvgCpu1 = 0
        End If
        If AvgUsageCount > 100 Then
            AvgCpu0 = AvgCpu0 / AvgUsageCount
            AvgCpu1 = AvgCpu1 / AvgUsageCount
            AvgUsageCount = 1
        End If
    
        PdhCollectQueryData (hQuery)
    
        dbl = PdhVbGetDoubleCounterValue(Counters(0).hCounter, pdhStatus)
        If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
            AvgCpu0 = AvgCpu0 + dbl
        End If
        
        If NumOfCores = 2 Then
            dbl = PdhVbGetDoubleCounterValue(Counters(1).hCounter, pdhStatus)
            If (pdhStatus = PDH_CSTATUS_VALID_DATA) Or (pdhStatus = PDH_CSTATUS_NEW_DATA) Then
                AvgCpu1 = AvgCpu1 + dbl
            End If
        End If
    
     
        AvgUsageCount = AvgUsageCount + 1
    
    txtProcess0 = CStr(Int(AvgCpu0 / AvgUsageCount)) + " %"
    txtProcess1 = CStr(Int(AvgCpu1 / AvgUsageCount)) + " %"
    txtProcessSum.Text = CStr(Int(((AvgCpu0 / AvgUsageCount) + (AvgCpu1 / AvgUsageCount)) / 2)) + " %"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Sub

Private Sub WaveBlur()

  Dim intX As Integer, intY, i As Integer
  Dim intI As Integer, intJ As Integer
  Dim intWidth As Integer, intHeight As Integer

    intWidth = 1024 'picBuff.Width
    intHeight = 768 'picBuff.Height
    Randomize
      For i = 1 To 1000
                 intX = (intWidth) * Rnd
                 intY = (intHeight) * Rnd
                 intI = M1 * Rnd - z1
                 intJ = M2 * Rnd - Z2
                 K1 = K1 * Rnd
                 K2 = K2 * Rnd
                 BitBlt picBuff.hdc, intX + intI, intY + intJ, K1, K2, picBuff.hdc, intX, intY, vbSrcCopy
      Next i

End Sub


Private Sub Timer_AutoSave_Timer()

On Error Resume Next

 TiS = TiS + 1
    If TiS >= txtspm(8) * 10 And chkAutoShot Then
        TiS = 0
        cmdSF_Click
        txtspm(8).BackColor = txtspm(8).BackColor Xor vbBlue
        txtspm(8).ForeColor = txtspm(8).ForeColor Xor vbRed
        TiS = 0
    End If

End Sub

Private Sub Timer_Seconds_Timer()
    St_Time = St_Time + 1
End Sub


Private Sub txtFpS_Change()

    txtFrm = Int(txtFpS)
    If Val(txtFpS.Text) < 16 Then txtFpS.BackColor = &HFF&: GoTo en
    If Val(txtFpS.Text) < 20 Then txtFpS.BackColor = &H64FB&: GoTo en
    If Val(txtFpS.Text) < 24 Then txtFpS.BackColor = &H1E7FA: GoTo en
    If Val(txtFpS.Text) < 28 Then txtFpS.BackColor = &H1FABC: GoTo en
    txtFpS.BackColor = &HFF00&
en:
End Sub


Private Sub txtMaxC_Change(Index As Integer)
'    MaxC(Index) = Val(txtMaxC(Index).Text)
End Sub

Private Sub txtMinC_Change(Index As Integer)
    MinC(Index) = Val(txtMinC(Index).Text)
End Sub

Private Sub txtShotCount_DblClick()
    txtShotCount = 0
End Sub

Private Sub txtspm_Change(Index As Integer)
 
    If Index = 0 Then chkAHeight.Value = 1
    If txtspm(9) < 1 And Index = 9 Then txtspm(9) = 1
'    If Index = 16 And chkLock Then txtspm(17) = txtspm(16)
End Sub

Private Sub txtspm_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
'Dim x, y, z
'    z = txtspm(Index)
'    If Not IsNumeric(Chr(KeyCode)) Then
'        txtspm(Index) = Replace(z, Chr(KeyCode), "")
'        Exit Sub
'    End If
    DoEvents
End Sub

