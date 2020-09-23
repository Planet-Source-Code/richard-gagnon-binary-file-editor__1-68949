VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FileEditor 
   Caption         =   "File Editor"
   ClientHeight    =   2730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3900
   Icon            =   "FileEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   3900
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar SBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2461
            MinWidth        =   2469
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Tbar 
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   1
      Left            =   2940
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   722
      TabIndex        =   41
      Top             =   210
      Width           =   10830
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   6
         Left            =   3780
         Picture         =   "FileEditor.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Clipboard"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   14
         Left            =   3150
         Picture         =   "FileEditor.frx":0BAC
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Insert Block"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   13
         Left            =   2520
         Picture         =   "FileEditor.frx":1316
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Insert Row"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   7
         Left            =   1890
         Picture         =   "FileEditor.frx":1A80
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Delete"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   5
         Left            =   1260
         Picture         =   "FileEditor.frx":21EA
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Paste"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   4
         Left            =   630
         Picture         =   "FileEditor.frx":2954
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Copy"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   3
         Left            =   0
         Picture         =   "FileEditor.frx":30BE
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Cut"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   9
         Left            =   5250
         Picture         =   "FileEditor.frx":3828
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Print Preview"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   8
         Left            =   4620
         Picture         =   "FileEditor.frx":3F92
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Print"
         Top             =   0
         Width           =   435
      End
      Begin VB.TextBox txbGoTo 
         BackColor       =   &H00E0E0E0&
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
         Left            =   7770
         MaxLength       =   9
         TabIndex        =   43
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   40
         Width           =   990
      End
      Begin VB.TextBox txbPage 
         BackColor       =   &H00E0E0E0&
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
         Left            =   10080
         MaxLength       =   5
         TabIndex        =   42
         TabStop         =   0   'False
         Text            =   "1"
         Top             =   40
         Width           =   675
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Goto File Location"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   5985
         TabIndex        =   45
         Top             =   75
         Width           =   1725
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Goto Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   8925
         TabIndex        =   44
         Top             =   75
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   2
         X1              =   392
         X2              =   392
         Y1              =   0
         Y2              =   28
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   1
         X1              =   294
         X2              =   294
         Y1              =   0
         Y2              =   28
      End
   End
   Begin VB.PictureBox Tbar 
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   2
      Left            =   210
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   904
      TabIndex        =   37
      Top             =   840
      Width           =   13560
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   12
         Left            =   12495
         Picture         =   "FileEditor.frx":46FC
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Find First"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   11
         Left            =   13125
         Picture         =   "FileEditor.frx":4E66
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Find Next"
         Top             =   0
         Width           =   435
      End
      Begin VB.TextBox txbSearch 
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2625
         TabIndex        =   40
         TabStop         =   0   'False
         Text            =   "Search String"
         Top             =   0
         Width           =   9675
      End
      Begin VB.CommandButton SearchType 
         Caption         =   "TEXT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1680
         TabIndex        =   39
         TabStop         =   0   'False
         Tag             =   "1"
         ToolTipText     =   "Click to Search HEX"
         Top             =   0
         Width           =   855
      End
      Begin VB.CheckBox sCase 
         Caption         =   "Match Case"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   0
         TabIndex        =   38
         TabStop         =   0   'False
         ToolTipText     =   "Match Case"
         Top             =   0
         Width           =   1485
      End
   End
   Begin VB.PictureBox Tbar 
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   0
      Left            =   210
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   176
      TabIndex        =   36
      Top             =   210
      Width           =   2640
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   10
         Left            =   1890
         Picture         =   "FileEditor.frx":55D0
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Help"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   2
         Left            =   1260
         Picture         =   "FileEditor.frx":5D3A
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Save"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   1
         Left            =   630
         Picture         =   "FileEditor.frx":64A4
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Open"
         Top             =   0
         Width           =   435
      End
      Begin VB.CommandButton cmdBut 
         Height          =   435
         Index           =   0
         Left            =   0
         Picture         =   "FileEditor.frx":6C0E
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "New"
         Top             =   0
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Index           =   0
         X1              =   168
         X2              =   168
         Y1              =   0
         Y2              =   28
      End
   End
   Begin VB.Frame SbarFrame 
      BorderStyle     =   0  'None
      Height          =   2010
      Left            =   735
      TabIndex        =   29
      Top             =   1470
      Width           =   330
      Begin VB.CommandButton cmdScroll 
         Height          =   345
         Index           =   1
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FileEditor.frx":7378
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Next Block"
         Top             =   1470
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton cmdScroll 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   3
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FileEditor.frx":76D9
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Next Row"
         Top             =   1155
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton cmdScroll 
         Appearance      =   0  'Flat
         Height          =   345
         Index           =   0
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FileEditor.frx":7A33
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Prev Block"
         Top             =   105
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.CommandButton cmdScroll 
         Height          =   345
         Index           =   2
         Left            =   0
         MaskColor       =   &H00FFFFFF&
         Picture         =   "FileEditor.frx":7D91
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Prev Row"
         Top             =   420
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.PictureBox SbarBox 
         Height          =   435
         Left            =   0
         ScaleHeight     =   375
         ScaleWidth      =   255
         TabIndex        =   30
         Top             =   735
         Width           =   315
         Begin VB.CommandButton Sbar 
            BackColor       =   &H0000FF00&
            Height          =   195
            Left            =   -100
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   0
            Width           =   500
         End
      End
   End
   Begin VB.Timer DEBUGtimer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   13965
      Top             =   1470
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   105
      Top             =   1470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "*.*"
      DialogTitle     =   "Open File"
      Filter          =   $"FileEditor.frx":80EB
   End
   Begin VB.Frame HexFrame 
      Caption         =   "Hex Grid"
      Height          =   645
      Left            =   2310
      TabIndex        =   11
      Top             =   1470
      Width           =   1065
      Begin VB.Label HexGrid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   12
         Tag             =   "0"
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.Frame TextFrame 
      Caption         =   "Text Grid"
      Height          =   645
      Left            =   3465
      TabIndex        =   9
      Top             =   1470
      Width           =   960
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Reference Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   4
         Left            =   840
         TabIndex        =   28
         Top             =   2415
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         Caption         =   "Search Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Index           =   3
         Left            =   840
         TabIndex        =   27
         Top             =   1995
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000012&
         Caption         =   "Selected Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   840
         TabIndex        =   26
         Top             =   1575
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Start of Page Marker"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   540
         Index           =   1
         Left            =   840
         TabIndex        =   25
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "Legend"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   840
         TabIndex        =   24
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label PGID 
         Caption         =   "Page 1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   840
         TabIndex        =   23
         Top             =   315
         Width           =   1170
      End
      Begin VB.Label TextGrid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Terminal"
            Size            =   12
            Charset         =   255
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   0
         Left            =   105
         TabIndex        =   10
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.Frame PreviewPrint 
      Caption         =   "Print Preview"
      Height          =   645
      Left            =   4515
      TabIndex        =   15
      Top             =   1470
      Visible         =   0   'False
      Width           =   1170
      Begin VB.CommandButton PrevOK 
         Caption         =   "OK"
         Height          =   330
         Left            =   210
         TabIndex        =   22
         Top             =   2415
         Width           =   750
      End
      Begin VB.Frame Prange 
         Caption         =   "Print Range"
         Height          =   1380
         Left            =   210
         TabIndex        =   17
         Top             =   945
         Width           =   3795
         Begin VB.TextBox PRpages 
            Height          =   330
            Left            =   1155
            TabIndex        =   20
            Top             =   735
            Width           =   2430
         End
         Begin VB.OptionButton PR 
            Caption         =   "Pages"
            Height          =   225
            Index           =   1
            Left            =   105
            TabIndex        =   19
            Top             =   780
            Width           =   855
         End
         Begin VB.OptionButton PR 
            Caption         =   "All"
            Height          =   225
            Index           =   0
            Left            =   105
            TabIndex        =   18
            Top             =   315
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.TextBox PVtext 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   210
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   210
         Width           =   330
      End
      Begin VB.Label PofP 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Page 0 of 0"
         Height          =   345
         Left            =   1470
         TabIndex        =   21
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.Frame FileLoc 
      Caption         =   "File Location"
      Height          =   645
      Left            =   1260
      TabIndex        =   13
      Top             =   1470
      Width           =   960
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   14
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   645
      Index           =   1
      Left            =   105
      Top             =   735
      Width           =   13770
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      Height          =   645
      Index           =   0
      Left            =   105
      Top             =   105
      Width           =   13770
   End
   Begin VB.Label DEBUGtext 
      Height          =   225
      Index           =   3
      Left            =   14490
      TabIndex        =   8
      Top             =   735
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label DEBUGtext 
      Height          =   225
      Index           =   2
      Left            =   14490
      TabIndex        =   7
      Top             =   525
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label DEBUGtext 
      Height          =   225
      Index           =   1
      Left            =   14490
      TabIndex        =   6
      Top             =   315
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label DEBUGtext 
      Height          =   225
      Index           =   0
      Left            =   14490
      TabIndex        =   5
      Top             =   105
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label DEBUGLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "FL:"
      Height          =   225
      Index           =   3
      Left            =   13965
      TabIndex        =   4
      Top             =   735
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label DEBUGLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "SFP:"
      Height          =   225
      Index           =   2
      Left            =   13965
      TabIndex        =   3
      Top             =   525
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label DEBUGLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "SL:"
      Height          =   225
      Index           =   1
      Left            =   13965
      TabIndex        =   2
      Top             =   315
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label DEBUGLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "FF:"
      Height          =   225
      Index           =   0
      Left            =   13965
      TabIndex        =   1
      Top             =   105
      UseMnemonic     =   0   'False
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Menu mnuChoose 
      Caption         =   "Choose"
      Visible         =   0   'False
      Begin VB.Menu cmdPop 
         Caption         =   "Cut"
         Index           =   0
      End
      Begin VB.Menu cmdPop 
         Caption         =   "Copy"
         Index           =   1
      End
      Begin VB.Menu cmdPop 
         Caption         =   "Paste"
         Index           =   2
      End
      Begin VB.Menu cmdPop 
         Caption         =   "Delete"
         Index           =   3
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu popInsert 
         Caption         =   "Insert Row"
         Index           =   0
      End
      Begin VB.Menu popInsert 
         Caption         =   "Insert Block"
         Index           =   1
      End
   End
End
Attribute VB_Name = "FileEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
'----------------------------------------------------------\
'Author: Richard E. Gagnon.                                |
'URL:    http://members.cox.net/reg501/                    |
'Email:  reg501@cox.net                                    |
'Copyright © 2007 Richard E. Gagnon. All Rights Reserved.  |
'----------------------------------------------------------/

'----------------------------------------------------------\
'Don't forget to set:                                      |
'Project.. References... for "Microsoft Scripting Runtime" |
'This will enable the FileSystemObject                     |
'----------------------------------------------------------/

'Functions
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Dest As Any, Source As Any, ByVal Length As Long)

'Constants
Const KEY_F1 = &H70: Const KEY_F2 = &H71
Const KEY_F3 = &H72: Const KEY_F4 = &H73
Const KEY_F5 = &H74: Const KEY_F6 = &H75
Const KEY_F7 = &H76: Const KEY_F8 = &H77
Const KEY_F9 = &H78: Const KEY_F10 = &H79
Const KEY_F11 = &H7A: Const KEY_F12 = &H7B
Const KEY_SHIFT = &H10: Const KEY_ESCAPE = &H1B
Const KEY_PRIOR = &H21: Const KEY_NEXT = &H22
Const KEY_HOME = &H24: Const KEY_END = &H23
Const KEY_DELETE = &H2E: Const KEY_CTRL = &H11
Const KEY_LEFT = &H25: Const KEY_RIGHT = &H27
Const KEY_UP = &H26: Const KEY_DOWN = &H28
Const zGray = &HE0E0E0: Const zAqua = &HFFFFC0
Const zWhite = &HFFFFFF: Const zBlack = &H80000012
Const zRed = &H8080FF

'Printer variables and constants
Const CPL = 16          'Characters Per Line
Const RPP = 52          'Rows Per Page
Const CPP = CPL * RPP   'Characters Per Page (CPL*RPP)
Private PrintPages() As Boolean
'=======================
Private WorkSpace() As Byte 'Open File Array
'-----------------------
Private SbarPos As Long     'Scroll Bar Position
Private SL As Integer       '# of selected bytes
Private SFP As Long         'Byte pointer
Private FL As Long          'End of file pointer
Private FF As Long          'Top of Page pointer
Private GP As Integer       'Grid pointer
Private DragStart As Boolean
Private TypeOK As Boolean
Private GridType As Boolean 'TRUE=HEX grid, FALSE=TEXT grid
Private OpenFileName As String
Private FSO As New FileSystemObject

Private Function GetKeyInput(mKey As Long) As Integer
GetKeyInput = GetKeyState(mKey)
End Function

Private Sub ExportToClipboard()
Dim I As Integer, J As Integer
Dim CopyString As String
CopyString = Clipboard.GetText()
    For I = 0 To 511
        MyClip.TextGrid(I).Caption = ""
    Next I
    If GridType Then
        For I = 1 To Len(CopyString) Step 2
            MyClip.TextGrid(J).Caption = Mid(CopyString, I, 2)
            J = J + 1
        Next I
    Else
        For I = 1 To Len(CopyString)
            MyClip.TextGrid(I - 1).Caption = Mid(CopyString, I, 1)
        Next I
    End If
End Sub

Private Sub CutPaste(OP As Boolean, PasteBytes() As Byte)
'TRUE=Cut..FALSE=Paste
If SFP >= 0 And SL > 0 Then
    Dim nBytes As Long
    Dim I As Long
    If OP = False Then  'Paste
        nBytes = UBound(PasteBytes)
        ReDim Preserve WorkSpace(1 To FL + nBytes)
        ' Move the bytes DOWN in the Array--Dest,Src,Bytes to move
        If FL - SFP > 0 Then Call CopyMemory(ByVal VarPtr(WorkSpace(nBytes + SFP + 1)), ByVal VarPtr(WorkSpace(SFP + 1)), FL - SFP)
        ' Copy the selected data to the Array--Dest,Src,Bytes to move
        Call CopyMemory(ByVal VarPtr(WorkSpace(SFP + 1)), PasteBytes(1), nBytes)
        FL = FL + nBytes
    Else            'Cut
        If FL - SL > 0 Then
            ' Move the Data UP in the Array--Dest,Src,Bytes to move
            If FL - (SL + SFP) > 0 Then Call CopyMemory(ByVal VarPtr(WorkSpace(SFP + 1)), ByVal VarPtr(WorkSpace(SL + SFP + 1)), FL - (SL + SFP))
            FL = FL - SL
            ReDim Preserve WorkSpace(1 To FL)
        Else
            ReDim WorkSpace(1 To 1)
            FL = 1
        End If
    End If
    If FF > FL Then FF = FL
    If FF < 0 Then FF = 0
    ShowFileInfo
End If
End Sub

Private Sub UpdatePageCount()
If FL / CPP <> Int(FL / CPP) Then
    ReDim PrintPages(1 To Int(FL / CPP) + 1)
Else
    ReDim PrintPages(1 To Int(FL / CPP))
End If
End Sub

Private Sub goClip()
MyClip.Show
End Sub

Private Sub goCopy()
If SFP >= 0 And SL > 0 Then
    Dim saveok As Integer
    Dim cBad As Boolean
    Dim I As Integer, J As Integer
    Dim CopyString As String
    cBad = False
    J = SFP - FF
    Clipboard.Clear
    For I = J To (J + SL) - 1
        If GridType Then
            CopyString = CopyString + HexGrid(I).Caption
        Else
            If HexGrid(I).Caption = "00" Then cBad = True
            CopyString = CopyString + TextGrid(I).Caption
        End If
    Next I
    saveok = True
    If cBad Then
        If MsgBox("The text you have selected contains" & vbCrLf & _
                  "null characters and will not paste correctly." & vbCrLf & _
                  "Copy from HEX Grid instead" & vbCrLf & vbCrLf & _
                  " Want to continue with the current operation?" _
                  , vbExclamation + vbYesNo, " File Editor") = vbNo Then saveok = False
    End If
    If saveok Then
        Clipboard.SetText CopyString
        ExportToClipboard
    End If
End If
End Sub

Private Sub goCut()
If SFP >= 0 And SL > 0 Then
    Dim I As Integer, J As Integer
    Dim CopyString As String
    Dim saveok As Integer
    Dim cBad As Boolean
    cBad = False
    Clipboard.Clear
    J = SFP - FF
    For I = J To (J + SL) - 1
        If GridType Then
             CopyString = CopyString + HexGrid(I).Caption
        Else
             If HexGrid(I).Caption = "00" Then cBad = True
             CopyString = CopyString + TextGrid(I).Caption
        End If
    Next I
    saveok = True
    If cBad Then
        If MsgBox("The text you have selected contains" & vbCrLf & _
                  "null characters and will not paste correctly." & vbCrLf & _
                  "Copy from HEX Grid instead" & vbCrLf & vbCrLf & _
                  " Want to continue with the current operation?" _
                  , vbExclamation + vbYesNo, " File Editor") = vbNo Then saveok = False
    End If
    If saveok Then
        Dim dummy() As Byte
        Clipboard.SetText CopyString
        ExportToClipboard
        CutPaste True, dummy()
        UpdateGridData
        SelectGrid SFP - FF
    End If
End If
End Sub

Private Sub goDelete()
If SFP >= 0 And SL > 0 Then
    Dim dummy() As Byte
    CutPaste True, dummy()
    UpdateGridData
    SelectGrid SFP - FF
End If
End Sub

Private Sub goHelp()
FileEditHelp.Show
End Sub

Private Sub goNew()
Dim I As Integer
On Error GoTo NEWerr
Dim Fnum As Integer
Dim Fname As String
Dim saveok As Boolean
CommonDialog1.DialogTitle = "SAVE FILE"
CommonDialog1.FileName = OpenFileName
CommonDialog1.Filter = "All files (*.*)|*.*|WAV files (*.wav)|*.wav|COM files (*.com)|*.com|BIN files (*.bin)|*.bin |SYS files (*.sys)|*.sys|DLL files (*.dll)|*.dll|HEX files (*.hex)|*.hex"
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.ShowSave
Fname = CommonDialog1.FileName
saveok = True
If Dir(Fname) <> "" Then
    If MsgBox("Do you want to overwrite file '" & Fname & " ' ?", vbQuestion + vbYesNo, " File Editor") = vbNo Then saveok = False
End If
If saveok Then
    ShowGrids True
    If FSO.FileExists(Fname) Then FSO.DeleteFile (Fname)
    ReDim WorkSpace(1 To 1)
    Fnum = FreeFile
    Open Fname For Binary Access Write As Fnum
    Put Fnum, , WorkSpace()
    Close Fnum
    Me.Caption = "File Editor"
    SFP = 0: SL = 1: FF = 0
    FillLabels
    ClearGrid
    For I = 0 To 511
        HexGrid(I).Caption = ""
        HexGrid(I).Tag = ""
        TextGrid(I).Caption = ""
    Next I
    HexGrid(0).Caption = "00"
    OpenFileName = Fname
    Me.Caption = OpenFileName
    ShowFileInfo
    FL = FSO.GetFile(OpenFileName).Size
End If
If Tbar(1).Enabled Then txbSearch.SetFocus: TypeOK = False
Exit Sub
NEWerr:
If Tbar(1).Enabled Then txbSearch.SetFocus: TypeOK = False
If Err <> 32755 Then MsgBox (Error & vbCr & vbCr & "Error Number: " & Str(Err)), vbCritical, "! ERROR !"
End Sub

Private Sub goOpen()
Dim FilePath As String
Dim Fnum As Integer
On Error GoTo OPNerr
CommonDialog1.Filter = "All files (*.*)|*.*|WAV files (*.wav)|*.wav|COM files (*.com)|*.com|BIN files (*.bin)|*.bin |SYS files (*.sys)|*.sys|DLL files (*.dll)|*.dll|HEX files (*.hex)|*.hex"
CommonDialog1.DialogTitle = "OPEN FILE"
CommonDialog1.FileName = ""
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.ShowOpen
OpenFileName = CommonDialog1.FileName
If OpenFileName <> "" Then
    FL = FSO.GetFile(OpenFileName).Size
    'Max file size = 2,147,483,647 bytes
    If FL > 0 Then
        ShowGrids True
        Fnum = FreeFile
        Open OpenFileName For Binary Access Read As Fnum
        ReDim WorkSpace(1 To FL)
        Get Fnum, , WorkSpace
        Close Fnum
        Me.Caption = OpenFileName
        FF = 0
        UpdateGridData
        SelectGrid 0, True
        ShowFileInfo
    Else
        MsgBox ("File contains no data!..."), vbExclamation, "No Data"
    End If
End If
If Tbar(1).Enabled Then txbSearch.SetFocus: TypeOK = False
Exit Sub
OPNerr:
If Tbar(1).Enabled Then txbSearch.SetFocus: TypeOK = False
If Err <> 32755 Then MsgBox (Error & vbCr & vbCr & "Error Number: " & Str(Err)), vbCritical, "! ERROR !"
End Sub

Private Sub goPaste()
On Error Resume Next
Dim nBytes As Long
Dim TempByte As Byte
Dim CopyString As String
Dim arrPaste() As Byte
CopyString = Clipboard.GetText()
MousePointer = 11
SBar1.Panels(4).Text = "Processing....."
If CopyString <> "" Then
    Dim I As Long
    Dim zSL As Integer, zSFP As Long
    TempByte = CByte("&h" & Mid(CopyString, 1, 2))
    If Not GridType Then
        nBytes = Len(CopyString)
        ReDim arrPaste(1 To nBytes)
        For I = 1 To nBytes
            arrPaste(I) = Asc(Mid(CopyString, I, 1))
            If I = 65536 Then DoEvents
        Next I
    Else
        nBytes = Len(CopyString) / 2
        ReDim arrPaste(1 To nBytes)
        For I = 1 To nBytes
            arrPaste(I) = CByte("&h" & Mid(CopyString, I * 2 - 1, 2))
            If I = 65536 Then DoEvents
            If Err = 13 Then
                MsgBox ("Paste data is invalid Hex data." & vbCr & "Try pasting in Text Grid instead"), vbInformation, " File Editor"
                GoTo ErrOut
            End If
        Next I
    End If
    If SL > 1 Then
        zSL = SL: zSFP = SFP    'Preserve pointers
        CutPaste True, arrPaste()
        SL = zSL: SFP = zSFP    'Restore pointers
    End If
    CutPaste False, arrPaste()
    UpdateGridData
    SelectGrid SFP - FF
Else
    MsgBox ("Nothing to Paste"), vbInformation, " File Editor"
End If
ErrOut:
MousePointer = 0
SBar1.Panels(4).Text = ""
End Sub

Private Sub goPreview()
If FL > 0 Then
    Tbar(0).Visible = False: Tbar(1).Visible = False
    Tbar(2).Visible = False
    Shape1(0).Visible = False: Shape1(1).Visible = False
    FileLoc.Visible = False: SbarFrame.Visible = False
    HexFrame.Visible = False: TextFrame.Visible = False
    UpdatePageCount
    PreviewPrintHex
    PreviewPrint.Visible = True
    PreviewPrint.Left = (Me.Width / 2) - (PreviewPrint.Width / 2)
End If
End Sub

Private Sub cmdBut_Click(Index As Integer)
TypeOK = True
Select Case Index
    Case 0: goNew
    Case 1: goOpen
    Case 2: goSave
    Case 3: goCut
    Case 4: goCopy
    Case 5: goPaste
    Case 6: goClip
    Case 7: goDelete
    Case 8: goPrint
    Case 9: goPreview
    Case 10: goHelp
    Case 11: goSearch 0
    Case 12: goSearch 1
    Case 13: goInsert 0
    Case 14: goInsert 1
End Select
End Sub

Private Sub cmdPop_Click(Index As Integer)
Select Case Index
    Case 0: goCut
    Case 1: goCopy
    Case 2: goPaste
    Case 3: goDelete
End Select
End Sub

Private Sub goPrint()
If FL > 1 Then
    Dim I As Long, I1 As Long
    Dim J As Byte
    Dim PTP As Long
    Dim MyMsg As String
    Dim D1 As String
    Dim P1() As String
    Dim R1 As String, R2 As String
    Dim PrintOK As Boolean
    PrintOK = True
    If FL / CPP <> Int(FL / CPP) Then
        ReDim PrintPages(1 To Int(FL / CPP) + 1)
    Else
        ReDim PrintPages(1 To Int(FL / CPP))
    End If
    For I = 1 To UBound(PrintPages): PrintPages(I) = False: Next I
    If PR(0).Value Then For I = 1 To UBound(PrintPages): PrintPages(I) = True: Next I
    If PR(1).Value Then
        If Len(PRpages.Text) = 0 Then
            MyMsg = "No pages have been selected"
            PrintOK = False
        Else
            For I = 1 To Len(PRpages.Text)
                J = Asc(Mid(PRpages.Text, I, 1))
                If (J < 48 Or J > 58) And (J < 44 Or J > 45) Then
                    MyMsg = "Invalid page selection"
                    PrintOK = False
                End If
            Next I
        End If
        If PrintOK Then
            P1() = Split(PRpages.Text, ",")
            For I = 0 To UBound(P1)
                If InStr(1, P1(I), ":") Or InStr(1, P1(I), "-") Then
                    If InStr(1, P1(I), ":") Then D1 = ":" Else D1 = "-"
                    R1 = Mid(P1(I), 1, InStr(1, P1(I), D1) - 1)
                    R2 = Mid(P1(I), InStr(1, P1(I), D1) + 1, Len(P1(I)))
                    If Val(R1) < 1 Or Val(R1) > UBound(PrintPages) Then PrintOK = False
                    If Val(R2) < 1 Or Val(R2) > UBound(PrintPages) Then PrintOK = False
                    If PrintOK Then
                        For I1 = Val(R1) To Val(R2): PrintPages(I1) = True: Next I1
                    Else
                        MyMsg = "Invalid page selection"
                        PrintOK = False
                    End If
                Else
                    If Val(P1(I)) < 1 Or Val(P1(I)) > UBound(PrintPages) Then
                        MyMsg = "Invalid page selection"
                        PrintOK = False
                    Else
                        PrintPages(Val(P1(I))) = True
                    End If
                End If
            Next I
        End If
    End If
    If PrintOK Then
        For I = 1 To UBound(PrintPages)
            If PrintPages(I) Then PTP = PTP + 1
        Next I
        If MsgBox("Are you sure you want to print " & PTP & " Pages?", vbQuestion + vbYesNo, " Print") = vbYes Then
            For I = 1 To UBound(PrintPages)
                If PrintPages(I) = True Then PrintHex I
            Next
        End If
    Else
        MsgBox (MyMsg), vbInformation, " File Editor"
    End If
End If
End Sub

Private Sub goSave()
On Error GoTo CLSerr
Dim Fnum As Integer
Dim Fname As String
Dim saveok As Boolean
CommonDialog1.DialogTitle = "SAVE FILE"
CommonDialog1.FileName = OpenFileName
CommonDialog1.Filter = "All files (*.*)|*.*|WAV files (*.wav)|*.wav|COM files (*.com)|*.com|BIN files (*.bin)|*.bin |SYS files (*.sys)|*.sys|DLL files (*.dll)|*.dll|HEX files (*.hex)|*.hex"
CommonDialog1.Flags = cdlOFNFileMustExist
CommonDialog1.ShowSave
Fname = CommonDialog1.FileName
saveok = True
If Dir(Fname) <> "" Then
    If MsgBox("Do you want to overwrite file '" & Fname & " ' ?", vbQuestion + vbYesNo, " File Editor") = vbNo Then saveok = False
End If
If saveok Then
    ShowGrids True
    If FSO.FileExists(Fname) Then FSO.DeleteFile (Fname)
    Fnum = FreeFile
    Open Fname For Binary Access Write As Fnum
    Put Fnum, , WorkSpace()
    Close Fnum
    OpenFileName = Fname
    Me.Caption = OpenFileName
    ShowFileInfo
    FL = FSO.GetFile(OpenFileName).Size
End If
If Tbar(1).Enabled Then txbSearch.SetFocus: TypeOK = False
Exit Sub
CLSerr:
If Tbar(1).Enabled Then txbSearch.SetFocus: TypeOK = False
If Err <> 32755 Then MsgBox (Error & vbCr & vbCr & "Error Number: " & Str(Err)), vbCritical, "! ERROR !"
End Sub

Private Sub cmdScroll_Click(Index As Integer)
Select Case Index
    Case 0: If FF - 512 < 0 Then FF = 0 Else FF = FF - 512
    Case 1: If FF + 512 > FL Then FF = FL Else FF = FF + 512
    Case 2: If FF - 16 < 0 Then FF = 0 Else FF = FF - 16
    Case 3: If FF + 16 > FL Then FF = FL Else FF = FF + 16
End Select
UpdateGridData
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Static mKeylock As Boolean
If TypeOK And (Not mKeylock) Then
    mKeylock = True
    Select Case KeyCode
        Case KEY_F1:     goOpen
        Case KEY_F2:     goSave
        Case KEY_F3:     goSearch 1
        Case KEY_F4:     goCopy
        Case KEY_F5:     goPaste
        Case KEY_F6:     goCut
        Case KEY_F10:    goHelp
        Case KEY_HOME:   FF = 0: UpdateGridData: SelectGrid 0, True
        Case KEY_NEXT:   cmdScroll_Click 1
        Case KEY_PRIOR:  cmdScroll_Click 0
        Case KEY_END
            FF = FL - 1
            UpdateGridData False
            SFP = (FL - FF) - 1
            SelectGrid (FL - FF) - 1
        Case KEY_DELETE: goDelete
        Case KEY_UP:     If (SFP - FF) - 16 >= 0 Then SelectGrid ((SFP - FF) - 16), True Else cmdScroll_Click 2
        Case KEY_DOWN:   If (SFP - FF) + 16 <= 511 Then SelectGrid ((SFP - FF) + 16), True Else cmdScroll_Click 3
        Case KEY_LEFT
            If (SFP - FF) - 1 >= 0 Then
                SelectGrid ((SFP - FF) - 1), True
            Else
                If SFP > 0 Then
                    cmdScroll_Click 2
                    SelectGrid 15, True
                End If
            End If
        Case KEY_RIGHT
            If (SFP - FF) + 1 <= 511 Then
                SelectGrid ((SFP - FF) + 1), True
            Else
                If SFP < FL Then
                    cmdScroll_Click 3
                    SelectGrid 496, True
                End If
            End If
    End Select

End If
mKeylock = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If TypeOK And OpenFileName <> "" And KeyAscii > 31 Then
    Dim T1 As String, T2 As String
    Dim kCHR As String
    If FF + GP >= FL Then
        ReDim Preserve WorkSpace(1 To FL + 1)
        FL = FL + 1
        ShowFileInfo
    End If
    kCHR = Chr(KeyAscii)
    If GridType = False Then
        HexGrid(GP).Caption = Hex(KeyAscii)
        TextGrid(GP).Caption = kCHR
        WorkSpace(FF + GP + 1) = KeyAscii
        GP = GP + 1
    Else
        T2 = HexGrid(GP).Tag & kCHR
        If Len(T2) = 2 Then
            T1 = EvaluateHex(T2)
            If T1 = " " Then T1 = "00"
            HexGrid(GP).Caption = T1
            TextGrid(GP).Caption = Chr("&h" & T1)
            WorkSpace(FF + GP + 1) = "&h" & T1
            HexGrid(GP).Tag = ""
            GP = GP + 1
        Else
            HexGrid(GP).Caption = "0" & kCHR
            HexGrid(GP).Tag = kCHR
        End If
    End If
    If GP > 511 Then: GP = 496: cmdScroll_Click 3
    SelectGrid GP
End If
End Sub

Private Sub Form_Load()
Clipboard.Clear
CreateGrids
FL = 0
FF = 0
TypeOK = True
SelectGrid 0, True
End Sub

Private Sub Form_Resize()
Dim I As Long
With SBar1
    I = Me.Width - .Panels(1).Width - .Panels(2).Width - .Panels(3).Width - .Panels(5).Width
End With
If I > 800 Then SBar1.Panels(4).Width = I - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload MyClip
Unload FileEditHelp
ReDim WorkSpace(1 To 1)
End Sub

Private Sub HexGrid_Click(Index As Integer)
TypeOK = True
Me.SetFocus
HexGrid(Index).Tag = ""
GridType = True
SelectGrid Index, True
End Sub

Private Sub HexGrid_DblClick(Index As Integer)
DragStart = True
End Sub

Private Sub HexGrid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    GridType = True
    If SL > 0 Then Me.PopupMenu mnuChoose
End If
End Sub

Private Sub HexGrid_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If DragStart Then
    SelectGrid Index
Else
    If (FF + Index) Mod CPP = 0 Then
        HexGrid(Index).ToolTipText = "Start of Page " & Str(Int((FF + Index) / CPP) + 1)
    Else
        HexGrid(Index).ToolTipText = ""
    End If
End If
UpdateGridPointer Index
End Sub

Private Sub HexGrid_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
DragStart = False
End Sub

Private Sub SbarBox_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If FL > 0 Then
    Sbar.Enabled = True
    Sbar.Top = y - Sbar.Height \ 2
End If
End Sub

Private Sub SbarBox_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Sbar.Enabled Then
    Dim DD As Long
    DD = y - Sbar.Height \ 2
    If DD < 1 Then DD = 1
    If DD > SbarBox.Height - Sbar.Height - 40 Then DD = SbarBox.Height - Sbar.Height - 40
    Sbar.Top = DD
    DD = Int((y / SbarBox.Height) * FL)
    If DD < 0 Then DD = 1
    If DD > FL Then DD = FL
    SBar1.Panels(5).Text = "Byte: " & Str(DD)
End If
End Sub

Private Sub SbarBox_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If FL > 0 Then
    FF = (y / SbarBox.Height) * FL
    If FF < 0 Then FF = 0
    If FF > FL Then FF = FL
    UpdateGridData
End If
End Sub

Private Sub popInsert_Click(Index As Integer)
goInsert Index
End Sub

Private Sub goInsert(Index As Integer)
Dim I As Integer, J As Integer
Dim zSL As Integer, zSFP As Long
Dim CopyString As String
Dim dummy() As Byte
SL = 1
If Index = 0 Then ReDim dummy(1 To 16) Else ReDim dummy(1 To 512)
CutPaste False, dummy()
UpdateGridData
End Sub

Private Sub PrevOK_Click()
Dim I As Integer
Dim J As Byte
If PR(1).Value Then
    If Len(PRpages.Text) = 0 Then
        MsgBox ("No paged have been selected"), vbInformation, " File Editor"
        Exit Sub
    Else
        For I = 1 To Len(PRpages.Text)
            J = Asc(Mid(PRpages.Text, I, 1))
            If (J < 48 Or J > 58) And (J < 44 Or J > 45) Then
                MsgBox ("Invalid page selection"), vbInformation, " File Editor"
                Exit Sub
            End If
        Next I
    End If
End If
PreviewPrint.Visible = False
Tbar(0).Visible = True
Tbar(1).Visible = True
Tbar(2).Visible = True
Shape1(0).Visible = True
Shape1(1).Visible = True
FileLoc.Visible = True
HexFrame.Visible = True
TextFrame.Visible = True
SbarFrame.Visible = True
End Sub

Private Sub TextGrid_DblClick(Index As Integer)
DragStart = True
End Sub

Private Sub TextGrid_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
DragStart = False
End Sub

Private Sub txbGoTo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
TypeOK = False
End Sub

Private Sub txbPage_Change()
Static PageChange As Boolean
If OpenFileName <> "" And txbPage.Text <> "" And (Not PageChange) Then
    Dim PN As Long
    PN = Val(txbPage.Text)
    If PN > UBound(PrintPages) Or PN < 1 Then txbPage.BackColor = zRed Else txbPage.BackColor = zGray
    FF = (CPP * (PN - 1))
    If FF > FL Then FF = FL - 1
    If FF < 0 Then FF = 0
    PageChange = True
    UpdateGridData
    PageChange = False
End If
End Sub

Private Sub TextGrid_Click(Index As Integer)
TypeOK = True
Me.SetFocus
GridType = False
SelectGrid Index, True
End Sub

Private Sub TextGrid_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
    GridType = False
    If SL > 0 Then Me.PopupMenu mnuChoose
End If
End Sub

Private Sub TextGrid_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If DragStart Then
    SelectGrid Index
Else
    If (FF + Index) Mod CPP = 0 Then
        TextGrid(Index).ToolTipText = "Start of Page " & Str(Int((FF + Index) / CPP) + 1)
    Else
        TextGrid(Index).ToolTipText = ""
    End If
End If
UpdateGridPointer Index
End Sub

Private Sub SearchType_Click()
With SearchType
    If .Tag = 1 Then
        .Tag = 0: .Caption = "HEX"
        sCase.Tag = sCase.Value
        sCase.Value = 1
        sCase.Enabled = False
        .ToolTipText = "Click to Search TEXT"
    Else
        .Tag = 1: .Caption = "TEXT"
        sCase.Enabled = True
        sCase.Value = sCase.Tag
        .ToolTipText = "Click to Search HEX"
    End If
End With
End Sub

Private Function EvaluateHex(Num As String) As String
Dim Hok As Boolean
Dim I As Integer
Dim LS As String
Hok = True
For I = 1 To 2
    LS = Mid(UCase(Num), I, 1)
    If Asc(Left(LS, 1)) < Asc("0") Or Asc(Left(LS, 1)) > Asc("9") And _
    Asc(Left(LS, 1)) < Asc("A") Or Asc(Left(LS, 1)) > Asc("F") _
    Then Hok = False
Next I
If Hok Then EvaluateHex = UCase(Num) Else EvaluateHex = " "
End Function

Private Sub SetSbarTop(SbarLoc)
Dim Z As Long
Sbar.Top = ((SbarLoc / FL) * SbarBox.Height) - Sbar.Height \ 2
Z = SbarBox.Height - Sbar.Height - 40
If Sbar.Top > Z Then Sbar.Top = Z
If Sbar.Top < 0 Then Sbar.Top = 1
End Sub

Private Sub DEBUGtimer_Timer()
DEBUGtext(0).Caption = FF
DEBUGtext(1).Caption = SL
DEBUGtext(2).Caption = SFP
DEBUGtext(3).Caption = FL
End Sub

Private Sub txbGoTo_Change()
Static GoTochange As Boolean
If OpenFileName <> "" And txbGoTo.Text <> "" And (Not GoTochange) Then
    Dim FP As Long
    FP = Val(txbGoTo.Text)
    If FP > FL Or FP < 1 Then
        txbGoTo.BackColor = zRed
    Else
        txbGoTo.BackColor = zGray
        FF = FP - 1
        GoTochange = True
        UpdateGridData False
        SFP = (FP - FF) - 1
        SelectGrid (FP - FF) - 1
        GoTochange = False
    End If
End If
End Sub

Private Sub ClearGrid()
Dim I As Integer
For I = 0 To 511
    HexGrid(I).BackColor = zWhite
    HexGrid(I).ForeColor = zBlack
    TextGrid(I).BackColor = zWhite
    TextGrid(I).ForeColor = zBlack
Next I
End Sub

Private Sub UpdateGridData(Optional mSelect As Boolean = True)
If FL > 0 Then
    Dim I As Integer
    Dim J As Byte
    Dim K As Boolean
    Dim Z As Long
    Dim Fend As Long
    Dim Fstart As Long
    UpdatePageCount
    SetSbarTop FF
    Sbar.Enabled = False
    DoEvents
    If FF + 495 > FL Then Fstart = FL - 495 Else Fstart = FF + 1
    If Fstart < 0 Then Fstart = 1
    FF = Fstart - 1
    Fend = Fstart + 511
    If Fend > FL Then Fend = FL
    SFP = GP + FF
    If SFP > FL Then SFP = FF
    FillLabels
    For Z = Fstart To Fend
        If (Z - 1) Mod CPP = 0 Then
            TextGrid(Z - Fstart).BorderStyle = 1
            HexGrid(Z - Fstart).BorderStyle = 1
        Else
            TextGrid(Z - Fstart).BorderStyle = 0
            HexGrid(Z - Fstart).BorderStyle = 0
        End If
        J = WorkSpace(Z)
        HexGrid(Z - Fstart).Caption = FillZeroByte(J)
        TextGrid(Z - Fstart).Caption = Chr(J)
        HexGrid(Z - Fstart).Tag = ""
    Next Z
    If Z - Fstart <> 512 Then
        For I = (Z - Fstart) To 511
            HexGrid(I).Tag = ""
            HexGrid(I).Caption = ""
            TextGrid(I).Caption = ""
            TextGrid(I).BorderStyle = 0
            HexGrid(I).BorderStyle = 0
        Next I
    End If
    PGID.Caption = "Page " & Str(Int(FF / CPP) + 1)
    If mSelect Then SelectGrid SFP - FF
End If
End Sub

Private Sub txbGoTo_Click()
TypeOK = False
End Sub

Private Sub txbPage_Click()
TypeOK = False
End Sub

Private Sub txbPage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
TypeOK = False
End Sub

Private Sub txbSearch_Click()
TypeOK = False
End Sub

Private Sub txbSearch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then goSearch 0
End Sub

Private Sub PreviewPrintHex()
Dim HexData As String
Dim TextData As String
Dim J As Byte
Dim Z As Long           'File Byte counter
Dim Y1 As Integer       'Character Count
Dim S1 As String
Dim FP As Long
PofP.Caption = UBound(PrintPages) & " Pages"
S1 = Space(10)
If CPP > FL Then FP = FL Else FP = CPP
With PVtext
    'Setup printer
    .FontName = "Courier New"
    .FontSize = 4
    .FontBold = False
    .FontItalic = False
    'Print First Page header
    PVtext.Text = ""
    .SelText = vbCrLf & vbCrLf
    .SelText = S1 & "FILE: " & FSO.GetFile(OpenFileName).Name
    .SelText = Space(36) & "Page 1" & vbCrLf & vbCrLf
    .SelText = S1 & "Byte: 1 of" & Str(FL) & vbCrLf & vbCrLf
    Y1 = 0
    'Print file
    For Z = 1 To FP
        Y1 = Y1 + 1       'Character Count
        J = WorkSpace(Z)
        HexData = HexData & FillZeroByte(J) & Chr(32)
        If J < 32 Then J = 133
        TextData = TextData & Chr(J)
        If Len(TextData) > 0 And Z = FL Then Y1 = CPL
        If Y1 = 16 Then
            Y1 = 0
            .SelText = S1 & HexData
            .SelText = Space(51 - Len(HexData)) & TextData & vbCrLf
            HexData = ""
            TextData = ""
        End If
    Next Z
End With
End Sub

Private Sub PrintHex(PageStart As Long)
If OpenFileName <> "" Then
    Dim HexData As String
    Dim TextData As String
    Dim BlockStart As Long
    Dim BlockEnd As Long
    Dim J As Byte
    Dim Z As Long           'File Byte counter
    Dim Y1 As Integer       'Character Count
    
    BlockStart = (CPP * (PageStart - 1)) + 1
    BlockEnd = (CPP + BlockStart) - 1
    If BlockEnd > FL Then BlockEnd = FL
    MousePointer = 11
    SBar1.Panels(4).Text = "Printing " & FSO.GetFile(OpenFileName).Name
    With Printer
        'Setup printer
        .ScaleMode = 6
        .FontName = "Courier New"
        .FontSize = 11
        .FontBold = False
        .FontItalic = False
        .ColorMode = vbPRCMMonochrome
        'Print First Page header
        .CurrentY = 15
        .CurrentX = 25
        Printer.Print "FILE: " & FSO.GetFile(OpenFileName).Name;
        .CurrentX = 145
        Printer.Print "Page " & Str(PageStart)
        Printer.Print
        .CurrentX = 25
        Printer.Print "Byte: "; BlockStart & " of" & Str(FL)
        Printer.Print
        Y1 = 0
        'Print file
        For Z = BlockStart To BlockEnd
            Y1 = Y1 + 1       'Character Count
            J = WorkSpace(Z)
            HexData = HexData & FillZeroByte(J) & Chr(32)
            If J < 32 Then J = 133
            TextData = TextData & Chr(J)
            If Len(TextData) > 0 And Z = FL Then Y1 = CPL
            If Y1 = CPL Then
                Y1 = 0
                .CurrentX = 25
                Printer.Print HexData;
                .CurrentX = 145
                Printer.Print TextData
                HexData = ""
                TextData = ""
            End If
        Next Z
        .EndDoc
    End With
    MousePointer = 0
    SBar1.Panels(4).Text = ""
End If
End Sub

Private Sub ShowFileInfo()
If OpenFileName <> "" Then
    SBar1.Panels(1).Text = FSO.GetFile(OpenFileName).Name
    SBar1.Panels(2).Text = FSO.GetFile(OpenFileName).DateLastModified
    SBar1.Panels(3).Text = FSO.GetFile(OpenFileName).Size & " Bytes / " & FL & " Bytes "
    Form_Resize
End If
End Sub

Private Sub FillLabels()
Dim I As Byte
For I = 0 To 31: Label3(I).Caption = Format((FF + (16 * I)) + 1, "000000000"): Next I
End Sub

Private Sub CreateGrids()
Dim I As Integer, J As Integer
Dim cT1 As Long     'Array Cell Top
Dim cL1 As Long     'Array Cell Left
Dim cW As Long      'Cell Width
Dim FT As Long      'Frame Top
Const cL As Long = 100      'Cell Left
Const cT As Long = 200      'Cell Top
Const cH As Long = 240      'Cell Height
Const Thick As Long = 20    'Line Thickness

'Create, Size and place the 32 Row Labels
FT = Tbar(2).Top + Tbar(2).Height + 150
cW = 1000
cT1 = cT
For I = 0 To 31
    If I > 0 Then Load Label3(I) 'Create labels
    Label3(I).Visible = True
    Label3(I).Width = cW
    Label3(I).Height = cH
    Label3(I).Top = cT1
    Label3(I).Left = cL
    cT1 = cT1 + cH + Thick
Next I
FillLabels
FileLoc.Top = FT
FileLoc.Left = 100
FileLoc.Height = Label3(31).Top + Label3(31).Height + 100
FileLoc.Width = Label3(31).Left + Label3(31).Width + 100

'Create, Size and place the 512 Hex and Text grids
cW = 300
cT1 = cT
For I = 0 To 31
    cL1 = cL
    For J = I * 16 To I * 16 + 15
        If J > 0 Then Load HexGrid(J) 'Create hex labels
        If J > 0 Then Load TextGrid(J) 'Create text labels
        HexGrid(J).Visible = True: TextGrid(J).Visible = True
        HexGrid(J).Caption = "": TextGrid(J).Caption = ""
        HexGrid(J).Width = cW: TextGrid(J).Width = cW
        HexGrid(J).Height = cH: TextGrid(J).Height = cH
        HexGrid(J).Top = cT1: TextGrid(J).Top = cT1
        HexGrid(J).Left = cL1: TextGrid(J).Left = cL1
        cL1 = cL1 + cW + Thick
    Next J
    cT1 = cT1 + cH + Thick
Next I
HexGrid(0).Caption = "00"
HexFrame.Top = FT
HexFrame.Left = FileLoc.Width + FileLoc.Left + 100
HexFrame.Height = HexGrid(511).Top + HexGrid(511).Height + 100
HexFrame.Width = HexGrid(511).Left + HexGrid(511).Width + 150

'Size and position the Scroll bar
cL1 = HexFrame.Left + HexFrame.Width + 90

SbarFrame.Left = cL1
SbarFrame.Top = HexFrame.Top: SbarFrame.Height = HexFrame.Height
cmdScroll(2).Top = cmdScroll(0).Top + cmdScroll(0).Height
SbarBox.Top = cmdScroll(2).Top + cmdScroll(2).Height + 40
cmdScroll(1).Top = SbarFrame.Height - cmdScroll(1).Height
cmdScroll(3).Top = cmdScroll(1).Top - cmdScroll(1).Height
SbarBox.Height = (cmdScroll(3).Top - SbarBox.Top) - 50

'Position the text frame
PGID.Left = TextGrid(15).Left + TextGrid(15).Width + 100
PGID.Top = TextGrid(15).Top
TextFrame.Top = FT
TextFrame.Left = SbarFrame.Left + SbarFrame.Width + 100
TextFrame.Height = TextGrid(511).Top + TextGrid(511).Height + 100
TextFrame.Width = PGID.Left + PGID.Width + 150

'Position the Print Preview frame
PreviewPrint.Top = Me.Top + 400
PreviewPrint.Width = HexFrame.Width
PVtext.Top = 300:
PVtext.Width = 3800
PVtext.Height = 5800
PVtext.Left = (PreviewPrint.Width / 2) - (PVtext.Width / 2)
PofP.Top = PVtext.Height + PVtext.Top + 100
PofP.Left = PVtext.Left
PofP.Width = PVtext.Width
Prange.Top = PofP.Height + PofP.Top + 200
Prange.Left = PVtext.Left
Prange.Width = PVtext.Width
PrevOK.Left = (Prange.Width / 2 + Prange.Left) - (PrevOK.Width / 2)
PrevOK.Top = Prange.Top + Prange.Height + 200
PreviewPrint.Height = PrevOK.Top + PrevOK.Height + 100

'Position the Legend Stuff
Label7(0).Left = PGID.Left: Label7(0).Top = HexGrid(15 * 9).Top
For I = 1 To 4
    Label7(I).Left = PGID.Left
    Label7(I).Top = Label7(I - 1).Top + Label7(I - 1).Height + 100
Next I
Sbar.Enabled = True: Sbar.Top = 1: Sbar.Enabled = False
ShowGrids False
End Sub

Private Sub ShowGrids(GridsShow As Boolean)
Tbar(1).Enabled = GridsShow: Tbar(2).Enabled = GridsShow
FileLoc.Visible = GridsShow: SbarFrame.Visible = GridsShow
HexFrame.Visible = GridsShow: TextFrame.Visible = GridsShow
cmdBut(2).Enabled = GridsShow
End Sub

Private Function FillZeroByte(DecNum As Byte) As String
Dim rL As String
rL = Hex(DecNum)
Do Until Len(rL) >= 2
    rL = "0" & rL
Loop
FillZeroByte = rL
End Function

Private Function FillZeroLong(DecNum As Long) As String
Dim rL As String
rL = Hex(DecNum)
Do Until Len(rL) >= 6
    rL = "0" & rL
Loop
FillZeroLong = rL
End Function

Private Sub UpdateGridPointer(GridNo As Integer)
SBar1.Panels(5).Text = "Byte: " & Str(FF + GridNo + 1)
End Sub

Private Sub SelectGrid(GridNo As Integer, Optional mSel As Boolean = False)
'GridType-TRUE=HEX grid, GridType-FALSE=TEXT grid
Dim I As Integer
Dim Lnum As String
Dim HGBC As Long 'Hex grid back color
Dim TGBC As Long 'Text grid back color
Dim HGFC As Long 'Hex grid fore color
Dim TGFC As Long 'Text grid fore color
Dim HS As Long, HF As Long
If GridType Then
    HGBC = zBlack: TGBC = zGray
    HGFC = zWhite: TGFC = zBlack
Else
    HGBC = zGray: TGBC = zBlack
    HGFC = zBlack: TGFC = zWhite
End If

If (GetKeyInput(KEY_SHIFT) < 0 And mSel) Or DragStart Then
    If SFP >= 0 And FF + GridNo <= FL - 1 Then
        If GridNo > GP Then
            SL = (GridNo - GP) + 1
            SFP = GP + FF
        Else
            SL = (GP - GridNo) + 1
            SFP = GridNo + FF
        End If
        If SFP + SL <= FL Then
            HS = SFP - FF: HF = (SFP - FF + SL) - 1
            For I = 0 To 511
                If I >= HS And I <= HF Then
                    HexGrid(I).BackColor = HGBC
                    HexGrid(I).ForeColor = HGFC
                    TextGrid(I).BackColor = TGBC
                    TextGrid(I).ForeColor = TGFC
                Else
                    HexGrid(I).BackColor = zWhite
                    HexGrid(I).ForeColor = zBlack
                    TextGrid(I).BackColor = zWhite
                    TextGrid(I).ForeColor = zBlack
                End If
            Next I
            If GridType And SL < 9 Then
                For I = HF To HS Step -1
                    Lnum = Lnum & HexGrid(I).Caption
                Next I
                Select Case SL
                    Case 1: SBar1.Panels(4).Text = CByte("&h" & Lnum)
                    Case 2: SBar1.Panels(4).Text = Format(CInt("&h" & Lnum), "##,###")
                    Case 3, 4: SBar1.Panels(4).Text = Format(CLng("&h" & Lnum), "#,###,###,###")
                    Case 5, 6, 7, 8: SBar1.Panels(4).Text = Hex2Dec(Lnum)
                End Select
            Else
                SBar1.Panels(4).Text = ""
            End If
        End If
    End If
Else
    If GridNo + FF <= FL Then
        ClearGrid
        SL = 1
        SFP = GridNo + FF: GP = GridNo
        HexGrid(GP).BackColor = HGBC
        HexGrid(GP).ForeColor = HGFC
        TextGrid(GP).BackColor = TGBC
        TextGrid(GP).ForeColor = TGFC
    End If
    SBar1.Panels(4).Text = ""
End If
End Sub

Private Sub goSearch(Index As Integer)
'Index-0=Start, Index-1=Next
Dim sMatch As Boolean
Dim STXT As String
Dim arrhexbyte() As Byte
Dim CaseSel As Byte
Dim HexCtn As Integer
Dim SP As Long
Dim I As Long
Dim J As Integer
Dim EV As Long
On Error Resume Next
EV = 32768
If SP = 0 Then SP = 1
STXT = txbSearch.Text
If txbSearch.Text = "" Or FF > FL Then Exit Sub
If Index = 0 Then
    SP = 0
Else
    If SFP > 0 Then SP = SFP + 1 Else SP = FF + 1
End If
If sCase Then CaseSel = 0 Else CaseSel = 32

If SearchType.Tag = 0 Then  'Search HEX
    HexCtn = Len(STXT) / 2
    ReDim arrhexbyte(1 To HexCtn)
    For I = 1 To HexCtn: arrhexbyte(I) = CByte("&h" & (Mid(STXT, (I * 2 - 1), 2))): Next I
    If Err Then
        MsgBox ("Invalid search type, try TEXT"), vbExclamation, "Bad Search Type"
        Exit Sub
    End If
Else                        'Search TEXT
    HexCtn = Len(STXT)
    ReDim arrhexbyte(1 To HexCtn)
    For I = 1 To HexCtn: arrhexbyte(I) = CByte(Asc(Mid(STXT, (I), 1))): Next I
End If

MousePointer = 11
SBar1.Panels(4).Text = "Searching for " & STXT & "....."
For I = SP To UBound(WorkSpace) - HexCtn + 1
    SbarPos = I
    If WorkSpace(I + 1) = arrhexbyte(1) Or _
    (WorkSpace(I + 1) Xor CaseSel) = arrhexbyte(1) Then
        sMatch = True
        ' Compare the rest of the bytes
        For J = 1 To (HexCtn) - 1
            If WorkSpace(I + J + 1) <> arrhexbyte(J + 1) And _
            (WorkSpace(I + J + 1) Xor CaseSel) <> arrhexbyte(J + 1) Then
                sMatch = False
                Exit For
            End If
        Next J
    End If
    If sMatch Then Exit For
    If I Mod EV = 0 Then SetSbarTop SbarPos: DoEvents
Next I
If sMatch Then
    FF = I: GP = 0
    ClearGrid
    UpdateGridData False
    SFP = I
    SelectGrid SFP - FF
    For I = I - FF To (I - FF) + (HexCtn - 1)
        HexGrid(I).BackColor = zAqua
        HexGrid(I).ForeColor = zBlack
        TextGrid(I).BackColor = zAqua
        TextGrid(I).ForeColor = zBlack
    Next I
    
Else
    MsgBox ("Reached end of file"), vbInformation, "Search Complete"
End If
MousePointer = 0
SBar1.Panels(4).Text = ""
End Sub

Private Sub txbSearch_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
TypeOK = False
End Sub

Private Function Hex2Dec(ByVal HexString As String) As Variant
  Dim x As Integer
  Dim BinStr As String
  If Left$(HexString, 2) Like "&[hH]" Then
    HexString = Mid$(HexString, 3)
  End If
  If Len(HexString) <= 23 Then
    Const BinValues = "0000000100100011" & _
                      "0100010101100111" & _
                      "1000100110101011" & _
                      "1100110111101111"
    For x = 1 To Len(HexString)
      BinStr = BinStr & Mid$(BinValues, 4 * Val("&h" & Mid$(HexString, x, 1)) + 1, 4)
    Next
    Hex2Dec = CDec(0)
    For x = 0 To Len(BinStr) - 1
      Hex2Dec = Hex2Dec + Val(Mid(BinStr, _
                Len(BinStr) - x, 1)) * 2 ^ x
    Next
  Else
    ' Number is too big, handle error here
  End If
End Function

