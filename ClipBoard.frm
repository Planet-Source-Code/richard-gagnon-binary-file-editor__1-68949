VERSION 5.00
Begin VB.Form MyClip 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Clip Board"
   ClientHeight    =   480
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   1035
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleWidth      =   1035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Label TextGrid 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   9
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   750
   End
End
Attribute VB_Name = "MyClip"
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

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Sub CreateGrids()
Dim I As Integer, J As Integer
Dim cT1 As Long     'Array Cell Top
Dim cL As Long      'Cell Left
Dim cW As Long      'Cell Width
Const cT As Long = 100      'Cell Top
Const cH As Long = 200      'Cell Height
Const Thick As Long = 20    'Line Thickness

'Create, Size and place the 32 Row Labels
cW = 600: cL = 40
'Create, Size and place the 512 Text grids
cW = 300
cT1 = cT
For I = 0 To 31
    cL = 40
    For J = I * 16 To I * 16 + 15
        If J > 0 Then Load TextGrid(J) 'Create labels
        TextGrid(J).Visible = True
        TextGrid(J).Caption = ""
        TextGrid(J).Width = cW
        TextGrid(J).Height = cH
        TextGrid(J).Top = cT1
        TextGrid(J).Left = cL
        cL = cL + cW + Thick
    Next J
    cT1 = cT1 + cH + Thick
Next I
Me.Width = TextGrid(15).Left + TextGrid(15).Width + 170
Me.Height = TextGrid(511).Top + TextGrid(511).Height + 450
End Sub

Private Sub Form_Load()
Dim MC As Long
CreateGrids
ImportData
MC = SetWindowPos(Me.hwnd, -1, 0, 0, 0, 0, 2 Or 1)
End Sub
Private Sub ImportData()
'Dim CopyString As String
'CopyString = Clipboard.GetText()
'If CopyString <> "" Then
'    Dim I As Integer, X As Integer
'    For I = 1 To Len(CopyString) Step 2
'        TextGrid(X).Caption = Chr("&h" & Mid(CopyString, I, 2))
'        X = X + 1
'    Next I
'End If
End Sub
Private Function FillZeroLong(DecNum As Long) As String
Dim rL As String
rL = Hex(DecNum)
Do Until Len(rL) >= 6
    rL = "0" & rL
Loop
FillZeroLong = rL
End Function
