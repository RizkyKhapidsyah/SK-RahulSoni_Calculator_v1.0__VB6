VERSION 5.00
Begin VB.Form frmRealCalc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4155
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4305
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Realcalc.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShortCut 
      Caption         =   "ShortCuts"
      DownPicture     =   "Realcalc.frx":030A
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6000
      TabIndex        =   42
      ToolTipText     =   "Hides the help text box"
      Top             =   240
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Hide"
      DownPicture     =   "Realcalc.frx":074C
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      TabIndex        =   41
      ToolTipText     =   "Hides the help text box"
      Top             =   240
      Width           =   1605
   End
   Begin VB.TextBox Text1 
      Height          =   3285
      Left            =   4320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   40
      Text            =   "Realcalc.frx":0B8E
      Top             =   780
      Width           =   3285
   End
   Begin VB.Frame Frame2 
      Height          =   1245
      Left            =   120
      TabIndex        =   26
      Top             =   2820
      Width           =   4095
      Begin VB.CommandButton cmdXPowerY 
         Caption         =   "x^y"
         DownPicture     =   "Realcalc.frx":0CC0
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2100
         TabIndex        =   38
         ToolTipText     =   "(Use PAD to input a number) -> Click -> (Select the power to be raised) -> Press ""="""
         Top             =   720
         Width           =   550
      End
      Begin VB.CommandButton cmdSquareRoot 
         Caption         =   "Ö"
         DownPicture     =   "Realcalc.frx":1102
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   37
         ToolTipText     =   "(Use PAD to input a number) -> Click to get the SQUARE ROOT of the number"
         Top             =   240
         Width           =   550
      End
      Begin VB.CommandButton cmdMPlus 
         Caption         =   "M+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   750
         MaskColor       =   &H80000009&
         TabIndex        =   36
         ToolTipText     =   "(Use PAD to input a number) -> Click to Add in Memory"
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmdMemoryRecall 
         Caption         =   "MR"
         DownPicture     =   "Realcalc.frx":1544
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2790
         TabIndex        =   35
         ToolTipText     =   "Memory Recall - Displays the memory content"
         Top             =   240
         Width           =   550
      End
      Begin VB.CommandButton cmdPi 
         Caption         =   "p"
         BeginProperty Font 
            Name            =   "Symbol"
            Size            =   14.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3450
         TabIndex        =   34
         ToolTipText     =   "Show the value of PI"
         Top             =   240
         Width           =   550
      End
      Begin VB.CommandButton cmdMminus 
         Caption         =   "M-"
         DownPicture     =   "Realcalc.frx":1986
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1440
         TabIndex        =   33
         ToolTipText     =   "(Use PAD to input a number) -> Click to Subtract from Memory"
         Top             =   240
         Width           =   550
      End
      Begin VB.CommandButton cmdMC 
         Caption         =   "MC"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2100
         TabIndex        =   32
         ToolTipText     =   "Clear the Memory"
         Top             =   240
         Width           =   550
      End
      Begin VB.CommandButton cmdFactorial 
         Caption         =   "n!"
         DownPicture     =   "Realcalc.frx":1DC8
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   31
         ToolTipText     =   "(Use PAD to input a number) -> Click to find factorial of the number ( 1<= X <= 170 )"
         Top             =   720
         Width           =   550
      End
      Begin VB.CommandButton cmdSquare 
         Caption         =   "x^2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   750
         MaskColor       =   &H80000009&
         TabIndex        =   30
         ToolTipText     =   "(Use PAD to input a number) -> Click to Find square of a number"
         Top             =   720
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmdIntegerDivision 
         Caption         =   "\"
         DownPicture     =   "Realcalc.frx":220A
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2790
         TabIndex        =   29
         ToolTipText     =   "Integer Division - (Use PAD to input a number) -> Click -> (Select Divisor) -> Press ""="" sign"
         Top             =   720
         Width           =   550
      End
      Begin VB.CommandButton cmdMod 
         Caption         =   "Mod"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3450
         TabIndex        =   28
         ToolTipText     =   "Returns only remainder of a division - (Use PAD to input a number) -> Click -> (Select another number) -> Press ""="" sign"
         Top             =   720
         Width           =   550
      End
      Begin VB.CommandButton cmdCube 
         Caption         =   "x^3"
         DownPicture     =   "Realcalc.frx":264C
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1440
         TabIndex        =   27
         ToolTipText     =   "(Use PAD to input a number) -> Click to Find cube of a number"
         Top             =   720
         Width           =   550
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2145
      Left            =   120
      TabIndex        =   25
      Top             =   660
      Width           =   4095
      Begin VB.CommandButton cmdBackSpace 
         Caption         =   "ç"
         DownPicture     =   "Realcalc.frx":2A8E
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   11.25
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3450
         TabIndex        =   20
         ToolTipText     =   "Delete last digit"
         Top             =   225
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1410
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1635
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   735
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   2
         Left            =   1410
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   3
         Left            =   2070
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   210
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   7
         Left            =   735
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1155
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   8
         Left            =   1410
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1155
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   9
         Left            =   2070
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1155
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmdDecimal 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2070
         TabIndex        =   11
         Top             =   1635
         Width           =   550
      End
      Begin VB.CommandButton cmdAC 
         Caption         =   "&AC"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   16
         ToolTipText     =   "All Clear"
         Top             =   225
         Width           =   550
      End
      Begin VB.CommandButton cmdPlus 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2775
         TabIndex        =   15
         ToolTipText     =   "(Use PAD to input a number) -> Click -> (Select another number) -> Press ""="" sign"
         Top             =   1635
         Width           =   550
      End
      Begin VB.CommandButton cmdMinus 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2775
         TabIndex        =   14
         ToolTipText     =   "(Use PAD to input a number) -> Click -> (Select another number) -> Press ""="" sign"
         Top             =   1155
         Width           =   550
      End
      Begin VB.CommandButton cmdDivide 
         Caption         =   "÷"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2775
         TabIndex        =   12
         ToolTipText     =   "(Use PAD to input a number) -> Click -> (Select another number) -> Press ""="" sign"
         Top             =   225
         Width           =   550
      End
      Begin VB.CommandButton cmdPlusMinus 
         Caption         =   "± "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3450
         TabIndex        =   22
         ToolTipText     =   "(Use PAD to input a number) -> Click"
         Top             =   1155
         Width           =   550
      End
      Begin VB.CommandButton cmdEquals 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3450
         TabIndex        =   24
         ToolTipText     =   "Display the results"
         Top             =   1635
         Width           =   550
      End
      Begin VB.CommandButton cmdOff 
         Caption         =   "&Off"
         Height          =   390
         Left            =   90
         TabIndex        =   18
         ToolTipText     =   "Close this calculator"
         Top             =   1155
         Width           =   550
      End
      Begin VB.CommandButton cmdShowHide 
         Caption         =   "ô"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   19
         ToolTipText     =   "Show / Hide extra functions..."
         Top             =   1635
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   10
         Left            =   735
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1635
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton cmdC 
         Caption         =   "&C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   90
         TabIndex        =   17
         ToolTipText     =   "Clear Last"
         Top             =   690
         Width           =   550
      End
      Begin VB.CommandButton cmdPercent 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   3450
         TabIndex        =   21
         ToolTipText     =   "(Use PAD to input a number) -> (Use any operator) -> (Select another number) -> Click"
         Top             =   690
         Width           =   550
      End
      Begin VB.CommandButton cmdMultiply 
         Caption         =   "x"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2775
         TabIndex        =   13
         ToolTipText     =   "(Use PAD to input a number) -> Click -> (Select another number) -> Press ""="" sign"
         Top             =   690
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   6
         Left            =   2070
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   690
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   5
         Left            =   1410
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   690
         UseMaskColor    =   -1  'True
         Width           =   550
      End
      Begin VB.CommandButton digits 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   4
         Left            =   735
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   690
         UseMaskColor    =   -1  'True
         Width           =   550
      End
   End
   Begin VB.Label lblMemory 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   390
      Left            =   150
      TabIndex        =   39
      ToolTipText     =   "Displays the Memory status... (Click to get the value of memory)"
      Top             =   240
      Width           =   330
   End
   Begin VB.Label Display 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   390
      Left            =   540
      TabIndex        =   23
      Top             =   240
      Width           =   3690
   End
   Begin VB.Menu mExit 
      Caption         =   "&Exit"
   End
   Begin VB.Menu mEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu mHelp 
      Caption         =   "&Help"
      Begin VB.Menu mShowHide 
         Caption         =   "&Show/Hide"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mShortcuts 
         Caption         =   "S&hortcuts"
      End
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmRealCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'[][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]
'Calculator Project
'
'This software is FREEWARE. You may use it as you see fit for your own
'projects but you may not re-sell the original or the source code. Do
'not copy this sample to a collection, such as a CD-ROM archive.
'
'No warranty express or implied, is given as to the use of this
'program. Use at your own risk.
'
'I have a lot of ideas to work upon and I am seeking to make a good
'circle of online friends who could work at tandem to produce something
'good for everybody and make some money for them too...

'I want to make contacts with lots of people and I would be thankful if
'anyone of you viewing this code drops me a line stating the bugs,
'modifications or anything else including your friendship hand :-)
'
'Copyright © Jan 2002 - Rahul Soni
'
'email  - rahul1717@yahoo.com
'Tel    - 0091-612-427881
'         IAS Colony, Bailey Road
'         Patna - 801503 (Bihar)
'         India
'
'Happy Programming, Enjoy !!!!
'[][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][][]

Dim DotPushed As Boolean
Dim OperatorPushed As Boolean
Dim Operand1 As Double
Dim Operand2 As Double
Dim Operation As String
Dim LOGO As String
Dim EqualPressed As Boolean
Dim MemoryVar As Double

Private Sub cmdAC_Click()
    Display.Caption = "0."
    DotPushed = False
    Operand1 = 0
    Operand2 = 0
    MemoryVar = 0
    Operation = ""
    lblMemory.Caption = ""
End Sub

Private Sub cmdBackSpace_Click()
    If Display.Caption = "0." Then
        DotPushed = False
        Exit Sub
    End If
        
    If Len(Display.Caption) = 1 Then
        DotPushed = False
        Display.Caption = "0."
    Else
        With Display
            .Caption = Left(.Caption, Len(.Caption) - 1)
        End With
    End If

End Sub

Private Sub cmdC_Click()
    Display.Caption = "0."
End Sub

Private Sub cmdCube_Click()
On Error GoTo ErrorHandler
    Display.Caption = Val(Display.Caption) ^ 3

Exit Sub
ErrorHandler:
MsgBox Err.Description, vbCritical, LOGO
Display.Caption = "Error"
End Sub

Private Sub cmdDecimal_Click()
    DotPushed = True
    If InStr(Display.Caption, ".") = 0 Then
        Display.Caption = Display.Caption + "."
    End If
    If OperatorPushed Or EqualPressed Then
        OperatorPushed = False
        Display.Caption = "0."
    End If
End Sub

Private Sub cmdDivide_Click()
    OperatorPushed = True
    Operand1 = Val(Display.Caption)
    Operation = "/"
End Sub

Private Sub cmdEquals_Click()
On Error GoTo ErrorHandler

    If Not EqualPressed Then
        Operand2 = Val(Display.Caption)
    Else
        Operand1 = Val(Display.Caption)
    End If
    Select Case Operation
    Case "+": Display.Caption = Operand1 + Operand2
    Case "-": Display.Caption = Operand1 - Operand2
    Case "x": Display.Caption = Operand1 * Operand2
    Case "/":
        If Operand2 = 0 Then
            Display.Caption = "Error"
        Else
            Display.Caption = Operand1 / Operand2
        End If
    Case "^": Display.Caption = Operand1 ^ Operand2
    Case "\":
        If Operand2 = 0 Then
            Display.Caption = "Error"
        Else
            Display.Caption = Int(Operand1 / Operand2)
        End If
    Case "Mod":
        If Operand2 = 0 Then
            Display.Caption = "Error"
        Else
            Display.Caption = Operand1 - (Int(Operand1 / Operand2) * Operand2)
        End If
    End Select
    
    OperatorPushed = False
    DotPushed = False
    EqualPressed = True
    Exit Sub
    
ErrorHandler:
    MsgBox Err.Description, vbCritical, LOGO
    Display.Caption = "Error"
End Sub

Private Sub cmdFactorial_Click()
    OperatorPushed = False
    Operation = ""
    EqualPressed = True
    If Val(Display.Caption) > 170 Or Val(Display.Caption) <= 0 Then
        MsgBox "Sorry, Factorial of numbers only between 1 & 170 is possible in this calculator !!", vbCritical, LOGO
        Exit Sub
    End If
    Display.Caption = Factorial(Val(Display.Caption))
End Sub

Private Sub cmdIntegerDivision_Click()
    OperatorPushed = True
    Operand1 = Val(Display.Caption)
    Operation = "\"
End Sub

Private Sub cmdMC_Click()
    lblMemory.Caption = ""
    MemoryVar = 0
End Sub

Private Sub cmdMemoryRecall_Click()
    Display.Caption = MemoryVar
    OperatorPushed = True
End Sub

Private Sub cmdMinus_Click()
    OperatorPushed = True
    Operand1 = Val(Display.Caption)
    Operation = "-"
End Sub

Private Sub cmdMminus_Click()
    lblMemory.Caption = "M"
    MemoryVar = MemoryVar - Val(Display.Caption)
    OperatorPushed = True
End Sub

Private Sub cmdMod_Click()
    OperatorPushed = True
    Operand1 = Val(Display.Caption)
    Operation = "Mod"
End Sub

Private Sub cmdMPlus_Click()
    lblMemory.Caption = "M"
    MemoryVar = MemoryVar + Val(Display.Caption)
    OperatorPushed = True
End Sub

Private Sub cmdMultiply_Click()
    OperatorPushed = True
    Operand1 = Val(Display.Caption)
    Operation = "x"
End Sub

Private Sub cmdOff_Click()
    End
End Sub

Private Sub cmdPercent_Click()
On Error GoTo ErrorHandler
    OperatorPushed = False
    
    Operand2 = Val(Display.Caption)
    
    Select Case Operation
    Case "x": Display.Caption = Operand1 * Operand2 / 100
    Case "+": Display.Caption = Operand1 + (Operand2 / 100) * Operand1
    Case "-": Display.Caption = Operand1 - (Operand2 / 100) * Operand1
    Case "/": Display.Caption = Operand1 / ((Operand2 / 100) * Operand1)
    End Select
    
    Operation = ""
    EqualPressed = True
    Exit Sub
    
ErrorHandler:
MsgBox Err.Description, vbCritical, LOGO
Display.Caption = "Error"

End Sub

Private Sub cmdPi_Click()
    Display.Caption = 3.14159265358979
End Sub

Private Sub cmdPlus_Click()
    OperatorPushed = True
    Operand1 = Val(Display.Caption)
    Operation = "+"
End Sub

Private Sub cmdPlusMinus_Click()
    Display.Caption = -Display.Caption
End Sub

Private Sub cmdShortCut_Click()
MsgBox "The shortcuts used in this Calculator are :- " & vbCrLf & vbCrLf & _
    "All Clear = Esc" & vbCrLf & vbCrLf & _
    "End = Shift+Esc " & vbCrLf & vbCrLf & _
    "Delete Last Digit = Backspace" & vbCrLf & vbCrLf & _
    "Integer Division = \" & vbCrLf & vbCrLf & _
    "M+ = F5" & vbCrLf & vbCrLf & _
    "M- = F6" & vbCrLf & vbCrLf & _
    "MC = F7" & vbCrLf & vbCrLf & _
    "MR = F8" & vbCrLf & vbCrLf & _
    "Pi = F10" & vbCrLf & vbCrLf & _
    "Factorial = Shift + 1" & vbCrLf & vbCrLf & _
    "Square = Shift + 2" & vbCrLf & vbCrLf & _
    "Cube = Shift + 3" & vbCrLf & vbCrLf & _
    "Plus Minus = P" & vbCrLf & vbCrLf & _
    "Square Root = S" & vbCrLf & vbCrLf & _
    "Mod = M" & vbCrLf & vbCrLf & _
    "Show Result = Shift + Enter" & vbCrLf & vbCrLf, vbInformation, LOGO
End Sub

Private Sub cmdShowHide_Click()
If Me.Height = 3555 Then
    Me.Height = 4815
Else
    Me.Height = 3555
End If
End Sub

Private Sub cmdSquare_Click()
On Error GoTo ErrorHandler
    
    Display.Caption = Val(Display.Caption) ^ 2
    Exit Sub

ErrorHandler:
MsgBox Err.Description, vbCritical, LOGO
Display.Caption = "Error"
End Sub

Private Sub cmdSquareRoot_Click()
On Error GoTo ErrorHandler

If Val(Display.Caption) >= 0 Then
    Display.Caption = Sqr(Val(Display.Caption))
Else
    MsgBox "Can't calculate Square root of a Negative number!!", vbCritical, LOGO
End If
Exit Sub

ErrorHandler:
MsgBox Err.Description, vbCritical, LOGO
Display.Caption = "Error"

End Sub

Private Sub cmdXPowerY_Click()
    OperatorPushed = True
    Operand1 = Val(Display.Caption)
    Operation = "^"

End Sub

Private Sub Command1_Click()
    Me.Width = 4400
End Sub

Private Sub digits_Click(Index As Integer)
    If OperatorPushed Or Display.Caption = "Error" Or EqualPressed Then
        Display.Caption = "0."
        OperatorPushed = False
        EqualPressed = False
    End If
    
    If Val(Display.Caption) > 10 ^ 20 Then Exit Sub
    
    If (Index = 0 Or Index = 10) And Val(Display.Caption) = 0 Then
        If Not DotPushed Then
            Exit Sub
        Else
            Display.Caption = Display.Caption + Trim(digits(Index).Caption)
            Exit Sub
        End If
    End If

    If Display.Caption = "0." And Not DotPushed Then
        Display.Caption = Trim(digits(Index).Caption)
    Else
        Display.Caption = Display.Caption + Trim(digits(Index).Caption)
    End If
    
    
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 1 Then
        End
    ElseIf KeyCode = vbKeyEscape Then
        cmdAC_Click
    ElseIf KeyCode >= vbKey0 And KeyCode <= vbKey9 And Shift = 0 Then
        digits_Click (Chr(KeyCode))
    ElseIf KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9 Then
        digits_Click (Chr(KeyCode - 48))    'Convert to corresponding 0 (See MSDN for keycodes)
    ElseIf KeyCode = vbKeyBack Then
        cmdBackSpace_Click
    ElseIf (KeyCode = 187 And Shift = 1) Or KeyCode = 107 Then
        cmdPlus_Click
    ElseIf KeyCode = 189 Or KeyCode = 109 Then
        cmdMinus_Click
    ElseIf KeyCode = 191 Or KeyCode = 111 Then
        cmdDivide_Click
    ElseIf (KeyCode = 56 And Shift = 1) Or KeyCode = 106 Then
        cmdMultiply_Click
    ElseIf KeyCode = 220 Then               'Shortcut="\"
        cmdIntegerDivision_Click
    ElseIf KeyCode = vbKeyF5 Then           'Shortcut=F5
        cmdMPlus_Click
    ElseIf KeyCode = vbKeyF6 Then           'Shortcut=F6
        cmdMminus_Click
    ElseIf KeyCode = vbKeyF7 Then           'Shortcut=F7
        cmdMC_Click
    ElseIf KeyCode = vbKeyF8 Then           'Shortcut=F8
        cmdMemoryRecall_Click
    ElseIf KeyCode = vbKeyF10 Then          'Shortcut=F10
        cmdPi_Click
    ElseIf KeyCode = 49 And Shift = 1 Then  'Shortcut=Shift+"1"
        cmdFactorial_Click
    ElseIf KeyCode = 50 And Shift = 1 Then  'Shortcut=Shift+"2"
        cmdSquare_Click
    ElseIf KeyCode = 51 And Shift = 1 Then  'Shortcut=Shift+"3"
        cmdCube_Click
    ElseIf KeyCode = 80 Then    'Shortcut=P
        cmdPlusMinus_Click
    ElseIf KeyCode = 83 Then    'Shortcut=S
        cmdSquareRoot_Click
    ElseIf KeyCode = 77 Then    'Shortcut=M
        cmdMod_Click
    ElseIf KeyCode = 13 And Shift = 1 Then  'Press Shift+Enter to display result
        cmdEquals_Click
    End If

End Sub

Private Sub Form_Load()
'Make the height equal to 3555 if you want the calculator to start in a compact mode
    frmRealCalc.KeyPreview = True
    LOGO = "Calculator"
    Me.Height = 4815    '=3555 for compact mode open
    Me.Width = 4400
    Me.Top = (Screen.Height - Me.Height) / 3
    Me.Left = (Screen.Width - Me.Width) / 3
    DotPushed = False
    MemoryVar = 0
End Sub

Private Sub lblMemory_Click()
    MsgBox MemoryVar, vbInformation, LOGO
End Sub

Private Sub mAbout_Click()
    About.Show 1
End Sub

Private Sub mCopy_Click()
    Clipboard.Clear
    Clipboard.SetText Val(Display.Caption)
End Sub

Private Sub mExit_Click()
    End
End Sub

Function Factorial(NUM As Double) As Double
If NUM = 1 Then
    Factorial = NUM
Else
    Factorial = NUM * Factorial(NUM - 1)
End If
End Function

Private Sub mPaste_Click()
    Display.Caption = Val(Clipboard.GetText)
End Sub

Private Sub mShortcuts_Click()
    cmdShortCut_Click
End Sub

Private Sub mShowHide_Click()
    If Me.Width = 4400 Then
        Me.Width = 7800
    Else
        Me.Width = 4400
    End If
End Sub
