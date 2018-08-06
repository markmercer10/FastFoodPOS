VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDisplayPanels 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Display Panels"
   ClientHeight    =   8895
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8895
   ScaleWidth      =   19635
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox BlockBG2 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   17640
      ScaleHeight     =   315
      ScaleWidth      =   795
      TabIndex        =   60
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox BlockGradient 
      Caption         =   "Gradient"
      Height          =   255
      Left            =   18000
      TabIndex        =   59
      Top             =   1150
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame TabContent 
      BorderStyle     =   0  'None
      Height          =   6855
      Index           =   3
      Left            =   14760
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   4935
      Begin VB.ComboBox tFont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox tFontSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1440
         TabIndex        =   53
         Text            =   "12"
         Top             =   1680
         Width           =   975
      End
      Begin VB.CheckBox tFontBold 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1440
         TabIndex        =   52
         Top             =   2100
         Width           =   1215
      End
      Begin VB.PictureBox tFontColor 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   1440
         ScaleHeight     =   315
         ScaleWidth      =   795
         TabIndex        =   51
         Top             =   2580
         Width           =   855
      End
      Begin VB.TextBox blockText 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   50
         Text            =   "Text"
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label20 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   58
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label19 
         Caption         =   "Font Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   57
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label18 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   56
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   55
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Text"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   49
         Top             =   120
         Width           =   3975
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   19080
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer LoadTimer 
      Interval        =   50
      Left            =   240
      Top             =   840
   End
   Begin VB.CommandButton ButnSave 
      BackColor       =   &H0080FFFF&
      Caption         =   "Save Changes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   11280
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   120
      Width           =   2535
   End
   Begin VB.Timer DrawTimer 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   720
      Top             =   840
   End
   Begin VB.ComboBox cboRes 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmDisplayPanels.frx":0000
      Left            =   5640
      List            =   "frmDisplayPanels.frx":0010
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox BlockBG 
      BackColor       =   &H00000000&
      Height          =   495
      Left            =   16920
      ScaleHeight     =   435
      ScaleWidth      =   915
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox labColor 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14760
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Background Color"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame TabContent 
      BorderStyle     =   0  'None
      Height          =   6855
      Index           =   0
      Left            =   1200
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   4935
      Begin VB.CheckBox hStrokeBool 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   67
         Top             =   3400
         Width           =   375
      End
      Begin VB.PictureBox hStrokeColor 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   2280
         ScaleHeight     =   315
         ScaleWidth      =   795
         TabIndex        =   66
         Top             =   3420
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox hBackColor2 
         BackColor       =   &H00FF0000&
         Height          =   375
         Left            =   3840
         ScaleHeight     =   315
         ScaleWidth      =   795
         TabIndex        =   64
         Top             =   3900
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.PictureBox hBackColor1 
         BackColor       =   &H00FFFF00&
         Height          =   375
         Left            =   2640
         ScaleHeight     =   315
         ScaleWidth      =   795
         TabIndex        =   63
         Top             =   3900
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CheckBox hBackBool 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2280
         TabIndex        =   61
         Top             =   3880
         Width           =   375
      End
      Begin VB.TextBox Margin 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   40
         Text            =   "10"
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox SectionIDs 
         Height          =   255
         Left            =   3240
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox Section 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   120
         Width           =   2775
      End
      Begin VB.PictureBox iFontColor 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         ScaleHeight     =   315
         ScaleWidth      =   795
         TabIndex        =   32
         Top             =   6180
         Width           =   855
      End
      Begin VB.CheckBox iFontBold 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   5700
         Width           =   1215
      End
      Begin VB.TextBox iFontSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   30
         Text            =   "12"
         Top             =   5280
         Width           =   975
      End
      Begin VB.ComboBox iFont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   4800
         Width           =   2775
      End
      Begin VB.PictureBox hFontColor 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1920
         ScaleHeight     =   315
         ScaleWidth      =   795
         TabIndex        =   23
         Top             =   2940
         Width           =   855
      End
      Begin VB.CheckBox hFontBold 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1920
         TabIndex        =   22
         Top             =   2460
         Width           =   1215
      End
      Begin VB.TextBox hFontSize 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         TabIndex        =   21
         Text            =   "12"
         Top             =   2040
         Width           =   975
      End
      Begin VB.ComboBox hFont 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   2775
      End
      Begin VB.Label Label23 
         Caption         =   "Stroke"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   68
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   65
         Top             =   3900
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.Label Label21 
         Caption         =   "Background"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   62
         Top             =   3960
         Width           =   1455
      End
      Begin VB.Label Label15 
         Caption         =   "Heading"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   1080
         Width           =   2775
      End
      Begin VB.Label Label13 
         Caption         =   "Menu Section"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label12 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   29
         Top             =   6240
         Width           =   2655
      End
      Begin VB.Label Label11 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   28
         Top             =   5760
         Width           =   2655
      End
      Begin VB.Label Label10 
         Caption         =   "Font Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   27
         Top             =   5280
         Width           =   2655
      End
      Begin VB.Label Label9 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   26
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label Label8 
         Caption         =   "Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   4440
         Width           =   2775
      End
      Begin VB.Label Label7 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   20
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Bold"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Font Size"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   18
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label4 
         Caption         =   "Font"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label3 
         Caption         =   "Margin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   2775
      End
   End
   Begin VB.Frame TabContent 
      BorderStyle     =   0  'None
      Height          =   6855
      Index           =   1
      Left            =   9960
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   4935
      Begin VB.OptionButton opFit 
         Caption         =   "Fit to Box"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   43
         Top             =   5520
         Width           =   1575
      End
      Begin VB.OptionButton opStretch 
         Caption         =   "Stretch && Crop"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   42
         Top             =   5520
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.PictureBox PicturePreview 
         AutoRedraw      =   -1  'True
         Height          =   4575
         Left            =   120
         ScaleHeight     =   4515
         ScaleWidth      =   4515
         TabIndex        =   14
         Top             =   840
         Width           =   4575
      End
      Begin VB.CommandButton butnSelectPicture 
         Caption         =   "Select Picture"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame TabContent 
      BorderStyle     =   0  'None
      Height          =   6855
      Index           =   2
      Left            =   14520
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label Label2 
         Caption         =   "This feature can be added"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   12
         Top             =   120
         Width           =   3975
      End
   End
   Begin MSComctlLib.TabStrip BlockOptions 
      Height          =   8175
      Left            =   14520
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   14420
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Menu Section"
            Key             =   "blockMenu"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Picture"
            Key             =   "blockPicture"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Video"
            Key             =   "blockVideo"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Text"
            Key             =   "blockText"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameLayouts 
      BackColor       =   &H00FFFFFF&
      Height          =   8175
      Left            =   7800
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   6735
      Begin VB.VScrollBar VScroll1 
         CausesValidation=   0   'False
         Height          =   8175
         LargeChange     =   8175
         Left            =   6480
         Max             =   1825
         SmallChange     =   100
         TabIndex        =   46
         Top             =   0
         Width           =   255
      End
      Begin VB.Frame frameScroll 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Scroll"
         Height          =   12000
         Left            =   0
         TabIndex        =   44
         Top             =   0
         Width           =   6600
         Begin VB.PictureBox Layouts 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E3E3E3&
            Height          =   2895
            Index           =   0
            Left            =   1560
            ScaleHeight     =   2835
            ScaleWidth      =   435
            TabIndex        =   47
            Top             =   1440
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.OptionButton LayoutOptions 
            BackColor       =   &H00D8D8D8&
            Height          =   975
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   120
            Width           =   1935
         End
      End
   End
   Begin VB.CommandButton butnSelectLayout 
      Caption         =   "Select Layout"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.PictureBox Preview 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8025
      ScaleWidth      =   14295
      TabIndex        =   2
      Top             =   720
      Width           =   14320
      Begin VB.PictureBox Blocks 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   1695
         Index           =   0
         Left            =   0
         ScaleHeight     =   1695
         ScaleWidth      =   1935
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.ComboBox DisplaySelect 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      ItemData        =   "frmDisplayPanels.frx":0047
      Left            =   1200
      List            =   "frmDisplayPanels.frx":0051
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label14 
      Caption         =   "Resolution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   38
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label labBlockOptions 
      Caption         =   "Block Options"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14520
      TabIndex        =   35
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label Label1 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmDisplayPanels"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SelectedDisplay As Long
Private SelectedLayout As Long
Private SelectedBlock As Long
Private MouseOverBlock As Long
Private BlockSettings() As PanelBlock
Private PreviewScale As Double
Private BlockCount As Byte
Private MadeChanges As Boolean
Private ChangeTracking As Boolean
Private LastLayoutMousedOver As Long

Private Sub BlockBG_Click()
    BlockBG_DblClick
End Sub

Private Sub BlockBG_DblClick()
    BlockBG.BackColor = SelectColor(BlockBG.BackColor)
    BlockSettings(SelectedBlock).Background = BlockBG.BackColor
    PicturePreview.BackColor = BlockBG.BackColor
    PaintPic BlockSettings(SelectedBlock).PicturePath, PicturePreview, False
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub BlockBG2_Click()
    BlockBG2.BackColor = SelectColor(BlockBG2.BackColor)
    BlockSettings(SelectedBlock).Background2 = BlockBG2.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub BlockGradient_Click()
    BlockBG2.Visible = CBool(BlockGradient)
    BlockSettings(SelectedBlock).BackgroundGradient = CBool(BlockGradient)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub BlockOptions_Click()
    Dim i As Byte
    For i = 0 To 3
        If i = BlockOptions.SelectedItem.Index - 1 Then
            TabContent(i).Visible = True
            BlockSettings(SelectedBlock).ContentType = i
        Else
            TabContent(i).Visible = False
        End If
    Next i
    BlockSettings(SelectedBlock).DrawBlock Blocks(SelectedBlock), PreviewScale
    'ChangeTracking = True
End Sub

Private Sub BlockOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseOver
End Sub

Private Sub Blocks_Click(Index As Integer)
    If SelectedBlock <> Index Then
        ChangeTracking = False
        SelectedBlock = Index ' this sets the index of the selected block
        Block_BorderChanged
        BlockSelected Index ' this populates the block options for the block that was selected
        ChangeTracking = True
    End If
End Sub

Private Sub Blocks_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MouseOverBlock <> Index Then
        MouseOverBlock = Index
        Block_BorderChanged
    End If
End Sub

Private Sub Block_BorderChanged()
    Dim i As Long
    Dim state As String
    For i = 0 To Blocks.Count - 1
        state = "Default"
        If i = SelectedBlock Then state = "Selected"
        If i = MouseOverBlock Then state = "MouseOver"
        DrawBlockBorder i, state
    Next i
End Sub

Private Sub blockText_Change()
    BlockSettings(SelectedBlock).tText = blockText
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub ButnSave_Click()
    If Not CheckDBConnection Then Exit Sub
    db.Execute "UPDATE displays SET " & _
    "resolution_x = " & Left$(cboRes.text, 4) & _
    ", resolution_y = " & Mid$(cboRes.text, 8) & _
    ", layout = " & SelectedLayout & _
    " WHERE id = " & SelectedDisplay + 1
    
    Dim i As Long
    For i = 0 To BlockCount - 1
        BlockSettings(i).Save DisplaySelect.ListIndex + 1, i
    Next i
    MadeChanges = False
End Sub

Private Sub butnSelectLayout_Click()
    frameLayouts.Visible = Not frameLayouts.Visible
End Sub

Private Sub butnSelectPicture_Click()
    CommonDialog.Filter = "Jpeg (*.jpg)|*.jpg|PNG (*.png)|*.png|Bitmap (*.bmp)|*.bmp"
    CommonDialog.DefaultExt = "jpg"
    CommonDialog.DialogTitle = "Select Picture"
    CommonDialog.ShowOpen
    
    If CommonDialog.FileName <> "" Then
        BlockSettings(SelectedBlock).PicturePath = CommonDialog.FileName
        PicturePreview.Cls
        PaintPic BlockSettings(SelectedBlock).PicturePath, PicturePreview, False
        DrawTimer.Enabled = True
        If ChangeTracking Then MadeChanges = True
    End If
End Sub

Private Sub cboRes_Click()
    PreviewScale = (Preview.Width / Screen.TwipsPerPixelX) / val(Left$(cboRes.text, 4))
    DrawAllBlocks
    If ChangeTracking Then MadeChanges = True
End Sub


Private Sub DisplaySelect_Click()
    If DisplaySelect.ListIndex <> SelectedDisplay Then
        If MadeChanges Then
            If MsgBox("You have made changes to this display, by switching you will loose your changes unless you save them first.  Do you want to switch displays now?", vbYesNo) = vbYes Then
                LoadTimer.Enabled = True
                SelectedDisplay = DisplaySelect.ListIndex
            Else
                DisplaySelect.ListIndex = SelectedDisplay
            End If
        Else
            LoadTimer.Enabled = True
            SelectedDisplay = DisplaySelect.ListIndex
        End If
    End If
End Sub

Private Sub DrawTimer_Timer()
    DrawTimer.Enabled = False
    BlockSettings(SelectedBlock).DrawBlock Blocks(SelectedBlock), PreviewScale
    Block_BorderChanged
End Sub

Private Sub Form_Load()
    Const OFFSET = 120
    
    Dim i As Long
    Dim b As OptionButton
    Dim c As OLE_COLOR
    Dim Margin As Long
    Dim h As Long
    Dim w As Long
    Dim v() As Double
    Dim block As Long
    Dim l As Long
    Dim t As Long
    Dim q As ADODB.Recordset
    Dim li As ListItem
    
    c = &HBBBBBB
    Margin = 35
    SelectedDisplay = 0
    ChangeTracking = False
    VScroll1.LargeChange = frameLayouts.Height
    VScroll1.max = frameScroll.Height - frameLayouts.Height
    
    For i = 0 To 3
        TabContent(i).Left = BlockOptions.Left
        TabContent(i).Top = 1800
    Next i
    For i = 0 To Screen.FontCount - 1
        If Left$(Screen.Fonts(i), 1) <> "@" Then
            hFont.AddItem Screen.Fonts(i)
            iFont.AddItem Screen.Fonts(i)
            tFont.AddItem Screen.Fonts(i)
        End If
    Next
    
    ' have to save these values in the database
    DisplaySelect.ListIndex = 0
    cboRes.ListIndex = 1
    
    Layouts(0).Width = LayoutOptions(0).Width - 150
    Layouts(0).Height = LayoutOptions(0).Height - 150
    
    For i = 1 To 24
        Load LayoutOptions(i)
        Load Layouts(i)
    Next i
    For Each b In LayoutOptions
        LayoutOptions(b.Index).Picture = Layouts(b.Index).Image
        LayoutOptions(b.Index).Left = LayoutOptions(b.Index).Width * 1.1 * (b.Index Mod 3) + OFFSET
        LayoutOptions(b.Index).Top = LayoutOptions(b.Index).Height * 1.15 * Int(b.Index / 3) + OFFSET
        LayoutOptions(b.Index).Visible = True
    Next b
    Set b = Nothing
    
    If CheckDBConnection Then
        Set q = db.Execute("SELECT * FROM menu_sections")
        With q
            If Not (.EOF And .BOF) Then
                .MoveFirst
                Do Until .EOF
                    Section.AddItem !Title
                    SectionIDs.AddItem !id
                    .MoveNext
                Loop
            End If
        End With
        Set q = Nothing
        Set li = Nothing
    End If
    
    w = Layouts(0).Width
    h = Layouts(0).Height
    
    'Dim ran As Long
    'Randomize Timer
    'ran = Int(Rnd * 10000)
    'Rnd -1
    'Randomize ran
    'butnSelectLayout.Caption = ran
    
    '314
    '200
    '1700
    '6248
    '9595
    '4142
    '5642
    '5335
    '5955
    Rnd -1
    Randomize 5955

    For i = 0 To Layouts.Count - 1
        Let v = GetLayout(i)
        For block = 0 To UBound(v)
            l = w * v(block, 1) + Margin
            t = h * v(block, 2) + Margin
            Layouts(i).Line (l, t)-(l + w * v(block, 3) - Margin * 2, t + h * v(block, 4) - Margin * 2), Rnd * 16777216, BF
        Next block
    Next i
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseOver
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MadeChanges Then
        If MsgBox("You have made changes to this display, by closing you will loose your changes unless you save them first.  Are you sure you want to close this window now?", vbYesNo) = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub frameScroll_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    LayoutOptions(LastLayoutMousedOver).BackColor = &HD8D8D8
End Sub

Private Sub hBackBool_Click()
    BlockSettings(SelectedBlock).hBack = CBool(hBackBool)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
    
    hBackColor1.Visible = CBool(hBackBool)
    hBackColor2.Visible = CBool(hBackBool)
    Label22.Visible = CBool(hBackBool)
End Sub

Private Sub hBackColor1_Click()
    hBackColor1.BackColor = SelectColor(hBackColor1.BackColor)
    
    BlockSettings(SelectedBlock).hBackColor1 = hBackColor1.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hBackColor2_Click()
    hBackColor2.BackColor = SelectColor(hBackColor2.BackColor)
    
    BlockSettings(SelectedBlock).hBackColor2 = hBackColor2.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hFont_Click()
    BlockSettings(SelectedBlock).HFontName = hFont.text
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hFontBold_Click()
    BlockSettings(SelectedBlock).hFontBold = CBool(hFontBold)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hFontColor_Click()
    hFontColor_DblClick
End Sub

Private Sub hFontColor_DblClick()
    hFontColor.BackColor = SelectColor(hFontColor.BackColor)
    
    BlockSettings(SelectedBlock).hFontColor = hFontColor.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hFontSize_Change()
    If val(hFontSize) < MIN_FONT_SIZE Then hFontSize = MIN_FONT_SIZE
    BlockSettings(SelectedBlock).hFontSize = val(hFontSize)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub hStrokeBool_Click()
    BlockSettings(SelectedBlock).hStroke = CBool(hStrokeBool)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
    hStrokeColor.Visible = CBool(hStrokeBool)
End Sub

Private Sub hStrokeColor_Click()
    hStrokeColor.BackColor = SelectColor(hStrokeColor.BackColor)
    
    BlockSettings(SelectedBlock).hStrokeColor = hStrokeColor.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub iFont_Click()
    BlockSettings(SelectedBlock).IFontName = iFont.text
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub iFontBold_Click()
    BlockSettings(SelectedBlock).iFontBold = CBool(iFontBold)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub iFontColor_Click()
    iFontColor_DblClick
End Sub

Private Sub iFontColor_DblClick()
    iFontColor.BackColor = SelectColor(iFontColor.BackColor)
    
    BlockSettings(SelectedBlock).iFontColor = iFontColor.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub iFontSize_Change()
    If val(iFontSize) < MIN_FONT_SIZE Then iFontSize = MIN_FONT_SIZE
    BlockSettings(SelectedBlock).iFontSize = val(iFontSize)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub LayoutOptions_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LayoutOptions(LastLayoutMousedOver).BackColor = &HD8D8D8
    LayoutOptions(Index).BackColor = vbWhite
    LastLayoutMousedOver = Index
End Sub

Private Sub LayoutOptions_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    LoadLayout Index
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub LoadLayout(ByVal Index As Long)
    Dim v() As Double
    Dim block As Long
    SelectedLayout = Index
    Let v = GetLayout(Index)
    BlockCount = UBound(v) + 1
    ReDim BlockSettings(BlockCount - 1)
    
    For block = 0 To BlockCount - 1
        If block > Blocks.Count - 1 Then Load Blocks(block)
        Blocks(block).Left = Preview.Width * v(block, 1)
        Blocks(block).Top = Preview.Height * v(block, 2)
        Blocks(block).Width = Preview.Width * v(block, 3)
        Blocks(block).Height = Preview.Height * v(block, 4)
        
        InitializeBlockSettings block
        
        BlockSettings(block).DrawBlock Blocks(block), PreviewScale
        DrawBlockBorder block, "default"
        
        Blocks(block).Visible = True
    Next block
    frameLayouts.Visible = False
    Blocks_Click -1

End Sub

Private Sub InitializeBlockSettings(ByVal Index As Long)
    Set BlockSettings(Index) = New PanelBlock
    
    BlockSettings(Index).Load DisplaySelect.ListIndex + 1, Index
    If BlockSettings(Index).SectionID = -1 Then
        If SectionIDs.ListCount > 0 Then
            BlockSettings(Index).SectionID = val(SectionIDs.List(0))
        End If
    End If
    
End Sub

Private Sub DrawBlockBorder(ByVal Index As Long, ByVal state As String)
    Dim c As Long
    If state = "MouseOver" Then
        c = vbYellow
    ElseIf state = "Selected" Then
        c = vbRed
    Else '         "Default"
        c = &H666666 'vbBlack
    End If
    
    Blocks(Index).Line (0, 0)-(Blocks(Index).Width - 30, Blocks(Index).Height - 45), c, B
    Blocks(Index).Line (15, 15)-(Blocks(Index).Width - 45, Blocks(Index).Height - 60), c, B
End Sub

Private Sub BlockSelected(ByVal Index As Long)
    If Index = -1 Then
        BlockOptions.Visible = False
        TabContent(0).Visible = False
        TabContent(1).Visible = False
        TabContent(2).Visible = False
        labBlockOptions.Visible = False
        labColor.Visible = False
        BlockBG.Visible = False
        BlockBG2.Visible = False
        BlockGradient.Visible = False
    Else
        BlockOptions.Visible = True
        labBlockOptions.Visible = True
        labColor.Visible = True
        BlockBG.Visible = True
        BlockGradient.Visible = True
        
        BlockOptions.SelectedItem = BlockOptions.Tabs(BlockSettings(Index).ContentType + 1)
        BlockOptions_Click
        
        BlockBG.BackColor = BlockSettings(Index).Background
        PicturePreview.BackColor = BlockBG.BackColor
        BlockBG2.BackColor = BlockSettings(Index).Background2
        
        'load settings from blocksettings
        BlockGradient = -CLng(BlockSettings(Index).BackgroundGradient)
        Section.ListIndex = GetSectionIndex(BlockSettings(Index).SectionID)
        Margin = BlockSettings(Index).Margin
        hFont.ListIndex = GetFontIndex(BlockSettings(Index).HFontName)
        hFontSize = BlockSettings(Index).hFontSize
        hFontBold = -CLng(BlockSettings(Index).hFontBold)
        hFontColor.BackColor = BlockSettings(Index).hFontColor
        hStrokeBool = -CLng(BlockSettings(Index).hStroke)
        hStrokeColor.BackColor = BlockSettings(Index).hStrokeColor
        hBackBool = -CLng(BlockSettings(Index).hBack)
        hBackColor1.BackColor = BlockSettings(Index).hBackColor1
        hBackColor2.BackColor = BlockSettings(Index).hBackColor2
        iFont.ListIndex = GetFontIndex(BlockSettings(Index).IFontName)
        iFontSize = BlockSettings(Index).iFontSize
        iFontBold = -CLng(BlockSettings(Index).iFontBold)
        iFontColor.BackColor = BlockSettings(Index).iFontColor
        opStretch.value = BlockSettings(Index).Stretch
        opFit.value = Not BlockSettings(Index).Stretch
        PaintPic BlockSettings(SelectedBlock).PicturePath, PicturePreview, False
        tText = BlockSettings(Index).tText
        tFont.ListIndex = GetFontIndex(BlockSettings(Index).tFontName)
        tFontSize = BlockSettings(Index).tFontSize
        tFontBold = -CLng(BlockSettings(Index).tFontBold)
        tFontColor.BackColor = BlockSettings(Index).tFontColor
        
        BlockBG2.Visible = CBool(BlockGradient)
    End If
End Sub

Private Sub ClearMouseOver()
    MouseOverBlock = -1
    Block_BorderChanged
End Sub

Private Sub LoadTimer_Timer()
    LoadTimer.Enabled = False
    
    Dim i As Byte
    If Not CheckDBConnection Then Exit Sub
    Set q = db.Execute("SELECT * FROM displays WHERE id = " & DisplaySelect.ListIndex + 1)
    With q
        If Not (.EOF And .BOF) Then
            .MoveFirst
            If IsNull(!layout) Then
                LoadLayout 0
            Else
                LoadLayout !layout
            End If
            For i = 0 To BlockCount - 1
                BlockSettings(i).Load DisplaySelect.ListIndex + 1, i
                'BlockSettings(i).DrawBlock Blocks(i), PreviewScale
                'MsgBox BlockSettings(i).PicturePath
            Next i
            For i = 0 To cboRes.ListCount - 1
                If Left$(cboRes.List(i), 4) = !resolution_x Then
                    cboRes.ListIndex = i
                    cboRes_Click
                End If
            Next i
        End If
    End With
    
    MadeChanges = False
End Sub

Private Sub Margin_Change()
    BlockSettings(SelectedBlock).Margin = val(Margin)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub opFit_Click()
    BlockSettings(SelectedBlock).Stretch = False
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub opStretch_Click()
    BlockSettings(SelectedBlock).Stretch = True
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub Section_Click()
    BlockSettings(SelectedBlock).SectionID = val(SectionIDs.List(Section.ListIndex))
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub TabContent_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearMouseOver
End Sub

Private Function GetSectionIndex(ByVal id As Long) As Long
    Dim i As Long
    GetSectionIndex = -1
    For i = 0 To SectionIDs.ListCount - 1
        If SectionIDs.List(i) = id Then GetSectionIndex = i
    Next i
End Function

Private Function GetFontIndex(ByVal fontname As String) As Long
    Dim i As Long
    GetFontIndex = -1
    For i = 0 To hFont.ListCount - 1
        If hFont.List(i) = fontname Then
            GetFontIndex = i
            Exit For
        End If
    Next i
End Function

Private Sub DrawAllBlocks()
    For block = 0 To BlockCount - 1
        If Blocks.Count > block Then
            BlockSettings(block).DrawBlock Blocks(block), PreviewScale
            DrawBlockBorder block, "default"
        End If
    Next block
End Sub

Private Sub tFont_Change()
    BlockSettings(SelectedBlock).tFontName = tFont.text
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub tFontBold_Click()
    BlockSettings(SelectedBlock).tFontBold = CBool(tFontBold)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub tFontColor_Click()
    tFontColor.BackColor = SelectColor(tFontColor.BackColor)
    
    BlockSettings(SelectedBlock).tFontColor = tFontColor.BackColor
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub tFontSize_Change()
    If val(tFontSize) < MIN_FONT_SIZE Then tFontSize = MIN_FONT_SIZE
    BlockSettings(SelectedBlock).tFontSize = val(tFontSize)
    DrawTimer.Enabled = True
    If ChangeTracking Then MadeChanges = True
End Sub

Private Sub VScroll1_Scroll()
    frameScroll.Top = -VScroll1.value
End Sub

Private Function SelectColor(current As OLE_COLOR) As OLE_COLOR
    Load frmColorSelect
    frmColorSelect.Tag = current
    frmColorSelect.Show 1
    SelectColor = csSelectedColor
End Function
