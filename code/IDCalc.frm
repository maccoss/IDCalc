VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form IDCalcfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Isotope Distribution Calculator"
   ClientHeight    =   8445
   ClientLeft      =   225
   ClientTop       =   705
   ClientWidth     =   12120
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   563
   ScaleMode       =   0  'User
   ScaleWidth      =   830.848
   Begin VB.TextBox AvgMZTxt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   66
      Top             =   600
      Width           =   1335
   End
   Begin VB.TextBox MonoMZTxt 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   7800
      Locked          =   -1  'True
      TabIndex        =   65
      Top             =   600
      Width           =   1335
   End
   Begin VB.Frame Optionsfrm 
      Caption         =   "Options"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   52
      Top             =   4320
      Width           =   2775
      Begin VB.TextBox IntegralTxt 
         Height          =   285
         Left            =   1320
         TabIndex        =   67
         Text            =   "0.01"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox DeltaMassTxt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   62
         Text            =   "0.5"
         Top             =   2400
         Width           =   975
      End
      Begin VB.TextBox AtMassTxt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   60
         Text            =   "500"
         Top             =   2010
         Width           =   975
      End
      Begin VB.TextBox ResTxt 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   58
         Text            =   "1000"
         Top             =   1650
         Width           =   975
      End
      Begin VB.CheckBox ProfileChk 
         Caption         =   "Profile"
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1200
         Width           =   1095
      End
      Begin VB.CheckBox CentChk 
         Caption         =   "Centroid"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   840
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox chargetxt 
         Height          =   285
         Left            =   960
         TabIndex        =   54
         Text            =   "1"
         Top             =   330
         Width           =   615
      End
      Begin VB.Label IntegralLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Integral:"
         Height          =   255
         Left            =   360
         TabIndex        =   68
         Top             =   2800
         Width           =   855
      End
      Begin VB.Label DeltaMassLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Delta Mass:"
         Height          =   255
         Left            =   360
         TabIndex        =   61
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label AtMassLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "At Mass:"
         Height          =   255
         Left            =   480
         TabIndex        =   59
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label ResLbl 
         Alignment       =   1  'Right Justify
         Caption         =   "Resolution (FWHM):"
         Height          =   495
         Left            =   360
         TabIndex        =   57
         Top             =   1500
         Width           =   855
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   2520
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Chargelbl 
         Caption         =   "Charge:"
         Height          =   255
         Left            =   240
         TabIndex        =   53
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame comptypefrm 
      Caption         =   "Isotope Distribution From:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   49
      Top             =   120
      Width           =   6735
      Begin VB.CheckBox AACompchk 
         Caption         =   "Amino Acid Composition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   51
         Top             =   360
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox elecompchk 
         Caption         =   "Elemental Composition"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3720
         TabIndex        =   50
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame AAfrm 
      Caption         =   "Amino Acid Composition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   47
      Top             =   1080
      Width           =   2775
      Begin VB.TextBox AAcomptxt 
         Height          =   2535
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   48
         Top             =   360
         Width           =   2535
      End
   End
   Begin MSComDlg.CommonDialog cdbPrint 
      Left            =   6000
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComDlg.CommonDialog dbsaveresults 
      Left            =   6000
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame Elementfrm 
      Caption         =   "Elemental Composition"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   3000
      TabIndex        =   11
      Top             =   1080
      Width           =   3855
      Begin VB.TextBox enrich 
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   2760
         TabIndex        =   46
         Text            =   "94.1"
         Top             =   5760
         Width           =   855
      End
      Begin VB.TextBox enrich 
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   2760
         TabIndex        =   45
         Text            =   "99"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox enrich 
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   2760
         TabIndex        =   44
         Text            =   "99"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox enrich 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   2760
         TabIndex        =   43
         Text            =   "99"
         Top             =   4680
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   15
         Left            =   1680
         TabIndex        =   41
         Text            =   "0"
         Top             =   5760
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   14
         Left            =   1680
         TabIndex        =   40
         Text            =   "0"
         Top             =   5400
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   13
         Left            =   1680
         TabIndex        =   39
         Text            =   "0"
         Top             =   5040
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   12
         Left            =   1680
         TabIndex        =   38
         Text            =   "0"
         Top             =   4680
         Width           =   855
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "18O"
         Height          =   255
         Index           =   15
         Left            =   360
         TabIndex        =   37
         Top             =   5760
         Width           =   615
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "15N"
         Height          =   255
         Index           =   14
         Left            =   360
         TabIndex        =   36
         Top             =   5400
         Width           =   975
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "2H"
         Height          =   255
         Index           =   13
         Left            =   360
         TabIndex        =   35
         Top             =   5040
         Width           =   735
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "13C"
         Height          =   255
         Index           =   12
         Left            =   360
         TabIndex        =   34
         Top             =   4680
         Width           =   735
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Carbon"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   33
         Top             =   360
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Hydrogen"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   32
         Top             =   720
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Oxygen"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   31
         Top             =   1080
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Nitrogen"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   30
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Sulfur"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   29
         Top             =   1800
         Width           =   975
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Phosphorus"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   28
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Bromine"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   27
         Top             =   2640
         Width           =   1215
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Chlorine"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   26
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Florine"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   25
         Top             =   3360
         Width           =   1095
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Boron"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   24
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox Atom 
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   23
         Text            =   "6"
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Height          =   285
         Index           =   2
         Left            =   1680
         TabIndex        =   22
         Text            =   "12"
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Height          =   285
         Index           =   3
         Left            =   1680
         TabIndex        =   21
         Text            =   "6"
         Top             =   1080
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   1680
         TabIndex        =   20
         Text            =   "0"
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   1680
         TabIndex        =   19
         Text            =   "0"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   6
         Left            =   1680
         TabIndex        =   18
         Text            =   "0"
         Top             =   2160
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   7
         Left            =   1680
         TabIndex        =   17
         Text            =   "0"
         Top             =   2640
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   8
         Left            =   1680
         TabIndex        =   16
         Text            =   "0"
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   9
         Left            =   1680
         TabIndex        =   15
         Text            =   "0"
         Top             =   3360
         Width           =   855
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   10
         Left            =   1680
         TabIndex        =   14
         Text            =   "0"
         Top             =   3720
         Width           =   855
      End
      Begin VB.CheckBox NAtom 
         Caption         =   "Silicon"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   13
         Top             =   4080
         Width           =   975
      End
      Begin VB.TextBox Atom 
         Enabled         =   0   'False
         Height          =   285
         Index           =   11
         Left            =   1680
         TabIndex        =   12
         Text            =   "0"
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label Enrichlbl 
         Alignment       =   2  'Center
         Caption         =   "Atom %"
         Height          =   255
         Left            =   2640
         TabIndex        =   42
         Top             =   4440
         Width           =   1095
      End
      Begin VB.Line Line2 
         X1              =   360
         X2              =   2520
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line1 
         X1              =   360
         X2              =   2520
         Y1              =   2520
         Y2              =   2520
      End
   End
   Begin VB.Frame Specfrm 
      Caption         =   "Spectrum"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   6960
      TabIndex        =   4
      Top             =   1080
      Width           =   5055
      Begin VB.PictureBox SpecBox 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   720
         ScaleHeight     =   1000
         ScaleMode       =   0  'User
         ScaleWidth      =   1000
         TabIndex        =   5
         Top             =   480
         Width           =   3975
      End
      Begin VB.Label Maxlbl 
         Alignment       =   1  'Right Justify
         Caption         =   "100 -"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Minlbl 
         Alignment       =   1  'Right Justify
         Caption         =   "0 -"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label LowMasslbl 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label HighMasslbl 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   4080
         TabIndex        =   7
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label mzlbl 
         Alignment       =   2  'Center
         Caption         =   "m/z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   6
         Top             =   3240
         Width           =   615
      End
   End
   Begin VB.Frame resultsfrm 
      Caption         =   "Results"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   6960
      TabIndex        =   2
      Top             =   4800
      Width           =   5055
      Begin VB.TextBox ResultsBox 
         Height          =   3015
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "IDCalc.frx":0000
         Top             =   360
         Width           =   4815
      End
   End
   Begin VB.CommandButton Clear_cmd 
      Caption         =   "Clear"
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
      Left            =   5160
      TabIndex        =   1
      Top             =   7680
      Width           =   1455
   End
   Begin VB.CommandButton Calc_cmd 
      BackColor       =   &H8000000A&
      Caption         =   "Calculate"
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
      Left            =   3360
      MaskColor       =   &H00808080&
      TabIndex        =   0
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Label AvgMZLbl 
      Alignment       =   2  'Center
      Caption         =   "Average M/Z:"
      Height          =   255
      Left            =   9840
      TabIndex        =   64
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label MonoMZLbl 
      Alignment       =   2  'Center
      Caption         =   "Mono Isotopic M/Z:"
      Height          =   255
      Left            =   7680
      TabIndex        =   63
      Top             =   240
      Width           =   1575
   End
   Begin VB.Menu Filemnu 
      Caption         =   "File"
      Begin VB.Menu Exitmnu 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu Outputmnu 
      Caption         =   "Output"
      Begin VB.Menu TextOutmnu 
         Caption         =   "Text File"
      End
      Begin VB.Menu Printermnu 
         Caption         =   "Printer"
      End
   End
   Begin VB.Menu helpmnu 
      Caption         =   "Help"
      Begin VB.Menu aboutmnu 
         Caption         =   "About IDCalc"
      End
   End
End
Attribute VB_Name = "IDCalcfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------'
' IDCalc -- Michael J. MacCoss, The University of Washington                  '
'                                                                             '
' The Isotope Distribution Calculator predicts the isotope                    '
' distribution for a compound containing selected elements.                   '
'                                                                             '
' The algorithm used for the calculation of the isotopes is based on          '
' the method reported by Kubinyi, Analytica Chimica Acta, 247 (1991) 107-119. '
' The C, H, O, N isotope abundances used in the calculations are from         '
' Dwight E. Matthews (personal communication).  All other values are from     '
' Biemann, Mass Spectrometry, Organic Chemical Applications, McGraw-Hill      '
' New York, 1962, p.59                                                        '
'                                                                             '
'                                                                             '
' Version 0.2 -- Added enriched elements to the options.                      '
'                Added fractional abundance to the output.                    '
' Version 0.3 -- Added elemental composition to the printer and text output.  '
'                Added the ability to calculate isotope distribution from     '
'                   an amino acid sequence.                                   '
'                The printer output columns are no longer staggered.          '
' Version 0.4 -- Added profile spectrum display option                                 '
'                                                                             '
'                                                                             '
'-----------------------------------------------------------------------------'

'Define Global Variables
Option Explicit
Dim Max As Double
Dim final, formula, PrntA As String
Dim VerNum As String
Dim P, Q, J, i, K, L As Long
Dim AAcount(1 To 20, 1 To 7) As Variant
Dim CPATT(10000), D(10000) As Double
Dim Frac(10000) As Double
Dim ProfileSpec() As Double
Dim Carbon, Hydrogen, Oxygen, Nitrogen, Sulfur As String
Dim Phosphorous, Bromine, Chlorine, Florine, Boron, Silicon As String
Dim C13, H2, N15, O18 As String
Dim MonoMass As Double, AvgMass As Double

Const GapWidth = 1.00235 'Average distance between isotope peaks
'Const GapWidth = 1


Private Sub CentChk_Click()
    If CentChk = 1 Then
        ProfileChk = 0
        ResTxt.Enabled = False
        AtMassTxt.Enabled = False
        DeltaMassTxt.Enabled = False
    Else
        ProfileChk = 1
        ResTxt.Enabled = True
        AtMassTxt.Enabled = True
        DeltaMassTxt.Enabled = True
    End If
End Sub



Public Sub Form_Load()
    'Define version number here (also define it in the frmAbout code)
    VerNum = "0.3"
        
If AACompchk = 1 Then
ResultsBox.Text = "Type the single letter amino acid sequence and 'click' calculate."
elecompchk.Value = 0

 For i = 1 To 11
    Atom(i).Enabled = False
    NAtom(i).Enabled = False
 Next i

 For i = 12 To 15
    Atom(i).Enabled = False
    enrich(i).Enabled = False
    NAtom(i).Enabled = False
 Next i
 
Else
ResultsBox.Text = "Select the elements present in your compound and 'click' calculate."
 
End If

'----------------------------------------------------------
' Load data into AAcount array
' 1,1 to 20,1 are the single letter amino acid strings
' 1,2 to 20,2 are the corresponding # of residues in the
'       peptides (not added until program is run)
' 1,3 to 20,3 are the # of carbons in the residue
' 1,4 to 20,4 are the # of hydrogens in the residue
' 1,5 to 20,5 are the # of nitrogens in the residue
' 1,6 to 20,6 are the # of oxygens in the residue
' 1,7 to 20,7 are the # of sulfurs in the residue
'----------------------------------------------------------

AAcount(1, 1) = "A"
    AAcount(1, 3) = 3   ' carbons
    AAcount(1, 4) = 5   ' hydrogens
    AAcount(1, 5) = 1   ' nitrogens
    AAcount(1, 6) = 1   ' oxygens
    AAcount(1, 7) = 0   ' sulfurs

AAcount(2, 1) = "R"
    AAcount(2, 3) = 6
    AAcount(2, 4) = 12
    AAcount(2, 5) = 4
    AAcount(2, 6) = 1
    AAcount(2, 7) = 0
    
AAcount(3, 1) = "N"
    AAcount(3, 3) = 4
    AAcount(3, 4) = 6
    AAcount(3, 5) = 2
    AAcount(3, 6) = 2
    AAcount(3, 7) = 0
    
AAcount(4, 1) = "D"
    AAcount(4, 3) = 4
    AAcount(4, 4) = 5
    AAcount(4, 5) = 1
    AAcount(4, 6) = 3
    AAcount(4, 7) = 0
    
AAcount(5, 1) = "C"
    AAcount(5, 3) = 3
    AAcount(5, 4) = 5
    AAcount(5, 5) = 1
    AAcount(5, 6) = 1
    AAcount(5, 7) = 1
    
AAcount(6, 1) = "Q"
    AAcount(6, 3) = 5
    AAcount(6, 4) = 8
    AAcount(6, 5) = 2
    AAcount(6, 6) = 2
    AAcount(6, 7) = 0
    
AAcount(7, 1) = "E"
    AAcount(7, 3) = 5
    AAcount(7, 4) = 7
    AAcount(7, 5) = 1
    AAcount(7, 6) = 3
    AAcount(7, 7) = 0
    
AAcount(8, 1) = "G"
    AAcount(8, 3) = 2
    AAcount(8, 4) = 3
    AAcount(8, 5) = 1
    AAcount(8, 6) = 1

AAcount(9, 1) = "H"
    AAcount(9, 3) = 6
    AAcount(9, 4) = 7
    AAcount(9, 5) = 3
    AAcount(9, 6) = 1
    AAcount(9, 7) = 0

AAcount(10, 1) = "I"
    AAcount(10, 3) = 6
    AAcount(10, 4) = 11
    AAcount(10, 5) = 1
    AAcount(10, 6) = 1
    AAcount(10, 7) = 0
    
AAcount(11, 1) = "L"
    AAcount(11, 3) = 6
    AAcount(11, 4) = 11
    AAcount(11, 5) = 1
    AAcount(11, 6) = 1
    AAcount(11, 7) = 0
    
AAcount(12, 1) = "K"
    AAcount(12, 3) = 6
    AAcount(12, 4) = 12
    AAcount(12, 5) = 2
    AAcount(12, 6) = 1
    AAcount(12, 7) = 0
    
AAcount(13, 1) = "M"
    AAcount(13, 3) = 5
    AAcount(13, 4) = 9
    AAcount(13, 5) = 1
    AAcount(13, 6) = 1
    AAcount(13, 7) = 1
    
AAcount(14, 1) = "F"
    AAcount(14, 3) = 9
    AAcount(14, 4) = 9
    AAcount(14, 5) = 1
    AAcount(14, 6) = 1
    AAcount(14, 7) = 0

AAcount(15, 1) = "P"
    AAcount(15, 3) = 5
    AAcount(15, 4) = 7
    AAcount(15, 5) = 1
    AAcount(15, 6) = 1
    AAcount(15, 7) = 0
    
AAcount(16, 1) = "S"
    AAcount(16, 3) = 3
    AAcount(16, 4) = 5
    AAcount(16, 5) = 1
    AAcount(16, 6) = 2
    AAcount(16, 7) = 0

AAcount(17, 1) = "T"
    AAcount(17, 3) = 4
    AAcount(17, 4) = 7
    AAcount(17, 5) = 1
    AAcount(17, 6) = 2
    AAcount(17, 7) = 0
    
AAcount(18, 1) = "W"
    AAcount(18, 3) = 11
    AAcount(18, 4) = 10
    AAcount(18, 5) = 2
    AAcount(18, 6) = 1
    AAcount(18, 7) = 0
    
AAcount(19, 1) = "Y"
    AAcount(19, 3) = 9
    AAcount(19, 4) = 9
    AAcount(19, 5) = 1
    AAcount(19, 6) = 2
    AAcount(19, 7) = 0
    
AAcount(20, 1) = "V"
    AAcount(20, 3) = 5
    AAcount(20, 4) = 9
    AAcount(20, 5) = 1
    AAcount(20, 6) = 1
    AAcount(20, 7) = 0
    
End Sub

Public Sub Calc_cmd_Click()

If AACompchk.Value = 1 Then

    If Len(AAcomptxt.Text) = 0 Then
    MsgBox "Must Enter AA Sequence", 64, "Notice"
    Else
    
    Dim i, A As Long
    Dim AA As String
        
    For A = 1 To 20
        AAcount(A, 2) = 0
    Next
        
    ' Calculate the number of each amino acid present in the string
    For A = 1 To 20
        For i = 1 To Len(AAcomptxt.Text)
        AA = Mid(AAcomptxt.Text, i, 1)
        If AAcount(A, 1) = UCase(AA) Then
            AAcount(A, 2) = AAcount(A, 2) + 1
        End If
                
        Next
    Next
           
           
     Dim Cnum, Hnum, Nnum, Onum, Snum As Long
     Cnum = 0
     Hnum = 0
     Nnum = 0
     Onum = 0
     Snum = 0
     
     ' Calculate the elemental composition from the AA composition of the sequence
     For A = 1 To 20
        Cnum = (AAcount(A, 2) * AAcount(A, 3)) + Cnum
        Hnum = (AAcount(A, 2) * AAcount(A, 4)) + Hnum
        Nnum = (AAcount(A, 2) * AAcount(A, 5)) + Nnum
        Onum = (AAcount(A, 2) * AAcount(A, 6)) + Onum
        Snum = (AAcount(A, 2) * AAcount(A, 7)) + Snum
     Next
           
     Atom(1).Text = Cnum
     NAtom(1) = 1
     Atom(2).Text = Hnum + 2 'For termini
     NAtom(2) = 1
     Atom(3).Text = Onum + 1 'For termini
     NAtom(3) = 1
     Atom(4).Text = Nnum
     NAtom(4) = 1
     Atom(5).Text = Snum
     NAtom(5) = 1
     
     For A = 6 To 15
        Atom(A).Text = 0
        NAtom(A) = 0
     Next
           
     For A = 1 To 15
        Atom(A).Enabled = False
        NAtom(A).Enabled = False
     Next
           
    Call IDCalc
    Call Print_Results
    
    End If
Else
    Call IDCalc
    Call Print_Results
    
End If


End Sub

Public Sub IDCalc()
Dim Abund(1 To 15, 1 To 5) As Single
Dim NPeak(1 To 15) As Integer
Dim Prec As Double


'Change pointer to hourglass
Screen.MousePointer = vbHourglass

'-------------------------------------------------------'
' Natural isotope abundances for biological samples     '
' The isotope abundances are placed into an             '
' multidimensional array.                               '
'                                                       '
' These values were obtained from D.E. Matthews.        '
'-------------------------------------------------------'

'Carbon
Abund(1, 1) = 100
Abund(1, 2) = 1.0958793
'Abund(1, 2) = 1.11
NPeak(1) = 2

'Hydrogen
Abund(2, 1) = 100
Abund(2, 2) = 0.0142
NPeak(2) = 2

'Oxygen
Abund(3, 1) = 100
Abund(3, 2) = 0.03799194
Abund(3, 3) = 0.20499609
NPeak(3) = 3

'Nitrogen
Abund(4, 1) = 100
Abund(4, 2) = 0.368351851
NPeak(4) = 2

'Sulfur
Abund(5, 1) = 100
Abund(5, 2) = 0.789308
Abund(5, 3) = 4.430646
Abund(5, 4) = 0
Abund(5, 5) = 0.021048
NPeak(5) = 5

'Phosphorus
Abund(6, 1) = 100
NPeak(6) = 1

'Bromine -- Need to double check these!!
Abund(7, 1) = 100
Abund(7, 2) = 0
Abund(7, 3) = 98
NPeak(7) = 3

'Chlorine
Abund(8, 1) = 100
Abund(8, 2) = 0
Abund(8, 3) = 31.978
NPeak(8) = 3

'Fluorine
Abund(9, 1) = 100
NPeak(9) = 1

'Boron
Abund(10, 1) = 24.394
Abund(10, 2) = 100
NPeak(10) = 2

'Silicon
Abund(11, 1) = 100
Abund(11, 2) = 5.097
Abund(11, 3) = 3.351
NPeak(11) = 3

'13C
Abund(12, 1) = 100 * (1 - enrich(12) / 100)
Abund(12, 2) = Abund(1, 2) * (1 - enrich(12) / 100) + enrich(12)
NPeak(12) = 2

'13C
'Abund(12, 1) = 100 - enrich(12)
'Abund(12, 2) = enrich(12) + Abund(12, 1) * Abund(1, 2) / 100
'NPeak(12) = 2


'2H
Abund(13, 1) = 100 - enrich(13)
Abund(13, 2) = enrich(13) + Abund(13, 1) * Abund(2, 2) / 100
NPeak(13) = 2

'15N
Abund(14, 1) = 100 - enrich(14)
Abund(14, 2) = enrich(14) + Abund(14, 1) * Abund(4, 2) / 100
NPeak(14) = 2

'18O
Abund(15, 1) = 100 - enrich(15)
Abund(15, 2) = Abund(15, 1) * Abund(3, 2) / 100
Abund(15, 3) = enrich(15) + Abund(15, 1) * Abund(3, 3) / 100
NPeak(15) = 3


'------------------------------------------------------'
'Calculation of Isotope Distributions                  '
'This algorithmn is borrowed heavily from              '
'Kubinyi, Analytica Chimica Acta, 247 (1991) 107-119   '
'------------------------------------------------------'

P = 1
Q = 1
Prec = 0.0000001

CPATT(1) = 1

For J = 1 To 15

If NAtom(J) = 1 Then

    For i = 1 To Int(Atom(J))
        Erase D
        
        ' Calculate Isotope distribution
        For K = P To Q
            For L = 1 To NPeak(J)
                D(K + L - 1) = D(K + L - 1) + CPATT(K) * Abund(J, L)
            Next
        Next
        
        Q = Q + NPeak(J) - 1
        Max = 0
        For K = P To Q
            If D(K) > Max Then
            Max = D(K)
            End If
        Next
        For K = P To Q
            D(K) = D(K) / Max
        Next
        
        ' Eliminate small peaks to the left
        For K = P To Q
            If D(K) > Prec Then
                P = K
                K = Q
            End If
        Next
        
        'Eliminate small peaks to the right
        K = Q
        Do Until D(K) > Prec
            Q = K
            K = K - 1
        Loop
        
        'Create new isotope pattern
        Erase CPATT
        For K = P To Q
            CPATT(K) = D(K)
        Next
        
    Next

End If
Next

'------------------------------------------------------'
' Calculate Monoisotopic Mass and Molecular Formula    '
'------------------------------------------------------'

MonoMass = 0
If NAtom(1) = 1 Then
MonoMass = MonoMass + Atom(1) * 12
Carbon = "C" & Str(Atom(1)) & " "
End If
If NAtom(2) = 1 Then
MonoMass = MonoMass + Atom(2) * 1.007825
Hydrogen = "H" & Str(Atom(2)) & " "
End If
If NAtom(3) = 1 Then
MonoMass = MonoMass + Atom(3) * 15.9949146
Oxygen = "O" & Str(Atom(3)) & " "
End If
If NAtom(4) = 1 Then
MonoMass = MonoMass + Atom(4) * 14.003074
Nitrogen = "N" & Str(Atom(4)) & " "
End If
If NAtom(5) = 1 Then
MonoMass = MonoMass + Atom(5) * 31.9720718
Sulfur = "S" & Str(Atom(5)) & " "
End If
If NAtom(6) = 1 Then
MonoMass = MonoMass + Atom(6) * 30.9737634
Phosphorous = "P" & Str(Atom(6)) & " "
End If
If NAtom(7) = 1 Then
MonoMass = MonoMass + Atom(7) * 79.9183361
Bromine = "Br" & Str(Atom(7)) & " "
End If
If NAtom(8) = 1 Then
MonoMass = MonoMass + Atom(8) * 34.9688527
Chlorine = "Cl" & Str(Atom(8)) & " "
End If
If NAtom(9) = 1 Then
MonoMass = MonoMass + Atom(9) * 18.9984033
Florine = "F" & Str(Atom(9)) & " "
End If
If NAtom(10) = 1 Then
MonoMass = MonoMass + Atom(10) * 10.0129
Boron = "B" & Str(Atom(10)) & " "
End If
If NAtom(11) = 1 Then
MonoMass = MonoMass + Atom(11) * 27.9769
Silicon = "Si" & Str(Atom(11)) & " "
End If
If NAtom(12) = 1 Then
MonoMass = MonoMass + Atom(12) * (12)
C13 = "(13C)" & Str(Atom(12)) & " "
End If
If NAtom(13) = 1 Then
MonoMass = MonoMass + Atom(13) * (1.007825)
H2 = "(2H)" & Str(Atom(13)) & " "
End If
If NAtom(14) = 1 Then
MonoMass = MonoMass + Atom(14) * (14.003074)
N15 = "(15N)" & Str(Atom(14)) & " "
End If
If NAtom(15) = 1 Then
MonoMass = MonoMass + Atom(15) * (15.9949146)
O18 = "(18O)" & Str(Atom(15)) & " "
End If




'Calculate Fractional Abundances
Dim sum As Double

AvgMass = 0
sum = 0
For K = P To Q
    sum = sum + CPATT(K)
Next

For K = P To Q
    Frac(K) = CPATT(K) / sum
    AvgMass = Frac(K) * (MonoMass + (K * GapWidth) - GapWidth) + AvgMass
Next





End Sub

Public Sub Print_Results()
Dim MZ As Double
Dim PrntB As String
Dim i As Long
    'Print Results in Window
MonoMass = Round(MonoMass, 4)
formula = Carbon & C13 & Hydrogen & H2 & Oxygen & O18 & Nitrogen & N15 & Sulfur & Phosphorous & Bromine & Chlorine & Florine & Boron & Silicon
resultsfrm.Caption = "Results for: " & formula

final = "Mass (da) " & vbTab & "M/Z" & vbTab & "Rel. Abu." & vbTab & "Frac. Abu." & vbCrLf
For K = P To Q
    'The 0.00055 is subtracted for the loss of an electron for each proton that is added
    If Val(chargetxt.Text) = 0 Then
        MZ = Str(Round((MonoMass + (K * GapWidth) - GapWidth), 4))
    Else
        MZ = Str(Round((MonoMass + (K * GapWidth) - GapWidth + ((1.00782 - 0.00055) * Val(chargetxt.Text))) / Val(chargetxt.Text), 4))
    End If
    PrntA = Round(MonoMass + (K * GapWidth) - GapWidth, 4) & vbTab & MZ & vbTab & Str(Round(CPATT(K) * 100, 4)) & vbTab & Str(Round(Frac(K) * 100, 4)) & vbCrLf
    final = final & PrntA
Next


If ProfileChk.Value = 1 Then
    Call GaussianFunct
    'For i = 1 To UBound(ProfileSpec)
    '    PrntB = ProfileSpec(i, 1) & vbTab & ProfileSpec(i, 2) & vbCrLf
    '    final = final & PrntB
    'Next
End If

ResultsBox.Text = final
If Val(chargetxt.Text) = 0 Then
    MonoMZTxt.Text = Round(MonoMass, 4)
    AvgMZTxt.Text = Round(AvgMass, 4)
Else
    MonoMZTxt.Text = Round((MonoMass + (1.00782 - 0.00055) * Val(chargetxt.Text)) / Val(chargetxt.Text), 4)
    AvgMZTxt.Text = Round((AvgMass + (1.00782 - 0.00055) * Val(chargetxt.Text)) / Val(chargetxt.Text), 4)
End If

Call PrintSpec
    
'Change pointer back to normal
Screen.MousePointer = vbArrow

End Sub

Public Sub GaussianFunct()
Dim SpecTemp() As Double
Dim SpecTempNorm() As Double
Dim Count As Long, CountEnd As Long
Dim CountTotal As Long
Dim Integral As Double, Sigma As Double
Dim MassTemp As Double
Dim MZTemp As Double
Dim MeanMZTemp As Double
Dim MaxInt As Double
Dim Part1 As Double, Part2 As Double

Sigma = Val(DeltaMassTxt.Text / 2)
Integral = Val(IntegralTxt.Text)

If Int(chargetxt.Text) = 0 Then
    CountTotal = Round((Q - P) / Integral)
Else
    CountTotal = Round((Q - P) / Integral / Int(chargetxt.Text))
End If
ReDim ProfileSpec(1 To CountTotal, 1 To 2) As Double
ReDim SpecTemp(1 To CountTotal, 1 To 2) As Double
ReDim SpecTempNorm(1 To CountTotal, 1 To 2) As Double

Part1 = 1 / (((2 * 3.14) ^ 0.5) * Sigma)

For K = P To Q
    MassTemp = (MonoMass + P - 1) - (Val(DeltaMassTxt.Text) * 2)
    MZTemp = (MassTemp + ((1.00782 - 0.00055) * Val(chargetxt.Text))) / Int(chargetxt.Text)
    MaxInt = 0.00000000000001
    MeanMZTemp = (MonoMass + (K * GapWidth) - GapWidth + ((1.00782 - 0.00055) * Val(chargetxt.Text))) / Int(chargetxt.Text)
    For Count = 1 To CountTotal
        SpecTemp(Count, 1) = MZTemp
        Part2 = Exp(-0.5 * ((MZTemp - MeanMZTemp) / Sigma) ^ 2)
        
        SpecTemp(Count, 2) = Part1 * Part2
        
        'Determine the Max
        If SpecTemp(Count, 2) > MaxInt Then
            MaxInt = SpecTemp(Count, 2)
        End If
        MZTemp = MZTemp + Integral
    Next

    ' Normalize the Profile
    For Count = 1 To CountTotal
        SpecTempNorm(Count, 2) = SpecTemp(Count, 2) * CPATT(K) * 100 / MaxInt
        ProfileSpec(Count, 2) = ProfileSpec(Count, 2) + SpecTempNorm(Count, 2)
    Next

Next

MaxInt = 0.00000000000001
For Count = 1 To CountTotal
    ProfileSpec(Count, 1) = SpecTemp(Count, 1)
    If ProfileSpec(Count, 2) > MaxInt Then
        MaxInt = ProfileSpec(Count, 2)
    End If
Next

' Normalize the Profile Again
For Count = 1 To CountTotal
    ProfileSpec(Count, 2) = ProfileSpec(Count, 2) * 100 / MaxInt
Next


End Sub



Public Sub PrintSpec()
Dim Xinterval As Single, Xpoint As Single
Dim X1 As Double, Y1 As Double, X2 As Double, Y2 As Double
Dim X1_Temp As Double, X2_Temp As Double, deltaY As Double, deltaX As Double
Dim Y1_Temp As Double, Y2_Temp As Double

 
'Clear Specbox
SpecBox.Line (0, 0)-(1000, 1000), SpecBox.BackColor, BF

If CentChk.Value = 1 Then
'Draw Spectrum
    If Val(chargetxt.Text) = 0 Then
        LowMasslbl.Caption = Str(Round(MonoMass + P - 1, 4))
        HighMasslbl.Caption = Str(Round(MonoMass + Q - 1, 4))
    
    Else
        LowMasslbl.Caption = Str(Round((MonoMass + P - 1 + ((1.00782 - 0.00055) * Val(chargetxt.Text))) / Val(chargetxt.Text), 4))
        HighMasslbl.Caption = Str(Round((MonoMass + Q - 1 + ((1.00782 - 0.00055) * Val(chargetxt.Text))) / Val(chargetxt.Text), 4))
    End If
    Xinterval = 950 / ((Q + 1) - (P - 1))
    Xpoint = 50
    For K = P To Q
        SpecBox.Line (Xpoint, 1000)-(Xpoint, 1000 - Round(CPATT(K) * 1000, 1)), vbBlue
        Xpoint = Xpoint + Xinterval
    Next
Else
    
    Dim XIntegral As Double
    Dim endmz As Long
    
    endmz = UBound(ProfileSpec)
    
    LowMasslbl.Caption = Round(ProfileSpec(1, 1), 3)
    HighMasslbl.Caption = Round(ProfileSpec(endmz, 1), 3)
    
    XIntegral = 1000 / endmz
    
    Y1_Temp = ProfileSpec(1, 2)
    Y1 = 1000 - (Y1_Temp * (1000 / 100))
         
    X1 = 0
     
    
        For i = 2 To endmz
                        
            Y2_Temp = ProfileSpec(i, 2)
            Y2 = 1000 - (Y2_Temp * (1000 / 100))
            
            X2 = X1 + XIntegral
                      
            If (X2 <= 1000) And (X2 > X1) Then
            
                If Y2 < 0 Then
                    Y2 = 0
                End If
                If Y2 > 1000 Then
                    Y2 = 1000
                End If
                If Y1 > 1000 Then
                    Y1 = 1000
                End If
                
                If X1 < 0 Then
                    X1 = 0
                End If
                If X1 > 1000 Then
                    X1 = 1000
                End If
                If X2 < 0 Then
                    X2 = 0
                End If
                If X2 > 1000 Then
                    X2 = 1000
                End If
                
                SpecBox.Line (X1, Y1)-(X2, Y2), vbBlue
                        
            End If
             
            Y1 = Y2
            X1 = X2
        Next
    
    
    

End If

End Sub



Private Sub Clear_cmd_Click()
    
    If elecompchk = 1 Then
    ResultsBox.Text = "Select the elements present in your compound and 'click' calculate."
    Else
    ResultsBox.Text = "Type the single letter amino acid sequence and 'click' calculate."
    End If
    
    resultsfrm.Caption = "Results"
'Clear Specbox
    SpecBox.Line (0, 0)-(100, 100), SpecBox.BackColor, BF
    
    AAcomptxt.Text = ""

End Sub

Private Sub Exitmnu_Click()
    End
End Sub

Private Sub NAtom_Click(Index As Integer)
 'This subroutine makes the appropriate text boxes disabled/enabled when checked.
Dim i As Integer

 For i = 1 To 11
    If NAtom(i) = 1 Then
    Atom(i).Enabled = True
    Else
    Atom(i).Enabled = False
    End If
 Next i

 For i = 12 To 15
    If NAtom(i) = 1 Then
    Atom(i).Enabled = True
    enrich(i).Enabled = True
    Else
    Atom(i).Enabled = False
    enrich(i).Enabled = False
    End If
 Next i


End Sub

Private Sub aboutmnu_Click()
    frmAbout.Show 1
End Sub

Private Sub Printermnu_Click()
' This subroutine uses the Windows' common dialog
' box to allow the user to print a report to a selected printer.
    
    Dim prntmass, prntabund, prntfrac As Single
    
    On Error GoTo Error_Handler
    cdbPrint.ShowPrinter
    Screen.MousePointer = vbHourglass
    
    Printer.ScaleMode = vbInches
    Printer.CurrentY = 0.5
    Printer.CurrentX = 0.5
    
    Printer.FontBold = True
    Printer.FontSize = 18
    
    Printer.Print "Isotope Distribution Calculator v. " & VerNum
    
    Printer.FontBold = True
    Printer.FontSize = 10
    
    Printer.CurrentY = 1
    Printer.CurrentX = 0.5
    
    Printer.Print "M.J. MacCoss, Department of Genome Sciences, University of Washington"
    Printer.Print
        
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentY = 1.2
    Printer.CurrentX = 0.5
    Printer.Print "Elemental Composition: " & formula
    
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentY = 1.4
    Printer.CurrentX = 0.5
    Printer.Print "Mass"
    
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentY = 1.4
    Printer.CurrentX = 1.2
    Printer.Print "Rel. Abund"
    
    Printer.FontSize = 10
    Printer.FontBold = True
    Printer.CurrentY = 1.4
    Printer.CurrentX = 2.2
    Printer.Print "Frac. Abund."
    
Dim Y As Single
Y = 1.6
For K = P To Q
    prntmass = (MonoMass + K - 1)
    If prntmass > 0 Then
    prntabund = Round(CPATT(K) * 100, 4)
    prntfrac = Round(Frac(K) * 100, 4)
    'PrntA = prntmass & vbTab & prntabund & vbTab & vbTab & prntfrac
    Printer.CurrentY = Y
    Printer.CurrentX = 0.5
    Printer.FontBold = False
    Printer.Print prntmass
    
    Printer.CurrentY = Y
    Printer.CurrentX = 1.2
    Printer.FontBold = False
    Printer.Print prntabund
    
    Printer.CurrentY = Y
    Printer.CurrentX = 2.2
    Printer.FontBold = False
    Printer.Print prntfrac
    
    End If
    Y = Y + 0.2


Next
          
       
     Printer.EndDoc
     Screen.MousePointer = vbArrow
     
Error_Handler:
    Exit Sub
    
    

End Sub


Private Sub ProfileChk_Click()
    If ProfileChk = 1 Then
        CentChk = 0
        ResTxt.Enabled = True
        AtMassTxt.Enabled = True
        DeltaMassTxt.Enabled = True
    Else
        CentChk = 1
        ResTxt.Enabled = False
        AtMassTxt.Enabled = False
        DeltaMassTxt.Enabled = False
    End If
    
End Sub

Private Sub ResTxt_Change()
    DeltaMassTxt.Text = Val(AtMassTxt.Text) / Val(ResTxt.Text)
End Sub

Private Sub AtMassTxt_Change()
    DeltaMassTxt.Text = Val(AtMassTxt.Text) / Val(ResTxt.Text)
End Sub

Private Sub DeltaMassTxt_Change()
    If Val(DeltaMassTxt.Text) > 0 Then
        ResTxt.Text = Val(AtMassTxt.Text) / Val(DeltaMassTxt.Text)
    End If
End Sub


Private Sub TextOutmnu_Click()
Dim filename As String
Dim fnum As Integer
Dim prntmass, prntabund, prntfrac As Single

' This subroutine uses the Windows' common dialog
' box to allow the user to save data to a specific file.

dbsaveresults.CancelError = True
 On Error GoTo dbCancel

dbsaveresults.Flags = &H4
dbsaveresults.Filter = "All Files (*.*)|*.*"
dbsaveresults.ShowSave
 
filename = dbsaveresults.filename

fnum = FreeFile()
Open filename For Output As #fnum

    Print #fnum, "Isotope Distribution Calculator v. " & VerNum
    Print #fnum, "Michael J. MacCoss -- University of Washington"
    Print #fnum, vbCr
    Print #fnum, "Elemental Composition: " & formula
    Print #fnum, "Mass" & vbTab & "Rel. Abund" & vbTab & "Frac. Abund."
    
If CentChk.Value = 1 Then
    
    For K = P To Q
        prntmass = (MonoMass + K - 1)
        If prntmass > 0 Then
            prntabund = Round(CPATT(K) * 100, 4)
            prntfrac = Round(Frac(K) * 100, 4)
            PrntA = prntmass & vbTab & prntabund & vbTab & prntfrac
            Print #fnum, PrntA
        End If
    Next
Else
    
    
    'Below is a quick fix to outputing the profile spectrum to the text file
    Dim XIntegral As Double
    Dim endmz As Long
    
    endmz = UBound(ProfileSpec)
    
    For i = 1 To endmz
        PrntA = ProfileSpec(i, 1) & vbTab & ProfileSpec(i, 2)
        Print #fnum, PrntA
    Next
    
    
End If
    
Close #1
 
    

dbCancel:
    'The user pressed cancel so
    'ignore file selection


End Sub

Private Sub AACompchk_Click()
Dim i As Integer

If AACompchk = 1 Then
ResultsBox.Text = "Type amino acid sequence and 'click' calculate."
elecompchk.Value = 0
AAcomptxt.Enabled = True

 For i = 1 To 11
    Atom(i).Enabled = False
    NAtom(i).Enabled = False
 Next i

 For i = 12 To 15
    Atom(i).Enabled = False
    enrich(i).Enabled = False
    NAtom(i).Enabled = False
 Next i
 
Else
ResultsBox.Text = "Select the elements present in your compound and 'click' calculate."
elecompchk.Value = 1
AAcomptxt.Enabled = False


For i = 1 To 11
    NAtom(i).Enabled = True
    
    If NAtom(i) = 1 Then
    Atom(i).Enabled = True
    Else
    Atom(i).Enabled = False
    End If
 Next i

 For i = 12 To 15
    NAtom(i).Enabled = True
    
    If NAtom(i) = 1 Then
    Atom(i).Enabled = True
    enrich(i).Enabled = True
    Else
    Atom(i).Enabled = False
    enrich(i).Enabled = False
    End If
 Next i
    
    

End If
 
End Sub

Private Sub elecompchk_Click()
Dim i As Integer

If elecompchk = 1 Then
ResultsBox.Text = "Select the elements present in your compound and 'click' calculate."
AACompchk.Value = 0
AAcomptxt.Enabled = False

For i = 1 To 11
    NAtom(i).Enabled = True
    
    If NAtom(i) = 1 Then
    Atom(i).Enabled = True
    Else
    Atom(i).Enabled = False
    End If
 Next i

 For i = 12 To 15
    NAtom(i).Enabled = True
    
    If NAtom(i) = 1 Then
    Atom(i).Enabled = True
    enrich(i).Enabled = True
    Else
    Atom(i).Enabled = False
    enrich(i).Enabled = False
    End If
 Next i
    

 
Else
ResultsBox.Text = "Type amino acid sequence and 'click' calculate."
AACompchk.Value = 1
AAcomptxt.Enabled = True

 For i = 1 To 11
    Atom(i).Enabled = False
    NAtom(i).Enabled = False
 Next i

 For i = 12 To 15
    Atom(i).Enabled = False
    enrich(i).Enabled = False
    NAtom(i).Enabled = False
 Next i



End If

End Sub
