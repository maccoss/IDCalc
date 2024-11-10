Attribute VB_Name = "Globals"
Global Const VersionNum = "0.4"
Global Const Copyright = "Michael J. MacCoss, University of Washington, Department of Genome Sciences"


' Declare External DLL Libraries
' For the IDCALC subroutine in ExtractChro.dll define and call the variables as follows.

'Dim element(1 to 9) as long
'Dim enrich as double
'Dim masslist() as double
'Dim fracabun() as double
'Dim relabun() as double
'Dim masses () as double
'Dim beginmass as double
'Dim endmass as double
'Call IDCALC(element(1), enrich, masslist(1), fracabun(1), relabun(1), masses(1), beginmass, endmass, CorrectionFact)

Declare Sub IDCALC Lib "ExtractChro.dll" _
                    (element As Long, MassList As Double, FracAbun As Double, RelAbun As Double, _
                    Masses As Double, ByVal thres As Double, ByVal enrichment As Double)



