Attribute VB_Name = "modRatings"
Option Explicit
Public Function Rating(Comps As Single, Atts As Single, TDs As Single, Ints As Single, TtlYds As Single) As Single
'Comps=Completions
'Atts=attempts
'TDs=touchdowns
'Ints=interceptions
'TtlYds=Total yards

Dim PassPerc As Single 'Pass percentage
Dim PntRtng As Single 'point rating
Dim YrdAtt As Single 'yards per attempt
Dim AvgYrdAtt As Single 'average yards per attempt
Dim TDPer As Single 'average TDs per attempt
Dim IntPer As Single 'Interception percentage
Dim Totals As Single 'totals of all catagories, then grand total

PassPerc = FormatNumber(Comps / Atts, 4) * 100  '4 after attempts
PntRtng = (PassPerc - 30)
PntRtng = PntRtng * 0.05
YrdAtt = FormatNumber(TtlYds / Atts, 3)
AvgYrdAtt = FormatNumber((YrdAtt - 3) * 0.25, 3)
TDPer = FormatNumber(TDs / Atts * 100, 3)
TDPer = FormatNumber(TDPer * 0.2, 3)
IntPer = FormatNumber(Ints / Atts * 100, 2)
IntPer = FormatNumber(IntPer * 0.25, 3)
IntPer = FormatNumber(2.375 - IntPer, 3)
Totals = PntRtng + AvgYrdAtt + TDPer + IntPer
Totals = FormatNumber((Totals / 6) * 100, 2)
Rating = Totals

End Function
