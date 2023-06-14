' punto d'ingresso del programma (entry point of program)
Sub Main()
  If Application.Name = "Microsoft Excel" Then
        Call Calendar() 
    ElseIf Application.Name = "OpenOffice.org Calc" Then
        Call CalcCalendar() ' Esegui lo script VBScript per Calc
    End If
Call Calendar()
Call CalcCalendar()
End Sub

' punto di controllo; sono su openoffice o sono su excel? ( check point am I on openoffice or on excel?)


Sub Calendar()
'
' Calendario Macro
'

'
    ActiveCell.FormulaR1C1 = "Gennaio"
    Range("C2").Select
    Columns("A:A").ColumnWidth = 18.44
    Range("A2").Select
    Columns("A:A").ColumnWidth = 23.11
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "2022"
    Range("G2").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("$A$3:$G$3").Value = Array("Lunedi", "Martedi", "Mercoledi", "Giovedi", "Venerdi", "Sabato", "Domenica")
    ActiveCell.FormulaR1C1 = "=mo"
    Range("B2").Select
    ActiveCell.FormulaR1C1 = "=MONTH(RC[-1]&1)"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=DATE(R[-2]C[6],R[-2]C,1)"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=DATE(R[-2]C[6],R[-2]C[1],1)"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = ""
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "=DATE(R[-2]C[6],R[-2]C[1],1)-5"
    Range("B4").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]+1"
    Range("B4").Select
    Selection.AutoFill Destination:=Range("B4:G4"), Type:=xlFillDefault
    Range("B4:G4").Select
    Columns("C:C").ColumnWidth = 20.44
    Columns("D:D").ColumnWidth = 18.67
    Columns("E:E").ColumnWidth = 14
    Range("F1").Select
    Columns("F:F").ColumnWidth = 14.89
    Columns("G:G").ColumnWidth = 16.56
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "=R[-1]C+7"
    Range("A5").Select
    Selection.AutoFill Destination:=Range("A5:C5"), Type:=xlFillDefault
    Range("A5:C5").Select
    Range("C5").Select
    Selection.AutoFill Destination:=Range("C5:G5"), Type:=xlFillDefault
    Range("C5:G5").Select
    Range("A5:G5").Select
    Selection.AutoFill Destination:=Range("A5:G9"), Type:=xlFillDefault
    Range("A5:G9").Select
    Range("A4:G9").Select
    Selection.NumberFormat = "d"
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MONTH(A4)<>MONTH($A$2&1)"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Bold = True
        .Italic = True
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("A1:H12").Select
    With Selection.Font
        .Name = "Calibri"
        .Size = 18
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Rows("4:4").RowHeight = 51
    Rows("5:5").RowHeight = 50.4
    Rows("6:6").RowHeight = 54
    ActiveWindow.SmallScroll Down:=3
    Rows("7:7").RowHeight = 48
    Rows("8:8").RowHeight = 47.4
    Rows("9:9").RowHeight = 52.8
    Columns("B:B").ColumnWidth = 18.67
    Range("C4").Select
    ActiveWindow.SmallScroll Down:=-6
    Range("A2").Select
  With Range("A2").Validation
         .Delete
    .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Gennaio,Febbraio,Marzo,Aprile,Maggio,Giugno,Luglio,Agosto,Settembre,Ottobre,Novembre,Dicembre"
    .IgnoreBlank = True
    .InCellDropdown = True
    .InputTitle = ""
    .ErrorTitle = ""
    .InputMessage = ""
    .ErrorMessage = ""
    .ShowInput = True
    .ShowError = True
    End With

   With Selection
    .Font.Color = -16776961
    .Font.TintAndShade = 0
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = 5287936
    .Interior.TintAndShade = 0
    .Interior.PatternTintAndShade = 0
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlBottom
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
    End With
    Range("G4:G8").Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Range("A3:G9").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Range("A3:D9").Borders(xlEdgeLeft)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With

With Range("A3:D9").Borders(xlEdgeTop)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With

With Range("A3:D9").Borders(xlEdgeRight)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With

With Range("A3:D9").Borders(xlEdgeBottom)
    .LineStyle = xlContinuous
    .ColorIndex = xlAutomatic
    .TintAndShade = 0
    .Weight = xlThin
End With
End Sub

' calc code
Sub CalcCalendar()

rem ----------------------------------------------------------------------
rem define variables
Dim document As Object
Dim activeCell As Object

document = ThisComponent.CurrentController.Frame
activeCell = ThisComponent.CurrentController.CurrentSelection


' ----------------------------------------------------------------------

dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())


rem ----------------------------------------------------------------------
rem documentarsi sintassi vb 
dim cellRange as Object
cellRange = document.Sheets(0).getCellRangeByName("A3:G3")
cellRange.String = "lunedi,martedi,mercoledi,giovedi,venerdi,sabato,domenica"





rem ----------------------------------------------------------------------
dim args18(0) as new com.sun.star.beans.PropertyValue
args18(0).Name = "ToPoint"
args18(0).Value = "$A$2"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args18())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args20(1) as new com.sun.star.beans.PropertyValue
args20(0).Name = "By"
args20(0).Value = 1
args20(1).Name = "Sel"
args20(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args20())

rem ----------------------------------------------------------------------
dim args21(0) as new com.sun.star.beans.PropertyValue
args21(0).Name = "StringName"
args21(0).Value = "gennaio"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args21())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args23(1) as new com.sun.star.beans.PropertyValue
args23(0).Name = "By"
args23(0).Value = 1
args23(1).Name = "Sel"
args23(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args23())

rem ----------------------------------------------------------------------
dim args24(1) as new com.sun.star.beans.PropertyValue
args24(0).Name = "By"
args24(0).Value = 1
args24(1).Name = "Sel"
args24(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args24())

rem ----------------------------------------------------------------------
dim args25(0) as new com.sun.star.beans.PropertyValue
args25(0).Name = "StringName"
args25(0).Value = "=data(G2;1;1)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args25())

rem ----------------------------------------------------------------------
dim args26(0) as new com.sun.star.beans.PropertyValue
args26(0).Name = "ToPoint"
args26(0).Value = "$A$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args26())

rem ----------------------------------------------------------------------
dim args27(0) as new com.sun.star.beans.PropertyValue
args27(0).Name = "ToPoint"
args27(0).Value = "$A$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args27())

rem ----------------------------------------------------------------------
dim args28(0) as new com.sun.star.beans.PropertyValue
args28(0).Name = "EndCell"
args28(0).Value = "$G$4"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args28())

rem ----------------------------------------------------------------------
dim args29(0) as new com.sun.star.beans.PropertyValue
args29(0).Name = "ToPoint"
args29(0).Value = "$A$4:$G$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args29())

rem ----------------------------------------------------------------------
dim args30(0) as new com.sun.star.beans.PropertyValue
args30(0).Name = "ToPoint"
args30(0).Value = "$D$9"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args30())

rem ----------------------------------------------------------------------
dim args31(0) as new com.sun.star.beans.PropertyValue
args31(0).Name = "ToPoint"
args31(0).Value = "$B$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args31())

rem ----------------------------------------------------------------------
dim args32(0) as new com.sun.star.beans.PropertyValue
args32(0).Name = "StringName"
args32(0).Value = "=A4+1"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args32())

rem ----------------------------------------------------------------------
dim args33(0) as new com.sun.star.beans.PropertyValue
args33(0).Name = "ToPoint"
args33(0).Value = "$B$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args33())

rem ----------------------------------------------------------------------
dim args34(0) as new com.sun.star.beans.PropertyValue
args34(0).Name = "EndCell"
args34(0).Value = "$G$4"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args34())

rem ----------------------------------------------------------------------
dim args35(0) as new com.sun.star.beans.PropertyValue
args35(0).Name = "ToPoint"
args35(0).Value = "$B$4:$G$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args35())

rem ----------------------------------------------------------------------
dim args36(0) as new com.sun.star.beans.PropertyValue
args36(0).Name = "ToPoint"
args36(0).Value = "$A$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args36())

rem ----------------------------------------------------------------------
dim args37(0) as new com.sun.star.beans.PropertyValue
args37(0).Name = "StringName"
args37(0).Value = "=DATA(G2;1;1)-5"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args37())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:Undo", "", 0, Array())

rem ----------------------------------------------------------------------
dim args39(0) as new com.sun.star.beans.PropertyValue
args39(0).Name = "ToPoint"
args39(0).Value = "$A$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args39())

rem ----------------------------------------------------------------------
dim args40(0) as new com.sun.star.beans.PropertyValue
args40(0).Name = "StringName"
args40(0).Value = "=DATA(G2;1;27)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args40())

rem ----------------------------------------------------------------------
dim args41(0) as new com.sun.star.beans.PropertyValue
args41(0).Name = "ToPoint"
args41(0).Value = "$A$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args41())

rem ----------------------------------------------------------------------
dim args42(0) as new com.sun.star.beans.PropertyValue
args42(0).Name = "EndCell"
args42(0).Value = "$A$5"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args42())

rem ----------------------------------------------------------------------
dim args43(0) as new com.sun.star.beans.PropertyValue
args43(0).Name = "ToPoint"
args43(0).Value = "$A$4:$A$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args43())

rem ----------------------------------------------------------------------
dim args44(0) as new com.sun.star.beans.PropertyValue
args44(0).Name = "ToPoint"
args44(0).Value = "$A$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args44())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())

rem ----------------------------------------------------------------------
dim args46(1) as new com.sun.star.beans.PropertyValue
args46(0).Name = "By"
args46(0).Value = 1
args46(1).Name = "Sel"
args46(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args46())

rem ----------------------------------------------------------------------
dim args47(0) as new com.sun.star.beans.PropertyValue
args47(0).Name = "StringName"
args47(0).Value = "=$A$4+1"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args47())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args49(1) as new com.sun.star.beans.PropertyValue
args49(0).Name = "By"
args49(0).Value = 1
args49(1).Name = "Sel"
args49(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args49())

rem ----------------------------------------------------------------------
dim args50(0) as new com.sun.star.beans.PropertyValue
args50(0).Name = "ToPoint"
args50(0).Value = "$A$6"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args50())

rem ----------------------------------------------------------------------
dim args51(0) as new com.sun.star.beans.PropertyValue
args51(0).Name = "ToPoint"
args51(0).Value = "$A$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args51())

rem ----------------------------------------------------------------------
dim args52(0) as new com.sun.star.beans.PropertyValue
args52(0).Name = "StringName"
args52(0).Value = "=A4+2"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args52())

rem ----------------------------------------------------------------------
dim args53(1) as new com.sun.star.beans.PropertyValue
args53(0).Name = "By"
args53(0).Value = 1
args53(1).Name = "Sel"
args53(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoUp", "", 0, args53())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())

rem ----------------------------------------------------------------------
dim args55(0) as new com.sun.star.beans.PropertyValue
args55(0).Name = "StringName"
args55(0).Value = ""

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args55())

rem ----------------------------------------------------------------------
dim args56(0) as new com.sun.star.beans.PropertyValue
args56(0).Name = "ToPoint"
args56(0).Value = "$B$10"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args56())

rem ----------------------------------------------------------------------
dim args57(0) as new com.sun.star.beans.PropertyValue
args57(0).Name = "ToPoint"
args57(0).Value = "$D$7"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args57())

rem ----------------------------------------------------------------------
dim args58(0) as new com.sun.star.beans.PropertyValue
args58(0).Name = "ToPoint"
args58(0).Value = "$A$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args58())

rem ----------------------------------------------------------------------
dim args59(0) as new com.sun.star.beans.PropertyValue
args59(0).Name = "StringName"
args59(0).Value = "=DATA(G2;12;27)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args59())

rem ----------------------------------------------------------------------
dim args60(0) as new com.sun.star.beans.PropertyValue
args60(0).Name = "ToPoint"
args60(0).Value = "$A$4"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args60())

rem ----------------------------------------------------------------------
dim args61(0) as new com.sun.star.beans.PropertyValue
args61(0).Name = "StringName"
args61(0).Value = "=DATA(2021;12;27)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args61())

rem ----------------------------------------------------------------------
dim args62(0) as new com.sun.star.beans.PropertyValue
args62(0).Name = "StringName"
args62(0).Value = "=G4+1"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args62())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args64(0) as new com.sun.star.beans.PropertyValue
args64(0).Name = "ToPoint"
args64(0).Value = "$A$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args64())

rem ----------------------------------------------------------------------
dim args65(0) as new com.sun.star.beans.PropertyValue
args65(0).Name = "EndCell"
args65(0).Value = "$G$5"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args65())

rem ----------------------------------------------------------------------
dim args66(0) as new com.sun.star.beans.PropertyValue
args66(0).Name = "ToPoint"
args66(0).Value = "$A$5:$G$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args66())

rem ----------------------------------------------------------------------
dim args67(0) as new com.sun.star.beans.PropertyValue
args67(0).Name = "ToPoint"
args67(0).Value = "$B$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args67())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextUnprotected", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextUnprotected", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextUnprotected", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextUnprotected", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextUnprotected", "", 0, Array())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:ClearContents", "", 0, Array())

rem ----------------------------------------------------------------------
dim args79(0) as new com.sun.star.beans.PropertyValue
args79(0).Name = "ToPoint"
args79(0).Value = "$B$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args79())

rem ----------------------------------------------------------------------
dim args80(1) as new com.sun.star.beans.PropertyValue
args80(0).Name = "By"
args80(0).Value = 1
args80(1).Name = "Sel"
args80(1).Value = false

dispatcher.executeDispatch(document, ".uno:GoLeft", "", 0, args80())

rem ----------------------------------------------------------------------
dim args81(0) as new com.sun.star.beans.PropertyValue
args81(0).Name = "StringName"
args81(0).Value = "=A5+1"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args81())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:JumpToNextCell", "", 0, Array())

rem ----------------------------------------------------------------------
dim args83(0) as new com.sun.star.beans.PropertyValue
args83(0).Name = "ToPoint"
args83(0).Value = "$B$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args83())

rem ----------------------------------------------------------------------
dim args84(0) as new com.sun.star.beans.PropertyValue
args84(0).Name = "EndCell"
args84(0).Value = "$G$5"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args84())

rem ----------------------------------------------------------------------
dim args85(0) as new com.sun.star.beans.PropertyValue
args85(0).Name = "ToPoint"
args85(0).Value = "$B$5:$G$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args85())

rem ----------------------------------------------------------------------
dim args86(0) as new com.sun.star.beans.PropertyValue
args86(0).Name = "ToPoint"
args86(0).Value = "$A$5:$G$5"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args86())

rem ----------------------------------------------------------------------
dim args87(0) as new com.sun.star.beans.PropertyValue
args87(0).Name = "EndCell"
args87(0).Value = "$G$9"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args87())

rem ----------------------------------------------------------------------
dim args88(0) as new com.sun.star.beans.PropertyValue
args88(0).Name = "ToPoint"
args88(0).Value = "$A$5:$G$9"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args88())

rem ----------------------------------------------------------------------
dim args89(0) as new com.sun.star.beans.PropertyValue
args89(0).Name = "ToPoint"
args89(0).Value = "$B$14"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args89())

rem ----------------------------------------------------------------------
dim args90(0) as new com.sun.star.beans.PropertyValue
args90(0).Name = "ToPoint"
args90(0).Value = "$A$4:$G$9"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args90())

rem ----------------------------------------------------------------------
dim args91(0) as new com.sun.star.beans.PropertyValue
args91(0).Name = "ToPoint"
args91(0).Value = "$A$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args91())

rem ----------------------------------------------------------------------
dim args92(0) as new com.sun.star.beans.PropertyValue
args92(0).Name = "ToPoint"
args92(0).Value = "$A$1:$H$11"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args92())

rem ----------------------------------------------------------------------
dim args93(2) as new com.sun.star.beans.PropertyValue
args93(0).Name = "FontHeight.Height"
args93(0).Value = 18
args93(1).Name = "FontHeight.Prop"
args93(1).Value = 100
args93(2).Name = "FontHeight.Diff"
args93(2).Value = 0

dispatcher.executeDispatch(document, ".uno:FontHeight", "", 0, args93())

rem ----------------------------------------------------------------------
dim args94(0) as new com.sun.star.beans.PropertyValue
args94(0).Name = "ToPoint"
args94(0).Value = "$E$11"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args94())

rem ----------------------------------------------------------------------
dim args95(0) as new com.sun.star.beans.PropertyValue
args95(0).Name = "ToPoint"
args95(0).Value = "$A$4:$G$9"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args95())

rem ----------------------------------------------------------------------
dim args96(0) as new com.sun.star.beans.PropertyValue
args96(0).Name = "NumberFormatValue"
args96(0).Value = 104

dispatcher.executeDispatch(document, ".uno:NumberFormatValue", "", 0, args96())

rem ----------------------------------------------------------------------
dim args97(0) as new com.sun.star.beans.PropertyValue
args97(0).Name = "ToPoint"
args97(0).Value = "$E$19"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args97())
end sub







