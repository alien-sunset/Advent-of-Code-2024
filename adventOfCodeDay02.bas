REM  *****  BASIC  *****



sub Main
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:PasteUnformatted", "", 0, Array())

rem ----------------------------------------------------------------------
dim args2(0) as new com.sun.star.beans.PropertyValue
args2(0).Name = "ToPoint"
args2(0).Value = "$A$1:$H$1001"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args2())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, Array())

rem ----------------------------------------------------------------------
dim args4(9) as new com.sun.star.beans.PropertyValue
args4(0).Name = "ByRows"
args4(0).Value = true
args4(1).Name = "HasHeader"
args4(1).Value = false
args4(2).Name = "CaseSensitive"
args4(2).Value = false
args4(3).Name = "NaturalSort"
args4(3).Value = false
args4(4).Name = "IncludeAttribs"
args4(4).Value = true
args4(5).Name = "UserDefIndex"
args4(5).Value = 0
args4(6).Name = "Col1"
args4(6).Value = 6
args4(7).Name = "Ascending1"
args4(7).Value = true
args4(8).Name = "IncludeComments"
args4(8).Value = false
args4(9).Name = "IncludeImages"
args4(9).Value = true

dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, args4())

rem ----------------------------------------------------------------------
dim args5(0) as new com.sun.star.beans.PropertyValue
args5(0).Name = "ToPoint"
args5(0).Value = "$G$744"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args5())

rem ----------------------------------------------------------------------
dim args6(0) as new com.sun.star.beans.PropertyValue
args6(0).Name = "ToPoint"
args6(0).Value = "$A$1:$H$744"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args6())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, Array())

rem ----------------------------------------------------------------------
dim args8(9) as new com.sun.star.beans.PropertyValue
args8(0).Name = "ByRows"
args8(0).Value = true
args8(1).Name = "HasHeader"
args8(1).Value = false
args8(2).Name = "CaseSensitive"
args8(2).Value = false
args8(3).Name = "NaturalSort"
args8(3).Value = false
args8(4).Name = "IncludeAttribs"
args8(4).Value = true
args8(5).Name = "UserDefIndex"
args8(5).Value = 0
args8(6).Name = "Col1"
args8(6).Value = 7
args8(7).Name = "Ascending1"
args8(7).Value = true
args8(8).Name = "IncludeComments"
args8(8).Value = false
args8(9).Name = "IncludeImages"
args8(9).Value = true

dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, args8())

rem ----------------------------------------------------------------------
dim args9(0) as new com.sun.star.beans.PropertyValue
args9(0).Name = "ToPoint"
args9(0).Value = "$H$479"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args9())

rem ----------------------------------------------------------------------
dim args10(0) as new com.sun.star.beans.PropertyValue
args10(0).Name = "ToPoint"
args10(0).Value = "$A$1:$H$479"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args10())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, Array())

rem ----------------------------------------------------------------------
dim args12(9) as new com.sun.star.beans.PropertyValue
args12(0).Name = "ByRows"
args12(0).Value = true
args12(1).Name = "HasHeader"
args12(1).Value = false
args12(2).Name = "CaseSensitive"
args12(2).Value = false
args12(3).Name = "NaturalSort"
args12(3).Value = false
args12(4).Name = "IncludeAttribs"
args12(4).Value = true
args12(5).Name = "UserDefIndex"
args12(5).Value = 0
args12(6).Name = "Col1"
args12(6).Value = 8
args12(7).Name = "Ascending1"
args12(7).Value = true
args12(8).Name = "IncludeComments"
args12(8).Value = false
args12(9).Name = "IncludeImages"
args12(9).Value = true

dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, args12())

rem ----------------------------------------------------------------------
dim args13(0) as new com.sun.star.beans.PropertyValue
args13(0).Name = "ToPoint"
args13(0).Value = "$J$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args13())

rem ----------------------------------------------------------------------
dim args14(0) as new com.sun.star.beans.PropertyValue
args14(0).Name = "StringName"
args14(0).Value = "=IF(AND(A1>B1,B1>C1,C1>D1,D1>E1,E1>F1,F1>G1,G1>H1) OR(AND(A1<B1,B1<C1,C1<D1,D1<E1,E1<F1,F1<G1,G1<H1)))"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args14())

rem ----------------------------------------------------------------------
dim args15(0) as new com.sun.star.beans.PropertyValue
args15(0).Name = "ToPoint"
args15(0).Value = "$J$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args15())

rem ----------------------------------------------------------------------
dim args16(0) as new com.sun.star.beans.PropertyValue
args16(0).Name = "EndCell"
args16(0).Value = "$J$249"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args16())

rem ----------------------------------------------------------------------
dim args17(0) as new com.sun.star.beans.PropertyValue
args17(0).Name = "ToPoint"
args17(0).Value = "$J$1:$J$249"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args17())

rem ----------------------------------------------------------------------
dim args18(0) as new com.sun.star.beans.PropertyValue
args18(0).Name = "ToPoint"
args18(0).Value = "$J$249"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args18())

rem ----------------------------------------------------------------------
dim args19(0) as new com.sun.star.beans.PropertyValue
args19(0).Name = "StringName"
args19(0).Value = "=IF(AND(A249>B249,B249>C249,C249>D249,D249>E249,E249>F249,F249>G249) OR(AND(A249<B249,B249<C249,C249<D249,D249<E249,E249<F249,F249<G249)))"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args19())

rem ----------------------------------------------------------------------
dim args20(0) as new com.sun.star.beans.PropertyValue
args20(0).Name = "ToPoint"
args20(0).Value = "$J$249"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args20())

rem ----------------------------------------------------------------------
dim args21(0) as new com.sun.star.beans.PropertyValue
args21(0).Name = "EndCell"
args21(0).Value = "$J$255"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args21())

rem ----------------------------------------------------------------------
dim args22(0) as new com.sun.star.beans.PropertyValue
args22(0).Name = "ToPoint"
args22(0).Value = "$J$249:$J$255"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args22())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:Undo", "", 0, Array())

rem ----------------------------------------------------------------------
dim args24(0) as new com.sun.star.beans.PropertyValue
args24(0).Name = "ToPoint"
args24(0).Value = "$J$249"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args24())

rem ----------------------------------------------------------------------
dim args25(0) as new com.sun.star.beans.PropertyValue
args25(0).Name = "EndCell"
args25(0).Value = "$J$490"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args25())

rem ----------------------------------------------------------------------
dim args26(0) as new com.sun.star.beans.PropertyValue
args26(0).Name = "ToPoint"
args26(0).Value = "$J$249:$J$490"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args26())

rem ----------------------------------------------------------------------
dim args27(0) as new com.sun.star.beans.PropertyValue
args27(0).Name = "ToPoint"
args27(0).Value = "$J$490"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args27())

rem ----------------------------------------------------------------------
dim args28(0) as new com.sun.star.beans.PropertyValue
args28(0).Name = "StringName"
args28(0).Value = "=IF(AND(A490>B490,B490>C490,C490>D490,D490>E490,E490>F490) OR(AND(A490<B490,B490<C490,C490<D490,D490<E490,E490<F490)))"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args28())

rem ----------------------------------------------------------------------
dim args29(0) as new com.sun.star.beans.PropertyValue
args29(0).Name = "ToPoint"
args29(0).Value = "$J$490"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args29())

rem ----------------------------------------------------------------------
dim args30(0) as new com.sun.star.beans.PropertyValue
args30(0).Name = "EndCell"
args30(0).Value = "$J$745"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args30())

rem ----------------------------------------------------------------------
dim args31(0) as new com.sun.star.beans.PropertyValue
args31(0).Name = "ToPoint"
args31(0).Value = "$J$490:$J$745"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args31())

rem ----------------------------------------------------------------------
dim args32(0) as new com.sun.star.beans.PropertyValue
args32(0).Name = "ToPoint"
args32(0).Value = "$J$745"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args32())

rem ----------------------------------------------------------------------
dim args33(0) as new com.sun.star.beans.PropertyValue
args33(0).Name = "StringName"
args33(0).Value = "=IF(AND(A745>B745,B745>C745,C745>D745,D745>E745) OR(AND(A745<B745,B745<C745,C745<D745,D745<E745)))"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args33())

rem ----------------------------------------------------------------------
dim args34(0) as new com.sun.star.beans.PropertyValue
args34(0).Name = "ToPoint"
args34(0).Value = "$J$745"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args34())

rem ----------------------------------------------------------------------
dim args35(0) as new com.sun.star.beans.PropertyValue
args35(0).Name = "EndCell"
args35(0).Value = "$J$1000"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args35())

rem ----------------------------------------------------------------------
dim args36(0) as new com.sun.star.beans.PropertyValue
args36(0).Name = "ToPoint"
args36(0).Value = "$J$745:$J$1000"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args36())

rem ----------------------------------------------------------------------
dim args37(0) as new com.sun.star.beans.PropertyValue
args37(0).Name = "ToPoint"
args37(0).Value = "$L$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args37())

rem ----------------------------------------------------------------------
dim args38(0) as new com.sun.star.beans.PropertyValue
args38(0).Name = "StringName"
args38(0).Value = "=IF(J1=1,AND((MAX($A1:$B1)-MIN($A1:$B1)<=3),(MAX($B1:$C1)-MIN($B1:$C1)<=3),(MAX($C1:$D1)-MIN($C1:$D1)<=3),(MAX($D1:$E1)-MIN($D1:$E1)<=3),(MAX($E1:$F1)-MIN($E1:$F1)<=3),(MAX($F1:$G1)-MIN($F1:$G1)<=3),(MAX($G1:$H1)-MIN($G1:$H1)<=3)))"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args38())

rem ----------------------------------------------------------------------
dim args39(0) as new com.sun.star.beans.PropertyValue
args39(0).Name = "ToPoint"
args39(0).Value = "$L$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args39())

rem ----------------------------------------------------------------------
dim args40(0) as new com.sun.star.beans.PropertyValue
args40(0).Name = "EndCell"
args40(0).Value = "$L$1000"

dispatcher.executeDispatch(document, ".uno:AutoFill", "", 0, args40())

rem ----------------------------------------------------------------------
dim args41(0) as new com.sun.star.beans.PropertyValue
args41(0).Name = "ToPoint"
args41(0).Value = "$L$1:$L$1000"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args41())

rem ----------------------------------------------------------------------
dim args42(0) as new com.sun.star.beans.PropertyValue
args42(0).Name = "ToPoint"
args42(0).Value = "$L$1004"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args42())

rem ----------------------------------------------------------------------
dim args43(0) as new com.sun.star.beans.PropertyValue
args43(0).Name = "ToPoint"
args43(0).Value = "$L$1002"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args43())

rem ----------------------------------------------------------------------
dim args44(0) as new com.sun.star.beans.PropertyValue
args44(0).Name = "StringName"
args44(0).Value = "=COUNTIF($L1:$L1000,"+CHR$(34)+"TRUE"+CHR$(34)+")"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args44())


end sub
