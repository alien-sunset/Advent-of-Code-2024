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
dispatcher.executeDispatch(document, ".uno:PasteTextImportDialog", "", 0, Array())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, Array())

rem ----------------------------------------------------------------------
dim args3(9) as new com.sun.star.beans.PropertyValue
args3(0).Name = "ByRows"
args3(0).Value = true
args3(1).Name = "HasHeader"
args3(1).Value = false
args3(2).Name = "CaseSensitive"
args3(2).Value = false
args3(3).Name = "NaturalSort"
args3(3).Value = false
args3(4).Name = "IncludeAttribs"
args3(4).Value = true
args3(5).Name = "UserDefIndex"
args3(5).Value = 0
args3(6).Name = "Col1"
args3(6).Value = 1
args3(7).Name = "Ascending1"
args3(7).Value = true
args3(8).Name = "IncludeComments"
args3(8).Value = false
args3(9).Name = "IncludeImages"
args3(9).Value = true

dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, args3())

rem ----------------------------------------------------------------------
rem dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, Array())

rem ----------------------------------------------------------------------
dim args5(9) as new com.sun.star.beans.PropertyValue
args5(0).Name = "ByRows"
args5(0).Value = true
args5(1).Name = "HasHeader"
args5(1).Value = false
args5(2).Name = "CaseSensitive"
args5(2).Value = false
args5(3).Name = "NaturalSort"
args5(3).Value = false
args5(4).Name = "IncludeAttribs"
args5(4).Value = true
args5(5).Name = "UserDefIndex"
args5(5).Value = 0
args5(6).Name = "Col1"
args5(6).Value = 2
args5(7).Name = "Ascending1"
args5(7).Value = true
args5(8).Name = "IncludeComments"
args5(8).Value = false
args5(9).Name = "IncludeImages"
args5(9).Value = true

dispatcher.executeDispatch(document, ".uno:DataSort", "", 0, args5())

rem ----------------------------------------------------------------------
dim args6(0) as new com.sun.star.beans.PropertyValue
args6(0).Name = "ToPoint"
args6(0).Value = "$C$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args6())

rem ----------------------------------------------------------------------
dim args7(0) as new com.sun.star.beans.PropertyValue
args7(0).Name = "StringName"
args7(0).Value = "=MAX(A1:B1)-MIN(A1:B1)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args7())

rem ----------------------------------------------------------------------
dim args8(0) as new com.sun.star.beans.PropertyValue
args8(0).Name = "ToPoint"
args8(0).Value = "$C$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args8())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args10(0) as new com.sun.star.beans.PropertyValue
args10(0).Name = "ToPoint"
args10(0).Value = "$C$2"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args10())

rem ----------------------------------------------------------------------
dim args11(0) as new com.sun.star.beans.PropertyValue
args11(0).Name = "ToPoint"
args11(0).Value = "$C$2:$C$1000"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args11())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim args13(0) as new com.sun.star.beans.PropertyValue
args13(0).Name = "ToPoint"
args13(0).Value = "$C$1002"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args13())

rem ----------------------------------------------------------------------
dim args14(0) as new com.sun.star.beans.PropertyValue
args14(0).Name = "ToPoint"
args14(0).Value = "$C$1001"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args14())

rem ----------------------------------------------------------------------
dim args15(0) as new com.sun.star.beans.PropertyValue
args15(0).Name = "StringName"
args15(0).Value = "=SUM(c1:c1000)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args15())

rem ----------------------------------------------------------------------
dim args16(0) as new com.sun.star.beans.PropertyValue
args16(0).Name = "ToPoint"
args16(0).Value = "$D$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args16())

rem ----------------------------------------------------------------------
dim args17(0) as new com.sun.star.beans.PropertyValue
args17(0).Name = "StringName"
args17(0).Value = "=COUNTIF(B$1:B$1000,A1)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args17())

rem ----------------------------------------------------------------------
dim args18(0) as new com.sun.star.beans.PropertyValue
args18(0).Name = "ToPoint"
args18(0).Value = "$D$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args18())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args20(0) as new com.sun.star.beans.PropertyValue
args20(0).Name = "ToPoint"
args20(0).Value = "$D$2"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args20())

rem ----------------------------------------------------------------------
dim args21(0) as new com.sun.star.beans.PropertyValue
args21(0).Name = "ToPoint"
args21(0).Value = "$D$2:$D$1000"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args21())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim args23(0) as new com.sun.star.beans.PropertyValue
args23(0).Name = "ToPoint"
args23(0).Value = "$D$1001"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args23())

rem ----------------------------------------------------------------------
dim args24(0) as new com.sun.star.beans.PropertyValue
args24(0).Name = "ToPoint"
args24(0).Value = "$E$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args24())

rem ----------------------------------------------------------------------
dim args25(0) as new com.sun.star.beans.PropertyValue
args25(0).Name = "StringName"
args25(0).Value = "=(A1*D1)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args25())

rem ----------------------------------------------------------------------
dim args26(0) as new com.sun.star.beans.PropertyValue
args26(0).Name = "ToPoint"
args26(0).Value = "$E$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args26())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Copy", "", 0, Array())

rem ----------------------------------------------------------------------
dim args28(0) as new com.sun.star.beans.PropertyValue
args28(0).Name = "ToPoint"
args28(0).Value = "$E$2"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args28())

rem ----------------------------------------------------------------------
dim args29(0) as new com.sun.star.beans.PropertyValue
args29(0).Name = "ToPoint"
args29(0).Value = "$E$2:$E$1000"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args29())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
dim args31(0) as new com.sun.star.beans.PropertyValue
args31(0).Name = "ToPoint"
args31(0).Value = "$E$1001"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args31())

rem ----------------------------------------------------------------------
dim args32(0) as new com.sun.star.beans.PropertyValue
args32(0).Name = "StringName"
args32(0).Value = "=SUM(e1:e1000)"

dispatcher.executeDispatch(document, ".uno:EnterString", "", 0, args32())


end sub
