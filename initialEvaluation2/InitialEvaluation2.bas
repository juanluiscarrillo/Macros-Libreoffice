rem Main Function
Sub makeEvalDocument2
	Dim oSheet As Object
	oSheet = ThisComponent.Sheets.getByIndex(0)
	
	Dim url
	Dim path 
	path = getPath(ThisComponent.getURL())
	url = path+"template2IE.odt"

	textDoc = OpenTemplate(url)	
	
	Dim oCell As Object
	Dim pupilName
	Dim row 
	Dim course, letter, group
	Dim tutor
	Dim day, month, year
	
	course = oSheet.getCellRangeByName("B1").getString()
	letter = oSheet.getCellRangeByName("C1").getString()
	group = course + " " + letter
	tutor = oSheet.getCellRangeByName("G2").getString()
	day = oSheet.getCellRangeByName("B2").getString()
	month = oSheet.getCellRangeByName("C2").getString()
	year = oSheet.getCellRangeByName("D2").getString()
	
	row = 3

	pupilName = Trim(oSheet.getCellByPosition( 0,row ).getString())
	Do While pupilName <> ""		
		if row > 3 then
			insertNewPage(textDoc)
		end if
		insertPupil(textDoc, pupilName, group, tutor, day, month, year)	
		row = row + 1
		pupilName = Trim(oSheet.getCellByPosition( 0,row ).getString())
	Loop
	
	save(textDoc, path+"EvaluacionInicial-"+year+"-"+course+letter+".odt")		
End Sub


rem Returns path of the folder
Function getPath(fullPath)
	a = Split(fullPath, GetPathSeparator())
	r=UBound(a)
	a(r) = ""
	getPath = Join(a(), GetPathSeparator())
end Function


rem Inserts pupil info on a new page
Sub insertPupil(doc, pupilName, group, tutorName, day, month, year)
	insertTitle(doc)
	insertNewLine(doc)
	insertText(doc, "ESTIMADA FAMILIA:")
	insertNewLine(doc)	
	insertText(doc, "De acuerdo con el artículo 25, punto 4 de la Orden 2398/2016, de 22 de julio, por la que se regulan determinados aspectos de organización, funcionamiento y evaluación en la Educación Secundaria Obligatoria, en el que se indica que tras la Evaluación inicial se dará cuenta de los resultados de la misma a las familias, le enviamos la siguiente información sobre el/la alumno/a ")
	insertBoldText(doc, pupilName)
	insertText(doc, ", del grupo ")	
	insertBoldText(doc, group)			
	insertText(doc, " de ESO:")	
	insertTable(doc)
	insertText(doc, "Para que conste que ustedes han recibido esta información rogamos devuelvan firmada la parte inferior de esta hoja a la mayor brevedad posible al tutor/a del grupo.")	
	insertNewLine(doc)			
	insertFirm(doc,tutorName, day, month, year)					
End Sub


rem Opens text template to be copied on each page
Function openTemplate(filePath)
	dim dispatcher as object
    targetURL = filePath
    textDoc = StarDesktop.loadComponentFromURL(targetURL, "_blank", 0, Array())

	rem get access to the document
	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(document, ".uno:Cut", "", 1, Array())
    cutTable(textDoc)

	dispatcher.executeDispatch(document, ".uno:ParaspaceDecrease", "", 0, Array())	
    openTemplate = textDoc
End Function


rem Saves the generated file on the current folder
Sub save(textDoc, filePath)
	targetURL = filePath
	Dim fileProperties(0) As New com.sun.star.beans.PropertyValue
	fileProperties(0).Name = "Overwrite"
	fileProperties(0).Value = True
	textDoc.storeAsURL(targetURL, fileProperties())
End Sub


rem Extracts table from template
sub cutTable(textDoc)
	dim dispatcher as object

	rem gets access to the document
	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	dim args(1) as new com.sun.star.beans.PropertyValue
	args(0).Name = "Count"
	args(0).Value = 1
	args(1).Name = "Select"
	args(1).Value = true
	
	rem goes through the table
	For iContador = 0 To 7
		dispatcher.executeDispatch(document, ".uno:GoDown", "", 0, args())
	Next iContador
	
	For iContador = 0 To 3
		dispatcher.executeDispatch(document, ".uno:GoRight", "", 0, args())
	Next iContador		
	
	rem ----------------------------------------------------------------------
	dispatcher.executeDispatch(document, ".uno:Cut", "", 0, Array())
end sub


rem Inserts text into the document
sub insertText(textDoc, textValue)
	dim dispatcher as object

	rem get access to the document
	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem Sets space
	dispatcher.executeDispatch(document, ".uno:SpacePara1", "", 0, Array())
	
	rem Sets font
	dim args1(4) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "CharFontName.StyleName"
	args1(0).Value = ""
	args1(1).Name = "CharFontName.Pitch"
	args1(1).Value = 2
	args1(2).Name = "CharFontName.CharSet"
	args1(2).Value = 0
	args1(3).Name = "CharFontName.Family"
	args1(3).Value = 3
	args1(4).Name = "CharFontName.FamilyName"
	args1(4).Value = "Times New Roman"
	
	dispatcher.executeDispatch(document, ".uno:CharFontName", "", 0, args1())
	
	rem Sets font heigh
	dim args2(2) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "FontHeight.Height"
	args2(0).Value = 12
	args2(1).Name = "FontHeight.Prop"
	args2(1).Value = 100
	args2(2).Name = "FontHeight.Diff"
	args2(2).Value = 0
	
	dispatcher.executeDispatch(document, ".uno:FontHeight", "", 0, args2())
	
	rem Sets no bold
	dim args3(0) as new com.sun.star.beans.PropertyValue
	args3(0).Name = "Bold"
	args3(0).Value = false
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args3())
	
	rem Sets font color
	dim args4(2) as new com.sun.star.beans.PropertyValue
	args4(0).Name = "Underline.LineStyle"
	args4(0).Value = 0
	args4(1).Name = "Underline.HasColor"
	args4(1).Value = false
	args4(2).Name = "Underline.Color"
	args4(2).Value = -1
	
	rem Sets justify style
	dim args5(0) as new com.sun.star.beans.PropertyValue
	args5(0).Name = "JustifyPara"
	args5(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:JustifyPara", "", 0, args5())
	
	
	rem Sets the text
	dim args6(0) as new com.sun.star.beans.PropertyValue
	args6(0).Name = "Text"
	args6(0).Value = textValue
	
	dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args6())
end sub


rem Inserts bold text into the document
sub insertBoldText(textDoc, textValue)
	dim dispatcher as object

	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem Sets space
	dispatcher.executeDispatch(document, ".uno:SpacePara1", "", 0, Array())
	
	rem Sets font
	dim args1(4) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "CharFontName.StyleName"
	args1(0).Value = ""
	args1(1).Name = "CharFontName.Pitch"
	args1(1).Value = 2
	args1(2).Name = "CharFontName.CharSet"
	args1(2).Value = 0
	args1(3).Name = "CharFontName.Family"
	args1(3).Value = 3
	args1(4).Name = "CharFontName.FamilyName"
	args1(4).Value = "Times New Roman"
	
	dispatcher.executeDispatch(document, ".uno:CharFontName", "", 0, args1())
	
	rem Sets font heigh
	dim args2(2) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "FontHeight.Height"
	args2(0).Value = 12
	args2(1).Name = "FontHeight.Prop"
	args2(1).Value = 100
	args2(2).Name = "FontHeight.Diff"
	args2(2).Value = 0
	
	dispatcher.executeDispatch(document, ".uno:FontHeight", "", 0, args2())
	
	rem Sets bold
	dim args3(0) as new com.sun.star.beans.PropertyValue
	args3(0).Name = "Bold"
	args3(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args3())
	
	rem Sets font color
	dim args4(2) as new com.sun.star.beans.PropertyValue
	args4(0).Name = "Underline.LineStyle"
	args4(0).Value = 0
	args4(1).Name = "Underline.HasColor"
	args4(1).Value = false
	args4(2).Name = "Underline.Color"
	args4(2).Value = -1
	
	rem Sets justify style
	dim args5(0) as new com.sun.star.beans.PropertyValue
	args5(0).Name = "JustifyPara"
	args5(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:JustifyPara", "", 0, args5())
	
	
	rem Sets text into doc file
	dim args6(0) as new com.sun.star.beans.PropertyValue
	args6(0).Name = "Text"
	args6(0).Value = textValue
	
	dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args6())
end sub


rem Inserts the title on the doc page
sub insertTitle(textDoc)
	dim dispatcher as object

	rem get access to the document
	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem Sets space
	dispatcher.executeDispatch(document, ".uno:SpacePara1", "", 0, Array())
	
	rem Sets font
	dim args1(4) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "CharFontName.StyleName"
	args1(0).Value = ""
	args1(1).Name = "CharFontName.Pitch"
	args1(1).Value = 2
	args1(2).Name = "CharFontName.CharSet"
	args1(2).Value = 0
	args1(3).Name = "CharFontName.Family"
	args1(3).Value = 3
	args1(4).Name = "CharFontName.FamilyName"
	args1(4).Value = "Times New Roman"
	
	dispatcher.executeDispatch(document, ".uno:CharFontName", "", 0, args1())
	
	rem Sets font heigh
	dim args2(2) as new com.sun.star.beans.PropertyValue
	args2(0).Name = "FontHeight.Height"
	args2(0).Value = 14
	args2(1).Name = "FontHeight.Prop"
	args2(1).Value = 100
	args2(2).Name = "FontHeight.Diff"
	args2(2).Value = 0
	
	dispatcher.executeDispatch(document, ".uno:FontHeight", "", 0, args2())
	
	rem Sets bold
	dim args5(0) as new com.sun.star.beans.PropertyValue
	args5(0).Name = "Bold"
	args5(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:Bold", "", 0, args5())
	
	rem Sets font color
	dim args6(2) as new com.sun.star.beans.PropertyValue
	args6(0).Name = "Underline.LineStyle"
	args6(0).Value = 1
	args6(1).Name = "Underline.HasColor"
	args6(1).Value = false
	args6(2).Name = "Underline.Color"
	args6(2).Value = -1
	
	dispatcher.executeDispatch(document, ".uno:Underline", "", 0, args6())
	
	rem Sets center style
	dim args10(0) as new com.sun.star.beans.PropertyValue
	args10(0).Name = "CenterPara"
	args10(0).Value = true
	
	dispatcher.executeDispatch(document, ".uno:CenterPara", "", 0, args10())
	
	
	rem Sets title text
	dim args7(0) as new com.sun.star.beans.PropertyValue
	args7(0).Name = "Text"
	args7(0).Value = "RESULTADOS DE LA EVALUACIÓN INICIAL"
	
	dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args7())
	
	rem Sets back color style
	dim args8(2) as new com.sun.star.beans.PropertyValue
	args8(0).Name = "Underline.LineStyle"
	args8(0).Value = 0
	args8(1).Name = "Underline.HasColor"
	args8(1).Value = false
	args8(2).Name = "Underline.Color"
	args8(2).Value = -1
	
	dispatcher.executeDispatch(document, ".uno:Underline", "", 0, args8())	
end sub


rem Inserts new blank line
Sub insertNewLine(textDoc)
	dim dispatcher as object

	rem get access to the document
	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

	rem Inserts new line
	dispatcher.executeDispatch(document, ".uno:InsertPara", "", 0, Array())
end sub


rem Inserts new page into document
sub insertNewPage(textDoc)
	dim dispatcher as object

	rem get access to the document
	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem Inserts new page
	dispatcher.executeDispatch(document, ".uno:InsertPagebreak", "", 0, Array())

end sub





rem Inserts the table to doc page
sub insertTable(textDoc)
	dim dispatcher as object

	rem get access to the document
	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem Pastes the table (it has been cutted in memory previosly) into doc
	dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())
end sub


rem Inserts the firm are into document
sub insertFirm(textDoc, firmName, day, month, year)
	dim document   as object
	dim dispatcher as object

	rem get access to the document
	document   = textDoc.CurrentController.Frame
	dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
	
	rem Inserts first part
	dim args1(0) as new com.sun.star.beans.PropertyValue
	args1(0).Name = "Text"
	args1(0).Value = CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+"Atentamente,"
	
	dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args1())
	
	For iContador = 0 To 7
		insertNewLine(textDoc)
	Next iContador

	
	rem Inserts second part
	dim args9(0) as new com.sun.star.beans.PropertyValue
	args9(0).Name = "Text"
	args9(0).Value = CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+"Fdo.: "+firmName
	
	dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args9())
	
	insertNewLine(textDoc)

	
	rem Inserts third part
	dim args10(0) as new com.sun.star.beans.PropertyValue
	args10(0).Name = "Text"
	args10(0).Value = CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+CHR$(9)+"Madrid, a "+day+" de "+month+" de "+year
	
	dispatcher.executeDispatch(document, ".uno:InsertText", "", 0, args10())
end sub
