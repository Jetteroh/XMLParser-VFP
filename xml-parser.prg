	*Written in Visual FoxPro 8.0
	*By Peter Stock, 2024
	
	******************************
	* Here comes the parse of XML
	******************************
	FUNCTION XMLParseToCursor
	LPARAMETERS pTableName, pXMLtext
	*pTableName		- DataTable name, XML made from 
	*pXMLtext		- the XML string (I got from C# Web API) what should de parsed
	*---------
	*Result is the error string;
	*	or "OK" and the Cursor table with same name as in pTableName
	*
	
	*StatusText		- Last status
	
	LOCAL loXmlAdapter, loXml
	LOCAL makro
	
	*IF there are no pTableName, we don't know what is the name of Node
	IF !EMPTY(pTableName)
		loXmlAdapter = CREATEOBJECT("XMLAdapter")		&& MSXML4 SP1 or newer
		*loXmlAdapter.IsDiffgram = .T.
		loXmlAdapter.LoadXML(STRCONV(pXMLtext,9), .F., .T.)			&& (from, isFile, shouldParse)
		*loXmlAdapter.LoadXML(cFile,.T.,.T.)					&& (from, isFile, shouldParse)	!Parse error, if no conversion to UTF-8
		IF loXmlAdapter.Tables.Count > 0
			WAIT WINDOW "Result table creation..." NOWAIT NOCLEAR
		    
			*Creating the CURSOR (shame, but "XMLAdapter" creates w/o data)
			IF USED(pTableName)						&& could not make CURSOR if it has opened
		    		USE IN (pTableName)
			ENDIF
	    		loXmlAdapter.Tables(1).ToCursor()		&& ... INTO CURSOR (pTableName)
		ELSE
			WAIT CLEAR
			StatusText = "There are no Table in the XML."
			
			RETURN StatusText
		ENDIF
	
		**********************************************
		*Loading of XML data into CURSOR we just made
		**********************************************
		loXml = CREATEOBJECT("MSXML2.DOMDocument.6.0")
		loXml.async = .F.
		loXml.LoadXML(pXMLtext)		&& STRCONV(pXMLtext,9) no need
		
		IF loXML.parseError.errorCode != 0
			WAIT CLEAR
       			StatusText = "XML format error: " + loXML.parseError.reason

       			RETURN StatusText
       		ELSE
			WAIT WINDOW "Result table data loading...    " NOWAIT NOCLEAR
			*Be sure, the CURSOR table already in selection
   			SELECT (pTableName)

   			loXML.setProperty("SelectionLanguage", "XPath") 
			*Table rows list   			
   			currNode = loXML.selectNodes("//" + pTableName)				&& find the Table node rows (objXMLDOMNodeList)
   			*? "Rows"
   			FOR i=0 TO currNode.length-1
   				*? ALLTRIM(STR(i))+".row:"
				WAIT WINDOW "Result table data loading..."+STR(i*100/currNode.length,3)+"%" NOWAIT NOCLEAR
				SCATTER MEMVAR BLANK

   				rowsColumns = currNode.item(i).selectNodes("*")
   				FOR j=0 TO rowsColumns.length-1
   					IF rowsColumns.item(j).specified
   						makro="m."+rowsColumns.item(j).nodeName+" = typeConversion(rowsColumns.item(j).nodeTypedValue, loXmlAdapter.Tables(1).Fields(j+1).DataType, loXmlAdapter.Tables(1).Fields(j+1).MaxLength)"
   						&makro
   					ENDIF
   				NEXT j

				APPEND BLANK IN (pTableName)
				GATHER MEMVAR
   			NEXT i
   			
			*SELECT (pTableName)
   			GO TOP IN (pTableName)

			*The CURSOR table already in selection
			RETURN "OK"
   		ENDIF
   	ENDIF
	StatusText = "Table name is emtpy!"
   	RETURN StatusText
   	ENDFUNC

	*------------------------------------------------------------------------------------------------

	FUNCTION typeConversion			&&(rowsColumns.item(j).nodeTypedValue, loXmlAdapter.Tables(1).Fields(j+1).DataType, loXmlAdapter.Tables(1).Fields(j+1).MaxLength)
	LPARAMETERS lcValue, lcType, lnHossz
	PRIVATE makro
	
	DO CASE
	CASE lcType = "C"
		RETURN lcValue
	CASE lcType = "T"
		*This is for a Hungarian style datetime: yyyy-mm-ddThh:mm:ss+hh:mm	(+01:00 GMT)
		makro = "DATETIME("+SUBSTR(lcValue,1,4)+","+SUBSTR(lcValue,6,2)+","+SUBSTR(lcValue,9,2)+",("+SUBSTR(lcValue,12,2)+SUBSTR(lcValue,20,3)+"),"+SUBSTR(lcValue,15,2)+","+SUBSTR(lcValue,18,2)+")"
		RETURN &makro
	CASE lcType = "D"
		RETURN CTOD(lcValue)
	CASE lcType = "B" AND lnHossz=8			&& float
		RETURN VAL(STRTRAN(lcValue,",","."))	&& VFP need decimal point: '.'
	CASE lcType = "L"
		RETURN IIF(lcValue="false", .F., .T.)
	OTHERWISE
		*Should handle if there are any other lcType!
		susp

		RETURN ""
	ENDCASE
	
	ENDFUNC
