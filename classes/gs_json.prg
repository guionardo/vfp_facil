*
* vfpjson
*
* ----------------------------------
* Ignacio Gutiérrez Torrero
* SAIT Software Administrativo
* www.sait.com.mx
* +52(653)534-8800
* Monterrey México
* -----------------------------------
*
* Libreria para el manejo de JSON en VFP
*
* Gracias a Google por el codigo de Json de Dart
* Thanks Google for the code in Json Dart
*     http://code.google.com/p/dart/source/browse/trunk/dart/lib/json/json.dart
*
*
* Codificar y Decodificar JSON
*
* Puedes usar las funciones:
*		json_encode(xExpr)	te regresa la cadena Json, que representa al objeto que se pasa como parametro
*		json_decode(cJson)  te regresa el objeto representado en cJson
*
* Tambien puedes usar directamente la clase:
*
*	oJson = newobject('json','json.prg')
*	oCliente = oJson.decode( ' { "nombre":"Ignacio" , "apellido":"Gutierrez", "edad":33 } ')
*	? oJson.encode(oCliente)
*	? oCliente.get('nombre')
*	? oCliente.get('apellido')
*
*
* VFPJSON  Encode and Decode JSON for VFP
* Examples:
*	oJson = newobject('json','json.prg')
*	oCustomer = oJson.decode( ' { "name":"Ignacio" , "lastname":"Gutierrez", "age":33 } ')
*	? oJson.encode(oCustomer)
*	? oCustomer.get('name')
*	? oCustomer.get('lastname')
*
*
lRunTest = .F.
IF lRunTest
	testJsonClass()
ENDIF
RETURN


FUNCTION json_encode(xExpr)
	IF VARTYPE(_json)<>'O'
		PUBLIC _json
		_json = NEWOBJECT('json')
	ENDIF
	RETURN _json.encode(@xExpr)


FUNCTION json_decode(cJson)
	LOCAL retval
	IF VARTYPE(_json)<>'O'
		PUBLIC _json
		_json = NEWOBJECT('json')
	ENDIF
	retval = _json.decode(cJson)
	IF NOT EMPTY(_json.cError)
		RETURN NULL
	ENDIF
	RETURN retval

FUNCTION json_getErrorMsg()
	RETURN _json.cError



	*
	* json class
	*
	*
DEFINE CLASS json AS CUSTOM


	nPos=0
	nLen=0
	cJson=''
	cError=''


	*
	* Genera el codigo cJson para parametro que se manda
	*
	FUNCTION encode(xExpr)
		LOCAL cTipo
		* Cuando se manda una arreglo,
		IF TYPE('ALen(xExpr)')=='N'
			cTipo = 'A'
		ELSE
			cTipo = VARTYPE(xExpr)
		ENDIF

		DO CASE
			CASE cTipo=='D'
				RETURN '"'+DTOS(xExpr)+'"'
			CASE cTipo=='N'
				RETURN TRANSFORM(xExpr)
			CASE cTipo=='L'
				RETURN IIF(xExpr,'true','false')
			CASE cTipo=='X'
				RETURN 'null'
			CASE cTipo=='C'
				xExpr = ALLT(xExpr)
				xExpr = STRTRAN(xExpr, '\', '\\' )
				xExpr = STRTRAN(xExpr, '/', '\/' )
				xExpr = STRTRAN(xExpr, CHR(13), '\n' )
				xExpr = STRTRAN(xExpr, CHR(10), '\f' )
				xExpr = STRTRAN(xExpr, '"', '\"' )
				RETURN '"'+xExpr+'"'

			CASE cTipo=='O'
				LOCAL cProp, cJsonValue, cRetVal, aProp[1]
				LOCAL lcExcludeProps
				IF UPPER(xExpr.class)="CUSTOM"
					m.lcExcludeProps = "baseclass,classlibrary,comment,controlcount,height,helpcontextid,left,name,parentclass,picture,tag,top,whatsthishelpid,width"
				ELSE 
				m.lcExcludeProps = ""
				endif
				=AMEMBERS(aProp,xExpr)
				cRetVal = ''
				FOR EACH cProp IN aProp
					*? cProp
					*? cRetVal
					IF TYPE('xExpr.'+cProp)=='U' OR cProp=='CLASS' OR LOWER(cProp)$m.lcExcludeProps
						* algunas propiedades pueden no estar definidas
						* como: activecontrol, parent, etc
						LOOP
					ENDIF
					IF TYPE( 'ALen(xExpr.'+cProp+')' ) == 'N'
						*
						* es un arreglo, recorrerlo usando los [ ] y eval
						*
						LOCAL i,nTotElem
						cJsonValue = ''
						nTotElem = EVAL('ALen(xExpr.'+cProp+')')
						FOR i=1 TO nTotElem
							cJsonValue = cJsonValue + ' , ' +  THIS.encode( EVAL( 'xExpr.'+cProp+'['+STR(i)+']' ))
						NEXT
						cJsonValue = '[' + SUBSTR(cJsonValue,4) + ']'
					ELSE
						*
						* es otro tipo de dato normal C, N, L
						*
						cJsonValue = THIS.encode( EVALUATE( 'xExpr.'+cProp ) )
					ENDIF
					IF LEFT(cProp,1)=='_'
						cProp = SUBSTR(cProp,2)
					ENDIF
					cRetVal = cRetVal + ',' + '"' + LOWER(cProp) + '":' + cJsonValue
				NEXT
				RETURN '{' + SUBSTR(cRetVal,2) + '}'

			CASE cTipo=='A'
				LOCAL valor, cRetVal
				cRetVal = ''
				FOR EACH valor IN xExpr
					cRetVal = cRetVal + ',' +  THIS.encode( @valor )
				NEXT
				RETURN  '[' + SUBSTR(cRetVal,2) + ']'

		ENDCASE

		RETURN ''





		*
		* regresa un elemento representado por la cadena json que se manda
		*

	FUNCTION decode(cJson)
		LOCAL retValue
		cJson = STRTRAN(cJson,CHR(9),'')
		cJson = STRTRAN(cJson,CHR(10),'')
		cJson = STRTRAN(cJson,CHR(13),'')
		cJson = THIS.fixUnicode(cJson)
		THIS.nPos  = 1
		THIS.cJson = cJson
		THIS.nLen  = LEN(cJson)
		THIS.cError = ''
		retValue = THIS.parsevalue()
		IF NOT EMPTY(THIS.cError)
			RETURN NULL
		ENDIF
		IF THIS.getToken()<>NULL
			THIS.setError('Junk at the end of JSON input')
			RETURN NULL
		ENDIF
		RETURN retValue


	FUNCTION parsevalue()
		LOCAL token
		token = THIS.getToken()
		IF token==NULL
			THIS.setError('Nothing to parse')
			RETURN NULL
		ENDIF
		DO CASE
			CASE token=='"'
				RETURN THIS.parseString()
			CASE ISDIGIT(token) OR token=='-'
				RETURN THIS.parseNumber()
			CASE token=='n'
				RETURN THIS.expectedKeyword('null',NULL)
			CASE token=='f'
				RETURN THIS.expectedKeyword('false',.F.)
			CASE token=='t'
				RETURN THIS.expectedKeyword('true',.T.)
			CASE token=='{'
				RETURN THIS.parseObject()
			CASE token=='['
				RETURN THIS.parseArray()
			OTHERWISE
				THIS.setError('Unexpected token')
		ENDCASE
		RETURN


	FUNCTION expectedKeyword(cWord,eValue)
		FOR i=1 TO LEN(cWord)
			cChar = THIS.getChar()
			IF cChar <> SUBSTR(cWord,i,1)
				THIS.setError("Expected keyword '" + cWord + "'")
				RETURN ''
			ENDIF
			THIS.nPos = THIS.nPos + 1
		NEXT
		RETURN eValue


	FUNCTION parseObject()
		LOCAL retval, cPropName, xValue
		retval = CREATEOBJECT('myObj')
		THIS.nPos = THIS.nPos + 1 && Eat {
		IF THIS.getToken()<>'}'
			DO WHILE .T.
				cPropName = THIS.parseString()
				IF NOT EMPTY(THIS.cError)
					RETURN NULL
				ENDIF
				IF THIS.getToken()<>':'
					THIS.setError("Expected ':' when parsing object")
					RETURN NULL
				ENDIF
				THIS.nPos = THIS.nPos + 1
				xValue = THIS.parsevalue()
				IF NOT EMPTY(THIS.cError)
					RETURN NULL
				ENDIF
				** Debug ? cPropName, type('xValue')
				retval.SET(cPropName, xValue)
				IF THIS.getToken()<>','
					EXIT
				ENDIF
				THIS.nPos = THIS.nPos + 1
			ENDDO
		ENDIF
		IF THIS.getToken()<>'}'
			THIS.setError("Expected '}' at the end of object")
			RETURN NULL
		ENDIF
		THIS.nPos = THIS.nPos + 1
		RETURN retval


	FUNCTION parseArray()
		LOCAL retval, xValue
		retval = CREATEOBJECT('MyArray')
		THIS.nPos = THIS.nPos + 1	&& Eat [
		IF THIS.getToken() <> ']'
			DO WHILE .T.
				xValue = THIS.parsevalue()
				IF NOT EMPTY(THIS.cError)
					RETURN NULL
				ENDIF
				retval.ADD( xValue )
				IF THIS.getToken()<>','
					EXIT
				ENDIF
				THIS.nPos = THIS.nPos + 1
			ENDDO
			IF THIS.getToken() <> ']'
				THIS.setError('Expected ] at the end of array')
				RETURN NULL
			ENDIF
		ENDIF
		THIS.nPos = THIS.nPos + 1
		RETURN retval


	FUNCTION parseString()
		LOCAL cRetVal, c
		IF THIS.getToken()<>'"'
			THIS.setError('Expected "')
			RETURN ''
		ENDIF
		THIS.nPos = THIS.nPos + 1 	&& Eat "
		cRetVal = ''
		DO WHILE .T.
			c = THIS.getChar()
			IF c==''
				RETURN ''
			ENDIF
			IF c == '"'
				THIS.nPos = THIS.nPos + 1
				EXIT
			ENDIF
			IF c == '\'
				THIS.nPos = THIS.nPos + 1
				IF (THIS.nPos>THIS.nLen)
					THIS.setError('\\ at the end of input')
					RETURN ''
				ENDIF
				c = THIS.getChar()
				IF c==''
					RETURN ''
				ENDIF
				DO CASE
					CASE c=='"'
						c='"'
					CASE c=='\'
						c='\'
					CASE c=='/'
						c='/'
					CASE c=='b'
						c=CHR(8)
					CASE c=='t'
						c=CHR(9)
					CASE c=='n'
						c=CHR(10)
					CASE c=='f'
						c=CHR(12)
					CASE c=='r'
						c=CHR(13)
					OTHERWISE
						******* FALTAN LOS UNICODE
						THIS.setError('Invalid escape sequence in string literal')
						RETURN ''
				ENDCASE
			ENDIF
			cRetVal = cRetVal + c
			THIS.nPos = THIS.nPos + 1
		ENDDO
		RETURN cRetVal


		**** Pendiente numeros con E
	FUNCTION parseNumber()
		LOCAL nStartPos,c, isInt, cNumero
		IF NOT ( ISDIGIT(THIS.getToken()) OR THIS.getToken()=='-')
			THIS.setError('Expected number literal')
			RETURN 0
		ENDIF
		nStartPos = THIS.nPos
		c = THIS.getChar()
		IF c == '-'
			c = THIS.nextChar()
		ENDIF
		IF c == '0'
			c = THIS.nextChar()
		ELSE
			IF ISDIGIT(c)
				c = THIS.nextChar()
				DO WHILE ISDIGIT(c)
					c = THIS.nextChar()
				ENDDO
			ELSE
				THIS.setError('Expected digit when parsing number')
				RETURN 0
			ENDIF
		ENDIF

		isInt = .T.
		IF c=='.'
			c = THIS.nextChar()
			IF ISDIGIT(c)
				c = THIS.nextChar()
				isInt = .F.
				DO WHILE ISDIGIT(c)
					c = THIS.nextChar()
				ENDDO
			ELSE
				THIS.setError('Expected digit following dot comma')
				RETURN 0
			ENDIF
		ENDIF

		cNumero = SUBSTR(THIS.cJson, nStartPos, THIS.nPos - nStartPos)
		RETURN VAL(cNumero)



	FUNCTION getToken()
		LOCAL char1
		DO WHILE .T.
			IF THIS.nPos > THIS.nLen
				RETURN NULL
			ENDIF
			char1 = SUBSTR(THIS.cJson, THIS.nPos, 1)
			IF char1==' '
				THIS.nPos = THIS.nPos + 1
				LOOP
			ENDIF
			RETURN char1
		ENDDO
		RETURN



	FUNCTION getChar()
		IF THIS.nPos > THIS.nLen
			THIS.setError('Unexpected end of JSON stream')
			RETURN ''
		ENDIF
		RETURN SUBSTR(THIS.cJson, THIS.nPos, 1)

	FUNCTION nextChar()
		THIS.nPos = THIS.nPos + 1
		IF THIS.nPos > THIS.nLen
			RETURN ''
		ENDIF
		RETURN SUBSTR(THIS.cJson, THIS.nPos, 1)

	FUNCTION setError(cMsg)
		THIS.cError= 'ERROR parsing JSON at Position:'+ALLT(STR(THIS.nPos,6,0))+' '+cMsg
		RETURN


	FUNCTION fixUnicode(cStr)
		cStr = STRTRAN(cStr,'\u00e1','á')
		cStr = STRTRAN(cStr,'\u00e9','é')
		cStr = STRTRAN(cStr,'\u00ed','í')
		cStr = STRTRAN(cStr,'\u00f3','ó')
		cStr = STRTRAN(cStr,'\u00fa','ú')
		cStr = STRTRAN(cStr,'\u00c1','Á')
		cStr = STRTRAN(cStr,'\u00c9','É')
		cStr = STRTRAN(cStr,'\u00cd','Í')
		cStr = STRTRAN(cStr,'\u00d3','Ó')
		cStr = STRTRAN(cStr,'\u00da','Ú')
		cStr = STRTRAN(cStr,'\u00f1','ñ')
		cStr = STRTRAN(cStr,'\u00d1','Ñ')
		cStr = STRTRAN(cStr,'\u00e3','ã')
		cStr = STRTRAN(cStr,'\u00e2','â')
		cStr = STRTRAN(cStr,'\u00e0','à')
		cStr = STRTRAN(cStr,'\u00e4','ä')
		cStr = STRTRAN(cStr,'\u00c0','À')
		cStr = STRTRAN(cStr,'\u00c2','Â')
		cStr = STRTRAN(cStr,'\u00c3','Ã')
		cStr = STRTRAN(cStr,'\u00c4','è')
		cStr = STRTRAN(cStr,'\u00e8','è')
		cStr = STRTRAN(cStr,'\u00ea','ê')
		cStr = STRTRAN(cStr,'\u00c8','È')
		cStr = STRTRAN(cStr,'\u00ca','Ê')
		cStr = STRTRAN(cStr,'\u00cb','Ë')
		cStr = STRTRAN(cStr,'\u00ec','ì')
		cStr = STRTRAN(cStr,'\u00ee','î')
		cStr = STRTRAN(cStr,'\u00f4','ô')
		cStr = STRTRAN(cStr,'\u00f5','õ')
		cStr = STRTRAN(cStr,'\u00e7','ç')
		cStr = STRTRAN(cStr,'\u00c7','Ç')
		cStr = STRTRAN(cStr,'\u0026','&')
		cStr = STRTRAN(cStr,'\u0027',"'")
		cStr = STRTRAN(cStr,'\u00ef','ï')
		cStr = STRTRAN(cStr,'\u00cc','Ì')
		cStr = STRTRAN(cStr,'\u00ce','Î')
		cStr = STRTRAN(cStr,'\u00cf','Ï')
		cStr = STRTRAN(cStr,'\u00f2','ò')
		cStr = STRTRAN(cStr,'\u00f6','ö')
		cStr = STRTRAN(cStr,'\u00d2','Ò')
		cStr = STRTRAN(cStr,'\u00d4','Ô')
		cStr = STRTRAN(cStr,'\u00d5','Õ')
		cStr = STRTRAN(cStr,'\u00d6','Ö')
		cStr = STRTRAN(cStr,'\u00f9','ù')
		cStr = STRTRAN(cStr,'\u00fb','û')
		cStr = STRTRAN(cStr,'\u00fc','ü')
		cStr = STRTRAN(cStr,'\u00d9','Ù')
		cStr = STRTRAN(cStr,'\u00db','Û')
		RETURN cStr



ENDDEFINE





*
* class used to return an array
*
DEFINE CLASS myArray AS CUSTOM
	nSize = 0
	DIMENSION ARRAY[1]

	FUNCTION ADD(xExpr)
		THIS.nSize = THIS.nSize + 1
		DIMENSION THIS.ARRAY[this.nSize]
		THIS.ARRAY[this.nSize] = xExpr
		RETURN

	FUNCTION GET(N)
		RETURN THIS.ARRAY[n]

ENDDEFINE



*
* class used to simulate an object
* all properties are prefixed with 'prop' to permit property names like: error, init
* that already exists like vfp methods
*
DEFINE CLASS myObj AS CUSTOM
	HIDDEN ;
		CLASSLIBRARY,COMMENT, ;
		BASECLASS,CONTROLCOUNT, ;
		CONTROLS,OBJECTS,OBJECT,;
		HEIGHT,HELPCONTEXTID,LEFT,NAME, ;
		PARENT,PARENTCLASS,PICTURE, ;
		TAG,TOP,WHATSTHISHELPID,WIDTH

	FUNCTION SET(cPropName, xValue)
		cPropName = '_'+cPropName
		IF TYPE('this.'+cPropName)=='U'
			THIS.ADDPROPERTY(cPropName,xValue)
		ELSE
			LOCAL cmd
			cmd = 'this.'+cPropName+'=xValue'
			&cmd
		ENDIF
		RETURN

	PROCEDURE GET (cPropName)
		cPropName = '_'+cPropName
		IF TYPE('this.'+cPropName)=='U'
			RETURN ''
		ELSE
			LOCAL cmd
			cmd = 'return this.'+cPropName
			&cmd
		ENDIF
		RETURN ''
ENDDEFINE





FUNCTION testJsonClass
	CLEAR
	SET DECIMAL TO 10
	oJson = NEWOBJECT('json')


	? 'Test Basic Types'
	? '----------------'
	? oJson.decode('null')
	? oJson.decode('true')
	? oJson.decode('false')
	?
	? oJson.decode('791123')
	? oJson.decode('791123.45')
	? oJson.decode('791123.45.')
	? oJson.decode('"nacho gtz"')
	IF NOT EMPTY(oJson.cError)
		? oJson.cError
		RETURN
	ENDIF
	? oJson.decode('"nacho gtz\nEs \"bueno\"\nMuy Bueno\ba"')
	IF NOT EMPTY(oJson.cError)
		? oJson.cError
		RETURN
	ENDIF

	? 'Test Array'
	? '----------'
	arr = oJson.decode('[3.1416,"Ignacio",false,null]')
	? arr.GET(1), arr.GET(2), arr.GET(3), arr.GET(4)
	arr = oJson.decode('[ ["Hugo","Paco","Luis"] , [ 8,9,11] ] ')
	nombres = arr.GET(1)
	edades  = arr.GET(2)
	? nombres.GET(1), edades.GET(1)
	? nombres.GET(2), edades.GET(2)
	? nombres.GET(3), edades.GET(3)
	?
	? 'Test Object'
	? '-----------'
	obj = oJson.decode('{"nombre":"Ignacio", "edad":33.17, "isGood":true}')
	? obj.GET('nombre'), obj.GET('edad'), obj.GET('isGood')
	? obj._Nombre, obj._Edad, obj._IsGood
	obj = oJson.decode('{"jsonrpc":"1.0", "id":1, "method":"sumArray", "params":[3.1415,2.14,10],"version":1.0}')
	? obj.GET('jsonrpc'), obj._jsonrpc
	? obj.GET('id'), obj._id
	? obj.GET('method'), obj._method
	? obj._Params.ARRAY[1], obj._Params.GET(1)

	?
	? 'Test nested object'
	? '------------------'
	cJson = '{"jsonrpc":"1.0", "id":1, "method":"upload", "params": {"data":{ "usrkey":"288af77b", "sendto":"ignacio@sait.com.mx", "name":"Ignacio is \"Nacho\"","expires":"20120731" }}}'
	obj = oJson.decode(cJson)
	IF NOT EMPTY(oJson.cError)
		? oJson.cError
		RETURN
	ENDIF
	? cJson
	? 'method -->',obj._method
	? 'usrkey -->',obj._Params._data._usrkey
	? 'sendto -->',obj._Params._data._sendto
	? 'name  --->',obj._Params._data._name
	? 'expires ->',obj._Params._data._expires

	?
	? 'Test empty object'
	? '-----------------'
	cJson = '{"result":null,"error":{"code":-3200.012,"message":"invalid usrkey","data":{}},"id":"1"}'
	obj = oJson.decode(cJson)
	IF NOT EMPTY(oJson.cError)
		? oJson.cError
		RETURN
	ENDIF
	? cJson
	? 'result -->',obj._result, obj.GET('result')
	oError = obj.GET('error')
	? 'ErrCode ->',obj._error._code, oError.GET('code')
	? 'ErrMsg -->',obj._error._message, oError.GET('message')
	? 'id  ----->',obj._id, obj.GET('id')
	?  TYPE("oError._code")

	?
	? 'Probar decode-enconde-decode-encode'
	? '------------------------------------'
	cJson = ' {"server":"", "user":"", "password":"" ,'+;
		' "port":0, "auth":false, "ssl":false, "timeout":20, "error":404}'
	? cJson
	oSmtp = json_decode(cJson)
	cJson =  json_encode(oSmtp)
	? cJson
	oSmtp = json_decode(cJson)
	cJson =  json_encode(oSmtp)
	? cJson

	* Probar falla
	?
	? 'Probar una falla en el json'
	? '---------------------------'
	cJson = ' {"server":"", "user":"", "password":"" ,'
	oSmtp = json_decode(cJson)
	IF NOT EMPTY(json_getErrorMsg())
		? json_getErrorMsg()
	ENDIF

	?
	? 'Pruebas Finalizadas'
	RETURN

