************************************************
* CEP: Consulta e Gerenciamento
*
DEFINE CLASS Logradouro AS CUSTOM
	CEP = SPACE(8)			 			&& CEP 			C 8
	ESTADO = SPACE(2)				&& UF			C 2
	CIDADE = SPACE(40)			 		&& CIDADE		C 40
	Logradouro = SPACE(60)	&& LOGRADOURO 	C 60
	BAIRRO = SPACE(40) 					&& BAIRRO 		C 40
*PUBLIC TPLOGRADOURO AS STRING	&& Tipo			C
	CONTEUDO = ""
	ACENTUADO = .T.
	ORIGEM = ""

	TipoConsulta = 1 && Consulta = 	1 (Correios)
*												2 (WS Time Informática)
*												3 (WS Guionardo.Info)

	FUNCTION INIT
		LPARAMETERS lxCEP
		THIS.TipoConsulta = 2
		IF VARTYPE(m.lxCEP)=="N"
			THIS.CEP = STR(m.lxCEP,8,0)
		ELSE
			IF VARTYPE(m.lxCEP)=="C"
				THIS.CEP = PADL(SoNumero(m.lxCEP),8)
			ELSE
				THIS.CEP = ""
			ENDIF
		ENDIF
		STORE '' TO THIS.ESTADO,THIS.CIDADE,THIS.Logradouro, THIS.BAIRRO
		IF !EMPTY(THIS.CEP)
			THIS.Carregar(this.CEP)
		ENDIF
	ENDFUNC

	FUNCTION CarregarDB
		IF EMPTY(THIS.CEP) OR !USED("CEP") OR !SEEK(VAL(THIS.CEP),"CEP",1)
			RETURN .F.
		ENDIF
		THIS.Logradouro = CEP.DESLOG
		THIS.BAIRRO = CEP.NOMBAI
		THIS.CIDADE = CEP.NOMMUN
		THIS.ESTADO = CEP.SIGEST
		THIS.ORIGEM = "D"
		RETURN .T.
	ENDFUNC

	FUNCTION ProcessaAcentos
		IF THIS.ACENTUADO
			RETURN
		ENDIF
*!*			THIS.Logradouro = SYS(15,M.ACANSI,THIS.Logradouro)
*!*			THIS.BAIRRO = SYS(15,M.ACANSI,THIS.BAIRRO)
*!*			THIS.CIDADE = SYS(15,M.ACANSI,THIS.CIDADE)
*!*			THIS.ESTADO = SYS(15,M.ACANSI,THIS.ESTADO)
	ENDFUNC

	FUNCTION GravaDB
		IF EMPTY(THIS.CEP) OR !USED("CEP")
			RETURN .F.
		ENDIF
		IF !SEEK(VAL(THIS.CEP),"CEP",1)	
				INSERT INTO CEP ;
					(NUMCEP,DESLOG,NOMBAI,NOMMUN,SIGEST,DATUPD) VALUES ;
					(VAL(THIS.CEP),THIS.Logradouro,THIS.BAIRRO,THIS.CIDADE,THIS.ESTADO,DATETIME())
		ELSE
			UPDATE CEP SET ;
				NUMCEP = VAL(THIS.CEP),;
				DESLOG = THIS.Logradouro,;
				NOMBAI = THIS.BAIRRO,;
				NOMMUN = THIS.CIDADE,;
				SIGEST = THIS.ESTADO,;
				DATUPD = DATETIME() ;
				WHERE NUMCEP=VAL(THIS.CEP)
		ENDIF
		RETURN _TALLY>0
	ENDFUNC

	FUNCTION Download
		LPARAMETERS lcURL, lcParams, lcMethod
		LOCAL loHttp AS OBJECT, loExc AS EXCEPTION, llErro
		m.llErro = .F.
		WAIT WINDOW "Consultando CEP para "+THIS.CEP+"..." NOWAIT
		TRY
			loHttp = CREATEOBJECT("MSXML2.ServerXMLHTTP.6.0")
		CATCH TO loExc
			m.llErro = .T.
		ENDTRY
		IF m.llErro
			TRY
				loHttp = CREATEOBJECT("MSXML2.ServerXMLHTTP.4.0")
			CATCH TO loExc
				m.llErro = .T.
				WAIT WINDOW "Consulta de CEP falhou!"+CHR(13)+;
					"Objeto MSXML2 não pôde ser criado!"+CHR(13)+;
					"Erro: "+TRANSFORM(m.loExc.ERRORNO)+" "+m.loExc.MESSAGE TIMEOUT 15
			ENDTRY
		ENDIF
		IF m.llErro
			RETURN .F.
		ENDIF
		TRY
			loHttp.OPEN(m.lcMethod,m.lcURL,.F.)
			loHttp.setRequestHeader("Content-type","application/x-www-form-urlencoded")
			loHttp.SEND(m.lcParams)
		CATCH TO loExc
			LOCAL lcM
			m.lcM = UPPER(m.loExc.MESSAGE)
			DO CASE
				CASE "TIMED OUT"$m.lcM OR "TIMEOUT"$m.lcM
					WAIT WINDOW "Consulta de CEP falhou!"+CHR(13)+;
						"Servidor de correios está demorando muito para responder!" TIMEOUT 15
				OTHERWISE
					WAIT WINDOW "Consulta de CEP falhou!"+CHR(13)+;
						"Erro: "+TRANSFORM(m.loExc.ERRORNO)+" "+m.loExc.MESSAGE TIMEOUT 15
			ENDCASE
		ENDTRY
		THIS.CONTEUDO = ''
		WAIT CLEAR
		IF loHttp.STATUS#200
			RETURN .F.
		ENDIF
		THIS.CONTEUDO = loHttp.responsetext()
		RETURN .T.
	ENDFUNC

	FUNCTION CarregarCorreios
		IF !THIS.Download("http://www.buscacep.correios.com.br/servicos/dnec/consultaLogradouroAction.do",;
				"relaxation=" + THIS.CEP + ;
				"&Metodo=listaLogradouro&TipoConsulta=relaxation&StartRow=1&EndRow=10",;
				"POST")
			RETURN .F.
		ENDIF
		LOCAL lcRet
		m.lcRet = THIS.CONTEUDO
		lcRet = STREXTRACT(lcRet,"<?xml","</table>")
		LOCAL ARRAY aRET(5)
		THIS.Logradouro = STREXTRACT(lcRet,'px">',"</",1) && Endereço
		THIS.BAIRRO		= STREXTRACT(lcRet,'px">',"</",2) && Bairro
		THIS.CIDADE		= STREXTRACT(lcRet,'px">',"</",3) && Cidade
		THIS.ESTADO		= STREXTRACT(lcRet,'px">',"</",4) && UF
		THIS.ORIGEM		= "C"
		THIS.ProcessaAcentos()
		THIS.GravaDB()
	ENDFUNC

	FUNCTION CarregarWSGuiosoft
		IF !THIS.Download("http://ws.guiosoft.info/cep/"+THIS.CEP,"","GET")
			RETURN .F.
		ENDIF
		LOCAL oJSON AS OBJECT , obj AS OBJECT
		m.oJSON = CREATEOBJECT("JSON")
		obj = oJSON.decode(THIS.CONTEUDO)
		IF VARTYPE(m.obj)#"O"
			Erros("Não foi possível obter a informação do webservice de CEP.\n\n"+;
				TRANSFORM(m.oJSON.cError))
			RETURN .F.
		ENDIF

		THIS.Logradouro = m.obj.GET("tipo_logradouro")+" "+m.obj.GET("logradouro")
		THIS.BAIRRO = m.obj.GET("bairro")
		THIS.CIDADE = m.obj.GET("cidade")
		THIS.ESTADO = m.obj.GET("uf")
		THIS.ORIGEM = "G"
		THIS.ProcessaAcentos()
		THIS.GravaDB()
		RETURN .T.
	ENDFUNC

	FUNCTION Carregar
		LPARAMETERS lxCEP
		THIS.TipoConsulta = 2
		IF VARTYPE(m.lxCEP)=="N"
			THIS.CEP = STR(m.lxCEP,8,0)
		ELSE
			IF VARTYPE(m.lxCEP)=="C"
				THIS.CEP = PADL(SoNumero(m.lxCEP),8)
			ELSE
				THIS.CEP = ""
			ENDIF
		ENDIF
		STORE '' TO THIS.ESTADO,THIS.CIDADE,THIS.Logradouro, THIS.BAIRRO, THIS.ORIGEM
		IF EMPTY(THIS.CEP)
			RETURN .F.
		ENDIF
		IF THIS.CarregarDB()
			RETURN .T.
		ENDIF
		IF THIS.CarregarWSGuiosoft()
			RETURN .T.
		ENDIF
		IF THIS.CarregarCorreios()
			RETURN .T.
		ENDIF
		RETURN .F.

	ENDFUNC

ENDDEFINE



