****
* Função Análoga ao MESSAGEBOX, com fonte maior e opção de alerta com cores diferenciadas
*
* lcMensagem	Mensagem texto com uso de \n para separar as linhas
* lnTipo		  0 OK
* 				  1 OK + CANCELAR
*				  2 Abortar + Repetir + Cancelar
*				  3 Sim + Não + Cancelar
*				  4 Sim + Não
*				  5 Repetir + Cancelar
*				  8 Alerta Verde
*				 16 Alerta Amarelo
*				 32 Alerta Laranja
*				 64 Alerta Vermelho
*				128 Alerta Azul
*				256 Botão 2 Default
*				512 Botão 3 Default
* lcCaption		Título da mensagem
* lnTimeOut		Timeout em segundos
*
* Retorno		0	Timeout
*				1 	Ok
*				2	Cancelar
*				3	Abortar
*				4	Repetir
*				5 	Ignorar
*				6	Sim
*				7	Não
****
FUNCTION AlertBox
	LPARAMETERS lcMensagem, lnTipo, lcCaption, lnTimeOut
	LOCAL loAB AS AlertBox, lnRes
	PUBLIC PAB_Result
	m.loAB = CREATEOBJECT("AlertBox",m.lcMensagem,m.lnTipo,m.lcCaption,m.lnTimeOut)
	m.loAB.SHOW()
	m.lnRes = PAB_Result
	RELEASE PAB_Result
	RETURN m.lnRes
ENDFUNC



**************************************************
*-- Class:        AlertBox
*-- ParentClass:  form
*-- BaseClass:    form
*-- Time Stamp:   05/27/15 10:51:10 AM
*
DEFINE CLASS AlertBox AS FORM
	DOCREATE 		= .T.
	AUTOCENTER 		= .T.
	BORDERSTYLE 	= 1
	CAPTION 		= "Form1"
	CONTROLBOX 		= .F.
	ButtonDefault 	= 1
	Margem 			= 12
	CorFundo 		= .F.
	CorFrente 		= .F.
	RefCaption 		= .F.
	TIMEOUT 		= 0
	NAME 			= "AlertBox"
	SHOWWINDOW 		= 1
	WINDOWTYPE 		= 1
	DESKTOP 		= .T.
	ALWAYSONTOP 	= .T.

	ADD OBJECT lblMsg AS LABEL WITH ;
		AUTOSIZE = .T., ;
		FONTBOLD = .T., ;
		FONTSIZE = 13, ;
		WORDWRAP = .T., ;
		BACKSTYLE = 0, ;
		CAPTION = "lblMsg", ;
		HEIGHT = 22, ;
		LEFT = 24, ;
		TOP = 24, ;
		WIDTH = 51, ;
		NAME = "lblMsg"


	ADD OBJECT cmd1 AS COMMANDBUTTON WITH ;
		HEIGHT = 27, ;
		WIDTH = 84, ;
		NAME = "cmd1"


	ADD OBJECT cmd2 AS COMMANDBUTTON WITH ;
		HEIGHT = 27, ;
		WIDTH = 84, ;
		NAME = "cmd2"


	ADD OBJECT cmd3 AS COMMANDBUTTON WITH ;
		HEIGHT = 27, ;
		WIDTH = 84, ;
		NAME = "cmd3"


	ADD OBJECT tmrtimeout AS TIMER WITH ;
		ENABLED = .F., ;
		INTERVAL = 1000, ;
		NAME = "tmrTimeOut"


	PROCEDURE clickcmd
		LPARAMETERS lcCmd
		THISFORM.TAG = STR(AT(m.lcCmd,"OCARISN"),1,0)
		THISFORM.RELEASE()
	ENDPROC


	PROCEDURE ACTIVATE
		IF !EMPTY(THIS.TAG)
			RETURN
		ENDIF
		LOCAL lnL, lnMW, laL(1), lnW, lnMH
		LOCAL lcFont, lnSize, lcStyle
		m.lcFont = THIS.lblMsg.FONTNAME
		m.lnSize = THIS.lblMsg.FONTSIZE
		m.lcStyle = IIF(THIS.lblMsg.FONTBOLD,"B","")+;
			IIF(THIS.lblMsg.FONTITALIC,"I","")

		m.lnL = ALINES(laL,THISFORM.lblMsg.CAPTION,2)
		m.lnMW = 0
		FOR m.lnI = 1 TO m.lnL
			m.lnW = TXTWIDTH(m.laL(m.lnI), m.lcFont, m.lnSize, m.lcStyle) * ;
				FONTMETRIC(6, m.lcFont, m.lnSize, m.lcStyle)
			m.lnMW=MAX(m.lnMW,m.lnW)
		NEXT
		m.lnMH = MAX(m.lnI,1+GETWORDCOUNT(THISFORM.lblMsg.CAPTION,CHR(13)))*FONTMETRIC(1,m.lcFont,m.lnSize,m.lcStyle)
		THIS.lblMsg.WIDTH = m.lnMW

		THIS.lblMsg.REFRESH()
		LOCAL lnB
		m.lnB = ICASE(THIS.cmd3.VISIBLE,3,THIS.cmd2.VISIBLE,2,1)
		THIS.WIDTH = MAX(lnMW + 2*THIS.Margem,m.lnB*THIS.cmd1.WIDTH+(m.lnB+1)*THIS.Margem)
		THIS.HEIGHT = m.lnMH + 3*THIS.Margem + THIS.cmd1.HEIGHT
		THIS.LEFT = (SYSMETRIC(21)-THIS.WIDTH)/2
		THIS.TOP = (SYSMETRIC(22)-THIS.HEIGHT)/2
		THIS.TAG = "0"
		IF THIS.cmd1.DEFAULT
			THIS.cmd1.SETFOCUS()
		ELSE
			IF THIS.cmd2.DEFAULT
				THIS.cmd2.SETFOCUS()
			ELSE
				IF THIS.cmd3.DEFAULT
					THIS.cmd3.SETFOCUS()
				ENDIF
			ENDIF
		ENDIF

		IF THISFORM.TIMEOUT>0
			THISFORM.tmrtimeout.ENABLED = .T.
		ENDIF
	ENDPROC


	PROCEDURE UNLOAD
		m.PAB_Result = VAL(THIS.TAG)
		RETURN INT(VAL(THIS.TAG))
	ENDPROC


	PROCEDURE INIT
		LPARAMETERS lcMessage, lnType, lcCaption, lnTimeOut
		IF VARTYPE(m.lcMessage)#"C"
			m.lcMessage = TRANSFORM(m.lcMessage)
		ENDIF
		IF VARTYPE(m.lnType)#"N"
			m.lnType = 0
		ENDIF
		DO CASE
			CASE BITAND(m.lnType,5) = 5	&& Repedir & Cancelar
				THIS.cmd1.CAPTION = "Repetir"
				THIS.cmd1.TAG = "R"
				THIS.cmd2.CAPTION = "Cancelar"
				THIS.cmd2.TAG = "C"
				THIS.cmd3.VISIBLE = .F.
			CASE BITAND(m.lnType,4) = 4	&& Sim e Não
				THIS.cmd1.CAPTION = "Sim"
				THIS.cmd1.TAG = "S"
				THIS.cmd2.CAPTION = "Não"
				THIS.cmd2.TAG = "N"
				THIS.cmd3.VISIBLE = .F.
			CASE BITAND(m.lnType,3) = 3	&& Sim, Não e Cancelar
				THIS.cmd1.CAPTION = "Sim"
				THIS.cmd1.TAG = "S"
				THIS.cmd2.CAPTION = "Não"
				THIS.cmd2.TAG = "N"
				THIS.cmd3.CAPTION = "Cancelar"
				THIS.cmd3.TAG = "C"
			CASE BITAND(m.lnType,2) = 2	&& Abortar, Repetir, Ignorar
				THIS.cmd1.CAPTION = "Abortar"
				THIS.cmd1.TAG = "A"
				THIS.cmd2.CAPTION = "Repetir"
				THIS.cmd2.TAG = "R"
				THIS.cmd3.CAPTION = "Ignorar"
				THIS.cmd3.TAG = "I"
			CASE BITAND(m.lnType,1) = 1	&& Ok e Cancelar
				THIS.cmd1.CAPTION = "OK"
				THIS.cmd1.TAG = "O"
				THIS.cmd2.CAPTION = "Cancelar"
				THIS.cmd2.TAG = "C"
				THIS.cmd3.VISIBLE = .F.
			OTHERWISE 	&& Ok
				THIS.cmd1.CAPTION = "OK"
				THIS.cmd1.TAG = "O"
				THIS.cmd2.VISIBLE = .F.
				THIS.cmd3.VISIBLE = .F.
		ENDCASE


		DO CASE
			CASE BITAND(m.lnType,8) = 8		&& Alerta Verde
				THIS.CorFundo = RGB(0,128,64)
				THIS.CorFrente = RGB(255,255,255)
			CASE BITAND(m.lnType,16) = 16	&& Alerta Amarelo
				THIS.CorFundo = RGB(255,255,0)
				THIS.CorFrente = RGB(0,0,0)
			CASE BITAND(m.lnType,32) = 32	&& Alerta Laranja
				THIS.CorFundo = RGB(255,128,0)
				THIS.CorFrente = RGB(0,0,0)
			CASE BITAND(m.lnType,64) = 64	&& Alerta Vermelho
				THIS.CorFundo = RGB(255,0,0)
				THIS.CorFrente = RGB(255,255,255)
			CASE BITAND(m.lnType,128) = 128	&& Alerta Azul
				THIS.CorFundo = RGB(0,64,128)
				THIS.CorFrente = RGB(255,255,255)
			OTHERWISE
				THIS.CorFundo = THIS.BACKCOLOR
				THIS.CorFrente = THIS.FORECOLOR
		ENDCASE

		THISFORM.BACKCOLOR = THIS.CorFundo
		THISFORM.FORECOLOR = THIS.CorFrente
		THISFORM.lblMsg.FORECOLOR = THIS.CorFrente

		IF (BITAND(m.lnType,256) = 256) AND THIS.cmd2.VISIBLE	&& Segundo botão default
			THIS.cmd2.DEFAULT = .T.
		ENDIF
		IF (BITAND(m.lnType,512) = 512) AND THIS.cmd3.VISIBLE	&& Terceiro botão default
			THIS.cmd3.DEFAULT = .T.
		ENDIF

		LOCAL lnR, lnM, lnB
		m.lnR = THIS.WIDTH
		m.lnB = 1

		IF THIS.cmd3.VISIBLE
			THIS.cmd3.LEFT = m.lnR - THIS.Margem - THIS.cmd3.WIDTH
			THIS.cmd3.TOP = THIS.HEIGHT - THIS.Margem - THIS.cmd3.HEIGHT
			m.lnR = THIS.cmd3.LEFT
			m.lnB = 3
		ENDIF

		IF THIS.cmd2.VISIBLE
			THIS.cmd2.LEFT = m.lnR - THIS.Margem - THIS.cmd2.WIDTH
			THIS.cmd2.TOP = THIS.HEIGHT - THIS.Margem - THIS.cmd2.HEIGHT
			m.lnR = THIS.cmd2.LEFT
			m.lnB = MAX(2,m.lnB)
		ENDIF

		THIS.cmd1.LEFT = m.lnR - THIS.Margem - THIS.cmd1.WIDTH
		THIS.cmd1.TOP = THIS.HEIGHT - THIS.Margem - THIS.cmd1.HEIGHT

		THIS.SETALL("ANCHOR",12,"COMMANDBUTTON")

		THIS.lblMsg.AUTOSIZE = .T.
		THIS.lblMsg.CAPTION = STRTRAN(m.lcMessage,'\n',CHR(13))
		THIS.lblMsg.LEFT = THIS.Margem
		THIS.lblMsg.TOP = THIS.Margem

		IF VARTYPE(m.lcCaption)#"C"
			IF TYPE("M.SISTEMA")="C"
				m.lcCaption = M.Sistema
			ELSE
				IF (APPLICATION.FORMS.COUNT > 0)
					m.lcCaption = APPLICATION.FORMS[1].CAPTION
				ELSE
					m.lcCaption = APPLICATION.CAPTION
				ENDIF
			ENDIF
		ENDIF

		THIS.RefCaption = m.lcCaption
		THIS.CAPTION = m.lcCaption

		IF (VARTYPE(m.lnTimeOut)="N") AND (m.lnTimeOut>0)
			THIS.TIMEOUT = m.lnTimeOut
		ENDIF
	ENDPROC


	PROCEDURE cmd1.CLICK
		THISFORM.clickcmd(THIS.TAG)
	ENDPROC


	PROCEDURE cmd2.CLICK
		THISFORM.clickcmd(THIS.TAG)
	ENDPROC


	PROCEDURE cmd3.CLICK
		THISFORM.clickcmd(THIS.TAG)
	ENDPROC


	PROCEDURE tmrtimeout.TIMER
		THISFORM.CAPTION = THISFORM.RefCaption+" ("+TRANSFORM(THISFORM.TIMEOUT)+")"
		THISFORM.TIMEOUT = THISFORM.TIMEOUT - 1
		IF THISFORM.TIMEOUT = -2
			THISFORM.TAG = "0"
			THISFORM.RELEASE()
		ENDIF
	ENDPROC


ENDDEFINE
*
*-- EndDefine: alertbox
**************************************************
