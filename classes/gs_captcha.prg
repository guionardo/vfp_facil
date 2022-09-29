****
* Função de confirmação com CAPTCHA numérico
*
* lcMsg		Mensagem a ser apresentada (\n = quebra de linha)
* lcCaption	Caption do form
*
* Retorno 	Boolean
*
* ? ConfirmaCaptcha("MENSAGEM de captcha\nSegunda linha\nLinha número 3 comprida pra cacete ou nem tanto\nLinha 4","CAPTION")

FUNCTION ConfirmaCaptcha
	LPARAMETERS lcMsg, lcCaption
	LOCAL loF as CaptchaForm
	PUBLIC _PubCaptchaOk
	m.loF = CREATEOBJECT("CaptchaForm",m.lcMsg,m.lcCaption)
	m.loF.Show(1)
	LOCAL llPC
	m.llPC = m._PubCaptchaOk
	RELEASE m._PubCaptchaOk
	RETURN m.llPC
ENDFUNC


**************************************************
*-- Class:        CaptchaForm
*-- Time Stamp:   11/05/15 12:00:08 PM
*-- Programador:  Guionardo Furlan
*
DEFINE CLASS CaptchaForm AS form


	DoCreate = .T.
	AutoCenter = .T.
	Caption = "CaptchaForm"
	mensagem = ""
	captcha = .F.
	Name = "CaptchaForm"
	ShowWindow = 1
	WindowType = 1
	BorderStyle = 1
	ControlBox = .f.


	ADD OBJECT lblmsg AS label WITH ;
		BackStyle = 0, ;
		Height = 24, ;
		Left = 24, ;
		Top = 24, ;
		Width = 336, ;
		Name = "lblMsg"


	ADD OBJECT txcaptcha AS textbox WITH ;
		FontBold = .T., ;
		FontSize = 12, ;
		Value = "", ;
		Format = "RK", ;
		Height = 26, ;
		InputMask = "9999", ;
		Left = 24, ;
		SpecialEffect = 1, ;
		Top = 96, ;
		Width = 54, ;
		IntegralHeight = .T., ;
		Name = "txCaptcha"


	ADD OBJECT cmdok AS commandbutton WITH ;
		Top = 144, ;
		Left = 60, ;
		Height = 27, ;
		Width = 84, ;
		Caption = "OK", ;
		Name = "cmdOk"


	ADD OBJECT cmdcancelar AS commandbutton WITH ;
		Top = 144, ;
		Left = 180, ;
		Height = 27, ;
		Width = 84, ;
		Caption = "Cancelar", ;
		Name = "cmdCancelar"


	ADD OBJECT lblcaptcha AS label WITH ;
		FontBold = .T., ;
		FontSize = 12, ;
		Alignment = 2, ;
		BackStyle = 0, ;
		Caption = "Label1", ;
		Height = 24, ;
		Left = 24, ;
		Top = 60, ;
		Width = 336, ;
		Name = "lblCaptcha", ;
		ForeColor = RGB(128,0,0)
		


	PROCEDURE mensagem_access
		RETURN this.lblMsg.Caption 
	ENDPROC


	PROCEDURE mensagem_assign
		LPARAMETERS vNewVal
		IF VARTYPE(m.vNewVal)#"C"
			m.vNewVal = ""
		ENDIF
		THIS.lblMsg.CAPTION = LEFT(CHRTRAN(STRTRAN(TRANSFORM(m.vNewVal),"\n",CHR(10)),CHR(13),CHR(10)),255)
		*
		* Conta linhas
		*
		LOCAL lnL, lnMW, lnI, lnMW
		m.lnL = GETWORDCOUNT(THIS.lblMsg.CAPTION,CHR(10))
		STORE 0 TO m.lnMW, m.lnMH
		FOR m.lnI = 1 TO m.lnL
			m.lnMW = MAX(m.lnMW,THIS.TEXTWIDTH(GETWORDNUM(THIS.lblMsg.CAPTION,m.lnI,CHR(10))))
			m.lnMH = m.lnMH + THIS.TEXTHEIGHT(GETWORDNUM(THIS.lblMsg.CAPTION,m.lnI,CHR(10)))
		NEXT
		THIS.lblMsg.WIDTH = m.lnMW
		THIS.lblMsg.HEIGHT = m.lnMH
		THIS.WIDTH = MAX(THIS.WIDTH,THIS.lblMsg.WIDTH + 2*THIS.lblMsg.LEFT)
		THIS.lblCaptcha.WIDTH = THIS.WIDTH - 2*THIS.lblCaptcha.LEFT
		THIS.txCaptcha.LEFT = (THIS.WIDTH - THIS.txCaptcha.WIDTH)/2

		THIS.lblCaptcha.TOP = THIS.lblMsg.TOP + THIS.lblMsg.HEIGHT + 32
		THIS.txCaptcha.TOP = THIS.lblCaptcha.TOP + THIS.lblCaptcha.HEIGHT + 16
		THIS.cmdOk.TOP = THIS.txCaptcha.TOP + THIS.txCaptcha.HEIGHT + 32
		THIS.cmdCancelar.TOP = THIS.cmdOk.TOP

		THIS.cmdCancelar.LEFT = THIS.WIDTH - THIS.lblCaptcha.LEFT - THIS.cmdCancelar.WIDTH
		THIS.cmdOk.LEFT = THIS.cmdCancelar.LEFT - THIS.cmdOk.WIDTH - 8

		THIS.HEIGHT = THIS.cmdOk.TOP + THIS.cmdOk.HEIGHT + THIS.lblMsg.TOP
	ENDPROC


	PROCEDURE Init
		LPARAMETERS lcMsg, lcCaption
		THIS.Mensagem = m.lcMsg
		IF VARTYPE(m.lcCaption)#"C"
			m.lcCaption = "CAPTCHA"
		ENDIF
		THIS.CAPTION = m.lcCaption

		LOCAL lcN, laN(4), lnI, lcCap, lcRep
		m.lcN = "ZERO UM DOIS TRÊS QUATRO CINCO SEIS SETE OITO NOVE DEZ"
		RAND(SECONDS())
		STORE "" TO THIS.Captcha,m.lcCap
		FOR m.lnI = 1 TO 4
			m.laN(m.lnI) = FLOOR(RAND()*9)+1
			m.lcCap = m.lcCap + GETWORDNUM(m.lcN,m.laN(m.lnI))+" "
			THIS.Captcha = THIS.Captcha + TRANSFORM(m.laN(m.lnI)-1)
		NEXT

		THIS.lblCaptcha.CAPTION = m.lcCap
	ENDPROC


	PROCEDURE Unload
		m._PubCaptchaOk = !EMPTY(THIS.TAG)
	ENDPROC


	PROCEDURE cmdok.Click
		IF THISFORM.txCaptcha.VALUE == THISFORM.Captcha
			THISFORM.TAG = "1"
			THISFORM.RELEASE()
		ELSE
			MESSAGEBOX("Captcha informado está incorreto!",48,"ERRO")
			THISFORM.txCaptcha.SETFOCUS()
		ENDIF
	ENDPROC


	PROCEDURE cmdcancelar.Click
		THISFORM.RELEASE()
	ENDPROC


ENDDEFINE
*
*-- EndDefine: captcha
**************************************************
