****
*
* ProgressForm - Controle de progresso por itens
*
****

****
* 
* Inicializa Progress Form
*
****
FUNCTION GSProgressFormInit
	LPARAMETERS lnTotProc	&& Número de processos
	IF VARTYPE(m.lnTotProc)#"N"
		m.lnTotProc = 0
	ENDIF
	IF TYPE("M._GSProgressForm")="U"
		PUBLIC _GSProgressForm AS GSProgressForm
	ENDIF
	IF VARTYPE(M._GSProgressForm)#"O"
		IF m.lnTotProc = 0
			RETURN .F.
		ENDIF
		m._GSProgressForm = CREATEOBJECT("GSProgressForm",m.lnTotProc)
	ELSE
		IF m.lnTotProc > 1
			m._GSProgressForm.TotProc = m.lnTotProc
		ENDIF
	ENDIF
	RETURN VARTYPE(m._GSProgressForm)="O"
ENDFUNC
****
*
* Adiciona processo a lista
*
****
FUNCTION GSProgressNovoProcesso
	LPARAMETERS lcProcesso
	IF (!GSProgressFormInit()) OR VARTYPE(m.lcProcesso)#"C"
		RETURN .F.
	ENDIF
	m._GSProgressForm.NovoProcesso(m.lcProcesso)
	RETURN .T.
ENDFUNC
****
*
* Fecha Processos
*
****
FUNCTION GSProgressFormClose
	IF (!GSProgressFormInit())
		RETURN .F.
	ENDIF
	m._GSProgressForm.HIDE()
	m._GSProgressForm.DESTROY()
	RELEASE _GSProgressForm
	RETURN .T.
ENDFUNC
****
*
* Registra progresso de um processo
*
****
FUNCTION GSProgressProcessa
	LPARAMETERS lnProgresso
	IF (!GSProgressFormInit()) OR VARTYPE(m.lnProgresso)#"N" OR !BETWEEN(m.lnProgresso,0,100)
		RETURN .F.
	ENDIF
	m._GSProgressForm.Progresso = m.lnProgresso
	IF !m._GSProgressForm.VISIBLE
		m._GSProgressForm.SHOW()
	ENDIF
	RETURN .T.
ENDFUNC
****
*
* Atualiza Processo
*
****
FUNCTION GSProgressUpdate
	LPARAMETERS lcProcesso
	IF (!GSProgressFormInit())
		RETURN .F.
	ENDIF
	m._GSProgressForm.UpdateProcesso(m.lcProcesso)
	IF !m._GSProgressForm.VISIBLE
		m._GSProgressForm.SHOW()
	ENDIF
	RETURN .T.
ENDFUNC

DEFINE CLASS GSProgressForm AS form
	Height = 320
	Width = 320
	ShowWindow = 1
	DoCreate = .T.
	BorderStyle = 0
	Caption = "Form"
	ControlBox = .F.
	AlwaysOnTop = .T.
	*-- Total de Processos
	totproc = 0
	*-- Progresso do processo atual
	progresso = 0
	*-- XML Metadata for customizable properties
	_memberdata = [<VFPData><memberdata name="novoprocesso" type="method" display="NovoProcesso"/><memberdata name="progresso" type="property" display="Progresso"/><memberdata name="totproc" type="property" display="TotProc"/></VFPData>]
	Name = "progressform"


	ADD OBJECT pbprocesso AS progressbar WITH ;
		Top = 252, ;
		Left = 12, ;
		Width = 300, ;
		Name = "pbProcesso", ;
		shpPrg.Name = "shpPrg", ;
		lblPos.Name = "lblPos"


	ADD OBJECT pbtodos AS progressbar WITH ;
		Top = 288, ;
		Left = 12, ;
		Width = 300, ;
		Height = 24, ;
		Name = "pbTodos", ;
		shpPrg.Name = "shpPrg", ;
		lblPos.Name = "lblPos"


	ADD OBJECT lstproc AS listbox WITH ;
		ColumnCount = 0, ;
		FirstElement = 1, ;
		Height = 228, ;
		Left = 10, ;
		NumberOfElements = 0, ;
		SpecialEffect = 1, ;
		Top = 12, ;
		Width = 300, ;
		AutoHideScrollbar = 1, ;
		Name = "lstProc"


	PROCEDURE progresso_assign
		LPARAMETERS vNewVal
		IF VARTYPE(m.vNewVal)#"N" OR !BETWEEN(m.vNewVal,0,100)
			RETURN
		ENDIF

		THISFORM.pbProcesso.Valor = m.vNewVal
		THISFORM.pbTodos.Valor = INT(100*((THIS.lstProc.LISTCOUNT-1)/THIS.TotProc+m.vNewVal/100/THIS.TotProc))

		THIS.Progresso = m.vNewVal
		IF !THIS.VISIBLE
			THIS.SHOW()
		ENDIF
	ENDPROC


	*-- Inicia novo processo
	PROCEDURE novoprocesso
		LPARAMETERS lcProcesso
		THIS.lstProc.ADDITEM(TRANSFORM(m.lcProcesso))
		THIS.lstProc.LISTINDEX = THIS.lstProc.LISTCOUNT
		THIS.Progresso = 0
	ENDPROC


	PROCEDURE totproc_assign
		LPARAMETERS vNewVal
		IF VARTYPE(m.vNewVal)#"N"
			RETURN
		ENDIF

		THIS.TotProc = m.vNewVal
		THIS.lstProc.CLEAR()
		THIS.Progresso = 0
	ENDPROC


	PROCEDURE Init
		LPARAMETERS lnTotProc
		IF VARTYPE(m.lnTotProc)#"N" OR m.lnTotProc<=0
			RETURN .F.
		ENDIF
		THISFORM.BORDERSTYLE = 0
		THISFORM.TotProc = m.lnTotProc
		THISFORM.Progresso = 0
		IF TYPE("M.SISTEMA")="C"
			THIS.CAPTION = M.SISTEMA
		ELSE
			THIS.CAPTION = JUSTSTEM(APPLICATION.SERVERNAME)
		ENDIF
		THIS.AUTOCENTER = .F.
		THIS.LEFT = SYSMETRIC(21) - THIS.WIDTH - SYSMETRIC(5)
		THIS.TOP = SYSMETRIC(22) -  THIS.HEIGHT - SYSMETRIC(6) * 4
	ENDPROC


	*-- Atualiza nome do último processo
	PROCEDURE updateprocesso
		LPARAMETERS lcNovoProcesso
		IF THIS.lstProc.LISTCOUNT = 0
			RETURN
		ENDIF
		THIS.lstProc.LIST(THIS.lstProc.LISTCOUNT) = TRANSFORM(m.lcNovoProcesso)
	ENDPROC


ENDDEFINE
*
*-- EndDefine: progressform
**************************************************

**************************************************
*-- Class:        progressbar 
*-- ParentClass:  container
*-- BaseClass:    container
*-- Time Stamp:   08/01/14 03:37:12 PM
*
DEFINE CLASS ProgressBar AS container


	Width = 200
	Height = 24
	*-- Valor mínimo
	valormin = 0
	*-- Valor Máximo
	valormax = 100
	*-- Valor corrente
	valor = 0
	Name = "progressbar"


	ADD OBJECT shpprg AS shape WITH ;
		Top = 2, ;
		Left = 2, ;
		Height = 20, ;
		Width = 100, ;
		Anchor = 7, ;
		BackColor = RGB(0,128,192), ;
		Name = "shpPrg"


	ADD OBJECT lblpos AS label WITH ;
		Alignment = 2, ;
		BackStyle = 0, ;
		Caption = "Label1", ;
		Height = 17, ;
		Left = 48, ;
		Top = 4, ;
		Width = 40, ;
		Name = "lblPos"


	PROCEDURE valormin_assign
		LPARAMETERS vNewVal
		IF VARTYPE(m.vNewVal)#"N"
			RETURN
		ENDIF
		IF m.vNewVal > THIS.ValorMax
			m.vNewVal = THIS.ValorMax
		ENDIF
		THIS.ValorMin = INT(m.vNewVal)
	ENDPROC


	PROCEDURE valormax_assign
		LPARAMETERS vNewVal
		IF VARTYPE(m.vNewVal)#"N"
			RETURN
		ENDIF
		IF m.vNewVal < THIS.ValorMin
			m.vNewVal = THIS.ValorMin
		ENDIF
		THIS.ValorMax = INT(m.vNewVal)
	ENDPROC


	PROCEDURE valor_assign
		PARAMETERS vNewVal
		IF VARTYPE(m.vNewVal)#"N"
			RETURN
		ENDIF
		IF m.vNewVal > THIS.ValorMax
			m.vNewVal = THIS.ValorMax
		ENDIF
		IF m.vNewVal < THIS.ValorMin
			m.vNewVal = THIS.ValorMin
		ENDIF
		THIS.Valor = m.vNewVal
		this.Desenha()
	ENDPROC


	PROCEDURE desenha
		WITH THIS.shpPrg
			.LEFT = 2
			.TOP = 2
			.HEIGHT = THIS.HEIGHT - 4
			.WIDTH = INT((THIS.WIDTH-4)*(THIS.Valor-THIS.ValorMin)/(THIS.ValorMax-THIS.ValorMin))
		ENDWITH
		WITH THIS.lblPos
			.LEFT = 2
			.TOP = INT((THIS.HEIGHT-.HEIGHT)/2)
			.WIDTH = THIS.WIDTH - 4
			.CAPTION = TRANSFORM(THIS.Valor)
		ENDWITH
	ENDPROC


	PROCEDURE Refresh
		this.Desenha()
	ENDPROC


ENDDEFINE
*
*-- EndDefine: progressbar
**************************************************

