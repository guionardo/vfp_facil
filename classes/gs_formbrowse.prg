FUNCTION GS_FormBrowse
*
* lcAlias
* lcCaption		Caption do form
* lcFields		Campos separados por ;
* lcCaption 	Captions separados por ;
* lcInputMasks	InputMasks separados por ;
*
	LPARAMETERS lcAlias, lcCaption, lcFields, lcCaptions, lcInputMasks
	LOCAL loForm AS FORM
	m.loForm = CREATEOBJECT("GS_FormBrowse",m.lcAlias,m.lcCaption,m.lcFields,m.lcCaptions,m.lcInputMasks)
	PUBLIC _GS_FormBrowseRet
	m.loForm.SHOW(1)
	LOCAL llRet
	m.llRet = m._GS_FormBrowseRet
	RELEASE _GS_FormBrowseRet
	RETURN m.llRet
ENDFUNC

**************************************************
*-- Class:        frmbrowse
*-- ParentClass:  form
*-- BaseClass:    form
*-- Time Stamp:   11/09/15 09:37:05 PM
*
DEFINE CLASS GS_FormBrowse AS FORM


	DESKTOP = .T.
	SHOWWINDOW = 1
	DOCREATE = .T.
	CAPTION = "Form1"
	CONTROLBOX = .F.
	WINDOWTYPE = 1
	ALWAYSONTOP = .T.
*-- Specifies the name used to reference an object in VFP.
	ALIAS = ""
*-- XML Metadata for customizable properties
	_MEMBERDATA = [<VFPData><memberdata name="alias" type="property" display="Alias"/></VFPData>]
	NAME = "Form1"


	ADD OBJECT grdbro AS GRID WITH ;
		DELETEMARK = .F., ;
		GRIDLINES = 0, ;
		HEIGHT = 200, ;
		LEFT = 12, ;
		RECORDMARK = .F., ;
		TOP = 12, ;
		WIDTH = 320, ;
		HIGHLIGHTFORECOLOR = RGB(0,0,0), ;
		HIGHLIGHTSTYLE = 1, ;
		ALLOWCELLSELECTION = .F., ;
		NAME = "grdBro"


	PROCEDURE INIT
		#DEFINE Borda 12
*
* lcAlias
* lcCaption		Caption do form
* lcFields		Campos separados por ;
* lcCaption 	Captions separados por ;
* lcInputMasks	InputMasks separados por ;
*
		LPARAMETERS lcAlias, lcCaption, lcFields, lcCaptions, lcInputMasks
		IF VARTYPE(m.lcAlias)#"C"
			ERROR "frmBrowse: Alias não informado"
			RETURN .F.
		ENDIF
		IF !USED(m.lcAlias)
			ERROR "frmBrowse: Alias não encontrado ["+m.lcAlias+"]"
			RETURN .F.
		ENDIF
		THIS.ALIAS = m.lcAlias
		IF VARTYPE(m.lcCaption)#"C"
			m.lcCaption = m.lcAlias
		ENDIF

		LOCAL lnI, lcCap, lnW

* Valida campos
		IF VARTYPE(m.lcFields)#"C" OR EMPTY(m.lcFields)
			m.lcFields=""
			FOR m.lnI = 1 TO FCOUNT(m.lcAlias)
				m.lcFields = m.lcFields + FIELD(m.lnI,m.lcAlias)+";"
			NEXT
		ENDIF

		FOR m.lnI = 1 TO GETWORDCOUNT(m.lcFields,';')
			IF EMPTY(FIELD(GETWORDNUM(m.lcFields,m.lnI,';'),m.lcAlias))
				ERROR "frmBrowse: FIELD ["+GETWORDNUM(m.lcFields,m.lnI,';')+"] não existente em ["+m.lcAlias+"]"
				RETURN .F.
			ENDIF
		NEXT

		IF VARTYPE(m.lcCaptions)#"C"
			m.lcCaptions = ""
		ENDIF
		IF VARTYPE(m.lcInputMasks)#"C"
			m.lcInputMasks = ""
		ENDIF

		THIS.CAPTION = m.lcCaption
		m.lnW = 0

		WITH THIS.grdbro
			.RECORDSOURCE = m.lcAlias
			.COLUMNCOUNT = GETWORDCOUNT(m.lcFields,';')
			FOR m.lnI = 1 TO .COLUMNCOUNT
				.COLUMNS(m.lnI).CONTROLSOURCE = m.lcAlias+"."+GETWORDNUM(m.lcFields,m.lnI,';')
				m.lcCap = GETWORDNUM(m.lcCaptions,m.lnI,';')
				IF EMPTY(m.lcCap)
					m.lcCap = GETWORDNUM(m.lcFields,m.lnI,';')
				ENDIF
				.COLUMNS(m.lnI).Header1.CAPTION = m.lcCap
				m.lcCap = GETWORDNUM(m.lcInputMasks,m.lnI,';')
				IF !EMPTY(m.lcCap)
					.COLUMNS(m.lnI).INPUTMASK = m.lcCap
				ENDIF
			NEXT
			.AUTOFIT()
			FOR m.lnI = 1 TO .COLUMNCOUNT
				m.lnW = m.lnW + .COLUMNS(m.lnI).WIDTH + .GRIDLINEWIDTH
			NEXT
			.LEFT = Borda
			.TOP = Borda
			.WIDTH = THIS.WIDTH - Borda * 2
			.HEIGHT = THIS.HEIGHT - Borda * 2
			.ANCHOR = 15
			THIS.WIDTH = MIN(THIS.WIDTH - .WIDTH + m.lnW + SYSMETRIC(5)*2,;
				IIF(_SCREEN.VISIBLE,SYSMETRIC(1),SYSMETRIC(21)))
			LOCAL lnRC(1)
			SELECT COUNT(*) FROM (m.lcAlias) WHERE !DELETED() INTO ARRAY lnRC

			m.lnH = m.lnRC * (.ROWHEIGHT + .GRIDLINEWIDTH) + ;
				.HEADERHEIGHT +;
				2 * Borda +;
				SYSMETRIC(9)

			THIS.HEIGHT = MIN(SYSMETRIC(22)*0.8,m.lnH)

		ENDWITH

		THIS.LEFT = (SYSMETRIC(21)-THIS.WIDTH) / 2
		THIS.TOP = (SYSMETRIC(22)-THIS.HEIGHT) / 2
	ENDPROC


	PROCEDURE KEYPRESS
		LPARAMETERS nKeyCode, nShiftAltCtrl
		IF INLIST(m.nKeyCode,13,27)
			THISFORM.RELEASE()
		ENDIF
	ENDPROC


	PROCEDURE UNLOAD
		m._GS_FormBrowseRet = !(EOF(THIS.ALIAS) OR BOF(THIS.ALIAS))
	ENDPROC


	PROCEDURE grdbro.KEYPRESS
		LPARAMETERS nKeyCode, nShiftAltCtrl
		THISFORM.KEYPRESS(m.nKeyCode,m.nShiftAltCtrl)
	ENDPROC


	PROCEDURE grdbro.DBLCLICK
		THISFORM.RELEASE()
	ENDPROC


ENDDEFINE
*
*-- EndDefine: frmbrowse
**************************************************
