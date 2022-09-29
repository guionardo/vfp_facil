PROCEDURE CommandLine
  LOCAL loCL 
  m.loCL = CREATEOBJECT("GS_CommandLine")
  m.loCL = Show()
ENDPROC 

**************************************************
*-- Class:        janelacomando (a:\tbyte-2015\classes\tools.vcx)
*-- ParentClass:  form
*-- BaseClass:    form
*-- Time Stamp:   06/18/15 10:42:01 AM
*-- Simula a Janela de Comandos do VFP
*
DEFINE CLASS GS_CommandLine AS form


	Height = 46
	Width = 779
	DoCreate = .T.
	BorderStyle = 3
	Caption = "Janela de Comando"
	ControlBox = .F.
	WindowType = 1
	Name = "janelacomando"


	ADD OBJECT shape1 AS shape WITH ;
		Top = 2, ;
		Left = 1, ;
		Height = 41, ;
		Width = 778, ;
		Name = "Shape1"


	ADD OBJECT cbocommand AS combobox WITH ;
		FontName = "Arial", ;
		FontSize = 9, ;
		ControlSource = "gcCommand", ;
		Height = 25, ;
		Left = 10, ;
		SpecialEffect = 1, ;
		Style = 0, ;
		Top = 11, ;
		Width = 722, ;
		Format = "K", ;
		Name = "cboCommand"


	ADD OBJECT cmdsair AS commandbutton WITH ;
		Top = 12, ;
		Left = 734, ;
		Height = 23, ;
		Width = 39, ;
		Caption = "Sair", ;
		TabIndex = 3, ;
		TabStop = .F., ;
		SpecialEffect = 1, ;
		Name = "cmdSair"


	PROCEDURE Destroy
		Release gcCommand,;
		        glUpArrow,; 
		        glEsc,;     
		        glExecutar  
	ENDPROC


	PROCEDURE Init
		*!*	Public gcCommand,;  && Armazena o comando digitado
		*!*	       glUpArrow,;  && Informa se foi pressionada tecla Seta para Cima
		*!*	       glEsc,;      && Informa se foi pressionada tecla Esc
		*!*	       glExecutar   && Informa se pode executar o comando ou não

		PUBLIC glUpArrow,;  && Informa se foi pressionada tecla Seta para Cima
		glEsc,;      && Informa se foi pressionada tecla Esc
		glExecutar   && Informa se pode executar o comando ou não


		gcCommand          = ''
		glUpArrow          = .F.
		glEsc              = 0
		glExecutar         = .T.

		STORE 4 TO ;
			THIS.cboCommand.TOP,;
			THIS.cmdSair.TOP,;
			THIS.cboCommand.LEFT

		STORE 2 TO ;
			THIS.Shape1.TOP,;
			THIS.Shape1.LEFT

		WITH THIS.Shape1
			.WIDTH = THIS.WIDTH - 4
			.HEIGHT = THIS.cboCommand.HEIGHT + 4
		ENDWITH

		WITH THIS.cmdSair
			.LEFT = THIS.WIDTH - THIS.cmdSair.WIDTH - 8
			.HEIGHT = THIS.cboCommand.HEIGHT
		ENDWITH

		THIS.cboCommand.WIDTH = THIS.cmdSair.LEFT - THIS.cboCommand.LEFT - 2

		WITH THIS
			.TOP = _SCREEN.HEIGHT - .HEIGHT - SYSMETRIC(9) - SYSMETRIC(8)  && 435
			.LEFT = 0
		ENDWITH
		THIS.cboCommand.ANCHOR = 15
		THIS.Shape1.ANCHOR = 15
		THIS.cmdSair.ANCHOR = 13
		** 26,131
		_SCREEN.VISIBLE=.T.
		LOCAL lnW2, lnH2
		m.lnW2 = SCOLS() - 1
		m.lnH2 = (THIS.TOP - 4)/FONTMETRIC(1)+ 1
		* FROM 1,1 TO 30,158;

		DEFINE WINDOW WOUT;
			FROM 1,1 TO (m.lnH2),(m.lnW2);
			TITLE "Janela de Comandos ";
			FONT "arial", 9 STYLE "B";
			FLOAT;
			NOCLOSE;
			NOMINIMIZE;
			GROW

		ACTIVATE WINDOW WOUT
		SET TALK WINDOW WOUT
		SET NOTIFY ON
		SET STATUS BAR ON
	ENDPROC


	PROCEDURE cbocommand.Error
		LPARAMETERS nError, cMethod, nLine

		#DEFINE CRLF CHR(13)+CHR(10)

		WAIT WINDOW "Erro: "+TRANSFORM(nError,'9999')+CRLF+;
			"MESSAGE(): "+MESSAGE()+CRLF+;
			"MESSAGE(1): "+MESSAGE(1)+CRLF+;
			"Método : "+cMethod + CRLF+;
			"Pressione qq. tecla para continuar ..."

		RETURN
	ENDPROC


	PROCEDURE cbocommand.KeyPress
		LPARAMETERS nKeyCode, nShiftAltCtrl

		WITH This
			DO CASE
				CASE LASTKEY()=13   && Enter
					 LOCAL lcCommand

					** Tratamento para quando o usuário pressionar a tecla Seta para Cima,
					** desejando assim chamar algum comando anteriormente digitado
					IF glUpArrow
						LOCAL lcComandoSelecionado
						glUpArrow = .F.
						lcComandoSelecionado = .VALUE
						.Style = 0    && Volta estilo para Combo
						.DisplayValue = lcComandoSelecionado
						glExecutar = .F.  && não executa o comando imediatemente, permitindo
						                  && ao usuário a edição do comando antes de reexecutá-lo
						.SetFocus
					ELSE
						glExecutar = .T.
					ENDIF

		            && Se não estiver vazia a linha de comando, adiciona ao histórico na combo
					IF !EMPTY(.Text)    
						.AddItem(.Text,1)
					ENDIF
					.Value = ALLTRIM(.Text)
					IF !EMPTY(.VALUE) AND glExecutar
						lcCommand=.VALUE
						Activate WINDOW WOUT
						? "C: "+lcCommand
						IF ALLTRIM(UPPER(lcCommand)) = 'QUIT'
						   	*-- Caixa de Mensagem "Informação"
							*	botões "Ok"
							*	default "1º Botão"
							= MESSAGEBOX( ;
								[O comando QUIT não é aceito neste programa.], ;
								64+0+0, ;
								[Atenção]) 
						   RETURN
						ENDIF
						&lcCommand
						.VALUE=""
						gcCommand = ''
						.SETFOCUS()
					ENDIF
				CASE LASTKEY() = 27  && Esc
		            ** Limpa o texto digitado pelo usuário;
		            ** se o usuário pressionar 2 vezes o Esc, a execução do form
		            ** será encerrada (o mesmo que clicar no botão Sair)
					IF EMPTY(.Text)
						glEsc = glEsc + 1
					ENDIF
					.DisplayValue = ''
					IF glEsc = 2
						thisform.cmdSair.Click
					ENDIF
				CASE LASTKEY() = 5   && UpArrow (Seta para Cima)
					glUpArrow = .T.
					.Style = 2       && passa para List
					KEYBOARD '{SPACEBAR}'   && Abre a lista dos comandos digitados durante a atual
					                        && sessão
			ENDCASE
		ENDWITH
	ENDPROC


	PROCEDURE cmdsair.Click
		DEACTIVATE WINDOW WOUT
		THISFORM.Release 
	ENDPROC


ENDDEFINE
