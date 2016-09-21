****
* Funções de Compactação de arquivos
****

****
* Cria arquivo zip vazio
****
FUNCTION GSZip_CriaVazio
	LPARAMETERS lcArq
	IF VARTYPE(m.lcArq)#"C"
		RETURN .F.
	ENDIF
	m.lcArq = FORCEEXT(m.lcArq,".ZIP")
	LOCAL lnH
	m.lnH = FCREATE(m.lcArq)
	IF m.lnH<=0
		RETURN .F.
	ENDIF
	LOCAL lcZip
	m.lcZip = "PK"+CHR(5)+CHR(6)+REPLICATE(CHR(0),18)
	IF FWRITE(m.lnH,m.lcZip,LEN(m.lcZip))#LEN(m.lcZip)
		FCLOSE(m.lnH)
		ERASE (m.lcArq)
		RETURN .F.
	ENDIF
	FCLOSE(m.lnH)
	RETURN .T.
ENDFUNC
****
* Adiciona arquivo a um ZIP
****
FUNCTION GSZip_Adiciona
	LPARAMETERS lcArqZip, lcArq, llMove, loApp, loAZip
	IF (VARTYPE(m.lcArqZip)#"C")
		RETURN .F.
	ENDIF
	IF (VARTYPE(m.lcArq)="C") AND (TYPE("m.lcArq[1]")="C")
		FOR EACH m.lcA IN m.lcArq
			GSZip_Adiciona(m.lcArqZip,m.lcA,m.llMove,m.loApp,m.loAZip)
		ENDFOR
		RETURN .T.
	ENDIF
	IF ("?"$m.lcArq) OR ("*"$m.lcArq)
		LOCAL ARRAY laArqs(1,1)
		LOCAL lcPath,llErro
		IF ADIR(laArqs,m.lcArq)=0
			RETURN .F.
		ENDIF
		m.lcPath = ADDBS(JUSTPATH(m.lcArq))

		FOR m.lnI = 1 TO ALEN(m.laArqs,1)
			m.llErro = m.llErro OR !GSZip_Adiciona(m.lcArqZip,m.lcPath+m.laArqs(m.lnI,1),m.llMove,m.loApp, m.loAZip)
		NEXT
		RETURN .T.
	ELSE
		IF !(FILE(m.lcArq) OR DIRECTORY(m.lcArq))
			RETURN .F.
		ENDIF
	ENDIF
	m.lcArqZip = FORCEEXT(m.lcArqZip,".ZIP")
	IF (!FILE(m.lcArqZip)) AND !GSZip_CriaVazio(m.lcArqZip)
		RETURN .F.
	ENDIF

	LOCAL lnOpcoes

	IF VARTYPE(m.loApp)="L"
		m.loApp = CREATEOBJECT("SHELL.APPLICATION")
		m.loAZip = m.loApp.NameSpace(""+m.lcArqZip)
	ENDIF

	m.lnOpcoes = 8 + 16 + 2048

*!*		(4) 	Do not display a progress dialog box.
*!*		(8)		Give the file being operated on a new name in a move, copy, or rename operation if a file with the target name already exists.
*!*		(16)	Respond with "Yes to All" for any dialog box that is displayed.
*!*		(64)	Preserve undo information, if possible.
*!*		(128)	Perform the operation on files only if a wildcard file name (*.*) is specified.
*!*		(256)	Display a progress dialog box but do not show the file names.
*!*		(512)	Do not confirm the creation of a new directory if the operation requires one to be created.
*!*		(1024)	Do not display a user interface if an error occurs.
*!*		(2048)	Version 4.71. Do not copy the security attributes of the file.
*!*		(4096)	Only operate in the local directory. Do not operate recursively into subdirectories.
*!*		(8192)	Version 5.0. Do not copy connected files as a group. Only copy the specified files.
	IF VARTYPE(m.loAZip)#"O"
		WAIT WINDOW "Abrindo ZIP..." TIMEOUT 1
		LOCAL loExc AS EXCEPTION
		TRY
			IF m.llMove
				m.loApp.NameSpace(m.lcArqZip).MoveHere(m.lcArq,m.lnOpcoes)
			ELSE
				m.loApp.NameSpace(m.lcArqZip).CopyHere(m.lcArq,m.lnOpcoes)
			ENDIF
		CATCH TO loExc
			Erros("Arquivo "+m.lcArqZip+ " não foi atualizado"+CHR(13)+m.loExc.MESSAGE)
		ENDTRY
		RETURN .T.
	ENDIF

	IF m.llMove
		m.loAZip.MoveHere(m.lcArq,m.lnOpcoes)
	ELSE
		m.loAZip.CopyHere(m.lcArq,m.lnOpcoes)
	ENDIF
	RETURN .T.
ENDFUNC
