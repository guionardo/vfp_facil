****
*
* Classe de Busca de Arquivos em Pastas e SubPastas
* Guionardo / 2015
*
* Arquivos serão registrados na propriedade ITEM
* Pastas serão registradas na propriedade Pastas.ITEM
* Eventos OnAddFile e OnAddFolder 
*		Devem ser o nome de uma função, que recebe um parâmetro string e 
*		retorna booleano caso a busca deva ser encerrada
*
****

DEFINE CLASS GS_BuscaArquivos AS COLLECTION
	Arquivo 	= ""
	PastaBase 	= ""
	OnAddFile 	= .F.		&& Nome do evento a ser chamado. Parâmetro é o arquivo adicionado
	OnAddFolder = .F.		&& Nome do evento a ser chamado. Parâmetro é a pasta adicionada
	Pastas 		= .F.		&& Pastas coletadas

	FUNCTION INIT
		LPARAMETERS lcArq, lcPB
		IF VARTYPE(m.lcPB)#"C"
			m.lcPB = CURDIR()
		ENDIF
		IF !DIRECTORY(m.lcPB)
			ACErros("Caminho não encontrado!\n\n"+m.lcPB,"Localização de Arquivos")
			RETURN .F.
		ENDIF
		THIS.PastaBase = ADDBS(m.lcPB)
		IF VARTYPE(m.lcArq)="C"
			THIS.Arquivo = m.lcArq
		ENDIF
		THIS.Pastas = CREATEOBJECT("COLLECTION")
	ENDFUNC

	FUNCTION DESTROY
		THIS.Pastas.DESTROY()
	ENDFUNC

	FUNCTION OnAddEvent
		LPARAMETERS lcTipo, lcArq
		IF (m.lcTipo="A" AND VARTYPE(THIS.OnAddFile)#"C") OR ;
				(m.lcTipo="D" AND VARTYPE(THIS.OnAddFolder)#"C")
			RETURN
		ENDIF
		LOCAL lcTA, loEx AS EXCEPTION, llRet
		TRY
			m.lcTA = IIF(m.lcTipo="A",THIS.OnAddFile,THIS.OnAddFolder)+"(m.lcArq)"
			m.llRet = EVALUATE(m.lcTA)

		CATCH TO loEx
&& Se houver uma exceção, desabilita o evento
			IF m.lcTipo="A"
				THIS.OnAddFile = .F.
			ELSE
				THIS.OnAddFolder = .F.
			ENDIF
			RETURN .T.
		ENDTRY
		RETURN m.llRet
	ENDFUNC

	FUNCTION Busca
		LPARAMETERS lcArq, lcPB
		IF VARTYPE(m.lcArq)#"C"
			m.lcArq = THIS.Arquivo
		ENDIF
		IF VARTYPE(m.lcPB)#"C" OR !DIRECTORY(m.lcPB)
			m.lcPB = THIS.PastaBase
		ENDIF

		IF VARTYPE(m.lcArq)#"C" OR EMPTY(m.lcArq)
			ACErros("Expressão de busca inválida: ["+TRANSFORM(m.lcArq)+"]")
			RETURN .F.
		ENDIF
		DO WHILE THIS.COUNT > 0
			THIS.REMOVE(THIS.COUNT)
		ENDDO

		DO WHILE THIS.Pastas.COUNT > 0
			THIS.Pastas.REMOVE(THIS.Pastas.COUNT)
		ENDDO

		THIS.FindFile(m.lcArq,m.lcPB)

		RETURN THIS.COUNT
	ENDFUNC

	FUNCTION FindFile
		LPARAMETERS lcArq, lcPB

		LOCAL laF(1), lnI, lnF
*
* Busca os arquivos
*
		m.lnF = ADIR(m.laF,ADDBS(m.lcPB)+m.lcArq)
		IF m.lnF>0
			THIS.Pastas.ADD(m.lcPB)
			IF !THIS.OnAddEvent("D",m.lcPB)
				RETURN .F.
			ENDIF
			FOR m.lnI = 1 TO m.lnF
				IF "D"$m.laF(m.lnI,5)
					LOOP
				ENDIF
				THIS.ADD(ADDBS(m.lcPB)+m.laF(m.lnI,1))
				IF !THIS.OnAddEvent("A",ADDBS(m.lcPB)+m.laF(m.lnI,1))
					RETURN .F.
				ENDIF
			NEXT
		ENDIF
*
* Busca as subpastas
*
		m.lnF = ADIR(m.laF,ADDBS(m.lcPB)+"*.*","D")
		FOR m.lnI = 1 TO m.lnF
			IF m.laF(m.lnI,1)="."
				LOOP
			ENDIF
			IF !THIS.FindFile(m.lcArq,ADDBS(m.lcPB)+m.laF(m.lnI,1))
				RETURN .F.
			ENDIF
		NEXT
	ENDFUNC


ENDDEFINE


****
*
* Teste
*
****
PROCEDURE TestaBuscaArquivo
	PUBLIC ba
	ba = CREATEOBJECT("BuscaArquivos","*.log","A:\MINIPA\PEDMINIPA\LOG")
	ba.OnAddFolder = "FAddFolder"
	? ba.COUNT
	? ba.Busca()

	RETURN
ENDPROC

FUNCTION FAddFolder
	LPARAMETERS lcFol
	LOCAL X
	m.X = ""
	WAIT WINDOW m.lcFol TO X
	? m.X
	RETURN m.X#"X"
ENDFUNC
