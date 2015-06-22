****
*
* WaitCenter - Wait window centralizado
*
* lcTxt		  Texto para exibição
* llWait    Aguarda fechamento do wait para continuar a execução
* lnTimeOut Timeout em segundos para o fechamento
*
****
FUNCTION WaitCenter
	LPARAMETERS lcTxt, llWait, lnTimeOut
*** Trata parâmetros
	IF VARTYPE(m.llWait)="N"
	  m.lnTimeOut = m.llWait
	  m.llWait = .f.
	ENDIF 
	IF VARTYPE(m.lnTimeOut)#"N"
	  m.lnTimeOut = .f.
	ENDIF 
	LOCAL lnTexLen, lnRows, lnAvgChar, lcDispText, lnCnt, lcLine, lnCol, lnRow
*** Calcular o tamanho da mensagem
	LOCAL lnMW, lnML
	m.lnMW = SET("Memowidth")
	m.lnML = _MLINE
	SET MEMOWIDTH TO SCOLS()*0.75 && Limita a largura em 75% da Screen
	_MLINE = 0
	lnTexLen = 0
	lnRows = MEMLINES(lcTxt)

*** Calcular o tamanho do texto para posicionamento
	lnAvgChar = FONTMETRIC(6, 'Arial', 8) / FONTMETRIC(6, _SCREEN.FONTNAME, _SCREEN.FONTSIZE )
	lcDispText = ''
*** Encontrar a linha mais longa na mensagem
	FOR lnCnt = 1 TO lnRows
		lcLine = ' ' + MLINE(lcTxt, lnCnt) + ' '
		lcDispText =lcDispText + lcLine + IIF(m.lnCnt<m.lnRows,CHR(13),"")
		lnTexLen = MAX( TXTWIDTH(lcLine,'MS Sans Serif',8,'B')+4, lnTexLen)  && 4 is border
	NEXT
	m.lcDispText = LEFT(m.lcDispText,255)
*** Posiciona em relação a maior linha da mensagem
	lnCol = INT((SCOLS() - lnTexLen * lnAvgChar )/2)
	lnRow = INT((SROWS() - lnRows)/2)

*** Mostra windows centralizada
	LOCAL lcCmd
	m.lcCmd = IIF(!m.llWait," NOWAIT","")+IIF(VARTYPE(m.lnTimeOut)="N"," TIMEOUT "+TRANSFORM(m.lnTimeOut),"")
	WAIT WINDOW lcDispText AT lnRow, lnCol &lcCmd
	SET MEMOWIDTH TO (m.lnMW)
	_MLINE = m.lnML

ENDFUNC
