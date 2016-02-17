*
* Cálculo de feriados
*

FUNCTION EhFeriado
	LPARAMETERS ldData
	IF VARTYPE(m.ldData)#"D"
		m.ldData = DATE()
	ENDIF
	LOCAL lcY, lcFer
	m.lcY = "/"+STR(YEAR(m.ldData),4,0)
	m.lcFer = "01/01"+m.lcY+" "+;
		DTOC(DomingoPascoa(m.ldData)-47)+" "+;
		DTOC(DomingoPascoa(m.ldData))+" "+;
		DTOC(DomingoPascoa(m.ldData)+60)+" "+;
		"21/04"+m.lcY+" "+;
		"01/05"+m.lcY+" "+;
		"07/09"+m.lcY+" "+;
		"12/10"+m.lcY+" "+;
		"02/11"+m.lcY+" "+;
		"15/11"+m.lcY+" "+;
		"25/12"+m.lcY
	RETURN DTOC(m.ldData)$m.lcFer
ENDFUNC

*!*	Para anos entre 1901 e 2099:
*!*	X=24
*!*	Y=5
*!*	a = ANO MOD 19
*!*	b= ANO MOD 4
*!*	c = ANO MOD 7
*!*	d = (19 * a + X) MOD 30
*!*	e = (2 * b + 4 * c + 6 * d + Y) MOD 7
*!*	Se (d + e) > 9 então DIA = (d + e - 9) e MES = Abril
*!*	senão DIA = (d + e + 22) e MES = Março

*!*	Há dois casos particulares que ocorrem duas vezes por século:

*!*	Quando o domingo de Páscoa cair em Abril e o dia for 26, corrige-se para uma semana antes, ou seja, vai para dia 19;
*!*	Quando o domingo de Páscoa cair em Abril e o dia for 25 e o termo "d" for igual a 28, simultaneamente com "a" maior que 10, então o dia é corrigido para 18.
*!*	Neste século estes dois casos particulares só acontecerão em 2049 e 2076.

*!*	Para calcular a Terça-feira de Carnaval, basta subtrair 47 dias do Domingo de Páscoa.
*!*	Para calcular a Quinta-feira de Corpus Christi, soma-se 60 dias ao Domingo de Páscoa.
FUNCTION DomingoPascoa
	LPARAMETERS ldData
	IF VARTYPE(m.ldData)="N" AND BETWEEN(m.ldData,1901,2099)
		m.ldData = DATE(m.ldData,1,1)
	ELSE
		IF VARTYPE(m.ldData)#"D" OR !BETWEEN(YEAR(m.ldData),1901,2099)
			m.ldData = DATE()
		ENDIF
	ENDIF
	LOCAL lnY, lnA, lnB, lnC, lnD, lnE
	m.lnY = YEAR(m.ldData)
	m.lnA = MOD(m.lnY,19)
	m.lnB = MOD(m.lnY,4)
	m.lnC = MOD(m.lnY,7)
	m.lnD = MOD(19 * m.lnA + 24,30)
	m.lnE = MOD(2 * m.lnB + 4 * m.lnC + 6 * m.lnD + 5,7)

	IF (m.lnD+m.lnE)>9
		m.lnD = m.lnD + m.lnE - 9
		m.lnM = 4
	ELSE
		m.lnD = m.lnD + m.lnE + 22
		m.lnM = 3
	ENDIF

	IF (m.lnM = 4) AND (m.lnD = 26)
		m.lnD = 19
	ENDIF
	IF (m.lnM = 4) AND (19*m.lnA + 24 = 28) AND (m.lnA>10)
		m.lnD = 18
	ENDIF
	RETURN DATE(m.lnY,m.lnM,m.lnD)
ENDFUNC

