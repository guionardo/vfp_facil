****
*
* Retorna informações da placa mãe
*
* Fabricante, Modelo e Número de Série
*
****
FUNCTION GetMotherBoardInfo
	LOCAL loW, LOOP, lcManufacturer, lcModel, lcSerialNumber, loExc as Exception 
	m.loW = GETOBJECT("winmgmts:\\")
	STORE "" TO m.lcManufacturer, m.lcModel, m.lcSerialNumber
	TRY
		m.loOp = m.loW.ExecQuery("Select * from Win32_BaseBoard")
		FOR EACH loIt IN m.loOp
			IF !ISNULL(m.loIt.Manufacturer)
				m.lcManufacturer = m.loIt.Manufacturer
			ENDIF
			IF !ISNULL(m.loIt.Model)
				m.lcModel = m.loIt.Model
			ENDIF
			IF !ISNULL(m.loIt.SerialNumber)
				m.lcSerialNumber = m.loIt.SerialNumber
			ENDIF
		NEXT
	CATCH TO loExc
		m.lcManufacturer = "Erro: "+m.loExc.Message
	ENDTRY

	RETURN m.lcManufacturer+" "+m.lcModel+" "+m.lcSerialNumber
ENDFUNC
