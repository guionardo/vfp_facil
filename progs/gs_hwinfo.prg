****
*
* Retorna informações da placa mãe
*
* Fabricante, Modelo e Número de Série
*
****
FUNCTION GetMotherBoardInfo
	LOCAL loW, LOOP, lcManufacturer, lcModel, lcSerialNumber, loExc AS EXCEPTION
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
		m.lcManufacturer = "Erro: "+m.loExc.MESSAGE
	ENDTRY

	RETURN m.lcManufacturer+" "+m.lcModel+" "+m.lcSerialNumber
ENDFUNC

****
*
* Retorna objeto com informações da do volume no disco
* lcRootPathName = Unidade local (C, D, etc) ou unidade de rede (\\SERVIDOR\COMPARTILHAMENTO\PASTA) ou unidade padrão caso não seja informada
*
* Retorna Objeto com propriedades:
*		.OK		= Status da consulta
*		.MSG	= Mensagem
* 		.RootPathName	= Raiz do volume
*		.VolumeName		= Nome do volume
*		.VolumeSerial	= Número de série do volume
*		.CaseSensitive
*		.CasePreserved
*		.UnicodeOnDisk
*		.PersistentACLS	= Permissões persistentes
*		.FileCompression
*		.VolIsCompressed
*
****
FUNCTION GetVolumeInfo
	LPARAMETERS lcRootPathName

	LOCAL lpRootPathName, ;
		lpVolumeNameBuffer, ;
		nVolumeNameSize, ;
		lpVolumeSerialNumber, ;
		lpMaximumComponentLength, ;
		lpFileSystemFlags, ;
		lpFileSystemNameBuffer, ;
		nFileSystemNameSize,;
		loRet

	IF VARTYPE(m.lcRootPathName)#"C" OR EMPTY(m.lcRootPathName)
		m.lcRootPathName = SYS(5)
	ENDIF
	IF LEN(m.lcRootPathName)<2
		m.lcRootPathName = m.lcRootPathName + ":"
	ENDIF

	lpRootPathName           = ADDBS(m.lcRootPathName)      && Drive and directory path
	lpVolumeNameBuffer       = SPACE(256) && lpVolumeName return buffer
	nVolumeNameSize          = 256        && Size of/lpVolumeNameBuffer
	lpVolumeSerialNumber     = 0          && lpVolumeSerialNumber buffer
	lpMaximumComponentLength = 256
	lpFileSystemFlags        = 0
	lpFileSystemNameBuffer   = SPACE(256)
	nFileSystemNameSize      = 256

	m.loRet = CREATEOBJECT("EMPTY")
	ADDPROPERTY(m.loRet,"OK",.F.)
	ADDPROPERTY(m.loRet,"MSG","")

	TRY
		DECLARE INTEGER GetVolumeInformation IN Win32API AS GetVolInfo ;
			STRING  @lpRootPathName, ;
			STRING  @lpVolumeNameBuffer, ;
			INTEGER nVolumeNameSize, ;
			INTEGER @lpVolumeSerialNumber, ;
			INTEGER @lpMaximumComponentLength, ;
			INTEGER @lpFileSystemFlags, ;
			STRING  @lpFileSystemNameBuffer, ;
			INTEGER nFileSystemNameSize

		RetVal=GetVolInfo(@lpRootPathName, @lpVolumeNameBuffer, ;
			nVolumeNameSize, @lpVolumeSerialNumber, ;
			@lpMaximumComponentLength, @lpFileSystemFlags, ;
			@lpFileSystemNameBuffer, nFileSystemNameSize)


		ADDPROPERTY(m.loRet,"RootPathName",		m.lpRootPathName)
		ADDPROPERTY(m.loRet,"VolumeName",		LEFT(ALLTRIM(lpVolumeNameBuffer),LEN(ALLTRIM(lpVolumeNameBuffer))-1))
		ADDPROPERTY(m.loRet,"VolumeSerial",		TRANSFORM(lpVolumeSerialNumber))
		ADDPROPERTY(m.loRet,"CaseSensitive",	BITTEST(m.lpFileSystemFlags,0))
		ADDPROPERTY(m.loRet,"CasePreserved",	BITTEST(m.lpFileSystemFlags,1))
		ADDPROPERTY(m.loRet,"UnicodeOnDisk",	BITTEST(m.lpFileSystemFlags,2))
		ADDPROPERTY(m.loRet,"PersistentACLS",	BITTEST(m.lpFileSystemFlags,3))
		ADDPROPERTY(m.loRet,"FileCompression",	BITTEST(m.lpFileSystemFlags,4))
		ADDPROPERTY(m.loRet,"VolIsCompressed",	BITTEST(m.lpFileSystemFlags,15))
		m.loRet.OK = .T.
		m.loRet.MSG = "OK"
	CATCH TO loExc
		m.loRet.MSG = m.loExc.MESSAGE
	ENDTRY
	RETURN m.loRet
ENDFUNC

****
*
* Retorna informações sobre o disco
* lnDiskIndex 	= Índice do disco instalado no sistema (default = 1)
*
* Retorna objeto com propriedades:
*	.OK		= Sucesso (.t./.f.)
*	.SerialNumber	= número de série
*	.Caption		= Descrição textual do objeto
*	.Description	= Descrição
*	.Name			= Nome
*
****
FUNCTION GetDiskInfo
	LPARAMETERS lnDiskIndex
	IF VARTYPE(m.lnDiskIndex)#"N"
		m.lnDiskIndex = 1
	ENDIF

	LOCAL loWMI, loDisks, lnI, loRet
	m.loWMI = GETOBJECT("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	m.loDisks = m.loWMI.ExecQuery ("Select SerialNumber, Caption, Description, Name from Win32_Physicalmedia")
	m.lnI = 0
	m.loRet = CREATEOBJECT("EMPTY")
	ADDPROPERTY(m.loRet,"OK",.F.)
	ADDPROPERTY(m.loRet,"INDEX",0)
	FOR EACH objDisk IN m.loDisks
		m.lnI = m.lnI + 1
		IF m.lnI = m.lnDiskIndex
			m.loRet.INDEX = m.lnDiskIndex
			ADDPROPERTY(m.loRet,"SerialNumber",IIF(!ISNULL(m.objDisk.SerialNumber),m.objDisk.SerialNumber,""))
			ADDPROPERTY(m.loRet,"Caption",IIF(!ISNULL(m.objDisk.CAPTION),m.objDisk.CAPTION,""))
			ADDPROPERTY(m.loRet,"Description",IIF(!ISNULL(m.objDisk.DESCRIPTION),m.objDisk.DESCRIPTION,""))
			ADDPROPERTY(m.loRet,"Name",IIF(!ISNULL(m.objDisk.NAME),m.objDisk.NAME,""))
			m.loRet.OK = .T.
			RETURN m.loRet
		ENDIF
	NEXT
	RETURN m.loRet
ENDFUNC
