********************************************************************************
* Funções de acesso a API do windows
*
* Guionardo Furlan / 22/09/2014
*
*

****
*
* Inicia declarações de APIs do windows
*
****
PROCEDURE GSAPIInit
	IF PEMSTATUS(_SCREEN,"GSAPI",5)
		RETURN
	ENDIF

	DECLARE INTEGER timeGetTime IN "winmm.dll"
	DECLARE INTEGER ShellExecute IN Shell32 ;
		INTEGER HWND,;
		STRING lpOperation,;
		STRING lpFile,;
		STRING lpParameters,;
		STRING lpDirectory,;
		INTEGER nShowCmd
	DECLARE INTEGER IsUserAnAdmin IN Shell32
	DECLARE INTEGER GetProcessHandleFromHwnd IN "OLEACC.DLL";
		INTEGER HWND

	LOCAL lcSW
	m.lcSW = "ShowWindow"
	DECLARE INTEGER &lcSW IN WIN32API INTEGER p1, INTEGER p2 && ShowWindow
	DECLARE INTEGER FindWindow IN WIN32API STRING P1, STRING P2
	DECLARE INTEGER BringWindowToTop IN WIN32API INTEGER P1
	DECLARE INTEGER IsWindow IN WIN32API INTEGER P1

	#DEFINE TH32CS_SNAPPROCESS  2
	#DEFINE TH32CS_SNAPTHREAD   4
	#DEFINE TH32CS_SNAPMODULE   8
	#DEFINE MAX_PATH          260
	#DEFINE PE32_SIZE  296

	DECLARE INTEGER CloseHandle IN kernel32 INTEGER hObject

	DECLARE INTEGER CreateToolhelp32Snapshot IN kernel32;
		INTEGER dwFlags, INTEGER th32ProcessID

	DECLARE INTEGER Process32First IN kernel32;
		INTEGER hSnapshot, STRING @ lppe

	DECLARE INTEGER Process32Next IN kernel32;
		INTEGER hSnapshot, STRING @ lppe

	DECLARE INTEGER timeGetTime IN "winmm.dll"

*
* Funções de exclusão, movimento e cópia de arquivos
*
	PUBLIC plFileOp_ShowError		&& Flag para mostrar erro nas operações de arquivos
	m.plFileOp_ShowError = .T.

	DECLARE INTEGER DeleteFile IN kernel32;
		STRING lpFileName
	DECLARE INTEGER MoveFile IN kernel32;
		STRING lpExistingFileName,;
		STRING lpNewFileName

	DECLARE INTEGER CopyFile IN kernel32 ;
		STRING lpExistingFileName,;
		STRING lpNewFileName,;
		INTEGER bFailIfExists

	DECLARE LONG GetLastError IN WIN32API

	DECLARE LONG FormatMessage IN kernel32 ;
		LONG dwFlags, LONG lpSource, LONG dwMessageId, ;
		LONG dwLanguageId, STRING @lpBuffer, ;
		LONG nSize, LONG Arguments

	ADDPROPERTY(_SCREEN,"GSAPI",.T.)

ENDPROC

****
*
* Obtém a mensagem de erro do sistema a partir do código
*
****
FUNCTION WinApiErrMsg
	LPARAMETERS tnErrorCode
	#DEFINE FORMAT_MESSAGE_FROM_SYSTEM 0x1000

	LOCAL lcErrBuffer, lnNewErr, lnFlag, lcErrorMessage
	lnFlag = FORMAT_MESSAGE_FROM_SYSTEM
	lcErrBuffer = REPL(CHR(0),1000)
	lnNewErr = FormatMessage(lnFlag, 0, tnErrorCode, 0, @lcErrBuffer,500,0)
	lcErrorMessage = TRANSFORM(tnErrorCode) + " " + LEFT(lcErrBuffer, AT(CHR(0),lcErrBuffer)- 1 )
	RETURN lcErrorMessage
ENDFUNC

****
*
* Executa função de operação de arquivo e retorna mensagem caso M.plFileOp_ShowError for .T.
*
****
FUNCTION FileOpMessage
	LPARAMETERS lnRet, lcMsg
	IF m.lnRet # 0
		RETURN .T.
	ENDIF
	IF m.plFileOp_ShowError
		Erros(m.lcMsg+"\n\n"+WinApiErrMsg(GetLastError()),"ERRO")
	ENDIF
	RETURN .F.
ENDFUNC

****
*
* Tenta Excluir arquivo
* Retorna .T. em caso de sucesso
*
****
FUNCTION ExcluirArquivo
	LPARAMETERS lcArq
	RETURN FileOpMessage(DeleteFile(m.lcArq),"Excluir "+m.lcArq)
ENDFUNC

****
*
* Tenta Mover Arquivo
* Retorna .T. em caso de sucesso
*
*****
FUNCTION MoverArquivo
	LPARAMETERS lcArq, lcDest
	RETURN FileOpMessage(MoveFile(m.lcArq,m.lcDest),"Mover "+m.lcArq+" para "+m.lcDest)
ENDFUNC

****
*
* Tenta Copiar Arquivo
* Retorna .T. em caso de sucesso
*
****
FUNCTION CopiarArquivo
	LPARAMETERS lcArq, lcDest, llFalharSeExistir
	RETURN FileOpMessage(CopyFile(m.lcArq,m.lcDest,m.llFalharSeExistir),"Copiar "+m.lcArq+" para "+m.lcDest)
ENDFUNC


*
* Retorna o processo referente a uma janela
*
*!*	FUNCTION ProcessoWindow
*!*		LPARAMETERS lcCaption

*!*		LOCAL hSnapshot, lcBuffer

*!*		hSnapshot = CreateToolhelp32Snapshot (TH32CS_SNAPPROCESS, 0)
*!*		lcBuffer = num2dword(PE32_SIZE) + REPLI(CHR(0), PE32_SIZE-4)

*!*		IF Process32First (hSnapshot, @lcBuffer) = 1
*!*			* storing process properties to the cursor
*!*			IF VerProcessExec (lcBuffer,"calc.exe")
*!*				=MESSAGEBOX("Calculadora esta rodando")
*!*			ENDIF

*!*			DO WHILE .T.
*!*				IF Process32Next (hSnapshot, @lcBuffer) = 1
*!*					IF VerProcessExec (lcBuffer,"calc.exe")
*!*						=MESSAGEBOX("Calculadora esta rodando")
*!*					ENDIF
*!*				ELSE
*!*					EXIT
*!*				ENDIF
*!*			ENDDO
*!*		ELSE
*!*			* 87 – ERROR_INVALID_PARAMETER
*!*		ENDIF

*!*		= CloseHandle (hSnapshot)
*!*		RETURN  && main

*!*	FUNCTION VerProcessExec(lcBuffer,lcExec)
*!*		m.execname = SUBSTR(lcBuffer, 37)
*!*		m.execname = SUBSTR(m.execname, 1, AT(CHR(0),m.execname)-1)
*!*		RETURN UPPER(m.lcExec)$m.execname

*!*	FUNCTION  num2dword (lnValue)
*!*		#DEFINE m0       256
*!*		#DEFINE m1     65536
*!*		#DEFINE m2  16777216
*!*		LOCAL b0, b1, b2, b3
*!*		b3 = INT(lnValue/m2)
*!*		b2 = INT((lnValue – b3*m2)/m1)
*!*		b1 = INT((lnValue – b3*m2 – b2*m1)/m0)
*!*		b0 = MOD(lnValue, m0)
*!*		RETURN CHR(b0)+CHR(b1)+CHR(b2)+CHR(b3)

****
* Obtém lista de processos rodando sob o nome lcProcesso
* @param string Nome do processo (teste.exe)
* @param array Array de retorno
* 			(Nome: "teste.exe",
*			 Linha de Comando: "A:\teste.exe parametro",
*			 Nome de Usuário: "Guionardo",
* 			 Domínio: "Guionardo-PC")
* @return int quantidade de processos
*
****
FUNCTION AProcessos
	LPARAMETERS lcProcesso, laResult
	IF VARTYPE(m.lcProcesso)#"C"
		RETURN 0
	ENDIF
	LOCAL loWMI, loWMIItens, lcNameUser, lcUserDomain, loColProp, loWMIItem
	LOCAL colProperties, lnC, lnI
	m.loWMI= GETOBJECT("winmgmts:")
	m.loWMIItens= loWMI.execquery("SELECT * FROM Win32_Process WHERE Name = '"+lcProcesso+"'")
	STORE "" TO m.lcNameUser,m.lcUserDomain
	m.lnC = m.loWMIItens.COUNT
	IF m.lnC=0
		RETURN 0
	ENDIF
	DIMENSION m.laResult(m.lnC,4)
	m.lnI = 1
	FOR EACH loWMIItem IN m.loWMIItens
		m.loColProp= loWMIItem.GetOwner(@lcNameUser,@lcUserDomain)
		m.laResult(m.lnI,1)=loWMIItem.NAME
		m.laResult(m.lnI,2)=loWMIItem.CommandLine
		m.laResult(m.lnI,3)=m.lcNameUser
		m.laResult(m.lnI,4)=m.lcUserDomain
		m.lnI = m.lnI + 1
	ENDFOR
	RETURN m.lnC
ENDFUNC

*
***********************************************************
* CreateShortCut
* Cria atalho do windows para o aplicativo
* cDestiny -> Localização onde será criado o atalho
*              (AllUsersDesktop, AllUsersStartMenu, AllUsersPrograms, AllUsersStartup, Desktop, StartMenu, Startup	)
* cAppName -> Nome do aplicativo
* cAppPath -> Caminho do aplicativo
* cAppArgs -> Argumentos
* cAppDescription -> Descrição
* http://technet.microsoft.com/en-us/library/ee156616.aspx
*
FUNCTION CreateShortCut

	LPARAMETERS cDestiny AS STRING,;
		cAppName AS STRING, ;
		cAppPath AS STRING, ;
		cAppArgs AS STRING, ;
		cAppDescription AS STRING

	LOCAL oShell, cDesktopPath, oShortcut, ex AS EXCEPTION, ex2 AS EXCEPTION, cShortCutFile

	IF EMPTY(m.cAppDescription)
		m.cAppDescription = ''
	ENDIF
	IF EMPTY(m.cAppArgs)
		m.cAppArgs = ''
	ENDIF
	IF EMPTY(m.cAppName)
		m.cAppName = JUSTFNAME(m.cAppPath)
	ENDIF
	TRY
		IF !FILE(m.cAppPath)
			Erros('Aplicativo do atalho não existe: '+m.cAppPath)
			THROW
		ENDIF

		oShell = CREATEOBJECT("Wscript.shell")

		IF !DIRECTORY(m.cDestiny) && Se o path não existe, identifica pelos SpecialFolders
* AllUsersDesktop	Shortcuts that appear on the desktop for all users
* AllUsersStartMenu	Shortcuts that appear on the Start menu for all users
* AllUsersPrograms	Shortcuts that appear on the Programs menu for all users
* AllUsersStartup	Shortcuts to programs that are run on startup for all users
* Desktop			Shortcuts that appear on the desktop for the current user
* StartMenu			Shortcuts that appear in the current users start menu
* Startup			Shortcuts to applications that run automatically when the current user logs on to the system
			m.cDesktopPath = m.oShell.SpecialFolders(m.cDestiny)
		ENDIF
		IF !DIRECTORY(m.cDesktopPath)
			Erros('Destino do atalho não existe: '+m.cDesktopPath)
			THROW
		ENDIF
		m.cShortCutFile = ADDBS(m.cDesktopPath)+m.cAppName+'.lnk'
		m.oShortcut = m.oShell.CreateShortCut(m.cShortCutFile)

		WITH m.oShortcut
			.TargetPath = m.cAppPath
			.WorkingDirectory = JUSTPATH(m.cAppPath)
			.DESCRIPTION = m.cAppDescription
			.WindowStyle = 4  && Maximized
			.Arguments = m.cAppArgs
*.IconLocation = ???
			.SAVE
		ENDWITH
	CATCH TO ex
		IF ex.ERRORNO<>2071
			=MESSAGEBOX(TRANSFORM(ex.ERRORNO) + ". " +;
				ex.MESSAGE+CHR(13)+;
				ex.PROCEDURE+'('+TRANSFORM(ex.LINENO)+') '+ex.LINECONTENTS;
				, 48, "Creating shortcut failed")
		ENDIF
	ENDTRY
	RETURN FILE(m.cShortCutFile)
ENDFUNC
*
***********************************************************
* GetFileSizeAPI
* Retorna o tamanho de um arquivo usando a API do Windows
*
FUNCTION GetFileSizeAPI
	LPARAMETERS tcFile
	LOCAL lnHandle AS NUMBER ,;
		lnHiByte AS LONG ,;
		loLoByte AS LONG ,;
		lnSize AS LONG

	DECLARE INTEGER _lopen ;
		IN "kernel32" ;
		AS lOpen ;
		STRING lpPathName, INTEGER nRights

	DECLARE INTEGER _lclose ;
		IN "kernel32" ;
		AS lClose ;
		INTEGER hFile

	DECLARE INTEGER GetFileSize ;
		IN "kernel32" ;
		AS lGetFileSize ;
		INTEGER hFile, INTEGER @lpFileSizeHigh

	lnHandle = lOpen(tcFile, 0)
	IF lnHandle = -1 THEN
		lnSize = -1
	ELSE
		lnHiByte = 0
		lnLoByte = lGetFileSize(lnHandle, @lnHiByte)
		lClose(lnHandle)
		lnSize = INT((lnLoByte + (lnHiByte*2^32)))
	ENDIF

	RETURN lnSize
ENDFUNC
*
***********************************************************
* GetWorkGroup - Retorna o WORKGROUP ou Domínio da máquina local
*
FUNCTION GetWorkGroup
	LOCAL lcRet
	m.lcRet = .F.
	TRY
		objWMISvc = GETOBJECT( "winmgmts:\\.\root\cimv2" )
		colItems = objWMISvc.execquery( "Select * from Win32_ComputerSystem", , 48 )
		FOR EACH objItem IN colItems
			strComputerDomain = objItem.Domain
			IF objItem.PartOfDomain
				m.lcRet = "Computer Domain: " + strComputerDomain
			ELSE
				m.lcRet = "Workgroup: " + strComputerDomain
			ENDIF
		NEXT
	CATCH
	ENDTRY
	RETURN m.lcRet
ENDFUNC
*
***********************************************************
* IsUserAdmin - Retorna se o usuário logado no windows é administrador
*
FUNCTION IsUserAdmin
	LOCAL lnResult
	DECLARE INTEGER IsUserAnAdmin IN Shell32
	TRY
		lnResult = IsUserAnAdmin()
	CATCH
*** OLD OLD Version of Windows assume .T.
		lnResult = 1
	ENDTRY
	RETURN m.lnResult#0
ENDFUNC
****
*
* GetWinVer		Retorna a versão do windows
*
****
FUNCTION GetWinVer
	LOCAL lcVer
	m.lcVer = ""
	TRY
* http://fox.wikis.com/wc.dll?Wiki~GetWindowsVersion
		loWMI = GETOBJECT("winmgmts://")
		loOSs = loWMI.InstancesOf("Win32_OperatingSystem")
		FOR EACH loOS IN loOSs
			m.lcVer = loOS.CAPTION
		ENDFOR
* https://www.berezniker.com/content/pages/visual-foxpro/how-detect-64-bit-os
		DECLARE LONG GetModuleHandle IN WIN32API STRING lpModuleName
		DECLARE LONG GetProcAddress IN WIN32API LONG hModule, STRING lpProcName
		llIsWow64ProcessExists = (GetProcAddress(GetModuleHandle("kernel32"),"IsWow64Process") <> 0)

		llIs64BitOS = .F.
		IF llIsWow64ProcessExists
			DECLARE LONG GetCurrentProcess IN WIN32API
			DECLARE LONG IsWow64Process IN WIN32API LONG hProcess, LONG @ Wow64Process
			lnIsWow64Process = 0
* IsWow64Process function return value is nonzero if it succeeds
* The second output parameter value will be nonzero if VFP application is running under 64-bit OS
			IF IsWow64Process( GetCurrentProcess(), @lnIsWow64Process) <> 0
				llIs64BitOS = (lnIsWow64Process <> 0)
			ENDIF
		ENDIF
		m.lcVer = m.lcVer + " ("+IIF(m.llIs64BitOS,"x64","x86")+")"
	CATCH
		m.lcVer = "ERRO:VERSÃO WINDOWS"
	ENDTRY
	RETURN m.lcVer

ENDFUNC
************************************************************************
* wwAPI :: Createprocess
****************************************
***  Function: Calls the CreateProcess API to run a Windows application
***    Assume: Gets around RUN limitations which has command line
***            length limits and problems with long filenames.
***            Can do everything EXCEPT REDIRECTION TO FILE!
***      Pass: lcExe - Name of the Exe
***            lcCommandLine - Any command line arguments
***    Return: .t. or .f.
************************************************************************

FUNCTION Createprocess(lcExe,lcCommandLine,lnShowWindow,llWaitForCompletion)
	LOCAL hProcess, cProcessInfo, cStartupInfo

	DECLARE INTEGER CreateProcess IN kernel32 AS _CreateProcess;
		STRING   lpApplicationName,;
		STRING   lpCommandLine,;
		INTEGER  lpProcessAttributes,;
		INTEGER  lpThreadAttributes,;
		INTEGER  bInheritHandles,;
		INTEGER  dwCreationFlags,;
		INTEGER  lpEnvironment,;
		STRING   lpCurrentDirectory,;
		STRING   lpStartupInfo,;
		STRING @ lpProcessInformation

	cProcessInfo = REPLICATE(CHR(0),128)
	cStartupInfo = GetStartupInfo(lnShowWindow)

	IF !EMPTY(lcCommandLine)
		lcCommandLine = ["] + lcExe + [" ]+ lcCommandLine
	ELSE
		lcCommandLine = ""
	ENDIF

	lnResult =  _CreateProcess(lcExe,lcCommandLine,0,0,1,0,0,;
		SYS(5)+CURDIR(),cStartupInfo,@cProcessInfo)

	lhProcess = CHARTOBIN( SUBSTR(cProcessInfo,1,4) )

	IF llWaitForCompletion
		#DEFINE WAIT_TIMEOUT 0x00000102
		DECLARE INTEGER WaitForSingleObject IN kernel32.DLL ;
			INTEGER hHandle, INTEGER dwMilliseconds

		DO WHILE .T.
*** Update every 100 milliseconds
			IF WaitForSingleObject(lhProcess, 100) != WAIT_TIMEOUT
				EXIT
			ELSE
				DOEVENTS
			ENDIF
		ENDDO
	ENDIF

	DECLARE INTEGER CloseHandle IN kernel32.DLL ;
		INTEGER hObject

	CloseHandle(lhProcess)

	RETURN IIF(lnResult=1,.T.,.F.)

FUNCTION GetStartupInfo(lnShowWindow)
	LOCAL lnFlags

* creates the STARTUP structure to specify main window
* properties if a new window is created for a new process

	IF (VARTYPE(m.lnShowWindow)#"N") OR !BETWEEN(m.lnShowWindow,0,11)
* SW_FORCEMINIMIZE		11	Minimizes a window, even if the thread that owns the window is not responding. This flag should only be used when minimizing windows from a different thread.
* SW_HIDE				 0	Hides the window and activates another window.
* SW_MAXIMIZE			3	Maximizes the specified window.
* SW_MINIMIZE			6	Minimizes the specified window and activates the next top-level window in the Z order.
* SW_RESTORE			9	Activates and displays the window. If the window is minimized or maximized, the system restores it to its original size and position. An application should specify this flag when restoring a minimized window.
* SW_SHOW				5	Activates the window and displays it in its current size and position.
* SW_SHOWDEFAULT		10	Sets the show state based on the SW_ value specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application.
* SW_SHOWMAXIMIZED		3	Activates the window and displays it as a maximized window.
* SW_SHOWMINIMIZED		2	Activates the window and displays it as a minimized window.
* SW_SHOWMINNOACTIVE	7	Displays the window as a minimized window. This value is similar to SW_SHOWMINIMIZED, except the window is not activated.
* SW_SHOWNA				8	Displays the window in its current size and position. This value is similar to SW_SHOW, except that the window is not activated.
* SW_SHOWNOACTIVATE		4	Displays a window in its most recent size and position. This value is similar to SW_SHOWNORMAL, except that the window is not activated.
* SW_SHOWNORMAL			1
		lnShowWindow = 1
	ENDIF

*| typedef struct _STARTUPINFO {
*| DWORD cb; 4
*| LPTSTR lpReserved; 4
*| LPTSTR lpDesktop; 4
*| LPTSTR lpTitle; 4
*| DWORD dwX; 4
*| DWORD dwY; 4
*| DWORD dwXSize; 4
*| DWORD dwYSize; 4
*| DWORD dwXCountChars; 4
*| DWORD dwYCountChars; 4
*| DWORD dwFillAttribute; 4
*| DWORD dwFlags; 4
*| WORD wShowWindow; 2
*| WORD cbReserved2; 2
*| LPBYTE lpReserved2; 4
*| HANDLE hStdInput; 4
*| HANDLE hStdOutput; 4
*| HANDLE hStdError; 4
*| } STARTUPINFO, *LPSTARTUPINFO; total: 68 bytes


	#DEFINE STARTF_USESTDHANDLES 0x0100
	#DEFINE STARTF_USESHOWWINDOW 1
	#DEFINE SW_HIDE 0
	#DEFINE SW_SHOWMAXIMIZED 3
	#DEFINE SW_SHOWNORMAL 1

	lnFlags = STARTF_USESHOWWINDOW

	RETURN binToChar(80) +;
		binToChar(0) + binToChar(0) + binToChar(0) +;
		binToChar(0) + binToChar(0) + binToChar(0) + binToChar(0) +;
		binToChar(0) + binToChar(0) + binToChar(0) +;
		binToChar(lnFlags) +;
		binToWordChar(lnShowWindow) +;
		binToWordChar(0) + binToChar(0) +;
		binToChar(0) + binToChar(0) + binToChar(0) + REPLICATE(CHR(0),30)


************************************************************************
FUNCTION CHARTOBIN(lcBinString,llSigned)
****************************************
***  Function: Binary Numeric conversion routine.
***            Converts DWORD or Unsigned Integer string
***            to Fox numeric integer value.
***      Pass: lcBinString -  String that contains the binary data
***            llSigned    -  if .T. uses signed conversion
***                           otherwise value is unsigned (DWORD)
***    Return: Fox number
************************************************************************
	LOCAL m.i, lnWord

	lnWord = 0
	FOR m.i = 1 TO LEN(lcBinString)
		lnWord = lnWord + (ASC(SUBSTR(lcBinString, m.i, 1)) * (2 ^ (8 * (m.i - 1))))
	ENDFOR

	IF llSigned AND lnWord > 0x80000000
		lnWord = lnWord - 1 - 0xFFFFFFFF
	ENDIF

	RETURN lnWord

*  wwAPI :: CharToBin

************************************************************************
FUNCTION binToChar(lnValue)
****************************************
***  Function: Creates a DWORD value from a number
***      Pass: lnValue - VFP numeric integer (unsigned)
***    Return: binary string
************************************************************************

	LOCAL byte(4)

	IF lnValue < 0
		lnValue = lnValue + 4294967296
	ENDIF

	byte(1) = lnValue % 256
	byte(2) = BITRSHIFT(lnValue, 8) % 256
	byte(3) = BITRSHIFT(lnValue, 16) % 256
	byte(4) = BITRSHIFT(lnValue, 24) % 256

	RETURN CHR(byte(1))+CHR(byte(2))+CHR(byte(3))+CHR(byte(4))

*  wwAPI :: BinToChar

************************************************************************
FUNCTION binToWordChar(lnValue)
****************************************
***  Function: Creates a DWORD value from a number
***      Pass: lnValue - VFP numeric integer (unsigned)
***    Return: binary string
************************************************************************

	RETURN CHR(MOD(m.lnValue,256)) + CHR(INT(m.lnValue/256))

*****************************************************************************************
* Function....:	 ReduceMemory()
* Author......:  Bernard Bout
* Date........:  05/12/2007 3:03:15 PM
* Returns.....:
* Parameters..:
* Notes.......:  reduces memory usage for vfp
* URL.........: http://www.foxite.com/faq/default.aspx?id=55
*****************************************************************************************

FUNCTION ReduceMemory
	DECLARE INTEGER SetProcessWorkingSetSize IN kernel32 AS SetProcessWorkingSetSize  ;
		INTEGER hProcess , ;
		INTEGER dwMinimumWorkingSetSize , ;
		INTEGER dwMaximumWorkingSetSize
	DECLARE INTEGER GetCurrentProcess IN kernel32 AS GetCurrentProcess
	nProc = GetCurrentProcess()
	bb = SetProcessWorkingSetSize(nProc,-1,-1)
ENDFUNC

****
*
* MoveArquivo	Tenta mover ou copiar arquivo(s)
* Origem,Destino,@Movimentos com suceso, @Total de arquivos a mover
* Return .T./.F.
*
****
#DEFINE ArqTentativas 20
FUNCTION MoveArquivo
	LPARAMETERS lcOrigem, lcDestino, lnMovSuc, lnMovTot
	IF VARTYPE(m.lcOrigem)#"C" OR (!FILE(m.lcOrigem)) OR VARTYPE(m.lcDestino)#"C"
		RETURN .F.
	ENDIF

	m.lnMovSuc = 0

	IF ("*"$m.lcOrigem) OR ("?"$m.lcOrigem)
* Se houver caracteres curinga, processa a lista de arquivos/pastas
		LOCAL ARRAY laArq(1,5)
		m.lnMovTot = ADIR(m.laArq,m.lcOrigem)
		IF m.lnMovTot = 0
			RETURN .F.
		ENDIF
		LOCAL lnI, lnR
		m.lnR = .T.
		FOR m.lnI = 1 TO ALEN(m.laArq,1)
			IF MoveArquivo(m.laArq(m.lnI,1),JUSTPATH(m.lcDestino))
				m.lnMovSuc = m.lnMovSuc + 1
			ELSE
				m.lnR = .F.
			ENDIF
		NEXT
		RETURN m.lnR
	ENDIF

	m.lnMovTot = 1

	m.lcOrigem = FULLPATH(m.lcOrigem)

	LOCAL lcPastaDestino
	IF DIRECTORY(m.lcDestino)
		m.lcPastaDestino = m.lcDestino
		m.lcDestino = ADDBS(m.lcDestino)+JUSTFNAME(m.lcOrigem)
	ELSE
		m.lcPastaDestino = JUSTPATH(FULLPATH(m.lcDestino))
	ENDIF

	IF !DIRECTORY(m.lcPastaDestino)
		MKDIR (m.lcPastaDestino)
	ENDIF
	IF !DIRECTORY(m.lcPastaDestino)
		MESSAGEBOX(CHR(13)+"Impossível mover o arquivo "+CHR(13)+CHR(13)+;
			m.lcOrigem+CHR(13)+CHR(13)+;
			"para"+CHR(13)+CHR(13)+;
			m.lcDestino+CHR(13)+CHR(13)+;
			"Pasta destino não existe e não pode ser criada."+CHR(13),48,"Erro")
		RETURN .F.
	ENDIF
	DECLARE INTEGER MoveFile IN kernel32;
		STRING lpExistingFileName,;
		STRING lpNewFileName
	LOCAL lnSuc, lnTenta
	m.lnSuc = 0
	m.lnTenta = ArqTentativas
	DO WHILE (m.lnSuc=0) AND (m.lnTenta>0)
		m.lnSuc = MoveFile(m.lcOrigem,m.lcDestino)
		IF m.lnSuc = 0
			m.lnTenta = m.lnTenta - 1
			WAIT WINDOW "Movendo "+m.lcOrigem+" para "+m.lcDestino+;
				" ... (Tentativa "+TRANSFORM(ArqTentativas-m.lnTenta)+"/"+TRANSFORM(ArqTentativas)+")" TIMEOUT 1
		ENDIF
	ENDDO
	IF m.lnSuc = 1
		m.lnMovSuc = 1
		RETURN .T.
	ENDIF
	WAIT WINDOW "Não foi possível mover o arquivo "+m.lcOrigem+CHR(13)+"Tentando copiar..." NOWAIT
	DECLARE INTEGER CopyFile IN kernel32 ;
		STRING lpExistingFileName,;
		STRING lpNewFileName
	IF CopyFile(m.lcOrigem,m.lcDestino)=0
		MESSAGEBOX(CHR(13)+"Impossível copiar o arquivo "+CHR(13)+CHR(13)+;
			m.lcOrigem+CHR(13)+CHR(13)+;
			"para"+CHR(13)+CHR(13)+;
			m.lcDestino+CHR(13),48,"Erro")
		WAIT CLEAR
		RETURN .F.
	ELSE
		=ExcluiArquivo(m.lcOrigem)
		RETURN .T.
	ENDIF
	WAIT WINDOW "Cópia de "+m.lcOrigem+" efetuada com sucesso para "+m.lcDestino TIMEOUT 5

	RETURN .T.
ENDFUNC
****
*
* ExcluiArquivo		Tenta excluir um arquivo
*
****
FUNCTION ExcluiArquivo
	LPARAMETERS lcArq, lnTenta
	IF !FILE(m.lcArq)
		RETURN .T.
	ENDIF
	LOCAL lnSuc
	m.lnSuc = 0
	IF VARTYPE(m.lnTenta)#"N" OR !BETWEEN(m.lnTenta,1,20)
		m.lnTenta = ArqTentativas
	ENDIF
	DECLARE INTEGER DeleteFile IN kernel32;
		STRING lpFileName
	DO WHILE (m.lnSuc=0) AND (m.lnTenta>0)
		m.lnSuc = DeleteFile(m.lcArq)
		IF m.lnSuc = 0
			m.lnTenta = m.lnTenta - 1
			WAIT WINDOW "Excluindo "+m.lcArq+;
				" ... (Tentativa "+TRANSFORM(ArqTentativas-m.lnTenta)+"/"+TRANSFORM(ArqTentativas)+")" TIMEOUT 1
		ENDIF
	ENDDO
	IF m.lnSuc = 0
		LOCAL lcDel, lnEspera
		m.lcDel = "/c DEL "+m.lcArq
		m.lnEspera = 2	&& 2 segundos
		WAIT WINDOW "Excluindo "+m.lcArq+" via CMD..." NOWAIT
		RUN &lcDel
		DO WHILE FILE(m.lcArq) AND (m.lnEspera>0)
			WAIT WINDOW "Excluindo "+m.lcArq+" via CMD (aguardando)..." TIMEOUT 0.2
			m.lnEspera = m.lnEspera - 0.2
		ENDDO
	ENDIF
	IF !FILE(m.lcArq)
		m.lnSuc = 1
	ENDIF

	IF m.lnSuc = 0
		MESSAGEBOX(CHR(13)+"Não foi possível excluir o arquivo "+CHR(13)+CHR(13)+;
			m.lcArq+CHR(13)+CHR(13)+;
			"Você pode tentar excluí-lo manualmente!",48,"Erro")
		RETURN .F.
	ENDIF
	RETURN .T.
ENDFUNC
****
*
* FileExtract	Extrai arquivos incorporados no executável
* lcArquivo é o nome do arquivo interno
* lcDestino, se não informado será gravado em TEMP
*
* Retorna o caminho do arquivo salvo ou vazio caso não seja possível gravá-lo
*
****
FUNCTION FileExtract
	LPARAMETERS lcArquivo, lcDestino
	IF VARTYPE(m.lcArquivo)#"C"
		RETURN ""
	ENDIF
	m.lcArquivo = JUSTFNAME(m.lcArquivo)
	IF !FILE(m.lcArquivo)
		RETURN ""
	ENDIF
	IF EMPTY(m.lcDestino)
		m.lcDestino = FORCEEXT(ADDBS(SYS(2023))+SYS(3),JUSTEXT(m.lcArquivo))
	ENDIF
	IF FILE(m.lcDestino) AND !ExcluiArquivo(m.lcDestino)
		RETURN ""
	ENDIF
	IF STRTOFILE(FILETOSTR(m.lcArquivo),m.lcDestino,0)=0
		RETURN ""
	ENDIF
	IF FILE(m.lcDestino)
		RETURN m.lcDestino
	ENDIF
	RETURN ""
ENDFUNC


****
*
* FileOpWithProgressBar 	Copia/Move arquivo usando o shell
*
* tcSource			Arquivo origem
* tcDestination		Destino
* tcAction			MOVE ou COPY ou DELETE
* lcUserCanceled	Usuário cancelou a operação (usar com referência @)
*
* Retorno			.T./.F.
*
****
* Shell File Operations

#DEFINE FO_MOVE           0x0001
#DEFINE FO_COPY           0x0002
#DEFINE FO_DELETE         0x0003
*#DEFINE FO_RENAME         0x0004

#DEFINE FOF_MULTIDESTFILES         0x0001
#DEFINE FOF_CONFIRMMOUSE           0x0002
#DEFINE FOF_SILENT                 0x0004  && don't create progress/report
#DEFINE FOF_RENAMEONCOLLISION      0x0008
#DEFINE FOF_NOCONFIRMATION         0x0010  && Don't prompt the user.
#DEFINE FOF_WANTMAPPINGHANDLE      0x0020  && Fill in SHFILEOPSTRUCT.hNameMappings
&& Must be freed using SHFreeNameMappings
#DEFINE FOF_ALLOWUNDO              0x0040  && DELETE - sends the file to the Recycle Bin
#DEFINE FOF_FILESONLY              0x0080  && on *.*, do only files
#DEFINE FOF_SIMPLEPROGRESS         0x0100  && don't show names of files
#DEFINE FOF_NOCONFIRMMKDIR         0x0200  && don't confirm making any needed dirs
#DEFINE FOF_NOERRORUI              0x0400  && don't put up error UI
#DEFINE FOF_NOCOPYSECURITYATTRIBS  0x0800  && dont copy NT file Security Attributes
#DEFINE FOF_NORECURSION            0x1000  && don't recurse into directories.

* http://www.berezniker.com/content/pages/visual-foxpro/file-operations-progressbar

FUNCTION FileOpWithProgressbar
	LPARAMETERS tcSource, tcDestination, tcAction, tlUserCanceled
	LOCAL lcSourceString, lcDestString, nStringBase, lcFileOpStruct, lnFlag, lnStringBase
	LOCAL loHeap, lcAction, lnRetCode, llCanceled, laActionList[1]

	DECLARE INTEGER SHFileOperation IN SHELL32.DLL STRING @ LPSHFILEOPSTRUCT
* Heap allocation class

	loHeap = NEWOBJECT('Heap',"GS_CLSHEAP.PRG")

	lcAction = UPPER(IIF( EMPTY( tcAction) OR VARTYPE(tcAction) <> "C", "COPY", tcAction))
* Convert Action name into function parameter
	ALINES(laActionList, "MOVE,COPY,DELETE",",")
	lnAction = ASCAN(laActionList, lcAction)
	IF lnAction = 0
* Unknown action
		RETURN NULL
	ENDIF

	lcSourceString = tcSource + CHR(0) + CHR(0)
	lnFlag = FOF_NOCONFIRMATION + FOF_NOCONFIRMMKDIR + FOF_NOERRORUI

	IF m.lnAction=3
	* TODO: Arquivo/pasta não é enviado para lixeira, verificar
		m.lnFlag = m.lnFlag + FOF_ALLOWUNDO
		m.tcDestination = ""
	ENDIF
	lcDestString   = tcDestination + CHR(0) + CHR(0)
	lnStringBase   = loHeap.AllocBlob(lcSourceString+lcDestString)

	lcFileOpStruct  = NumToLONG(_SCREEN.HWND) + ;
		NumToLONG(lnAction) + ;
		NumToLONG(lnStringBase) + ;
		NumToLONG(lnStringBase + LEN(lcSourceString)) + ;
		NumToWORD(lnFlag) + ;
		NumToLONG(0) + NumToLONG(0) + NumToLONG(0)

	lnRetCode = SHFileOperation(@lcFileOpStruct)

* Did user canceled operation?
	tlUserCanceled= ( SUBSTR(lcFileOpStruct, 19, 4) <> NumToLONG(0) )

	RETURN (lnRetCode = 0)
ENDFUNC
