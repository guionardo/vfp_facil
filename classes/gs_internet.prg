CLEAR
PUBLIC loI
LOCAL lcRetorno
loI = CREATEOBJECT("Internet")
*!*	m.loI.PROXYHOST = 'localhost'
*!*	m.loI.PROXYPORT = 8080
*!*	m.loI.PROXYUSER = 'userproxy'
*!*	m.loI.PROXYPASS = 'proxy'
*!*	m.loI.PROXYENABLE = .T.
m.loI.ProxyLoad('a:\tbyte\proxy.cfg')
? m.loI.ProxyTest()
? m.loI.ProxySave()
m.lcRetorno = ''

*? m.loI.HTTPGet('https://tbyte.com.br/tbyte/tbyte.zip',@lcRetorno,.F.,.T.)
*? STRTOFILE(m.lcRetorno,'A:\temp\tbyte.zip')

*
* Classe de comunica��o Internet
*
#DEFINE HTTPREQUEST_SETCREDENTIALS_FOR_SERVER 0
#DEFINE HTTPREQUEST_SETCREDENTIALS_FOR_PROXY 1
#DEFINE HTTPREQUEST_PROXYSETTING_PROXY 2
#DEFINE GENERIC_READ     2147483648
#DEFINE GENERIC_WRITE     1073741824
#DEFINE INTERNET_INVALID_PORT_NUMBER 0
#DEFINE INTERNET_OPEN_TYPE_DIRECT  1
#DEFINE INTERNET_OPEN_TYPE_PROXY  3
#DEFINE INTERNET_DEFAULT_FTP_PORT  21
#DEFINE INTERNET_FLAG_ASYNC    268435456
#DEFINE INTERNET_FLAG_FROM_CACHE  16777216
#DEFINE INTERNET_FLAG_OFFLINE   16777216
#DEFINE INTERNET_FLAG_CACHE_IF_NET_FAIL 65536
#DEFINE INTERNET_OPEN_TYPE_PRECONFIG 0
#DEFINE INTERNET_SERVICE_FTP   1
#DEFINE INTERNET_SERVICE_GOPHER   2
#DEFINE INTERNET_SERVICE_HTTP   3
#DEFINE FTP_TRANSFER_TYPE_ASCII   1
#DEFINE FTP_TRANSFER_TYPE_BINARY  2
#DEFINE FILE_ATTRIBUTE_NORMAL   128
#DEFINE FORMAT_MESSAGE_FROM_SYSTEM 0x1000
#DEFINE PROXY_TEST_URL "http://tbyte.com.br"
#DEFINE FTP_AGENT "VFP9_INTERNET"

DEFINE CLASS Internet AS CUSTOM
	PROXYHOST = ''
	PROXYPORT = 0
	PROXYUSER = ''
	PROXYPASS = ''
	PROXYENABLE = .F.
	PROXYARQ = ''

	FTPHost = 'ftp.tbyte.com.br'
	FTPUser = '***'
	FTPPass = '***'
	FTPServerPath = ''
	FTPBinary = .T.
	FTPOpen = 0
	FTPSession = 0

	ArqLog = .F.
	BasePath = ''

	HIDDEN BASECLASS, COMMENT, CONTROLCOUNT, CONTROLS, HEIGHT, HELPCONTEXTID, LEFT, OBJECTS, PICTURE, TOP, WHATSTHISHELP, WIDTH

	PROCEDURE INIT(lcArqProxy)
		DECLARE INTEGER InternetOpen IN WININET.DLL;
			STRING  sAgent,;
			INTEGER lAccessType,;
			STRING  sProxyName,;
			STRING  sProxyBypass,;
			STRING  lFlags
*
		DECLARE INTEGER InternetCloseHandle IN WININET.DLL;
			INTEGER hInet
*
		DECLARE INTEGER InternetConnect IN WININET.DLL;
			INTEGER hInternetSession,;
			STRING  sServerName,;
			INTEGER nServerPort,;
			STRING  sUsername,;
			STRING  sPassword,;
			INTEGER lService,;
			INTEGER lFlags,;
			INTEGER lContext
*
		DECLARE INTEGER FtpGetFile IN WININET.DLL;
			INTEGER hFtpSession,;
			STRING  lpszRemoteFile,;
			STRING  lpszNewFile,;
			INTEGER fFailIfExists,;
			INTEGER dwFlagsAndAttributes,;
			INTEGER dwFlags,;
			INTEGER dwContext
*
		DECLARE INTEGER FtpDeleteFile IN WININET.DLL ;
			INTEGER nFile,;
			STRING  lpszFileName
*
		DECLARE INTEGER InternetWriteFile IN WININET.DLL;
			INTEGER   hFile,;
			STRING  @ sBuffer,;
			INTEGER   lNumBytesToWrite,;
			INTEGER @ dwNumberOfBytesWritten

		DECLARE LONG InternetReadFile IN wininet.DLL ;
			LONG     hFtpSession, ;
			STRING  @lpBuffer, ;
			LONG     dwNumberOfBytesToRead, ;
			LONG    @lpNumberOfBytesRead
*
		DECLARE INTEGER FtpOpenFile IN WININET.DLL;
			INTEGER hFtpSession,;
			STRING  sFileName,;
			INTEGER lAccess,;
			INTEGER lFlags,;
			INTEGER lContext

		DECLARE LONG GetLastError IN WIN32API

		DECLARE LONG FormatMessage IN kernel32 ;
			LONG dwFlags, LONG lpSource, LONG dwMessageId, ;
			LONG dwLanguageId, STRING @lpBuffer, ;
			LONG nSize, LONG Arguments

		DECLARE INTEGER GetActiveWindow ;
			IN Win32API
		DECLARE INTEGER GetWindow IN Win32API ;
			INTEGER HWND, INTEGER nType
		DECLARE INTEGER GetWindowText IN Win32API ;
			INTEGER HWND, STRING @cText, INTEGER nType
		DECLARE INTEGER DeleteFile IN kernel32;
			STRING lpFileName

		IF "vfp9.exe"$LOWER(APPLICATION.SERVERNAME)
			THIS.BasePath = "A:\TEMP"
		ELSE
			THIS.BasePath = JUSTPATH(APPLICATION.SERVERNAME)
		ENDIF

		IF VARTYPE(m.lcArqProxy)='C' OR (VARTYPE(m.lcArqProxy)='L' AND m.lcArqProxy)
			THIS.ProxyLoad(m.lcArqProxy)
		ENDIF
	ENDFUNC

	PROCEDURE AddLog(lcMsg)
		IF ISNULL(THIS.ArqLog)
			RETURN
		ENDIF
		IF VARTYPE(THIS.ArqLog)#"C" OR EMPTY(THIS.ArqLog)
			THIS.ArqLog = ADDBS(THIS.BasePath)+"LOG\"+FORCEEXT(JUSTFNAME(APPLICATION.SERVERNAME),"LOG")
			IF !DIRECTORY(JUSTPATH(THIS.ArqLog))
				TRY
					MKDIR (JUSTPATH(THIS.ArqLog))
				CATCH
					THIS.ArqLog = .NULL.
				ENDTRY
			ENDIF
		ENDPROC
		IF ISNULL(THIS.ArqLog)
			RETURN
		ENDIF
		TRY
			m.lcMsg = DTOS(DATE())+" "+RIGHT(TTOC(DATETIME(),1),6)+" "+TRANSFORM(m.lcMsg)
			STRTOFILE(m.lcMsg+CHR(13)+CHR(10),THIS.ArqLog,.T.)
		CATCH
			THIS.ArqLog = .NULL.
		ENDTRY
	ENDPROC

	PROCEDURE PROXYHOST_ASSIGN(lcHost)
		IF VARTYPE(m.lcHost)#"C"
			RETURN
		ENDIF
		THIS.PROXYHOST = m.lcHost
	ENDPROC

	PROCEDURE PROXYPORT_ASSIGN(lnPort)
		IF VARTYPE(m.lnPort)="C"
			m.lnPort = VAL(m.lnPort)
		ENDIF
		IF VARTYPE(m.lnPort)#"N" OR !BETWEEN(m.lnPort,1,65535)
			RETURN
		ENDIF
		THIS.PROXYPORT = m.lnPort
	ENDPROC

	PROCEDURE PROXYUSER_ASSIGN(lcUser)
		IF VARTYPE(m.lcUser)#"C"
			RETURN
		ENDIF
		THIS.PROXYUSER = m.lcUser
	ENDPROC

	PROCEDURE PROXYPASS_ASSIGN(lcPass)
		IF VARTYPE(m.lcPass)#"C"
			RETURN
		ENDIF
		THIS.PROXYPASS = m.lcPass
	ENDPROC

	PROCEDURE PROXYENABLE_ASSIGN(llEnable)
		IF VARTYPE(m.llEnable)#"L"
			RETURN
		ENDIF
		THIS.PROXYENABLE = m.llEnable
	ENDPROC

	PROCEDURE MSG(lcMsg, llErro)
		MESSAGEBOX(m.lcMsg,IIF(m.llErro,16,64),'Conex�o internet')
	ENDPROC

	FUNCTION WinApiErrMsg
		LPARAMETERS tnErrorCode
		IF VARTYPE(m.tnErrorCode)#"N"
			m.tnErrorCode = GetLastError()
		ENDIF
		#DEFINE FORMAT_MESSAGE_FROM_SYSTEM 0x1000
		DECLARE LONG FormatMessage IN kernel32 ;
			LONG dwFlags, LONG lpSource, LONG dwMessageId, ;
			LONG dwLanguageId, STRING @lpBuffer, ;
			LONG nSize, LONG Arguments

		LOCAL lcErrBuffer, lnNewErr, lnFlag, lcErrorMessage
		lnFlag = FORMAT_MESSAGE_FROM_SYSTEM
		lcErrBuffer = REPL(CHR(0),1000)
		lnNewErr = FormatMessage(lnFlag, 0, tnErrorCode, 0, @lcErrBuffer,500,0)
		lcErrorMessage = TRANSFORM(tnErrorCode) + " " + LEFT(lcErrBuffer, AT(CHR(0),lcErrBuffer)- 1 )
		RETURN lcErrorMessage
	ENDFUNC

*
* Carregar informa��es de proxy
*
	FUNCTION ProxyLoad(lcArq)
		IF VARTYPE(m.lcArq)#"C"
			m.lcArq = ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+"PROXY.CFG"
		ENDIF
		IF !FILE(m.lcArq)
			RETURN .F.
		ENDIF
		LOCAL lnH, lcH
		m.lnH = FOPEN(m.lcArq,10)
		IF m.lnH = -1
			RETURN .F.
		ENDIF
		m.lcH = FGETS(m.lnH)
		FCLOSE(m.lnH)
		IF GETWORDCOUNT(m.lcH,'|')#5
			RETURN .F.
		ENDIF
		THIS.PROXYENABLE = IIF(GETWORDNUM(m.lcH,1,'|')='1',.T.,.F.)
		THIS.PROXYHOST = GETWORDNUM(m.lcH,2,'|')
		THIS.PROXYPORT = INT(VAL(GETWORDNUM(m.lcH,3,'|')))
		THIS.PROXYUSER = GETWORDNUM(m.lcH,4,'|')
		THIS.PROXYPASS = GETWORDNUM(m.lcH,5,'|')
		THIS.PROXYARQ = m.lcArq
	ENDFUNC

*
* Grava informa��es do proxy
*
	FUNCTION ProxySave(lcArq)
		IF VARTYPE(m.lcArq)#"C"
			IF EMPTY(THIS.PROXYARQ)
				m.lcArq = ADDBS(JUSTPATH(APPLICATION.SERVERNAME))+"PROXY.CFG"
			ELSE
				m.lcArq = THIS.PROXYARQ
			ENDIF
		ENDIF
		LOCAL lnH, lcH
		IF FILE(m.lcArq) AND DeleteFile(m.lcArq)=0
			THIS.MSG("N�o foi poss�vel excluir o arquivo de configura��o do proxy em "+m.lcArq,.T.)
			RETURN .F.
		ENDIF
		m.lnH = FCREATE(m.lcArq)
		IF m.lnH=-1
			RETURN .F.
		ENDIF

		FWRITE(m.lnH,IIF(THIS.PROXYENABLE,'1','0')+'|'+THIS.PROXYHOST+'|'+TRANSFORM(THIS.PROXYPORT)+'|'+THIS.PROXYUSER+'|'+THIS.PROXYPASS)
		FCLOSE(m.lnH)
		RETURN .T.
	ENDFUNC

*
* Testa conectividade com proxy
*
	FUNCTION ProxyTest
		IF EMPTY(THIS.PROXYHOST) OR ;
				(!BETWEEN(THIS.PROXYPORT,1,65535)) OR ;
				EMPTY(THIS.PROXYUSER) OR ;
				EMPTY(THIS.PROXYPASS)
			THIS.MSG('Configura��es de proxy est�o inv�lidas!',.T.)
			RETURN .F.
		ENDIF

		LOCAL llProxy, llOk,lcRetorno
		m.llProxy = THIS.PROXYENABLE
		THIS.PROXYENABLE = .T.

		IF !THIS.HTTPGet(PROXY_TEST_URL,@lcRetorno,.F.,.F.,'Teste de proxy')
			THIS.PROXYENABLE = .F.
			IF !THIS.HTTPGet(PROXY_TEST_URL,@lcRetorno,.F.,.F.,'Teste de proxy desativado')
				THIS.MSG('Teste de proxy falhou!',.T.)
				THIS.PROXYENABLE = m.llProxy
				RETURN .F.
			ELSE
				THIS.MSG('Conex�o OK sem necessidade de proxy!')
				RETURN .T.
			ENDIF
		ENDIF

		THIS.PROXYENABLE = m.llProxy
		THIS.MSG('Teste de proxy OK')

		RETURN .T.

	ENDFUNC


*
* Get via HTTP
*
	FUNCTION HTTPGet(lcURL, lcRetorno, lcMsg, llBinary, lcMsgAnt)
		IF VARTYPE(m.lcURL)#'C'
			RETURN .F.
		ENDIF
		LOCAL loHTTP, loExc AS EXCEPTION, llErro
		m.lcRetorno = ''
		IF VARTYPE(m.lcMsg)#"C"
			m.lcMsg = "Obtendo dados de "+m.lcURL
		ENDIF
		IF VARTYPE(m.lcMsgAnt)="C"
			m.lcMsg = LEFT(ALLTRIM(m.lcMsgAnt)+CHR(13)+m.lcMsg,255)
		ENDIF
		IF VARTYPE(m.llBinary)#"L"
			m.llBinary = .F.
		ENDIF
		WAIT WINDOW NOWAIT m.lcMsg
		LOCAL lcErro
		m.lcErro = 'Erro ao acessar a URL:'+CHR(13)+;
			m.lcURL+CHR(13)+;
			IIF(THIS.PROXYENABLE, 'Via proxy '+THIS.PROXYHOST+':'+TRANSFORM(THIS.PROXYPORT)+' (USUARIO '+THIS.PROXYUSER+')'+ CHR(13)+CHR(13),'')
		TRY
			m.lcRetorno = ''
			loHTTP = CREATEOBJECT("WinHttp.WinHttpRequest.5.1")
			LOCAL lcMsg
			m.lcMsg = ''

			loHTTP.OPEN("GET", m.lcURL, .F.)
			IF THIS.PROXYENABLE
				m.loHTTP.SetProxy(HTTPREQUEST_PROXYSETTING_PROXY,THIS.PROXYHOST+":"+TRANSFORM(THIS.PROXYPORT))
				m.loHTTP.SetCredentials(THIS.PROXYUSER,THIS.PROXYPASS,HTTPREQUEST_SETCREDENTIALS_FOR_PROXY)
			ENDIF
			loHTTP.SEND()
			IF m.loHTTP.STATUS = 200
				IF m.llBinary
					m.lcRetorno = m.loHTTP.ResponseBody
				ELSE
					m.lcRetorno = m.loHTTP.ResponseText
				ENDIF
			ELSE
				THIS.MSG(m.lcErro+TRANSFORM(m.loHTTP.STATUS)+" "+m.loHTTP.STATUSTEXT,.T.)
				m.llErro = .T.
			ENDIF
		CATCH TO loExc
			THIS.MSG(m.lcErro+m.loExc.MESSAGE,.T.)
			m.llErro = .T.
		ENDTRY

		RETURN !m.llErro
	ENDFUNC

	FUNCTION FTPDownload2(lcFTPFile, lcPath)
		IF !THIS.FTPConnect(THIS.FTPHost, THIS.FTPUser, THIS.FTPPass)
			RETURN .F.
		ENDIF

		LOCAL lnH
		m.lnH = FtpOpenFile(THIS.FTPSession,m.lcFTPFile,GENERIC_READ,FTP_TRANSFER_TYPE_BINARY,0)
		IF m.lnH = 0
			InternetCloseHandle(THIS.FTPOpen)
			THIS.MSG("Erro "+ THIS.WinApiErrMsg(),.T.)
			RETURN .F.
		ENDIF 
		

*** Realizar la conexi�n...
*** Abrir el fichero en el servidor
			nFichFTP = FtpOpenFile( ;
				nFtp, ;
				"/developr/fox/kb/index.txt", ;
				GENERIC_READ, ;
				FTP_TRANSFER_TYPE_ASCII, ;
				0 )
			IF nFichFTP = 0
				MESSAGEBOX( "Error: " ;
					+ LTRIM( STR( GetLastError() ) ) ;
					+ " en FtpOpenFile.", 16 )
				InternetCloseHandle( nFtp )
				InternetCloseHandle( nInternet )
				RETURN
			ENDIF
*** Abir el fichero en el cliente
			nFich = FCREATE( "index.txt" )
*** Contruir las variables necesarias
			nTama = 0
			nLen = 1
*** Bucle de lectura
			DO WHILE nLen # 0
				cBuffer = REPLICATE( CHR(0), 2048 )
*** Leer del fichero en el servidor
				InternetReadFile( ;
					nFichFTP, ;
					@cBuffer, ;
					LEN( cBuffer ), ;
					@nLen )
*** Escribir el fichero en el cliente
				FWRITE( ;
					nFich, ;
					SUBSTR( cBuffer, 1, nLen ) )
*** Aumentar el tama�o total
				nTama = nTama + nLen
				WAIT WIND "Recibidos " ;
					+ LTRIM( STR( nTama ) ) NOWAIT
			ENDDO
			WAIT CLEAR
*** Cerrar el fichero local
			FCLOSE( nFich )
*** Cerrar el fichero en el servidor
			InternetCloseHandle( nFichFTP )
*** Realizar la desconexi�n...
		ENDFUNC
*
* FTP
*
	FUNCTION FTPDownload
		LPARAMETERS lcPath, lcDestino, loProxy
		THIS.AddLog("FTP Download "+m.lcPath+" -> "+m.lcDestino)
		WAIT WINDOW NOWAIT "FTP DOWNLOAD"+CHR(13)+THIS.FTPHost+CHR(13)+m.lcPath
		LOCAL sAgent, sProxyName, sProxyBypass, lFlags, lpszRemoteFile, lpszNewFile, fFailIfExists, dwContext, llTransferType, lnResult, lcErro
		sProxyName = CHR(0)
		sProxyBypass = CHR(0)
		lFlags = 0
		IF VARTYPE(m.loProxy)="L"
			m.loProxy = GetProxy()
		ENDIF

		LOCAL lnH, lnS
		m.lnResult = 0
		IF !THIS.PROXYENABLE
			m.lnH = InternetOpen(FTP_AGENT, INTERNET_OPEN_TYPE_DIRECT, sProxyName, sProxyBypass, lFlags)
		ELSE
			sProxyName = THIS.PROXYHOST+':'+TRANSFORM(THIS.PROXYPORT)
			m.lnH = InternetOpen (sAgent, INTERNET_OPEN_TYPE_PROXY, sProxyName, sProxyBypass, lFlags)
			InternetSetOption(m.lnH, INTERNET_OPTION_PROXY_PASSWORD, THIS.PROXYPASS, LEN(THIS.PROXYPASS))
			InternetSetOption(m.lnH, INTERNET_OPTION_PROXY_USERNAME, THIS.PROXYUSER, LEN(THIS.PROXYUSER))
		ENDIF
		IF m.lnH > 0
			m.lnS = InternetConnect (m.lnH, THIS.FTPHost, INTERNET_INVALID_PORT_NUMBER, THIS.FTPUser, THIS.FTPPass, INTERNET_SERVICE_FTP, 0, 0)
			IF m.lnS > 0
				lpszRemoteFile = THIS.FTPServerPath+m.lcPath
				lpszNewFile    = m.lcDestino
				fFailIfExists  = 0
				dwContext      = 0
				IF THIS.FTPBinary
					m.llTransferType = FTP_TRANSFER_TYPE_BINARY
				ELSE
					m.llTransferType = FTP_TRANSFER_TYPE_ASCII
				ENDIF
				lnResult = FtpGetFile (m.lnS, lpszRemoteFile, lpszNewFile, fFailIfExists, FILE_ATTRIBUTE_NORMAL, m.llTransferType, dwContext)
				IF m.lnResult = 0
					m.lcErro = THIS.WinApiErrMsg()
					THIS.AddLog("!FtpGetFile = "+m.lcErro)
				ELSE
					THIS.AddLog("FtpGetFile OK")
				ENDIF
				= InternetCloseHandle (m.lnS)
			ELSE
				lcErro = THIS.WinApiErrMsg()
				THIS.AddLog("!InternetConnect = "+m.lcErro)
			ENDIF
			= InternetCloseHandle (m.lnH)
		ELSE
			lcErro = THIS.WinApiErrMsg()
			THIS.AddLog("!InternetOpen = "+m.lcErro)
		ENDIF
		WAIT CLEAR
		RETURN m.lnResult # 0
	ENDFUNC

	FUNCTION FTPUpload(lcArquivo)
		LOCAL lnFiles, laFiles(1), lnI, lnS, lnH
		m.lnFiles = ADIR(m.laFiles, m.lcArquivo)
		IF m.lnFiles = 0
			THIS.AddLog("!Erro upload FTP: sem arquivos")
			RETURN .F.
		ENDIF
		IF !THIS.FTPConnect(THIS.FTPHost, THIS.FTPUser, THIS.FTPPass)
			RETURN .F.
		ENDIF

		FOR m.lnI=1 TO m.lnFiles
			lcSource = ADDBS(JUSTPATH(m.lcArquivo)) + LOWER(m.laFiles[m.lnI, 1])
			lcTarget = THIS.ServerPath + LOWER(m.laFiles[m.lnI, 1])
			THIS.local2ftp (THIS.FTPSession, lcSource, lcTarget)
		ENDFOR
		= InternetCloseHandle(THIS.FTPSession)
		= InternetCloseHandle(THIS.FTPOpen)
	ENDFUNC

	FUNCTION DELETE
		LPARAMETERS lcArquivo
		LOCAL sAgent, sProxyName, sProxyPass, lFlags
		sAgent = "vfp6"
		sProxyName = CHR(0)
		sProxyBypass = CHR(0)
		lFlags = 0
		THIS.hOpen = InternetOpen (sAgent, INTERNET_OPEN_TYPE_DIRECT, sProxyName, sProxyBypass, lFlags)
		IF THIS.hOpen = 0
			RETURN
		ENDIF
		THIS.hFtpSession = InternetConnect (THIS.hOpen, THIS.HOST, INTERNET_INVALID_PORT_NUMBER, THIS.USER, THIS.PASS, INTERNET_SERVICE_FTP, 0, 0)
		IF THIS.hFtpSession = 0
			= InternetCloseHandle (hOpen)
			RETURN
		ENDIF
		lpszRemoteFile = THIS.ServerPath+vcFilename
		lpszNewFile    = vcLocalDir+lcArquivo
		fFailIfExists  = 0
		dwContext      = 0
		lnResult = FtpDeleteFile (hFtpSession, lpszRemoteFile)
		= InternetCloseHandle (hFtpSession)
		= InternetCloseHandle (hOpen)
		RETURN m.lnResult # 0
	ENDFUNC

	FUNCTION FTPConnect(strHost, strUser, strPwd)
		WAIT WINDOW NOWAIT "Conectando ao FTP "+strHost
		THIS.AddLog("FTP.Connect "+m.strHost)
		THIS.FTPOpen = InternetOpen ("vfp", 1, 0, 0, 0)
		IF THIS.FTPOpen = 0
			THIS.AddLog("!FTP.Connect => Sem acesso a WinInet.dll")
			THIS.MSG("Sem acceso a WinInet.Dll",.T.)
			RETURN .F.
		ENDIF
		THIS.FTPSession = InternetConnect (THIS.FTPOpen , strHost, 0, strUser, strPwd, 1,0,0)
		IF THIS.FTPSession = 0
			LOCAL lcErro
			m.lcErro = THIS.WinApiErrMsg()
			THIS.AddLog("!FTP.Connect => "+m.lcErro)
			= InternetCloseHandle (THIS.FTPOpen)
			THIS.FTPOpen = 0
			THIS.MSG("Erro de conex�o: "+m.lcErro,.T.)
			RETURN .F.
		ENDIF
		WAIT CLEAR
		THIS.AddLog("FTP.Connect OK")
		RETURN .T.

	FUNCTION local2ftp (hConnect, lcSource, lcTarget)
		LOCAL laF(1), lnTamanho
		IF ADIR(m.laF,m.lcSource)=0
			RETURN -1
		ENDIF
		m.lnTamanho = m.laF(1,2)
		THIS.AddLog("FTP.Upload "+m.lcSource)
		WAIT WINDOW NOWAIT "Enviando arquivo "+m.lcSource+" para FTP "+THIS.HOST
		LOCAL hSource
		hSource = FOPEN (lcSource)
		IF (hSource = -1)
			RETURN -1
		ENDIF
		LOCAL lcErro
		hTarget = FtpOpenFile(hConnect, lcTarget, GENERIC_WRITE, 2, 0)
		IF hTarget = 0
			m.lcErro = WinApiErrMsg()
			AddLog("!FTP.Upload = ERRO "+m.lcErro)
			= FCLOSE (hSource)
			WAIT CLEAR
			RETURN -2
		ENDIF
		lnBytesWritten = 0
		lnChunkSize =  128
		DO WHILE !FEOF(hSource)
			lcBuffer = FREAD (hSource, lnChunkSize)
			lnLength = LEN(lcBuffer)
			IF lnLength<>0
				WAIT WINDOW NOWAIT "Enviando arquivo "+m.lcSource+" para FTP "+THIS.FTPHost+CHR(13)+;
					TRANSFORM(100*m.lnBytesWritten/m.lnTamanho,"999.9%")
				IF InternetWriteFile (hTarget, @lcBuffer, lnLength, @lnLength) = 1
					lnBytesWritten = lnBytesWritten + lnLength
				ELSE
					EXIT
				ENDIF
			ELSE
				EXIT
			ENDIF
		ENDDO
		= InternetCloseHandle (hTarget)
		= FCLOSE (hSource)
		WAIT CLEAR
		THIS.AddLog("FTP.Upload "+lcSource+" OK")
		RETURN lnBytesWritten

	FUNCTION FTP2Local (hConnect, lcSource, lcTarget)
		LOCAL laF(1), lnTamanho
		IF ADIR(m.laF,m.lcSource)=0
			RETURN -1
		ENDIF
		m.lnTamanho = m.laF(1,2)
		THIS.AddLog("FTP.Download "+m.lcSource)
		WAIT WINDOW NOWAIT "Recebendo arquivo "+m.lcSource+" para FTP "+THIS.FTPHost
		LOCAL hSource
		hSource = FOPEN (lcSource)
		IF (hSource = -1)
			RETURN -1
		ENDIF
		LOCAL lcErro
		hTarget = FtpOpenFile(hConnect, lcTarget, GENERIC_READ, 2, 0)
		IF hTarget = 0
			m.lcErro = WinApiErrMsg()
			AddLog("!FTP.Upload = ERRO "+m.lcErro)
			= FCLOSE (hSource)
			WAIT CLEAR
			RETURN -2
		ENDIF
		lnBytesWritten = 0
		lnChunkSize =  128
		DO WHILE .T.
			IF InternetReadFile(hTarget, @lcBuffer,
				DO WHILE !FEOF(hSource)
					lcBuffer = FREAD (hSource, lnChunkSize)
					lnLength = LEN(lcBuffer)
					IF lnLength<>0
						WAIT WINDOW NOWAIT "Enviando arquivo "+m.lcSource+" para FTP "+THIS.FTPHost+CHR(13)+;
							TRANSFORM(100*m.lnBytesWritten/m.lnTamanho,"999.9%")
						IF InternetWriteFile (hTarget, @lcBuffer, lnLength, @lnLength) = 1
							lnBytesWritten = lnBytesWritten + lnLength
						ELSE
							EXIT
						ENDIF
					ELSE
						EXIT
					ENDIF
				ENDDO
				= InternetCloseHandle (hTarget)
				= FCLOSE (hSource)
				WAIT CLEAR
				THIS.AddLog("FTP.Upload "+lcSource+" OK")
				RETURN lnBytesWritten

ENDDEFINE
