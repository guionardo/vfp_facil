--------------------------------------------------
** FTP_DownLoad
**
** Download files from ftp server
**

PARAMETERS lcHost, lcUser, lcPwd, lcRemoteFile, lcNewFile, lnXFerType

*...........................................................................
......
*: Usage: DO ftpget WITH ;
*: 'ftp.host', 'name', 'password', 'source.file', 'target.file'[, 1
| 2]
*:
*: Where: lcHost = Host computer IP address or name
*: lcUser = user name - anonymous may be used
*: lcPwd = password
*: lcRemoteFile = source file name
*: lcNewFile = target file name
*: lnXFerType = 1 (default) for ascii, 2 for binary
*...........................................................................
......

*...set up API calls

DECLARE INTEGER InternetOpen IN wininet;
	STRING sAgent, INTEGER lAccessType, STRING sProxyName,;
	STRING sProxyBypass, STRING lFlags

DECLARE INTEGER InternetCloseHandle IN wininet INTEGER hInet

DECLARE INTEGER InternetConnect IN wininet.DLL;
	INTEGER hInternetSession,;
	STRING lcHost,;
	INTEGER nServerPort,;
	STRING lcUser,;
	STRING lcPassword,;
	INTEGER lService,;
	INTEGER lFlags,;
	INTEGER lContext

DECLARE INTEGER FtpGetFile IN wininet;
	INTEGER hftpSession, ;
	STRING lcRemoteFile,;
	STRING lcNewFile, ;
	INTEGER fFailIfExists,;
	INTEGER dwFlagsAndAttributes,;
	INTEGER dwFlags, ;
	INTEGER dwContext

lcHost = ALLTRIM(lcHost)
lcUser = ALLTRIM(lcUser)
lcPwd = ALLTRIM(lcPwd)
lcRemoteFile = ALLTRIM(lcRemoteFile)
lcNewFile = ALLTRIM(lcNewFile)

sAgent = "vfp"

sProxyName = CHR(0) &&... no proxy
sProxyBypass = CHR(0) &&... nothing to bypass
lFlags = 0 &&... no flags used

*... initialize access to Inet functions
hOpen = InternetOpen (sAgent, 1,;
	sProxyName, sProxyBypass, lFlags)

IF hOpen = 0
	WAIT WINDOW "Unable to get access to WinInet.Dll" TIMEOUT 2
	RETURN
ENDIF

*... The first '0' says use the default port, usually 21.
hftpSession = InternetConnect (hOpen, lcHost,;
	0, lcUser, lcPwd, 1, 0, 0) &&... 1 = ftp protocol

IF hftpSession = 0
*... close access to Inet functions and exit
	= InternetCloseHandle (hOpen)
	WAIT WINDOW "Unable to connect to " + lcHost + '.' TIMEOUT 2
	RETURN
ELSE
	WAIT WINDOW "Connected to " + lcHost + " as: [" + lcUser + "]" TIMEOUT 1
ENDIF

*... 0 to automatically overwrite file
*... 1 to fail if file already exists
fFailIfExists = 0
dwContext = 0 &&... used for callback

WAIT WINDOW 'Transferring ' + lcRemoteFile + ' to ' + lcNewFile + '...'
NOWAIT
lnResult = FtpGetFile (hftpSession, lcRemoteFile, lcNewFile,;
	fFailIfExists, 128, lnXFerType,;
	dwContext)

*... 128 = #define FILE_ATTRIBUTE_NORMAL 0x00000080
*... See CreateFile for other attributes

* close handles
= InternetCloseHandle (hftpSession)
= InternetCloseHandle (hOpen)

IF lnResult = 1
*... successful download, do what you want here
	WAIT WINDOW 'Completed.' NOWAIT
ELSE
	WAIT WINDOW "Unable to download selected file" TIMEOUT 2
ENDIF

RETURN




--------------------------------------------------
**
** Upload files to ftp server
**
#DEFINE GENERIC_READ 2147483648 && &H80000000
#DEFINE GENERIC_WRITE 1073741824 && &H40000000

LOCAL m.ftpServer, m.ftpUserName, m.ftpUserPass

PUBLIC hOpen, hftpSession
DECLARE INTEGER InternetOpen IN wininet.DLL;
	STRING sAgent,;
	INTEGER lAccessType,;
	STRING sProxyName,;
	STRING sProxyBypass,;
	STRING lFlags

DECLARE INTEGER InternetCloseHandle IN wininet.DLL;
	INTEGER hInet

DECLARE INTEGER InternetConnect IN wininet.DLL;
	INTEGER hInternetSession,;
	STRING sServerName,;
	INTEGER nServerPort,;
	STRING sUsername,;
	STRING sPassword,;
	INTEGER lService,;
	INTEGER lFlags,;
	INTEGER lContext

DECLARE INTEGER FtpOpenFile IN wininet.DLL;
	INTEGER hFtpSession,;
	STRING sFileName,;
	INTEGER lAccess,;
	INTEGER lFlags,;
	INTEGER lContext

DECLARE INTEGER InternetWriteFile IN wininet.DLL;
	INTEGER hFile,;
	STRING @ sBuffer,;
	INTEGER lNumBytesToWrite,;
	INTEGER @ dwNumberOfBytesWritten

m.ftpServer="klingon"
m.ftpServer="172.10.1.3"
m.ftpUserName="e2userdp"
m.ftpUserPass="e2user"

IF connect2ftp (m.ftpServer, m.ftpUserName, m.ftpUserPass)
	lcSourcePath = "C:\" && local folder
	lcTargetPath = "/home/e2userdp/" && remote folder (ftp server)

	lnFiles = ADIR (arr, lcSourcePath + "lolo.txt")

	FOR lnCnt=1 TO lnFiles
		lcSource = lcSourcePath + LOWER (arr [lnCnt, 1])
		lcTarget = lcTargetPath + LOWER (arr [lnCnt, 1])
		? lcSource + " -> " + lcTarget
		?? local2ftp (hftpSession, lcSource, lcTarget)
	ENDFOR

	= InternetCloseHandle (hftpSession)
	= InternetCloseHandle (hOpen)
ENDIF

FUNCTION connect2ftp (strHost, strUser, strPwd)
** Open the access
	hOpen = InternetOpen ("vfp", 1, 0, 0, 0)

	IF hOpen = 0
		? "No access to WinInet.Dll"
		RETURN .F.
	ENDIF

** Connect to FTP.
	hftpSession = InternetConnect (hOpen, strHost, 0, strUser, strPwd, 1, 0,
	0)

	IF hftpSession = 0
** Close
		= InternetCloseHandle (hOpen)
		? "FTP " + strHost + " not ready"
		RETURN .F.
	ELSE
		? "Connected to " + strHost + " as: [" + strUser + ", *****]"
	ENDIF
	RETURN .T.


**--------------------------------------------
** Copying files
**--------------------------------------------
FUNCTION local2ftp (hConnect, lcSource, lcTarget)
** Upload local file to ftp server
	hSource = FOPEN (lcSource)
	IF (hSource = -1)
		RETURN -1
	ENDIF

** New file in ftp server
	hTarget = FtpOpenFile(hConnect, lcTarget, GENERIC_WRITE, 2, 0)
	IF hTarget = 0
		= FCLOSE (hSource)
		RETURN -2
	ENDIF
	lnBytesWritten = 0
	lnChunkSize = 512 && 128, 512
	DO WHILE NOT FEOF(hSource)
		lcBuffer = FREAD (hSource, lnChunkSize)
		lnLength = LEN(lcBuffer)
		IF lnLength > 0
			IF InternetWriteFile (hTarget, @lcBuffer, lnLength, @lnLength) =
				1
				lnBytesWritten = lnBytesWritten + lnLength
				? lnBytesWritten
** Show Progress
			ELSE
				EXIT
			ENDIF
		ELSE
			EXIT
		ENDIF
	ENDDO

	= InternetCloseHandle (hTarget)
	= FCLOSE (hSource)

	RETURN lnBytesWritten


DECLARE INTEGER FtpRenameFile IN wininet;
    INTEGER hConnect,;
    STRING  lpszExisting,;
    STRING  lpszNew