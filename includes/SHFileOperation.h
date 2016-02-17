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
