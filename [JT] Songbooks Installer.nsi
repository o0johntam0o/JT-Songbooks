!include "FileFunc.nsh"
OutFile "[JT] Songbooks Installer.exe"
Name "[JT] Songbooks"
InstallDir "$PROGRAMFILES\[JT] Songbooks"
Caption "[JT] Songbooks Installer"
Icon "LOGO.ico"

!define VER "2.0.1"
!define /date INST_DATE "%Y%m%d"
!define UNINST_REG "Software\Microsoft\Windows\CurrentVersion\Uninstall\[JT] Songbooks"

Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

Section "Install"
	; Data -----------------
	SetOutPath "$INSTDIR\Languages"
	File "Languages\*.ini"
	SetOutPath "$INSTDIR"
	File "[JT] Songbooks.exe"
	File "license.txt"
	; ----------------------
	
	; Additional files -----
	WriteUninstaller "$INSTDIR\UnInstaller.exe"
	CreateShortCut "$DESKTOP\[JT] Songbooks.lnk" "$INSTDIR\[JT] Songbooks.exe" "" "" "" SW_SHOWNORMAL "" "Songbooks for guitarist"
	CreateShortCut "$SMPROGRAMS\[JT] Songbooks\[JT] Songbooks.lnk" "$INSTDIR\[JT] Songbooks.exe" "" "" "" SW_SHOWNORMAL "" "Songbooks for guitarist"
	CreateShortCut "$SMPROGRAMS\[JT] Songbooks\Help.lnk" "$INSTDIR\Help.txt" "" "" "" SW_SHOWMAXIMIZED "" "[JT] Songbooks Help"
	CreateShortCut "$SMPROGRAMS\[JT] Songbooks\Un-Install.lnk" "$INSTDIR\UnInstaller.exe" "" "" "" SW_SHOWNORMAL "" "Remove [JT] Songbooks"
	; ----------------------
	
	; Registry -------------
	WriteRegStr HKLM "${UNINST_REG}" "DisplayName" "[JT] Songbooks"
	WriteRegStr HKLM "${UNINST_REG}" "DisplayVersion" "${VER}"
	WriteRegStr HKLM "${UNINST_REG}" "DisplayIcon" "$\"$INSTDIR\[JT] Songbooks.exe$\""
	WriteRegStr HKLM "${UNINST_REG}" "Publisher" "o0johntam0o"
	WriteRegStr HKLM "${UNINST_REG}" "RegCompany" "[JT] Knowledge"
	WriteRegStr HKLM "${UNINST_REG}" "Contact" "o0johntam0o@gmail.com"
	WriteRegStr HKLM "${UNINST_REG}" "UrlUpdateInfo" "https://github.com/o0johntam0o/jt-songbooks"
	WriteRegStr HKLM "${UNINST_REG}" "Comments" "Songbooks for guitarist"
	WriteRegStr HKLM "${UNINST_REG}" "Readme" "$\"$INSTDIR\Help.txt$\""
	WriteRegStr HKLM "${UNINST_REG}" "InstallDate" "${INST_DATE}"
	WriteRegStr HKLM "${UNINST_REG}" "InstallLocation" "$\"$INSTDIR$\""
	${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
	IntFmt $0 "0x%08X" $0
	WriteRegDWORD HKLM "${UNINST_REG}" "EstimatedSize" "$0"
	WriteRegDWORD HKLM "${UNINST_REG}" "NoModify" "1"
	WriteRegDWORD HKLM "${UNINST_REG}" "NoRepair" "1"
	WriteRegStr HKLM "${UNINST_REG}" "UninstallString" "$\"$INSTDIR\UnInstaller.exe$\""
	WriteRegStr HKLM "${UNINST_REG}" "QuietUninstallString" "$\"$INSTDIR\UnInstaller.exe$\" /S"
	; ----------------------
SectionEnd

Section "Uninstall"
	MessageBox MB_YESNO|MB_DEFBUTTON1 "WARNING! All the songs that you've saved before will also be removed too. Do you want to continue?" IDYES +2
	Quit
	; Additional files -----
	SetShellVarContext current
	RMDir /r "$SMPROGRAMS\[JT] Songbooks"
	Delete "$DESKTOP\[JT] Songbooks.lnk"
	SetShellVarContext all
	RMDir /r "$SMPROGRAMS\[JT] Songbooks"
	Delete "$DESKTOP\[JT] Songbooks.lnk"
	SetShellVarContext current
	; ----------------------
	
	; Data -----------------
	RMDir /r /REBOOTOK "$INSTDIR"
	; ----------------------
	
	; Registry -------------
	DeleteRegKey HKLM "${UNINST_REG}"
	; ----------------------
SectionEnd

Function .onInit
	ReadRegStr $0 HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\[JT] Songbooks" "UninstallString"
	StrLen $1 $0
	IntOp $1 $1 - 2
	StrCpy $0 $0 $1 1
	IfFileExists "$0" +1 +4
	; +1
	MessageBox MB_YESNOCANCEL|MB_DEFBUTTON2 "You seem to have this program already installed. Would you like to remove it first?" IDNO +3 IDCANCEL +2
	ExecWait '"$0"'
	Quit
	; +4
	Return
FunctionEnd