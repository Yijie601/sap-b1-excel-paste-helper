Unicode True
SetCompressor /SOLID lzma

!include "MUI2.nsh"

!ifndef APP_VERSION
  !define APP_VERSION "0.1.0-beta.14"
!endif

!define APP_NAME "SAP B1 Excel Helper"
!define APP_EXE "SapB1ExcelHelper.exe"
!define APP_PUBLISHER "Yijie601"
!define APP_REGISTRY_KEY "Software\Microsoft\Windows\CurrentVersion\Uninstall\SapB1ExcelHelper"
!define SOURCE_ROOT ".."

Name "${APP_NAME}"
OutFile "${SOURCE_ROOT}\artifacts\installer\SapB1ExcelHelper-Setup-${APP_VERSION}-win-x64.exe"
InstallDir "$LOCALAPPDATA\Programs\SAP B1 Excel Helper"
InstallDirRegKey HKCU "${APP_REGISTRY_KEY}" "InstallLocation"
RequestExecutionLevel user
BrandingText "SAP B1 Excel Helper"

VIProductVersion "0.1.0.14"
VIAddVersionKey "ProductName" "${APP_NAME}"
VIAddVersionKey "ProductVersion" "${APP_VERSION}"
VIAddVersionKey "CompanyName" "${APP_PUBLISHER}"
VIAddVersionKey "LegalCopyright" "Copyright 2026 ${APP_PUBLISHER}"
VIAddVersionKey "FileDescription" "${APP_NAME} installer"
VIAddVersionKey "FileVersion" "${APP_VERSION}"

!define MUI_ABORTWARNING
!define MUI_FINISHPAGE_RUN "$INSTDIR\${APP_EXE}"
!define MUI_FINISHPAGE_RUN_TEXT "Launch ${APP_NAME}"

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_COMPONENTS
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM
!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "English"

Section "SAP B1 Excel Helper (required)" SEC_MAIN
  SectionIn RO
  SetShellVarContext current

  ; Close an older tray instance before replacing its executable during an update.
  nsExec::ExecToLog '"$SYSDIR\taskkill.exe" /IM ${APP_EXE} /F'

  SetOutPath "$INSTDIR"
  File /r /x "*.pdb" "${SOURCE_ROOT}\artifacts\publish\*"

  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr HKCU "${APP_REGISTRY_KEY}" "DisplayName" "${APP_NAME}"
  WriteRegStr HKCU "${APP_REGISTRY_KEY}" "DisplayVersion" "${APP_VERSION}"
  WriteRegStr HKCU "${APP_REGISTRY_KEY}" "Publisher" "${APP_PUBLISHER}"
  WriteRegStr HKCU "${APP_REGISTRY_KEY}" "InstallLocation" "$INSTDIR"
  WriteRegStr HKCU "${APP_REGISTRY_KEY}" "DisplayIcon" "$INSTDIR\${APP_EXE}"
  WriteRegStr HKCU "${APP_REGISTRY_KEY}" "UninstallString" '"$INSTDIR\Uninstall.exe"'
  WriteRegStr HKCU "${APP_REGISTRY_KEY}" "QuietUninstallString" '"$INSTDIR\Uninstall.exe" /S'
  WriteRegDWORD HKCU "${APP_REGISTRY_KEY}" "NoModify" 1
  WriteRegDWORD HKCU "${APP_REGISTRY_KEY}" "NoRepair" 1
  WriteRegDWORD HKCU "${APP_REGISTRY_KEY}" "EstimatedSize" 74000
SectionEnd

Section "Start Menu shortcut" SEC_START_MENU
  SetShellVarContext current
  CreateDirectory "$SMPROGRAMS\SAP B1 Excel Helper"
  CreateShortcut "$SMPROGRAMS\SAP B1 Excel Helper\SAP B1 Excel Helper.lnk" "$INSTDIR\${APP_EXE}"
  CreateShortcut "$SMPROGRAMS\SAP B1 Excel Helper\Uninstall.lnk" "$INSTDIR\Uninstall.exe"
SectionEnd

Section /o "Desktop shortcut" SEC_DESKTOP
  SetShellVarContext current
  CreateShortcut "$DESKTOP\SAP B1 Excel Helper.lnk" "$INSTDIR\${APP_EXE}"
SectionEnd

Section /o "Start with Windows" SEC_STARTUP
  SetShellVarContext current
  CreateShortcut "$SMSTARTUP\SAP B1 Excel Helper.lnk" "$INSTDIR\${APP_EXE}"
SectionEnd

LangString DESC_SEC_MAIN ${LANG_ENGLISH} "Installs the application for the current Windows user."
LangString DESC_SEC_START_MENU ${LANG_ENGLISH} "Creates Start Menu shortcuts."
LangString DESC_SEC_DESKTOP ${LANG_ENGLISH} "Creates a desktop shortcut."
LangString DESC_SEC_STARTUP ${LANG_ENGLISH} "Starts the helper automatically when you sign in."

!insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_MAIN} $(DESC_SEC_MAIN)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_START_MENU} $(DESC_SEC_START_MENU)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_DESKTOP} $(DESC_SEC_DESKTOP)
  !insertmacro MUI_DESCRIPTION_TEXT ${SEC_STARTUP} $(DESC_SEC_STARTUP)
!insertmacro MUI_FUNCTION_DESCRIPTION_END

Section "Uninstall"
  SetShellVarContext current
  nsExec::ExecToLog '"$SYSDIR\taskkill.exe" /IM ${APP_EXE} /F'

  Delete "$DESKTOP\SAP B1 Excel Helper.lnk"
  Delete "$SMSTARTUP\SAP B1 Excel Helper.lnk"
  RMDir /r "$SMPROGRAMS\SAP B1 Excel Helper"
  RMDir /r "$INSTDIR"
  DeleteRegKey HKCU "${APP_REGISTRY_KEY}"

  ; Mappings, calibration, and logs under %LOCALAPPDATA%\SapB1ExcelHelper are preserved.
SectionEnd
