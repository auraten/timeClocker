!define PRODUCT_NAME "Time Clocker"
!define PRODUCT_VERSION "2.0.2"
!define PY_VERSION "3.6.1"
!define PY_MAJOR_VERSION "3.6"
!define BITNESS "64"
!define ARCH_TAG ".amd64"
!define INSTALLER_NAME "Time_Clocker_2.0.2.exe"
!define PRODUCT_ICON "glossyorb.ico"
 
SetCompressor lzma

RequestExecutionLevel admin

; Modern UI installer stuff 
!include "MUI2.nsh"
!define MUI_ABORTWARNING
!define MUI_ICON "${NSISDIR}\Contrib\Graphics\Icons\modern-install.ico"

; UI pages
!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH
!insertmacro MUI_LANGUAGE "English"

Name "${PRODUCT_NAME} ${PRODUCT_VERSION}"
OutFile "${INSTALLER_NAME}"
InstallDir "$PROGRAMFILES${BITNESS}\${PRODUCT_NAME}"
ShowInstDetails show

Section -SETTINGS
  SetOutPath "$INSTDIR"
  SetOverwrite ifnewer
SectionEnd


Section "!${PRODUCT_NAME}" sec_app
  SetRegView 64
  SectionIn RO
  SetShellVarContext all
  File ${PRODUCT_ICON}
  SetOutPath "$INSTDIR\pkgs"
  File /r "pkgs\*.*"
  SetOutPath "$INSTDIR"

      ; Install files
    SetOutPath "$INSTDIR"
      File "glossyorb.ico"
      File "Time_Clocker.launch.py"
      File "chromedriver.exe"
      File "form.html"
  
  ; Install directories
    SetOutPath "$INSTDIR\Python"
    File /r "Python\*.*"


    ; Install MSVCRT if it's not already on the system
    IfFileExists "$SYSDIR\ucrtbase.dll" skip_msvcrt
    SetOutPath $INSTDIR\Python
    File msvcrt\api-ms-win-core-console-l1-1-0.dll
    File msvcrt\api-ms-win-core-datetime-l1-1-0.dll
    File msvcrt\api-ms-win-core-debug-l1-1-0.dll
    File msvcrt\api-ms-win-core-errorhandling-l1-1-0.dll
    File msvcrt\api-ms-win-core-file-l1-1-0.dll
    File msvcrt\api-ms-win-core-file-l1-2-0.dll
    File msvcrt\api-ms-win-core-file-l2-1-0.dll
    File msvcrt\api-ms-win-core-handle-l1-1-0.dll
    File msvcrt\api-ms-win-core-heap-l1-1-0.dll
    File msvcrt\api-ms-win-core-interlocked-l1-1-0.dll
    File msvcrt\api-ms-win-core-libraryloader-l1-1-0.dll
    File msvcrt\api-ms-win-core-localization-l1-2-0.dll
    File msvcrt\api-ms-win-core-memory-l1-1-0.dll
    File msvcrt\api-ms-win-core-namedpipe-l1-1-0.dll
    File msvcrt\api-ms-win-core-processenvironment-l1-1-0.dll
    File msvcrt\api-ms-win-core-processthreads-l1-1-0.dll
    File msvcrt\api-ms-win-core-processthreads-l1-1-1.dll
    File msvcrt\api-ms-win-core-profile-l1-1-0.dll
    File msvcrt\api-ms-win-core-rtlsupport-l1-1-0.dll
    File msvcrt\api-ms-win-core-string-l1-1-0.dll
    File msvcrt\api-ms-win-core-synch-l1-1-0.dll
    File msvcrt\api-ms-win-core-synch-l1-2-0.dll
    File msvcrt\api-ms-win-core-sysinfo-l1-1-0.dll
    File msvcrt\api-ms-win-core-timezone-l1-1-0.dll
    File msvcrt\api-ms-win-core-util-l1-1-0.dll
    File msvcrt\api-ms-win-crt-conio-l1-1-0.dll
    File msvcrt\api-ms-win-crt-convert-l1-1-0.dll
    File msvcrt\api-ms-win-crt-environment-l1-1-0.dll
    File msvcrt\api-ms-win-crt-filesystem-l1-1-0.dll
    File msvcrt\api-ms-win-crt-heap-l1-1-0.dll
    File msvcrt\api-ms-win-crt-locale-l1-1-0.dll
    File msvcrt\api-ms-win-crt-math-l1-1-0.dll
    File msvcrt\api-ms-win-crt-multibyte-l1-1-0.dll
    File msvcrt\api-ms-win-crt-private-l1-1-0.dll
    File msvcrt\api-ms-win-crt-process-l1-1-0.dll
    File msvcrt\api-ms-win-crt-runtime-l1-1-0.dll
    File msvcrt\api-ms-win-crt-stdio-l1-1-0.dll
    File msvcrt\api-ms-win-crt-string-l1-1-0.dll
    File msvcrt\api-ms-win-crt-time-l1-1-0.dll
    File msvcrt\api-ms-win-crt-utility-l1-1-0.dll
    File msvcrt\ucrtbase.dll
    skip_msvcrt:

  
  ; Install shortcuts
  ; The output path becomes the working directory for shortcuts
  SetOutPath "%HOMEDRIVE%\%HOMEPATH%"
    CreateShortCut "$SMPROGRAMS\Time Clocker.lnk" "$INSTDIR\Python\python.exe" \
      '"$INSTDIR\Time_Clocker.launch.py"' "$INSTDIR\glossyorb.ico"
  SetOutPath "$INSTDIR"

  
  ; Byte-compile Python files.
  DetailPrint "Byte-compiling Python modules..."
  nsExec::ExecToLog '"$INSTDIR\Python\python" -m compileall -q "$INSTDIR\pkgs"'
  WriteUninstaller $INSTDIR\uninstall.exe
  ; Add ourselves to Add/remove programs
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "DisplayName" "${PRODUCT_NAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "DisplayIcon" "$INSTDIR\${PRODUCT_ICON}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "DisplayVersion" "${PRODUCT_VERSION}"
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}" \
                   "NoRepair" 1

  ; Check if we need to reboot
  IfRebootFlag 0 noreboot
    MessageBox MB_YESNO "A reboot is required to finish the installation. Do you wish to reboot now?" \
                /SD IDNO IDNO noreboot
      Reboot
  noreboot:
SectionEnd

Section "Uninstall"
  SetRegView 64
  SetShellVarContext all
  Delete $INSTDIR\uninstall.exe
  Delete "$INSTDIR\${PRODUCT_ICON}"
  RMDir /r "$INSTDIR\pkgs"

  ; Remove ourselves from %PATH%

  ; Uninstall files
    Delete "$INSTDIR\glossyorb.ico"
    Delete "$INSTDIR\Time_Clocker.launch.py"
    Delete "$INSTDIR\chromedriver.exe"
    Delete "$INSTDIR\form.html"
  ; Uninstall directories
    RMDir /r "$INSTDIR\Python"

  ; Uninstall shortcuts
      Delete "$SMPROGRAMS\Time Clocker.lnk"
  RMDir $INSTDIR
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${PRODUCT_NAME}"
SectionEnd


; Functions

Function .onMouseOverSection
    ; Find which section the mouse is over, and set the corresponding description.
    FindWindow $R0 "#32770" "" $HWNDPARENT
    GetDlgItem $R0 $R0 1043 ; description item (must be added to the UI)

    StrCmp $0 ${sec_app} "" +2
      SendMessage $R0 ${WM_SETTEXT} 0 "STR:${PRODUCT_NAME}"
    
FunctionEnd