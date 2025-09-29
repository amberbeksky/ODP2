!include "MUI2.nsh"

!define APPNAME "Отделение дневного пребывания"
!define COMPANY "Полустационарное обслуживание"
!define DESCRIPTION "Программа для управления клиентами Отделения дневного пребывания"
!define VERSION "1.0"

Name "${APPNAME}"
OutFile "ODP-Installer.exe"
InstallDir "$PROGRAMFILES\${APPNAME}"

RequestExecutionLevel admin

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "Russian"

Icon "icon.ico"
UninstallIcon "icon.ico"

Section "Install"
  SetOutPath "$INSTDIR"
  File "dist\app.exe"
  File "credentials.json"
  File "icon.ico"

  ; Ярлык на рабочий стол
  CreateShortcut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\app.exe" "" "$INSTDIR\icon.ico"

  ; Папка в меню Пуск
  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\app.exe" "" "$INSTDIR\icon.ico"

  ; Uninstall
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; Запись в реестр для удаления через "Программы и компоненты"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "${COMPANY}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" "$INSTDIR\icon.ico"
SectionEnd

Section "Uninstall"
  Delete "$INSTDIR\app.exe"
  Delete "$INSTDIR\credentials.json"
  Delete "$INSTDIR\icon.ico"
  Delete "$DESKTOP\${APPNAME}.lnk"

  ; Удаление меню Пуск
  Delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
  RMDir "$SMPROGRAMS\${APPNAME}"

  ; Удаление файлов
  Delete "$INSTDIR\Uninstall.exe"
  RMDir "$INSTDIR"

  ; Чистим реестр
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
SectionEnd
