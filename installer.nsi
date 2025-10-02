Unicode true
ManifestDPIAware true

!include "MUI2.nsh"

!define APPNAME "Отделение дневного пребывания"
!define COMPANY "Полустационарное обслуживание"
!define DESCRIPTION "Программа для управления клиентами Отделения дневного пребывания"
!define VERSION "2.0"

Name "${APPNAME}"
OutFile "ODP-Installer-v${VERSION}.exe"
InstallDir "$PROGRAMFILES\${APPNAME}"

RequestExecutionLevel admin
ShowInstDetails show
ShowUnInstDetails show

!insertmacro MUI_PAGE_WELCOME
!insertmacro MUI_PAGE_DIRECTORY
!insertmacro MUI_PAGE_INSTFILES
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_INSTFILES

!insertmacro MUI_LANGUAGE "Russian"

!insertmacro MUI_RESERVEFILE_LANGDLL

Icon "icon.ico"
UninstallIcon "icon.ico"

Section "Install"
  SetOutPath "$INSTDIR"
  
  ; Основное приложение
  File "dist\app.exe"
  File "credentials.json"
  File "icon.ico"
  File "readme.txt"
  File "license.txt"
  
  ; Создаем папку для логов
  CreateDirectory "$INSTDIR\logs"
  
  ; Создаем папку для бэкапов
  CreateDirectory "$INSTDIR\backups"

  ; Ярлык на рабочий стол
  CreateShortcut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\app.exe" "" "$INSTDIR\icon.ico"

  ; Папка в меню Пуск
  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\app.exe" "" "$INSTDIR\icon.ico"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\Папка данных.lnk" "$APPDATA\MyApp"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\Прочитать меня.lnk" "$INSTDIR\readme.txt"

  ; Uninstall
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; Запись в реестр для удаления через "Программы и компоненты"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "${COMPANY}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" "$\"$INSTDIR\icon.ico$\""
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "HelpLink" "http://www.example.com"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "URLInfoAbout" "http://www.example.com"
  
  ; Записываем размер установки
  ${GetSize} "$INSTDIR" "/S=0K" $0 $1 $2
  IntFmt $0 "0x%08X" $0
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "EstimatedSize" "$0"
  
SectionEnd

Section "Uninstall"
  ; Удаляем файлы приложения
  Delete "$INSTDIR\app.exe"
  Delete "$INSTDIR\credentials.json"
  Delete "$INSTDIR\icon.ico"
  Delete "$INSTDIR\readme.txt"
  Delete "$INSTDIR\license.txt"
  Delete "$INSTDIR\Uninstall.exe"
  
  ; Удаляем ярлыки
  Delete "$DESKTOP\${APPNAME}.lnk"

  ; Удаление меню Пуск
  Delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
  Delete "$SMPROGRAMS\${APPNAME}\Папка данных.lnk"
  Delete "$SMPROGRAMS\${APPNAME}\Прочитать меня.lnk"
  RMDir "$SMPROGRAMS\${APPNAME}"

  ; Удаляем папки (если пустые)
  RMDir "$INSTDIR\logs"
  RMDir "$INSTDIR\backups"
  RMDir "$INSTDIR"

  ; Чистим реестр
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
  
  ; Сообщение о том, что данные пользователя сохранены
  MessageBox MB_OK|MB_ICONINFORMATION "Программа удалена.$\nВаши данные (база клиентов, настройки) сохранены в папке:$\n$APPDATA\MyApp"
SectionEnd

; Функция для получения размера папки
!include "FileFunc.nsh"
!insertmacro GetSize
