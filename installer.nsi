Unicode true
ManifestDPIAware true

!include "MUI2.nsh"

!define APPNAME "Отделение дневного пребывания"
!define COMPANY "Полустационарное обслуживание"
!define VERSION "2.0"

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
  
  ; Основные файлы приложения
  File "dist\app.exe"
  File "credentials.json"
  File "icon.ico"
  
  ; Создаем readme файл
  FileOpen $0 "$INSTDIR\readme.txt" w
  FileWrite $0 "Отделение дневного пребывания - Полустационарное обслуживание$\r$\nВерсия ${VERSION}$\r$\n$\r$\nДля работы с Google Sheets используется файл credentials.json"
  FileClose $0
  
  ; Создаем папки для логов и бэкапов
  CreateDirectory "$INSTDIR\logs"
  CreateDirectory "$INSTDIR\backups"

  ; Ярлык на рабочий стол
  CreateShortcut "$DESKTOP\${APPNAME}.lnk" "$INSTDIR\app.exe" "" "$INSTDIR\icon.ico"

  ; Папка в меню Пуск
  CreateDirectory "$SMPROGRAMS\${APPNAME}"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk" "$INSTDIR\app.exe" "" "$INSTDIR\icon.ico"
  CreateShortcut "$SMPROGRAMS\${APPNAME}\Прочитать меня.lnk" "$INSTDIR\readme.txt"

  ; Uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"

  ; Запись в реестр для удаления через "Программы и компоненты"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayName" "${APPNAME}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "UninstallString" "$\"$INSTDIR\Uninstall.exe$\""
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "Publisher" "${COMPANY}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayVersion" "${VERSION}"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "DisplayIcon" "$\"$INSTDIR\icon.ico$\""
  
  ; Фиксированный размер (примерно 1 МБ)
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "EstimatedSize" "0x00010000"
  
SectionEnd

Section "Uninstall"
  ; Удаляем файлы приложения
  Delete "$INSTDIR\app.exe"
  Delete "$INSTDIR\icon.ico"
  Delete "$INSTDIR\readme.txt"
  Delete "$INSTDIR\credentials.json"
  Delete "$INSTDIR\Uninstall.exe"
  
  ; Удаляем ярлыки
  Delete "$DESKTOP\${APPNAME}.lnk"
  Delete "$SMPROGRAMS\${APPNAME}\${APPNAME}.lnk"
  Delete "$SMPROGRAMS\${APPNAME}\Прочитать меня.lnk"
  
  ; Удаляем папку меню Пуск
  RMDir "$SMPROGRAMS\${APPNAME}"
  
  ; Удаляем папки приложения (если пустые)
  RMDir "$INSTDIR\logs"
  RMDir "$INSTDIR\backups"
  RMDir "$INSTDIR"

  ; Чистим реестр
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}"
  
  ; Сообщение о сохранении данных пользователя
  MessageBox MB_OK|MB_ICONINFORMATION "Программа удалена.$\nВаши данные (база клиентов, настройки) сохранены в папке:$\n$APPDATA\MyApp"
SectionEnd
