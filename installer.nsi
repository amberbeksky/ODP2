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

Icon "icon.ico"
UninstallIcon "icon.ico"

Section "Install"
  SetOutPath "$INSTDIR"
  
  ; Основное приложение
  File "dist\app.exe"
  File "icon.ico"
  
  ; Опциональные файлы (создаем если нужно)
  FileOpen $0 "$INSTDIR\readme.txt" w
  FileWrite $0 "Отделение дневного пребывания - Полустационарное обслуживание$\r$\nВерсия ${VERSION}$\r$\n$\r$\nДля работы с Google Sheets необходим файл credentials.json"
  FileClose $0
  
  FileOpen $1 "$INSTDIR\license.txt" w
  FileWrite $1 "Лицензионное соглашение$\r$\n$\r$\nПрограмма предоставляется как есть."
  FileClose $1
  
  ; Создаем пустой credentials.json если не существует
  FileOpen $2 "$INSTDIR\credentials.json" w
  FileWrite $2 "{}"
  FileClose $2
  
  ; Создаем папки
  CreateDirectory "$INSTDIR\logs"
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
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "HelpLink" "https://github.com/your-repo"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "URLInfoAbout" "https://github.com/your-repo"
  
  ; Фиксированный размер (упрощенно для GitHub Actions)
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APPNAME}" "EstimatedSize" "0x00010000"
  
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
