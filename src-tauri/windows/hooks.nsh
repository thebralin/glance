; Своя иконка для каждого типа файлов в Проводнике.
; Tauri регистрирует ProgId-ы из bundle.fileAssociations (Glance.PDF, Glance.Word, ...)
; с иконкой exe; этот хук выполняется после записи реестра и подменяет DefaultIcon.
; Сами .ico лежат в ресурсах (bundle.resources -> $INSTDIR\filetypes\).
; При удалении ключи ProgId-ов стирает сам Tauri вместе с DefaultIcon.

; Tauri (APP_ASSOCIATE) делает свой ProgId значением по умолчанию для расширения.
; Для .txt / .md / .json это отняло бы файлы у Блокнота / редактора пользователя, поэтому
; возвращаем прежнее значение (Tauri сохраняет его в "<ProgId>_backup") и оставляем
; Glance только в «Открыть с помощью».
!macro GLANCE_OPEN_WITH_ONLY EXT PROGID
  ; Если пользователь сам выбрал Glance для этого типа (см. GLANCE_FIX_USERCHOICE, $R1 = 1),
  ; оставляем расширение за Glance, иначе возвращаем прежнего владельца.
  ; Если до установки расширение уже принадлежало Glance (backup == ProgId) — тоже оставляем.
  ReadRegStr $R0 SHCTX "Software\Classes\.${EXT}" "${PROGID}_backup"
  ${If} $R1 == 1
  ${OrIf} $R0 == "${PROGID}"
    WriteRegStr SHCTX "Software\Classes\.${EXT}" "" "${PROGID}"
  ${ElseIf} $R0 == ""
    DeleteRegValue SHCTX "Software\Classes\.${EXT}" ""
  ${Else}
    WriteRegStr SHCTX "Software\Classes\.${EXT}" "" "$R0"
  ${EndIf}
  WriteRegStr SHCTX "Software\Classes\.${EXT}\OpenWithProgids" "${PROGID}" ""
!macroend

; Если раньше Glance выбрали через «Открыть с помощью → Всегда» (до появления ProgId-ов),
; Windows записала в UserChoice "Applications\glance.exe". Тогда Проводник берёт иконку
; прямо из exe — одну на все типы — и наши DefaultIcon игнорируются.
; Удаляем такой UserChoice: расширение и так указывает на наш ProgId, файлы продолжат
; открываться в Glance, но уже с иконкой своего типа. Чужие UserChoice не трогаем.
; На выходе $R1 = 1, если UserChoice указывал на Glance.
!macro GLANCE_FIX_USERCHOICE EXT
  StrCpy $R1 0
  StrCpy $R2 "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.${EXT}"
  ReadRegStr $R0 HKCU "$R2\UserChoice" "ProgId"
  ${If} $R0 == "Applications\${MAINBINARYNAME}.exe"
    DeleteRegKey HKCU "$R2\UserChoice"
    StrCpy $R1 1
  ${EndIf}
  ; Windows 11 24H2+: UserChoiceLatest (ProgId в подключе ProgId или прямо в ключе)
  ReadRegStr $R0 HKCU "$R2\UserChoiceLatest\ProgId" "ProgId"
  ${If} $R0 == ""
    ReadRegStr $R0 HKCU "$R2\UserChoiceLatest" "ProgId"
  ${EndIf}
  ${If} $R0 == "Applications\${MAINBINARYNAME}.exe"
    DeleteRegKey HKCU "$R2\UserChoiceLatest"
    StrCpy $R1 1
  ${EndIf}
  ; Windows 11 (сборки 26xxx) с драйвером UCPD молча запрещает сторонним процессам менять
  ; UserChoice/UserChoiceLatest. Проверяем, что ключ действительно удалён; если нет —
  ; копим список расширений в $R3, чтобы в конце подсказать пользователю ручной шаг.
  ReadRegStr $R0 HKCU "$R2\UserChoice" "ProgId"
  ${If} $R0 != "Applications\${MAINBINARYNAME}.exe"
    ReadRegStr $R0 HKCU "$R2\UserChoiceLatest\ProgId" "ProgId"
    ${If} $R0 == ""
      ReadRegStr $R0 HKCU "$R2\UserChoiceLatest" "ProgId"
    ${EndIf}
  ${EndIf}
  ${If} $R0 == "Applications\${MAINBINARYNAME}.exe"
    ${If} $R3 == ""
      StrCpy $R3 ".${EXT}"
    ${Else}
      StrCpy $R3 "$R3, .${EXT}"
    ${EndIf}
  ${EndIf}
!macroend

!macro NSIS_HOOK_POSTINSTALL
  WriteRegStr SHCTX "Software\Classes\Glance.PDF\DefaultIcon" "" "$INSTDIR\filetypes\pdf.ico"
  WriteRegStr SHCTX "Software\Classes\Glance.Word\DefaultIcon" "" "$INSTDIR\filetypes\word.ico"
  WriteRegStr SHCTX "Software\Classes\Glance.PowerPoint\DefaultIcon" "" "$INSTDIR\filetypes\powerpoint.ico"
  WriteRegStr SHCTX "Software\Classes\Glance.Excel\DefaultIcon" "" "$INSTDIR\filetypes\excel.ico"
  WriteRegStr SHCTX "Software\Classes\Glance.Markdown\DefaultIcon" "" "$INSTDIR\filetypes\markdown.ico"
  WriteRegStr SHCTX "Software\Classes\Glance.Text\DefaultIcon" "" "$INSTDIR\filetypes\text.ico"
  WriteRegStr SHCTX "Software\Classes\Glance.JSON\DefaultIcon" "" "$INSTDIR\filetypes\json.ico"
  StrCpy $R3 ""
  !insertmacro GLANCE_FIX_USERCHOICE "pdf"
  !insertmacro GLANCE_FIX_USERCHOICE "docx"
  !insertmacro GLANCE_FIX_USERCHOICE "docm"
  !insertmacro GLANCE_FIX_USERCHOICE "dotx"
  !insertmacro GLANCE_FIX_USERCHOICE "pptx"
  !insertmacro GLANCE_FIX_USERCHOICE "pptm"
  !insertmacro GLANCE_FIX_USERCHOICE "ppsx"
  !insertmacro GLANCE_FIX_USERCHOICE "potx"
  !insertmacro GLANCE_FIX_USERCHOICE "xlsx"
  !insertmacro GLANCE_FIX_USERCHOICE "xlsm"
  !insertmacro GLANCE_FIX_USERCHOICE "xltx"
  !insertmacro GLANCE_FIX_USERCHOICE "txt"
  !insertmacro GLANCE_OPEN_WITH_ONLY "txt" "Glance.Text"
  !insertmacro GLANCE_FIX_USERCHOICE "md"
  !insertmacro GLANCE_OPEN_WITH_ONLY "md" "Glance.Markdown"
  !insertmacro GLANCE_FIX_USERCHOICE "markdown"
  !insertmacro GLANCE_OPEN_WITH_ONLY "markdown" "Glance.Markdown"
  !insertmacro GLANCE_FIX_USERCHOICE "json"
  !insertmacro GLANCE_OPEN_WITH_ONLY "json" "Glance.JSON"
  ; SHCNE_ASSOCCHANGED + SHCNF_FLUSH: Проводник перечитает иконки без перезагрузки
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0x1000, p 0, p 0)'
  ; Сбросить кэш иконок Проводника (иначе старая иконка может держаться до перезагрузки)
  ${If} ${FileExists} "$SYSDIR\ie4uinit.exe"
    ExecWait '"$SYSDIR\ie4uinit.exe" -show'
  ${EndIf}
  ${If} $R3 != ""
    MessageBox MB_OK|MB_ICONINFORMATION "Windows не дал установщику сбросить старый выбор программы для: $R3.$\r$\n$\r$\nИз-за этого у этих файлов одна общая иконка Glance. Чтобы вернуть иконки по типам, для каждого из этих типов нажмите правой кнопкой на файл → «Открыть с помощью» → «Выбрать другое приложение» → Glance → «Всегда»." /SD IDOK
  ${EndIf}
!macroend

!macro NSIS_HOOK_POSTUNINSTALL
  DeleteRegValue SHCTX "Software\Classes\.txt\OpenWithProgids" "Glance.Text"
  DeleteRegValue SHCTX "Software\Classes\.md\OpenWithProgids" "Glance.Markdown"
  DeleteRegValue SHCTX "Software\Classes\.markdown\OpenWithProgids" "Glance.Markdown"
  DeleteRegValue SHCTX "Software\Classes\.json\OpenWithProgids" "Glance.JSON"
  System::Call 'shell32::SHChangeNotify(i 0x08000000, i 0, p 0, p 0)'
!macroend
