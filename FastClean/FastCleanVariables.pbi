;=======================================================================================================
;  Project  : FastClean
;  File     : FastCleanVariables.pb
;  Author   : Damian Daszkiewicz
;  WWW      : https://www.officeblog.pl/go/fastclean.php
;  Created  : 2025-11-11
;  Language : PureBasic 6.20
;-------------------------------------------------------------------------------------------------------
;  Description:
;     Program usuwa wpis ADS ZoneIdentifier z pliku przekazanego jako parametr
;     Jeśli parametrem jest folder, to porządkuje wszystkie pliki w folderze (z podfolderami)
;     przy okazji usuwa też z folderu różne śmieciowe pliki takie jak np. thumbs.db ~WRL*.tmp itp.
;-------------------------------------------------------------------------------------------------------
;  Revision History:
;     2021-11-11  Pierwsza publiczna wersja
;-------------------------------------------------------------------------------------------------------
;  License:
;     This software is provided "as is" without warranty of any kind.
;=======================================================================================================

; W tej zmiennej są trzymane logi
Global Logs.s
Global LangID
#LANG_PL = 1045

; Tablica z wzorcami plików/śmieci które usuwamy z folderu
Global Dim patterns.s(8)
patterns(0)=".DS_Store"
patterns(1)=".thumbs.db"
patterns(2)="desktop.ini"
patterns(3)="~$*.d*"
patterns(4)="~$*.x*"
patterns(5)="~$*.p*"
patterns(6)="~WRL*.tmp"
patterns(7)="~RF*.tmp"
patterns(8)="._*"
Global patternsSize = ArraySize(patterns())

Import "kernel32.lib"
  GetUserDefaultUILanguage()
EndImport
UsePNGImageDecoder()


; Stałe i struktury dla IShellLink
#CLSID_ShellLink = "{00021401-0000-0000-C000-000000000046}"
#IID_IShellLinkW = "{000214F9-0000-0000-C000-000000000046}"
#IID_IPersistFile = "{0000010B-0000-0000-C000-000000000046}"


; Definicje GUID na potrzeby funkcji utwórz skrót
DataSection
  CLSID_ShellLink:
  Data.l $00021401
  Data.w $0000, $0000
  Data.b $C0, $00, $00, $00, $00, $00, $00, $46
  
  IID_IShellLinkW:
  Data.l $000214F9
  Data.w $0000, $0000
  Data.b $C0, $00, $00, $00, $00, $00, $00, $46
  
  IID_IPersistFile:
  Data.l $0000010B
  Data.w $0000, $0000
  Data.b $C0, $00, $00, $00, $00, $00, $00, $46
EndDataSection
; IDE Options = PureBasic 6.30 beta 3 (Windows - x64)
; CursorPosition = 2
; EnableXP
; DPIAware
; SharedUCRT