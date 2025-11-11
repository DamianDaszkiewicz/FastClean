;=======================================================================================================
;  Project  : FastClean
;  File     : FastClean.pb
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



EnableExplicit
IncludeFile "frmMain.pbf"
IncludeFile "frmLogs.pbf"
IncludeFile "FastCleanVariables.pbi"
IncludeFile "FastCleanFunctions.pbi"


Main()
LoadLogsHelper()


; IDE Options = PureBasic 6.30 beta 3 (Windows - x64)
; CursorPosition = 4
; EnableXP
; DPIAware
; Executable = V:\FastClean.exe