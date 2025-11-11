;=======================================================================================================
;  Project  : FastClean
;  File     : FastCleanFunctions.pb
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

; odpowiednik funkcji Split z VBA
Procedure.i SplitFast(source.s, delimiter.s, List result.s())
  Protected *p.Character = @source
  Protected startPos = 0
  Protected length = Len(source)
  Protected delim.c = Asc(delimiter)
  Protected i, c.c
  
  ClearList(result())
  
  For i = 0 To length - 1
    c = PeekC(*p)
    If c = delim
      AddElement(result())
      result() = Mid(source, startPos + 1, i - startPos)
      startPos = i + 1
    EndIf
    *p + SizeOf(Character)
  Next
  
  ; Dodaj ostatni fragment jeśli istnieje
  If startPos < length
    AddElement(result())
    result() = Mid(source, startPos + 1, length - startPos)
  EndIf
  
  ProcedureReturn ListSize(result())
EndProcedure

; ładuj z pliku INI wzorce plików do usunięcia
Procedure LoadPatterns()
  Protected INIfile.s, i, pValue.s
  Static isLoaded.i
  
  ;Istnieje plik INI więc olewamy domyślne wzorce
  If isLoaded: ProcedureReturn: EndIf
  
  isLoaded = #True
  Debug "Load"
  INIfile.s = GetPathPart(ProgramFilename()) + "fastclean.ini"
  If FileSize(INIfile) >0
    i=-1
    OpenPreferences(INIfile)
    
    ; Liczenie rozmiaru tablicy
    PreferenceGroup("Patterns")
    ExaminePreferenceKeys()
    While NextPreferenceKey()
      i + 1
    Wend
    
    If i>-1
      ReDim patterns(i)
      i=-1
      ExaminePreferenceKeys()
      While NextPreferenceKey()
        pValue = Trim(PreferenceKeyValue())
        If pValue<>"" And pValue<>"*" And pValue<>"*.*"
          i + 1
          patterns(i) = pValue
        EndIf
      Wend      
    EndIf    
    ClosePreferences()
    patternsSize = i
  EndIf
EndProcedure


; Tworzy skrót do samego siebie na pulpicie (o ile go nie ma)
Procedure CreateDesktopShortcut()
  Protected ShortcutPath.s, TargetPath.s
  Protected *pShellLink.IShellLinkW
  Protected *pPersistFile.IPersistFile
  Protected result.i
  
  TargetPath = ProgramFilename()
  ShortcutPath = GetUserDirectory(#PB_Directory_Desktop) + "FastClean.lnk"  
  
  CoInitialize_(#Null)
  result = CoCreateInstance_(?CLSID_ShellLink, #Null, 1, ?IID_IShellLinkW, @*pShellLink)
  
  If result = 0 And *pShellLink
    *pShellLink\SetPath(TargetPath)
    
    Protected WorkingDir$ = GetPathPart(TargetPath)
    If WorkingDir$ <> ""
      *pShellLink\SetWorkingDirectory(WorkingDir$)
    EndIf
    
    result = *pShellLink\QueryInterface(?IID_IPersistFile, @*pPersistFile)
    
    If result = 0 And *pPersistFile
      *pPersistFile\Save(ShortcutPath, #True)
      *pPersistFile\Release()
    EndIf
    
    *pShellLink\Release()
  EndIf
  
  CoUninitialize_()
  ProcedureReturn Bool(FileSize(ShortcutPath) > 0)
EndProcedure


Procedure RunHelp()
  If LangID = #LANG_PL
    RunProgram("https://www.officeblog.pl/go/fastclean.php", "", "")
  Else
    RunProgram("https://www.officeblog.pl/go/fastclean_en.php", "", "")    
  EndIf
EndProcedure

; Funkcja usuwa ADS o nazwie ZoneIdentifier z pliku
Procedure.i DeleteZoneIdentifier(FileName.s)
  Protected StreamName.s = FileName + ":Zone.Identifier"
  If DeleteFile(StreamName)
    ProcedureReturn #True
  Else
    ProcedureReturn #False
  EndIf
EndProcedure

; funkcja sprawdza, czy dany plik posiada wpis ADS o nazwie ZoneIdentifier
Procedure.i HasZoneIdentifier(FileName.s)
  Protected StreamName.s = FileName + ":Zone.Identifier"
  Protected file = ReadFile(#PB_Any, StreamName)
  
  If file
    CloseFile(file)
    ProcedureReturn #True
  Else
    ProcedureReturn #False
  EndIf
EndProcedure

; Funkcja sprawdza rodzaj partycji (zwraca np. NTFS) - tylko na NTFS mogą być ADS
Procedure.s GetDriveFileSystem(DriveLetter.s)
  Protected buffer.s = Space(255)
  GetVolumeInformation_(DriveLetter, #Null, 0, #Null, #Null, #Null, @buffer, 255)
  ProcedureReturn buffer
EndProcedure

Procedure DeleteToRecycleBin(Path.s)
  Protected SH.SHFILEOPSTRUCT
  Protected *Buffer
  Protected length = StringByteLength(Path, #PB_Unicode) + 4 ; 2x Chr(0) = 4 bajty
  
  *Buffer = AllocateMemory(length)
  If *Buffer = 0
    ProcedureReturn #False
  EndIf
  
  ; Wpisz ścieżkę jako Unicode + dwa razy null
  PokeS(*Buffer, Path, -1, #PB_Unicode)
  PokeW(*Buffer + StringByteLength(Path, #PB_Unicode), 0)
  PokeW(*Buffer + StringByteLength(Path, #PB_Unicode) + 2, 0)
  
  SH\wFunc = #FO_DELETE
  SH\pFrom = *Buffer
  SH\fFlags = #FOF_ALLOWUNDO | #FOF_NOCONFIRMATION | #FOF_SILENT
  
  Protected result = SHFileOperation_(@SH)  
  FreeMemory(*Buffer)
  
  If result = 0 And SH\fAnyOperationsAborted = 0
    ProcedureReturn #True
  Else
    ProcedureReturn #False
  EndIf
EndProcedure


; Sprawdzanie dopasowania do wzorca
Procedure.i MatchWildcard(String$, Pattern$)
  Protected *s.Character = @String$
  Protected *p.Character = @Pattern$
  Protected *last_star.Character = 0
  Protected *last_s.Character = 0
  
  While *s\c
    Select *p\c
      Case '?'
        *s + SizeOf(Character)
        *p + SizeOf(Character)
        
      Case '*'
        *last_star = *p + SizeOf(Character)
        *last_s = *s
        *p + SizeOf(Character)
        
      Default
        If *p\c = *s\c
          *s + SizeOf(Character)
          *p + SizeOf(Character)
        ElseIf *last_star
          *p = *last_star
          *last_s + SizeOf(Character)
          *s = *last_s
        Else
          ProcedureReturn #False
        EndIf
    EndSelect
  Wend
  
  While *p\c = '*'
    *p + SizeOf(Character)
  Wend
  
  If *p\c = 0
    ProcedureReturn #True
  Else
    ProcedureReturn #False
  EndIf
EndProcedure

; Sprawdzamy czy dany plik można usunąć (czy jest na liście)
Procedure CanIDelete(file.s)
  Protected i.i
  
  For i=0 To patternsSize    
    If MatchWildcard(file, patterns(i))
      ProcedureReturn #True
    EndIf
  Next i
  
  ProcedureReturn #False
EndProcedure

; Rekurencyjne sprawdzanie wszystkich plików w danym folderze
Procedure ScanFolder(Path.s, Zone)
  Protected File.s  
  Protected dir
  
  If Right(Path, 1) <> "\": Path + "\": EndIf
  dir = ExamineDirectory(#PB_Any, Path, "*.*")
  
  If dir
    While NextDirectoryEntry(dir)
      File = DirectoryEntryName(dir)
      
      Select DirectoryEntryType(dir)
        Case #PB_DirectoryEntry_File
          If CanIDelete(LCase(File))            
            If DeleteToRecycleBin(Path + File)
              If LangID = #LANG_PL
                Logs + "[F] PRZENOSZENIE DO KOSZA: " + Path + File + #CRLF$
              Else
                Logs + "[F] MOVING TO TRASH: " + Path + File + #CRLF$
              EndIf  
            Else
              If LangID = #LANG_PL
                Logs + "[!] BŁĄD: " + Path + File + #CRLF$
              Else
                Logs + "[!] ERROR: " + Path + File + #CRLF$
              EndIf
            EndIf
          EndIf
          
          If Zone=#True And HasZoneIdentifier(Path + File)
            If DeleteZoneIdentifier(Path + File)
              If LangID = #LANG_PL
                Logs + "[Z] Usunięto ZoneIdentifier z pliku: " + Path + File + #CRLF$
              Else
                Logs + "[Z] Removed ZoneIdentifier from file: " + Path + File + #CRLF$
              EndIf
            Else
              If LangID = #LANG_PL
                Logs + "[!] Nie mogę usunąć ZoneIdentifier z pliku: " + Path + File + #CRLF$
              Else
                Logs + "[!] I can't remove ZoneIdentifier from file: " + Path + File + #CRLF$
              EndIf
            EndIf
          EndIf
          
        Case #PB_DirectoryEntry_Directory
          If File <> "." And File <> ".." And File <> "$RECYCLE.BIN" And File <> "System Volume Information"
            ScanFolder(Path + File, Zone)
          EndIf
      EndSelect
    Wend
    FinishDirectory(dir)
  Else
    Logs + "[!] Nie można otworzyć folderu: " + Path + #CRLF$
  EndIf
EndProcedure

; Główna funkcja do analizy folderu
; Funkcja sprawdza czy dysk jest NTFS i wywołuje rekurencyjne skanowanie folderu
Procedure ScanFolder_helper(dirPath.s)
  Protected DriveLetter.s=Left(dirPath, 3)
  Protected Zone.i
  
  If UCase(GetDriveFileSystem(DriveLetter)) = "NTFS"
    Zone=#True
  Else
    Zone=#False
  EndIf  
  
  ScanFolder(dirPath, Zone)
EndProcedure


Procedure CheckZoneIdentifier(fileName.s)
  If HasZoneIdentifier(filename)
    If DeleteZoneIdentifier(fileName)
      If LangID = #LANG_PL
        Logs + "[Z] Usunięto ZoneIdentifier z pliku: " + fileName + #CRLF$
      Else
        Logs + "[Z] Removed ZoneIdentifier from file: " + fileName + #CRLF$
      EndIf
    EndIf
  EndIf
EndProcedure

Procedure ReadDragFiles(Files.s)
  Protected NewList parts.s()
  Protected currentFile.s, fSize
  
  
  SplitFast(Files, Chr(10), parts())
  
  ForEach parts()
    currentFile = parts() 
    fSize=FileSize(currentFile)    
    
    If fsize>-1
      CheckZoneIdentifier(currentFile)
    EndIf
    
    If fSize=-2
      LoadPatterns() ;Ładuj z pliku INI jakie są wzorce plików śmieci
      ScanFolder_helper(currentFile)
    EndIf
    
    If fSize = -1
      If LangID = #LANG_PL
        MessageRequester("Błąd", "Plik (lub folder): " + currentFile + " nie istnieje!", #PB_MessageRequester_Error)  
      Else
        MessageRequester("Error", "file (or folder): " + currentFile + " does not exist!", #PB_MessageRequester_Error)  
      EndIf
    EndIf
  Next  
EndProcedure



; ----------------- LOGS ------------------------
Procedure frmLogs_Resize()
  ResizeGadget(txtLogs, 0, 0, WindowWidth(frmLogs)-2, WindowHeight(frmLogs)-20)
  RedrawWindow_(WindowID(frmLogs), #Null, #Null, #RDW_INVALIDATE | #RDW_ERASE | #RDW_ALLCHILDREN | #RDW_UPDATENOW)
EndProcedure

Procedure LoadLogs()
  Protected Event, Quit, text.s, lines, WindowHeight
  
  ; ustalamy wysokość okna, w przybliżeniu zakładamy, że 1 linia tekstu to 15 pikseli do tego na menu i pasek tytułowy doliczamy 60
  ; okno nie może być niższe niż 150 px i wyższe niż 600
  lines = CountString(Logs, #CRLF$)
  WindowHeight = 15 * lines + 60
  If WindowHeight <150: WindowHeight = 150: EndIf
  If WindowHeight >600: WindowHeight = 600: EndIf
  
  
  OpenfrmLogs(#PB_Ignore, #PB_Ignore, 800, WindowHeight)
  SetGadgetText(txtLogs, Logs)
  AddKeyboardShortcut(frmLogs, #PB_Shortcut_Escape, #MenuItem_quit)
  If LangID<>#LANG_PL
    SetMenuItemText(0, #MenuItem_copy, "Copy to clipboard")
    SetMenuItemText(0, #MenuItem_help, "Help")
    SetMenuItemText(0, #MenuItem_quit, "Quit")
    SetWindowTitle(frmLogs, "Report (press ESC to close the window)")
  EndIf
  BindEvent(#PB_Event_SizeWindow, @frmLogs_Resize(), frmLogs)
  PostEvent(#PB_Event_SizeWindow, frmLogs, 0)  
  HideWindow(frmLogs, #False)
  
  Repeat
    Event = WindowEvent()
    Select Event
      Case #PB_Event_CloseWindow
        Quit=1
        
        ;Case #PB_EVENT_  ;#PB_Event_SizeWindow
        ;  frmLogs_Resize()
        
      Case #PB_Event_Menu
        Select EventMenu()
          Case#MenuItem_copy
            text = GetGadgetText(txtLogs)
            SetClipboardText(text)
            MessageRequester("FastClean", "Skopiowano!", #PB_MessageRequester_Info)
          Case #MenuItem_help
            RunHelp()
          Case #MenuItem_quit
            Quit=1
        EndSelect
    EndSelect
    Delay(7)
  Until Quit=1
  
  End
EndProcedure

Procedure LoadLogsHelper()
  If Logs =""
    If LangID = #LANG_PL
      MessageRequester("FastClean", "Brak elementów do posprzątania", #PB_MessageRequester_Info)
    Else
      MessageRequester("FastClean", "No items to clean up", #PB_MessageRequester_Info)
    EndIf
  Else
    LoadLogs()
  EndIf
EndProcedure

; ---------- Obsługa zdarzeń formularza frmMain ---------------
Procedure frmMain_Events()
  Protected Event, Quit, DragFiles.s
  
  Repeat
    Event = WindowEvent()
    Select Event
      Case #PB_Event_CloseWindow
        Quit=1
        
      Case #PB_Event_Gadget
        Select EventGadget()
          Case cmdQuit
            Quit=1
            
          Case cmdHelp
            RunHelp()
            
          Case cmdShortcut
            CreateDesktopShortcut()
            HideGadget(cmdShortcut, #True)
        EndSelect
        
      Case #PB_Event_Menu
        Select EventMenu()
          Case 1, 2
            Quit=1
        EndSelect
        
      Case #PB_Event_GadgetDrop, #PB_Event_WindowDrop
        DragFiles = EventDropFiles()        
        If DragFiles <> ""
          ReadDragFiles(DragFiles)
          HideWindow(frmMain, #True)
          LoadLogsHelper()
          Quit=1
        EndIf        
    EndSelect
    Delay(7)
  Until Quit=1
EndProcedure



; ----------------- MAIN -------------------

; Funkcja startowa
Procedure Main()
  Protected cmd.s
  Protected a.i
  Protected currentParam.s
  Protected fSize.i
  Protected Desktop.s
  
  LangID = GetUserDefaultUILanguage()
  cmd=ProgramParameter()  
  If cmd=""
    Desktop = GetUserDirectory(#PB_Directory_Desktop) 
    
    OpenfrmMain()
    
    ;drag&drop
    EnableGadgetDrop(imgHelp, #PB_Drop_Files, #PB_Drag_Copy)
    EnableWindowDrop(frmMain, #PB_Drop_Files, #PB_Drag_Copy)
    
    ; Pokaż przycisk utwórz skrót jak nie ma na pulpicie sktótu albo samego siebie
    If GetPathPart(ProgramFilename())=Desktop Or FileSize(Desktop+"FastClean.lnk")>-1
      HideGadget(cmdShortcut, #True)
    EndIf
    
    ; Takie tam uproszczone ładowanie angielskich napisów
    If LangID<>#LANG_PL
      SetGadgetText(cmdShortcut, "Create shortcut")
      GadgetToolTip(cmdShortcut, "Creates an icon for the program on the desktop")
      SetGadgetText(cmdHelp, "Help")
      SetGadgetText(cmdQuit, "Quit")
      SetGadgetText(lblInfo, "Drag a file or folder onto a program icon (or shortcut)")
    EndIf
    
    HideWindow(frmMain, #False)
    AddKeyboardShortcut(frmMain, #PB_Shortcut_Return, 2)
    AddKeyboardShortcut(frmMain, #PB_Shortcut_Escape, 2)
    frmMain_Events()
    End
  EndIf
  
  For a=0 To CountProgramParameters()-1
    currentParam = ProgramParameter(a)    
    fSize=FileSize(currentParam)    
    
    If fsize>-1
      CheckZoneIdentifier(currentParam)
    EndIf
    
    If fSize=-2
      LoadPatterns() ;Ładuj z pliku INI jakie są wzorce plików śmieci
      ScanFolder_helper(currentParam)
    EndIf
    
    If fSize = -1
      If LangID = #LANG_PL
        MessageRequester("Błąd", "Plik (lub folder): " + currentParam + " nie istnieje!", #PB_MessageRequester_Error)  
      Else
        MessageRequester("Error", "file (or folder): " + currentParam + " does not exist!", #PB_MessageRequester_Error)  
      EndIf
    EndIf
  Next a
EndProcedure



; około 40% wolniejsze od tej brzydkiej wersji
; Structure WIN32_FIND_STREAM_DATA
;   StreamSize.q
;   cStreamName.w[296] ; MAX_PATH + 36
; EndStructure
; 
; Import ""
;   FindFirstStreamW(lpFileName.p-unicode, InfoLevel.l, *lpFindStreamData, dwFlags.l)
;   FindNextStreamW(hFindStream.i, *lpFindStreamData)
;   FindClose(hFindStream.i)
; EndImport
;
; Procedure.i HasZoneIdentifier2(FileName.s)
;   Protected streamData.WIN32_FIND_STREAM_DATA
;   Protected hFind, found.i = #False
;   Protected CurrentStreamName.s
;   Protected StreamName.s = "Zone.Identifier"
;   
;   hFind = FindFirstStreamW(FileName, 0, @streamData, 0)
;   If hFind <> #INVALID_HANDLE_VALUE
;     Repeat
;       CurrentStreamName = PeekS(@streamData\cStreamName[0], -1, #PB_Unicode)
;       
;       If FindString(CurrentStreamName, ":" + StreamName + ":", 1, #PB_String_NoCase)
;         found = #True
;         Break
;       EndIf
;       
;     Until FindNextStreamW(hFind, @streamData) = 0
;     FindClose(hFind)
;   EndIf
;   
;   ProcedureReturn found
; EndProcedure

; IDE Options = PureBasic 6.30 beta 3 (Windows - x64)
; CursorPosition = 2
; Folding = ----
; Markers = 126
; EnableXP
; DPIAware
; SharedUCRT