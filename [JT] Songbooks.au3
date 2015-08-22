#AutoIt3Wrapper_Icon=.\LOGO.ico
#AutoIt3Wrapper_Outfile=[JT] Songbooks.exe
; 0=Low 2=normal 4=High
#AutoIt3Wrapper_Compression=2
#AutoIt3Wrapper_Res_Field=Productname|[JT] Songbooks
#AutoIt3Wrapper_Res_Field=ProductVersion|2.0.1
#AutoIt3Wrapper_Res_Comment=Created by o0johntam0o
#AutoIt3Wrapper_Res_Description=Songbooks for guitarist
#AutoIt3Wrapper_Res_Fileversion=1.0.0
#AutoIt3Wrapper_Res_LegalCopyright=[JT] Knowledge
; Vietnamese=1066, English (United Kingdom)=2057 (Hex)
#AutoIt3Wrapper_Res_Language=2057
#AutoIt3Wrapper_Run_Tidy=Y
; I wrote here to remember the Obfuscation usage :D
#AutoIt3Wrapper_Run_Obfuscator=N
#Obfuscator_Parameters=/Convert_Vars=0
#Obfuscator_Parameters=/Convert_Funcs=0
#Obfuscator_Parameters=/Convert_Strings=0
#Obfuscator_Parameters=/Convert_Numerics=0
#Obfuscator_Parameters=/StripUnusedFunc=0
#Obfuscator_Parameters=/StripUnusedVars=0
#Obfuscator_On
#Obfuscator_Off
#include <ButtonConstants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <TabConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <TreeViewConstants.au3>
#include <GuiRichEdit.au3>
#include "Funcs.au3"

AutoItSetOption('MustDeclareVars', 1)
AutoItSetOption('GUICloseOnESC', 0)
AutoItSetOption('TrayIconHide', 1)
AutoItSetOption('GUIResizeMode', $GUI_DOCKALL)

#REGION ### GLOBAL VARIABLES ###
Global CONST $AppVersion = '2.0.1'
Global CONST $JdbUrl = 'https://google.com/SongbooksDB.mdb'
Global CONST $AboutFile = @ScriptDir & '\About.jpg'
Global CONST $HelpFile = @ScriptDir & '\Help.txt'
Global CONST $ConfigFile = @ScriptDir & '\Settings.ini'
Global CONST $ChordMapFile = @ScriptDir & '\ChordMap.ini'
Global CONST $LanguageDir = @ScriptDir & '\Languages'
Global $CurrentLanguage = IniRead($ConfigFile, 'Common', 'CustomLanguage', 'English')
Global $CurrentScreen = IniRead($ConfigFile, 'Common', 'Screen', 1024)
Global $CurrentSongInfo[4] = ['', '', '', '']
Global $tmpData[4] = ['', '', '', '']
Global $tmpString = ''
Global $tmpInt = 0
Global $CurrentTreeItem = ''
Global $CurrentSearch = ''
Global $UpdateMode = 0
Global $ToneChange = 0
Global $GuiMessages = 0

Global $RebuildTreeNeeded = 0

Global CONST $DB_Object = ObjCreate('DAO.DBEngine.36')
Global CONST $DB_Location = @ScriptDir & '\SongbooksDB.mdb'
Global CONST $SongDB_Table = 'JT_Songbooks'
Global CONST $SongDB_ColTitle = 'Title'
Global CONST $SongDB_ColAuthor = 'Author'
Global CONST $SongDB_ColCadence = 'Cadence'
Global CONST $SongDB_ColLyrics = 'Lyrics'
#ENDREGION ; <== SOME GLOBAL VARIABLES

#REGION ### FILE INSTALL ###
If (Not FileExists($LanguageDir & '\Vietnamese.ini')) Then
	DirCreate($LanguageDir)
	FileInstall('.\Languages\Vietnamese.ini', $LanguageDir & '\Vietnamese.ini', 1)
EndIf

If (Not FileExists('.\Help.txt')) Then
	FileInstall('.\Help.txt', $HelpFile, 1)
EndIf

If (Not FileExists('.\About.jpg')) Then
	FileInstall('.\About.jpg', $AboutFile, 1)
EndIf
#ENDREGION ; <== FILE INSTALL

JT_INIT()

SplashTextOn('', JT_TranSlate('Starting'), 300, 100, -1, -1, 0, 'Tahoma')

#REGION ### MAKE MAIN FORM ###
Global $FormMain = GUICreate('[JT] Songbooks')
GUISetStyle(BitOr($GUI_SS_DEFAULT_GUI, $WS_MINIMIZEBOX), -1, $FormMain)
GUISetFont(9, 400, 0, 'Tahoma', $FormMain)
GUISetHelp('notepad.exe "' & $HelpFile & '"', $FormMain)

; BEGIN - TREE
Global $Tree = GUICtrlCreateTreeView(0, 0)
GUICtrlSetStyle(-1, BitOR($TVS_HASBUTTONS, $TVS_HASLINES, $TVS_LINESATROOT, $TVS_DISABLEDRAGDROP, $TVS_SHOWSELALWAYS), $WS_EX_CLIENTEDGE)
For $i = 65 To 90
	Assign('TreeItem' & Chr($i), GUICtrlCreateTreeViewItem(Chr($i), $Tree), 2)
	GUICtrlSetColor(-1, 0x0000C0)
Next
Global $TreeItemOther = GUICtrlCreateTreeViewItem('', $Tree)
GUICtrlSetColor(-1, 0x0000C0)
JT_BuildTree(0)
; END - TREE

; BEGIN - TAB
Global $Tab = GUICtrlCreateTab(0, 0)
GUICtrlSetFont(-1, 9, 400, 0, 'Tahoma')
; TAB 1
Global $Tab1 = GUICtrlCreateTabItem('Tab1')
GUICtrlSetFont(-1, 9, 400, 0, 'Tahoma')
Global $LabelTitle = GUICtrlCreateLabel('', 0, 0)
Global $InputTitle = GUICtrlCreateInput('', 0, 0)
GUICtrlSetStyle(-1, BitOr($ES_READONLY, $ES_AUTOHSCROLL))
Global $LabelAuthor = GUICtrlCreateLabel('', 0, 0)
Global $InputAuthor = GUICtrlCreateInput('', 0, 0)
GUICtrlSetStyle(-1, BitOr($ES_READONLY, $ES_AUTOHSCROLL))
Global $LabelCadence = GUICtrlCreateLabel('', 0, 0)
Global $InputCadence = GUICtrlCreateInput('', 0, 0)
GUICtrlSetStyle(-1, BitOr($ES_READONLY, $ES_AUTOHSCROLL))

Global $ButtonNew = GUICtrlCreateButton('', 0, 0)
Global $ButtonEdit = GUICtrlCreateButton('', 0, 0)
Global $ButtonSave = GUICtrlCreateButton('', 0, 0)
Global $ButtonDelete = GUICtrlCreateButton('', 0, 0)

GUICtrlCreateTabItem('')
; TAB 2
Global $Tab2 = GUICtrlCreateTabItem('Tab2')
GUICtrlSetFont(-1, 9, 400, 0, 'Tahoma')
Global $LabelChord = GUICtrlCreateLabel('', 0, 0)
Global $ButtonChordBold = GUICtrlCreateButton('B', 0, 0)
Global $ButtonChordItalic = GUICtrlCreateButton('I', 0, 0)
Global $ButtonChordUnderline = GUICtrlCreateButton('U', 0, 0)
Global $ButtonChordBlack = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0x000000)
Global $ButtonChordWhite = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $ButtonChordGreen = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0x00FF00)
Global $ButtonChordBlue = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0x0000FF)
Global $ButtonChordRed = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0xFF0000)
Global $LabelLyrics = GUICtrlCreateLabel('', 0, 0)
Global $ButtonLyricsBold = GUICtrlCreateButton('B', 0, 0)
Global $ButtonLyricsItalic = GUICtrlCreateButton('I', 0, 0)
Global $ButtonLyricsUnderline = GUICtrlCreateButton('U', 0, 0)
Global $ButtonLyricsBlack = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0x000000)
Global $ButtonLyricsWhite = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0xFFFFFF)
Global $ButtonLyricsGreen = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0x00FF00)
Global $ButtonLyricsBlue = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0x0000FF)
Global $ButtonLyricsRed = GUICtrlCreateButton('', 0, 0)
GUICtrlSetBkColor(-1, 0xFF0000)
Global $LabelTone = GUICtrlCreateLabel('', 0, 0)
Global $ButtonToneFlat = GUICtrlCreateButton('b', 0, 0)
Global $ButtonToneDefault = GUICtrlCreateButton('O', 0, 0)
Global $ButtonToneSharp = GUICtrlCreateButton('#', 0, 0)
Global $LabelFontSize = GUICtrlCreateLabel('', 0, 0)
Global $ButtonFontSizeSmall = GUICtrlCreateButton('-', 0, 0)
Global $ButtonFontSizeNormal = GUICtrlCreateButton('A', 0, 0)
Global $ButtonFontSizeLarge = GUICtrlCreateButton('+', 0, 0)

Global $LabelViewChord = GUICtrlCreateLabel('', 0, 0)
Global $ButtonAViewChord = GUICtrlCreateButton('A', 0, 0)
Global $ButtonBViewChord = GUICtrlCreateButton('B', 0, 0)
Global $ButtonCViewChord = GUICtrlCreateButton('C', 0, 0)
Global $ButtonDViewChord = GUICtrlCreateButton('D', 0, 0)
Global $ButtonEViewChord = GUICtrlCreateButton('E', 0, 0)
Global $ButtonFViewChord = GUICtrlCreateButton('F', 0, 0)
Global $ButtonGViewChord = GUICtrlCreateButton('G', 0, 0)
Global $ButtonSViewChord = GUICtrlCreateButton('', 0, 0)

GUICtrlCreateTabItem('')
; TAB 3
Global $Tab3 = GUICtrlCreateTabItem('Tab3')
GUICtrlSetFont(-1, 9, 400, 0, 'Tahoma')
Global $ButtonExportLIB_JDB = GUICtrlCreateButton('', 0, 0)
Global $ButtonExportLIB_TXT = GUICtrlCreateButton('', 0, 0)
Global $ButtonPrint = GUICtrlCreateButton('', 0, 0)
Global $ButtonUpdate = GUICtrlCreateButton('', 0, 0)
Global $ButtonSetting = GUICtrlCreateButton('', 0, 0)
Global $ButtonAbout = GUICtrlCreateButton('', 0, 0)
GUICtrlCreateTabItem('')
; END - TAB

Global $InputLyrics = _GUICtrlRichEdit_Create($FormMain, '', 0, 0, 10, 10, BitOr($ES_MULTILINE, $ES_READONLY, $ES_AUTOVSCROLL, $WS_VSCROLL))

Global $ButtonCollapse = GUICtrlCreateButton('', 0, 0)
Global $LabelSearch = GUICtrlCreateLabel('', 0, 0)
Global $InputSearch = GUICtrlCreateInput('', 0, 0)
GUICtrlSetStyle(-1, 0)
Global $CheckboxScroll = GUICtrlCreateCheckbox('', 0, 0)
Global $InputScroll = GUICtrlCreateInput('2', 0, 0)
GUICtrlSetStyle(-1, BitOR($ES_CENTER, $ES_NUMBER))
GUICtrlCreateUpdown(-1)
GUICtrlSetLimit(-1, 50, 0)
Global $CheckboxOntop = GUICtrlCreateCheckbox('', 0, 0)

#ENDREGION ; <== MAKE MAIN FORM

HotKeySet('^{DEL}', 'JT_ExitApp')
HotKeySet('^`', 'JT_ViewChordPre')
HotKeySet('^+c', 'JT_AutoBracketPre')

$CurrentTreeItem = GUICtrlRead($Tree)

JT_SetMainFormSize($CurrentScreen, 1)
JT_SetMainFormText($CurrentLanguage, 1)
Sleep(1000)
GUISetState(@SW_SHOW)
SplashOff()

While 1
	$GuiMessages = GUIGetMsg()
	Switch $GuiMessages
		Case $GUI_EVENT_CLOSE
			_GUICtrlRichEdit_Destroy($InputLyrics)
			GuiDelete()
			Exit
			
		; Chord tuner
		Case $ButtonToneSharp
			JT_RichEditSetStyle($InputLyrics, 'ChordTuner', '#')
			$ToneChange = $ToneChange + 1
		Case $ButtonToneDefault
			JT_RichEditSetStyle($InputLyrics, 'ChordTuner', $ToneChange)
			$ToneChange = 0
		Case $ButtonToneFlat
			JT_RichEditSetStyle($InputLyrics, 'ChordTuner', 'b')
			$ToneChange = $ToneChange - 1
			
		; Font size
		Case $ButtonFontSizeSmall
			JT_RichEditSetStyle($InputLyrics, 'FontSize', -2)
		Case $ButtonFontSizeNormal
			JT_RichEditSetStyle($InputLyrics, 'FontSize', 0)
		Case $ButtonFontSizeLarge
			JT_RichEditSetStyle($InputLyrics, 'FontSize', 2)
			
		; Chord styles
		Case $ButtonChordBold
			JT_RichEditSetStyle($InputLyrics, 'ChordStyle', 'bo')
		Case $ButtonChordItalic
			JT_RichEditSetStyle($InputLyrics, 'ChordStyle', 'it')
		Case $ButtonChordUnderline
			JT_RichEditSetStyle($InputLyrics, 'ChordStyle', 'un')
		Case $ButtonChordBlack
			JT_RichEditSetStyle($InputLyrics, 'ChordStyle', '0x000000')
		Case $ButtonChordWhite
			JT_RichEditSetStyle($InputLyrics, 'ChordStyle', '0xFFFFFF')
		Case $ButtonChordRed
			JT_RichEditSetStyle($InputLyrics, 'ChordStyle', '0x0000FF')
		Case $ButtonChordGreen
			JT_RichEditSetStyle($InputLyrics, 'ChordStyle', '0x00FF00')
		Case $ButtonChordBlue
			JT_RichEditSetStyle($InputLyrics, 'ChordStyle', '0xFF0000')
			
		; Lyrics styles
		Case $ButtonLyricsBold
			JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', 'bo')
		Case $ButtonLyricsItalic
			JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', 'it')
		Case $ButtonLyricsUnderline
			JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', 'un')
		Case $ButtonLyricsBlack
			JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', '0x000000')
		Case $ButtonLyricsWhite
			JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', '0xFFFFFF')
		Case $ButtonLyricsRed
			JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', '0x0000FF')
		Case $ButtonLyricsGreen
			JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', '0x00FF00')
		Case $ButtonLyricsBlue
			JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', '0xFF0000')
			
		; View Chord
		Case $ButtonAViewChord
			JT_ViewChord(GUICtrlRead($ButtonAViewChord))
		Case $ButtonBViewChord
			JT_ViewChord(GUICtrlRead($ButtonBViewChord))
		Case $ButtonCViewChord
			JT_ViewChord(GUICtrlRead($ButtonCViewChord))
		Case $ButtonDViewChord
			JT_ViewChord(GUICtrlRead($ButtonDViewChord))
		Case $ButtonEViewChord
			JT_ViewChord(GUICtrlRead($ButtonEViewChord))
		Case $ButtonFViewChord
			JT_ViewChord(GUICtrlRead($ButtonFViewChord))
		Case $ButtonGViewChord
			JT_ViewChord(GUICtrlRead($ButtonGViewChord))
		Case $ButtonSViewChord
			JT_ViewChord(_GUICtrlRichEdit_GetSelText($InputLyrics))
			
		Case $ButtonExportLIB_JDB
			$tmpString = InputBox(JT_TranSlate('Export libraries (JDB)'), JT_TranSlate('Database version') & ':')
			If ($tmpString <> '') Then
				If (JT_ExportLIB_JDB(FileSaveDialog(JT_TranSlate('Choose path'), @ScriptDir, '[JT] Songbooks DB (*.JDB)', 16+2, 'SongbooksDB.JDB', $FormMain), $tmpString) == 1) Then
					JT_MessageBox($FormMain, 'Info', 'Export successful')
				Else
					JT_MessageBox($FormMain, 'Error', 'Export failed')
				EndIf
			EndIf
			
		Case $ButtonExportLIB_TXT
			$tmpString = FileSelectFolder(JT_TranSlate('Choose path'), '', 1+2, '', $FormMain)
			If (FileExists($tmpString)) Then
				If (JT_ExportLIB_TXT($tmpString) == 1) Then
					JT_MessageBox($FormMain, 'Info', 'Export successful')
				Else
					JT_MessageBox($FormMain, 'Error', 'Export failed')
				EndIf
			EndIf
			
		Case $ButtonPrint
			JT_PrintLyrics($CurrentSongInfo)
			
		Case $ButtonUpdate
			$RebuildTreeNeeded = 0
			JT_Update()
			If ($RebuildTreeNeeded == 1) Then
				If (JT_MessageBox($FormMain, 'Confirm', 'Do you want to reload the song lists?') == 6) Then
					GUISetState(@SW_DISABLE, $FormMain)
					SplashTextOn('', JT_TranSlate('Working'), 300, 100, -1, -1, 0, 'Tahoma')
					JT_BuildTree(1)
					SplashOff()
					GUISetState(@SW_ENABLE, $FormMain)
				EndIf
			EndIf
			
		Case $ButtonSetting
			JT_Settings()
			JT_SetMainFormText(IniRead($ConfigFile, 'Common', 'CustomLanguage', 'English'))
			JT_SetMainFormSize(IniRead($ConfigFile, 'Common', 'Screen', 1024))
			$CurrentLanguage = IniRead($ConfigFile, 'Common', 'CustomLanguage', 'English')
			$CurrentScreen = IniRead($ConfigFile, 'Common', 'Screen', 1024)
			
		Case $ButtonAbout
			JT_About()
			
		Case $ButtonCollapse
			For $i = 0 To 26
				ControlTreeView($FormMain, '', $Tree, 'Collapse', '#' & $i)
			Next
			
		Case $ButtonDelete
			If (StringInStr(ControlTreeView($FormMain, '', $Tree, 'GetSelected', 1), '|') <> 0) Then
				If (JT_MessageBox($FormMain, 'Warning', 'Do you really want to DELETE this song: "%s"', GUICtrlRead($CurrentTreeItem, 1)) == 6) Then
					If (JT_DelSong(GUICtrlRead($CurrentTreeItem, 1)) == 1) Then
						JT_MessageBox($FormMain, 'Info', 'Deleted')
						GUICtrlDelete($CurrentTreeItem)
					Else
						JT_MessageBox($FormMain, 'Error', 'Can not delete this song')
					EndIf
				EndIf
			EndIf
			
		Case $ButtonNew
			If ($UpdateMode == 0 Or JT_MessageBox($FormMain, 'Confirm', "The song haven't saved yet. Do you want to leave it?") == 6) Then
				$UpdateMode = 2
				GUICtrlSetStyle($InputTitle, $GUI_SS_DEFAULT_INPUT)
				GUICtrlSetState($InputTitle, $GUI_FOCUS)
				GUICtrlSetStyle($InputAuthor, $GUI_SS_DEFAULT_INPUT)
				GUICtrlSetStyle($InputCadence, $GUI_SS_DEFAULT_INPUT)
				_GUICtrlRichEdit_SetReadOnly($InputLyrics, False)
				GUICtrlSetData($InputTitle, '')
				GUICtrlSetData($InputAuthor, '')
				GUICtrlSetData($InputCadence, '')
				JT_RichEditSetText($InputLyrics, '')
			EndIf
			
		Case $ButtonEdit
			If ($UpdateMode <> 2) Then
				$UpdateMode = 1
				GUICtrlSetStyle($InputTitle, $GUI_SS_DEFAULT_INPUT)
				GUICtrlSetStyle($InputAuthor, $GUI_SS_DEFAULT_INPUT)
				GUICtrlSetStyle($InputCadence, $GUI_SS_DEFAULT_INPUT)
				_GUICtrlRichEdit_SetReadOnly($InputLyrics, False)
			EndIf
			
		Case $ButtonSave
			If ($UpdateMode <> 0) Then
				$tmpData[0] = GUICtrlRead($InputTitle)
				$tmpData[1] = GUICtrlRead($InputAuthor)
				$tmpData[2] = GUICtrlRead($InputCadence)
				$tmpData[3] = _GUICtrlRichEdit_GetText($InputLyrics)
				$tmpString = ''
				If ($UpdateMode == 1) Then
					$tmpString = JT_AddSong($tmpData, GUICtrlRead($CurrentTreeItem, 1))
				Else
					$tmpString = JT_AddSong($tmpData)
				EndIf
				
				GUICtrlSetData($InputTitle, $tmpData[0])
				GUICtrlSetData($InputAuthor, $tmpData[1])
				GUICtrlSetData($InputCadence, $tmpData[2])
				JT_RichEditSetText($InputLyrics, $tmpData[3])
				
				If ($tmpString == '') Then
					If ($UpdateMode == 1) Then
						GUICtrlDelete($CurrentTreeItem)
					EndIf
					If (GUICtrlCreateTreeViewItem($tmpData[0], Eval('TreeItem' & StringUpper(_JT_ToLatin(StringLeft($tmpData[0], 1))))) == 0) Then
						GUICtrlCreateTreeViewItem($tmpData[0], $TreeItemOther)
					EndIf
					GUICtrlSetStyle($InputTitle, $ES_READONLY)
					GUICtrlSetStyle($InputAuthor, $ES_READONLY)
					GUICtrlSetStyle($InputCadence, $ES_READONLY)
					_GUICtrlRichEdit_SetReadOnly($InputLyrics, True)
					ControlTreeView($FormMain, '', $Tree, 'Select', _JT_ToLatin(StringLeft($tmpData[0], 1)) & '|' & $tmpData[0])
					If (@error) Then ControlTreeView($FormMain, '', $Tree, 'Select', JT_TranSlate('Other') & '|' & $tmpData[0])
					JT_MessageBox($FormMain, 'Info', 'The song was added/updated successful')
					$UpdateMode = 0
					JT_GetDataFromTree(1)
				Else
					JT_MessageBox($FormMain, 'Error', $tmpString)
				EndIf
			EndIf
			
		Case Else
			JT_GetDataFromTree(0)
			If ($CurrentSearch <> GUICtrlRead($InputSearch)) Then
				$CurrentSearch = GUICtrlRead($InputSearch)
				JT_SearchSong($CurrentSearch)
			EndIf
			If (GUICtrlRead($CheckboxOntop) == 1) Then
				WinSetOnTop($FormMain, '', 1)
			Else
				WinSetOnTop($FormMain, '', 0)
			EndIf
			JT_RichEditScroll($InputLyrics, GUICtrlRead($CheckboxScroll))
	EndSwitch
WEnd

#REGION ### FUNCTIONS BLOCK - CHANGE APPEARANCE ###
FUNC JT_SetMainFormText($lang = 'English', $force = 0)
	If ($lang <> $CurrentLanguage Or $force == 1) Then
		GuiCtrlSetData($TreeItemOther, JT_TranSlate('Other'))
		
		GuiCtrlSetData($Tab1, JT_TranSlate('View'))
		GuiCtrlSetData($Tab2, JT_TranSlate('Style/Consultation'))
		GuiCtrlSetData($Tab3, JT_TranSlate('Misc'))

		GuiCtrlSetData($LabelTitle, JT_TranSlate('Title'))
		GuiCtrlSetData($LabelAuthor, JT_TranSlate('Author'))
		GuiCtrlSetData($LabelCadence, JT_TranSlate('Cadence'))

		GuiCtrlSetData($ButtonCollapse, JT_TranSlate('Collapse'))
		GuiCtrlSetData($LabelSearch, JT_TranSlate('Search'))
		GuiCtrlSetTip($InputSearch, JT_TranSlate('Enter keywords'))
		GuiCtrlSetData($CheckboxScroll, JT_TranSlate('Scroll'))
		GuiCtrlSetTip($InputScroll, JT_TranSlate('Scroll speed'))
		GuiCtrlSetData($CheckboxOntop, JT_TranSlate('Ontop'))

		GuiCtrlSetData($ButtonNew, JT_TranSlate('New'))
		GuiCtrlSetData($ButtonEdit, JT_TranSlate('Edit'))
		GuiCtrlSetData($ButtonSave, JT_TranSlate('Save'))
		GuiCtrlSetData($ButtonDelete, JT_TranSlate('Delete'))

		GuiCtrlSetData($LabelChord, JT_TranSlate('Chord'))

		GuiCtrlSetData($LabelTone, JT_TranSlate('Tone'))

		GuiCtrlSetData($LabelLyrics, JT_TranSlate('Lyrics'))

		GuiCtrlSetData($LabelFontSize, JT_TranSlate('Font size'))

		GuiCtrlSetData($LabelViewChord, JT_TranSlate('Chord'))

		GuiCtrlSetData($ButtonExportLIB_JDB, JT_TranSlate('Export libraries (JDB)'))
		GuiCtrlSetData($ButtonExportLIB_TXT, JT_TranSlate('Export libraries (TXT)'))
		GuiCtrlSetData($ButtonPrint, JT_TranSlate('Print'))
		GuiCtrlSetData($ButtonSetting, JT_TranSlate('Settings'))
		GuiCtrlSetData($ButtonUpdate, JT_TranSlate('Update libraries'))
		GuiCtrlSetData($ButtonAbout, JT_TranSlate('About'))
	EndIf
ENDFUNC

FUNC JT_SetMainFormSize($opt = 1024, $force = 0)
	If (($opt == 1024 And $force == 1) Or ($opt == 1024 And $CurrentScreen <> $opt)) Then
		WinMove($FormMain, '', 0, 0, 1000+6, 700+25)
		GuiCtrlSetPos($Tree, 5, 0, 275, 670)
		GuiCtrlSetPos($Tab, 285, 0, 710, 90)
		_WinAPI_MoveWindow($InputLyrics, 285, 95, 710, 575)
		_GUICtrlRichEdit_SetRECT($InputLyrics, 10, 10)
		
		GuiCtrlSetPos($LabelTitle, 295, 35, 60, 20)
		GuiCtrlSetPos($InputTitle, 365, 35, 460, 20)
		GuiCtrlSetPos($LabelAuthor, 295, 60, 60, 20)
		GuiCtrlSetPos($InputAuthor, 365, 60, 340, 20)
		GuiCtrlSetPos($LabelCadence, 720, 60, 60, 20)
		GuiCtrlSetPos($InputCadence, 785, 60, 40, 20)

		GuiCtrlSetPos($ButtonNew, 840, 35, 70, 20)
		GuiCtrlSetPos($ButtonEdit, 915, 35, 70, 20)
		GuiCtrlSetPos($ButtonSave, 840, 60, 70, 20)
		GuiCtrlSetPos($ButtonDelete, 915, 60, 70, 20)

		GuiCtrlSetPos($ButtonCollapse, 5, 675, 275, 20)
		GuiCtrlSetPos($LabelSearch, 285, 675, 70, 20)
		GuiCtrlSetPos($InputSearch, 360, 675, 125, 20)
		GuiCtrlSetPos($CheckboxScroll, 500, 675, 60, 20)
		GuiCtrlSetPos($InputScroll, 565, 675, 45, 20)
		GuiCtrlSetPos($CheckboxOntop, 915, 675, 80, 20)

		GuiCtrlSetPos($LabelChord, 295, 35, 60, 20)
		GuiCtrlSetPos($ButtonChordBold, 355, 35, 20, 20)
		GuiCtrlSetPos($ButtonChordItalic, 380, 35, 20, 20)
		GuiCtrlSetPos($ButtonChordUnderline, 405, 35, 20, 20)
		GuiCtrlSetPos($ButtonChordBlack, 430, 35, 20, 20)
		GuiCtrlSetPos($ButtonChordWhite, 455, 35, 20, 20)
		GuiCtrlSetPos($ButtonChordGreen, 480, 35, 20, 20)
		GuiCtrlSetPos($ButtonChordBlue, 505, 35, 20, 20)
		GuiCtrlSetPos($ButtonChordRed, 530, 35, 20, 20)

		GuiCtrlSetPos($LabelTone, 570, 35, 60, 20)
		GuiCtrlSetPos($ButtonToneFlat, 630, 35, 20, 20)
		GuiCtrlSetPos($ButtonToneDefault, 655, 35, 20, 20)
		GuiCtrlSetPos($ButtonToneSharp, 680, 35, 20, 20)

		GuiCtrlSetPos($LabelLyrics, 295, 60, 60, 20)
		GuiCtrlSetPos($ButtonLyricsBold, 355, 60, 20, 20)
		GuiCtrlSetPos($ButtonLyricsItalic, 380, 60, 20, 20)
		GuiCtrlSetPos($ButtonLyricsUnderline, 405, 60, 20, 20)
		GuiCtrlSetPos($ButtonLyricsBlack, 430, 60, 20, 20)
		GuiCtrlSetPos($ButtonLyricsWhite, 455, 60, 20, 20)
		GuiCtrlSetPos($ButtonLyricsGreen, 480, 60, 20, 20)
		GuiCtrlSetPos($ButtonLyricsBlue, 505, 60, 20, 20)
		GuiCtrlSetPos($ButtonLyricsRed, 530, 60, 20, 20)

		GuiCtrlSetPos($LabelFontSize, 570, 60, 60, 20)
		GuiCtrlSetPos($ButtonFontSizeSmall, 630, 60, 20, 20)
		GuiCtrlSetPos($ButtonFontSizeNormal, 655, 60, 20, 20)
		GuiCtrlSetPos($ButtonFontSizeLarge, 680, 60, 20, 20)

		GuiCtrlSetPos($LabelViewChord, 720, 35, 60, 20)
		GuiCtrlSetPos($ButtonAViewChord, 780, 35, 20, 20)
		GuiCtrlSetPos($ButtonBViewChord, 805, 35, 20, 20)
		GuiCtrlSetPos($ButtonCViewChord, 830, 35, 20, 20)
		GuiCtrlSetPos($ButtonDViewChord, 855, 35, 20, 20)
		GuiCtrlSetPos($ButtonEViewChord, 780, 60, 20, 20)
		GuiCtrlSetPos($ButtonFViewChord, 805, 60, 20, 20)
		GuiCtrlSetPos($ButtonGViewChord, 830, 60, 20, 20)
		GuiCtrlSetPos($ButtonSViewChord, 855, 60, 20, 20)

		GuiCtrlSetPos($ButtonExportLIB_JDB, 295, 35, 130, 20)
		GuiCtrlSetPos($ButtonExportLIB_TXT, 295, 60, 130, 20)
		GuiCtrlSetPos($ButtonPrint, 430, 35, 130, 20)
		GuiCtrlSetPos($ButtonUpdate, 430, 60, 130, 20)
		GuiCtrlSetPos($ButtonSetting, 565, 35, 130, 20)
		GuiCtrlSetPos($ButtonAbout, 565, 60, 130, 20)
	Else
		If (($opt == 800 And $force == 1) Or ($opt == 800 And $CurrentScreen <> $opt)) Then
			WinMove($FormMain, '', 0, 0, 790+6, 550+25)
			GuiCtrlSetPos($Tree, 5, 0, 185, 520)
			GuiCtrlSetPos($Tab, 195, 0, 590, 90)
			_WinAPI_MoveWindow($InputLyrics, 195, 95, 590, 425)
			_GUICtrlRichEdit_SetRECT($InputLyrics, 10, 10)

			GuiCtrlSetPos($LabelTitle, 205, 35, 60, 20)
			GuiCtrlSetPos($InputTitle, 260, 35, 370, 20)
			GuiCtrlSetPos($LabelAuthor, 205, 60, 60, 20)
			GuiCtrlSetPos($InputAuthor, 260, 60, 250, 20)
			GuiCtrlSetPos($LabelCadence, 525, 60, 60, 20)
			GuiCtrlSetPos($InputCadence, 590, 60, 40, 20)

			GuiCtrlSetPos($ButtonNew, 645, 35, 60, 20)
			GuiCtrlSetPos($ButtonEdit, 710, 35, 60, 20)
			GuiCtrlSetPos($ButtonSave, 645, 60, 60, 20)
			GuiCtrlSetPos($ButtonDelete, 710, 60, 60, 20)

			GuiCtrlSetPos($ButtonCollapse, 5, 525, 185, 20)
			GuiCtrlSetPos($LabelSearch, 195, 525, 80, 20)
			GuiCtrlSetPos($InputSearch, 280, 525, 100, 20)
			GuiCtrlSetPos($CheckboxScroll, 395, 525, 60, 20)
			GuiCtrlSetPos($InputScroll, 460, 525, 45, 20)
			GuiCtrlSetPos($CheckboxOntop, 705, 525, 80, 15)

			GuiCtrlSetPos($LabelChord, 205, 35, 60, 20)
			GuiCtrlSetPos($ButtonChordBold, 265, 35, 20, 20)
			GuiCtrlSetPos($ButtonChordItalic, 290, 35, 20, 20)
			GuiCtrlSetPos($ButtonChordUnderline, 315, 35, 20, 20)
			GuiCtrlSetPos($ButtonChordBlack, 340, 35, 20, 20)
			GuiCtrlSetPos($ButtonChordWhite, 365, 35, 20, 20)
			GuiCtrlSetPos($ButtonChordGreen, 415, 35, 20, 20)
			GuiCtrlSetPos($ButtonChordBlue, 440, 35, 20, 20)
			GuiCtrlSetPos($ButtonChordRed, 390, 35, 20, 20)

			GuiCtrlSetPos($LabelTone, 475, 35, 60, 20)
			GuiCtrlSetPos($ButtonToneFlat, 535, 35, 20, 20)
			GuiCtrlSetPos($ButtonToneDefault, 560, 35, 20, 20)
			GuiCtrlSetPos($ButtonToneSharp, 585, 35, 20, 20)

			GuiCtrlSetPos($LabelLyrics, 205, 60, 60, 20)
			GuiCtrlSetPos($ButtonLyricsBold, 265, 60, 20, 20)
			GuiCtrlSetPos($ButtonLyricsItalic, 290, 60, 20, 20)
			GuiCtrlSetPos($ButtonLyricsUnderline, 315, 60, 20, 20)
			GuiCtrlSetPos($ButtonLyricsBlack, 340, 60, 20, 20)
			GuiCtrlSetPos($ButtonLyricsWhite, 365, 60, 20, 20)
			GuiCtrlSetPos($ButtonLyricsGreen, 415, 60, 20, 20)
			GuiCtrlSetPos($ButtonLyricsBlue, 440, 60, 20, 20)
			GuiCtrlSetPos($ButtonLyricsRed, 390, 60, 20, 20)

			GuiCtrlSetPos($LabelFontSize, 475, 60, 60, 20)
			GuiCtrlSetPos($ButtonFontSizeSmall, 535, 60, 20, 20)
			GuiCtrlSetPos($ButtonFontSizeNormal, 560, 60, 20, 20)
			GuiCtrlSetPos($ButtonFontSizeLarge, 585, 60, 20, 20)

			GuiCtrlSetPos($LabelViewChord, 620, 35, 60, 20)
			GuiCtrlSetPos($ButtonAViewChord, 680, 35, 20, 20)
			GuiCtrlSetPos($ButtonBViewChord, 705, 35, 20, 20)
			GuiCtrlSetPos($ButtonCViewChord, 730, 35, 20, 20)
			GuiCtrlSetPos($ButtonDViewChord, 755, 35, 20, 20)
			GuiCtrlSetPos($ButtonEViewChord, 680, 60, 20, 20)
			GuiCtrlSetPos($ButtonFViewChord, 705, 60, 20, 20)
			GuiCtrlSetPos($ButtonGViewChord, 730, 60, 20, 20)
			GuiCtrlSetPos($ButtonSViewChord, 755, 60, 20, 20)

			GuiCtrlSetPos($ButtonExportLIB_JDB, 205, 35, 130, 20)
			GuiCtrlSetPos($ButtonExportLIB_TXT, 205, 60, 130, 20)
			GuiCtrlSetPos($ButtonPrint, 340, 35, 130, 20)
			GuiCtrlSetPos($ButtonUpdate, 340, 60, 130, 20)
			GuiCtrlSetPos($ButtonSetting, 475, 35, 130, 20)
			GuiCtrlSetPos($ButtonAbout, 475, 60, 130, 20)
		EndIf
	EndIf
ENDFUNC
#ENDREGION

#REGION ### FUNCTIONS BLOCK - RICH EDIT ###
FUNC JT_RichEditScroll($richEdit, $startScroll)
	Local $ScrollPos
	If ($startScroll == 1) Then
		$ScrollPos = _GUICtrlRichEdit_GetScrollPos($richEdit)
		Sleep(Int(1000/GUICtrlRead($InputScroll)))
		_GUICtrlRichEdit_SetScrollPos($richEdit, 0, $ScrollPos[1]+1)
	EndIf
ENDFUNC ; <== JT_RichEditScroll

FUNC JT_RichEditSetText($richEdit, $string)
	_GUICtrlRichEdit_SetText($richEdit, @CRLF & $string)
	_GUICtrlRichEdit_SetSel($richEdit, 0, -1)
	_GUICtrlRichEdit_SetFont($richEdit, IniRead($ConfigFile, 'ViewLyrics', 'FontSize', 10), IniRead($ConfigFile, 'ViewLyrics', 'FontName', 'Courier New'))
	_GUICtrlRichEdit_GotoCharPos($richEdit ,0)
ENDFUNC ; <== JT_RichEditSetText

FUNC JT_RichEditSetStyle($richEdit, $mode, $value)
	Local $GetFont = _GUICtrlRichEdit_GetFont($richEdit)
	Local $GetText = _GUICtrlRichEdit_GetText($richEdit)
	Local $CurrentPos[2] = [0, 0]
	Local $tmp
	
	Switch ($mode)
		Case 'FontSize'
			_GUICtrlRichEdit_SetSel($richEdit, 0, -1)
			If ($value == 0) Then
				_GUICtrlRichEdit_SetFont($richEdit, IniRead($ConfigFile, 'ViewLyrics', 'FontSize', 10))
			Else
				_GUICtrlRichEdit_SetFont($richEdit, $GetFont[0]+$value)
			EndIf
		Case 'ChordTuner'
			While (1)
				$tmp = _GUICtrlRichEdit_FindTextInRange($richEdit, '[', $CurrentPos[0], -1)
				If ($tmp[0] >= 0) Then
					$CurrentPos[0] = $tmp[0]
				Else
					ExitLoop(1)
				EndIf
				
				$tmp = _GUICtrlRichEdit_FindTextInRange($richEdit, ']', $CurrentPos[0], -1)
				If ($tmp[0] >= 0) Then
					$CurrentPos[1] = $tmp[0] + 1
				Else
					ExitLoop(1)
				EndIf
				
				_GUICtrlRichEdit_SetSel($richEdit, $CurrentPos[0] + 1, $CurrentPos[1] - 1)
				$tmp = _JT_ChordTuner(_GUICtrlRichEdit_GetSelText($richEdit), $value)
				_GUICtrlRichEdit_ReplaceText($richEdit, $tmp)
				
				$CurrentPos[0] = $CurrentPos[0] + StringLen($tmp)
			WEnd
		Case 'ChordStyle'
			While (1)
				$tmp = _GUICtrlRichEdit_FindTextInRange($richEdit, '[', $CurrentPos[0], -1)
				If ($tmp[0] >= 0) Then
					$CurrentPos[0] = $tmp[0]
				Else
					ExitLoop(1)
				EndIf
				
				$tmp = _GUICtrlRichEdit_FindTextInRange($richEdit, ']', $CurrentPos[0], -1)
				If ($tmp[0] >= 0) Then
					$CurrentPos[1] = $tmp[0] + 1
				Else
					ExitLoop(1)
				EndIf
				
				_GUICtrlRichEdit_SetSel($richEdit, $CurrentPos[0], $CurrentPos[1])
				$CurrentPos[0] = $CurrentPos[1]
				
				Switch ($value)
					Case 'bo', 'it', 'un'
						If (StringInStr(_GUICtrlRichEdit_GetCharAttributes($richEdit), $value) == 0) Then
							_GUICtrlRichEdit_SetCharAttributes($richEdit, '+' & $value)
						Else
							_GUICtrlRichEdit_SetCharAttributes($richEdit, '-' & $value)
						EndIf
					Case Else
						If (StringInStr($value, 'x') > 0) Then
							If (StringInStr($value, 'x') > 3) Then
								_GUICtrlRichEdit_SetCharAttributes($richEdit, StringLeft($value, StringInStr($value, 'x')-2))
							EndIf
							_GUICtrlRichEdit_SetCharColor($richEdit, StringMid($value, StringInStr($value, 'x')-1))
						Else
							_GUICtrlRichEdit_SetCharAttributes($richEdit, $value)
						EndIf
				EndSwitch
			WEnd
		Case 'LyricsStyle'
			While (1)
				$tmp = _GUICtrlRichEdit_FindTextInRange($richEdit, '[', $CurrentPos[0], -1)
				$CurrentPos[1] = $tmp[0]
				
				_GUICtrlRichEdit_SetSel($richEdit, $CurrentPos[0], $CurrentPos[1])
				
				Switch ($value)
					Case 'bo', 'it', 'un'
						If (StringInStr(_GUICtrlRichEdit_GetCharAttributes($richEdit), $value) == 0) Then
							_GUICtrlRichEdit_SetCharAttributes($richEdit, '+' & $value)
						Else
							_GUICtrlRichEdit_SetCharAttributes($richEdit, '-' & $value)
						EndIf
					Case Else
						If (StringInStr($value, 'x') > 0) Then
							If (StringInStr($value, 'x') > 3) Then
								_GUICtrlRichEdit_SetCharAttributes($richEdit, StringLeft($value, StringInStr($value, 'x')-2))
							EndIf
							If ($CurrentPos[0] <> $CurrentPos[1]) Then _GUICtrlRichEdit_SetCharColor($richEdit, StringMid($value, StringInStr($value, 'x')-1))
						Else
							_GUICtrlRichEdit_SetCharAttributes($richEdit, $value)
						EndIf
				EndSwitch
				
				$tmp = _GUICtrlRichEdit_FindTextInRange($richEdit, ']', $CurrentPos[1], -1)
				If ($tmp[0] > 0) Then
					$CurrentPos[0] = $tmp[0] + 1
				Else
					If ($CurrentPos[1] > 0) Then
						_GUICtrlRichEdit_SetSel($richEdit, $CurrentPos[1], -1)
						Switch ($value)
							Case 'bo', 'it', 'un'
								If (StringInStr(_GUICtrlRichEdit_GetCharAttributes($richEdit), $value) == 0) Then
									_GUICtrlRichEdit_SetCharAttributes($richEdit, '+' & $value)
								Else
									_GUICtrlRichEdit_SetCharAttributes($richEdit, '-' & $value)
								EndIf
							Case Else
								If (StringInStr($value, 'x') > 0) Then
									If (StringInStr($value, 'x') > 3) Then
										_GUICtrlRichEdit_SetCharAttributes($richEdit, StringLeft($value, StringInStr($value, 'x')-2))
									EndIf
									_GUICtrlRichEdit_SetCharColor($richEdit, StringMid($value, StringInStr($value, 'x')-1))
								Else
									_GUICtrlRichEdit_SetCharAttributes($richEdit, $value)
								EndIf
						EndSwitch
					EndIf
					ExitLoop(1)
				EndIf
			WEnd
	EndSwitch
	
	_GUICtrlRichEdit_GotoCharPos($richEdit ,0)
ENDFUNC ; <== JT_RichEditSetStyle
#ENDREGION

#REGION ### FUNCTIONS BLOCK ###
FUNC JT_INIT()
	If (Not FileExists($DB_Location)) Then
		Local $DB_Object = ObjCreate('DAO.DBEngine.36')
		$DB_Object.CreateDatabase($DB_Location, ';LANGID=0x0409;CP=1252;COUNTRY=0')
		Local $SongDB = $DB_Object.OpenDatabase($DB_Location, 0, 0)
		$SongDB.Execute('CREATE TABLE ' & $SongDB_Table & ' (ID AUTOINCREMENT, ' & _
					$SongDB_ColTitle & ' Text, ' & _
					$SongDB_ColAuthor & ' Text, ' & _
					$SongDB_ColCadence & ' Text, ' & _
					$SongDB_ColLyrics & ' Memo, ' & _
					'PRIMARY KEY (ID))')
		Local $Init[4] = ['[JT] Songbooks', 'o0johntam0o', '4/4', 'Initial']
		JT_AddSong($Init)
		$SongDB.Close
		Sleep(1000)
	EndIf
ENDFUNC ; <== JT_INIT

; $data (Required)				Array[<NewTitle>, <NewAuthor>, <NewCadence>, <NewLyrics>]
; $oldData						<OldTitle>
;
; Return						Result by string
FUNC JT_AddSong(ByRef $data, $oldData = '')
	$oldData = StringReplace(StringStripWS($oldData, 7), '"', "'")
	$data[0] = StringReplace(StringStripWS($data[0], 7), '"', "'")
	$data[1] = StringReplace(StringStripWS($data[1], 7), '"', "'")
	$data[2] = StringReplace(StringStripWS($data[2], 7), '"', "'")
	$data[3] = StringReplace(StringStripWS($data[3], 3), '"', "'")
	
	If (StringLen($data[0]) == 0) Then Return JT_TranSlate('Missing title')
	If (StringLen($data[1]) == 0) Then $data[1] = JT_TranSlate('Unknown')
	If (StringLen($data[2]) == 0) Then $data[2] = '0/0'
	
	Local $SongDB = $DB_Object.OpenDatabase($DB_Location, 0, 0)
	Local $RecordPointer = $SongDB.OpenRecordSet('SELECT ' & $SongDB_ColTitle & _
												' FROM ' & $SongDB_Table & _
												' WHERE ' & $SongDB_ColTitle & ' = "' & $data[0] & '"')

	If (StringLen($oldData) == 0) Then
		If ($RecordPointer.EOF <> -1 Or $RecordPointer.BOF <> -1) Then
			$SongDB.Close
			Return JT_TranSlate('The song "%s" is already exists!', $data[0])
			;$data[0] = $data[0] & ' - ' & $data[1] & ' (' & Random() & ')'
		EndIf
		$SongDB.Execute('INSERT INTO ' & $SongDB_Table & ' (' & _
			$SongDB_ColTitle & ', ' & $SongDB_ColAuthor & ', ' & $SongDB_ColCadence & ', ' & $SongDB_ColLyrics & _
			') VALUES ("' & $data[0] & '", "' & $data[1] & '", "' & $data[2] & '", "' & $data[3] & '")')
	Else
		If ($data[0] <> $oldData And ($RecordPointer.EOF <> -1 Or $RecordPointer.BOF <> -1)) Then
			$SongDB.Close
			Return JT_TranSlate('The song "%s" is already exists!', $data[0])
			;$data[0] = $data[0] & ' - ' & $data[1] & ' (' & Random() & ')'
		EndIf
		$SongDB.Execute('UPDATE ' & $SongDB_Table & ' SET ' & _
						$SongDB_ColTitle & ' = "' & $data[0] & '", ' & _
						$SongDB_ColAuthor & ' = "' & $data[1] & '", ' & _
						$SongDB_ColCadence & ' = "' & $data[2] & '", ' & _
						$SongDB_ColLyrics & ' = "' & $data[3] & _
						'" WHERE ' & $SongDB_ColTitle & ' = "' & $oldData & '"')
	EndIf

	$SongDB.Close
	
	Return ''
ENDFUNC ; <== JT_AddSong

FUNC JT_DelSong($data)
	Local $SongDB = $DB_Object.OpenDatabase($DB_Location, 0, 0)
	$SongDB.Execute('DELETE FROM ' & $SongDB_Table & ' WHERE ' & $SongDB_ColTitle & ' = "' & $data & '"')
	$SongDB.Close
	Return 1
ENDFUNC ; <== JT_DelSong

FUNC JT_BuildTree($rebuild = 0)
	Local $i, $SongLists = _JT_GetRecordLists($DB_Object, $DB_Location, $SongDB_Table, $SongDB_ColTitle, 0)
	
	If ($rebuild == 1) Then
		; Remove old items
		For $i = 65 To 90
			GUICtrlDelete(Eval('TreeItem' & Chr($i)))
		Next
		GUICtrlDelete($TreeItemOther)
		; Add new items
		For $i = 65 To 90
			Assign('TreeItem' & Chr($i), GUICtrlCreateTreeViewItem(Chr($i), $Tree), 2)
			GUICtrlSetColor(-1, 0x0000C0)
		Next
		$TreeItemOther = GUICtrlCreateTreeViewItem(JT_TranSlate('Other'), $Tree)
		GUICtrlSetColor(-1, 0x0000C0)
	EndIf

	If (IsArray($SongLists)) Then
		For $i = 0 To UBound($SongLists) - 1
			If (GUICtrlCreateTreeViewItem($SongLists[$i], Eval('TreeItem' & StringUpper(_JT_ToLatin(StringLeft(StringStripWS($SongLists[$i], 7), 1))))) == 0) Then
				GUICtrlCreateTreeViewItem($SongLists[$i], $TreeItemOther)
			EndIf
		Next
	EndIf
ENDFUNC ; <== JT_BuildTree

FUNC JT_GetDataFromTree($force = 0)
	If ($CurrentTreeItem <> GUICtrlRead($Tree) Or $force == 1) Then
		If ($UpdateMode <> 0) Then
			If (JT_MessageBox($FormMain, 'Confirm', "The song haven't saved yet. Do you want to leave it?") == 6) Then
				$UpdateMode = 0
			Else
				GUICtrlSetState($CurrentTreeItem, $GUI_FOCUS)
			EndIf
		Else
			$CurrentTreeItem = GUICtrlRead($Tree)
			$ToneChange = 0
			$UpdateMode = 0
			$CurrentSongInfo = JT_GetSongInfo(GUICtrlRead($CurrentTreeItem, 1))
			If (IsArray($CurrentSongInfo)) Then
				GUICtrlSetData($InputTitle, $CurrentSongInfo[0])
				GUICtrlSetData($InputAuthor, $CurrentSongInfo[1])
				GUICtrlSetData($InputCadence, $CurrentSongInfo[2])
				JT_RichEditSetText($InputLyrics, $CurrentSongInfo[3])

				; Load settings
				$tmpString = ''
				If (IniRead($ConfigFile, 'ViewLyrics', 'ChordBold', 0) == 1) Then $tmpString = '+bo'
				If (IniRead($ConfigFile, 'ViewLyrics', 'ChordItalic', 0) == 1) Then $tmpString = $tmpString & '+it'
				If (IniRead($ConfigFile, 'ViewLyrics', 'ChordUnderline', 0) == 1) Then $tmpString = $tmpString & '+un'
				If (IniRead($ConfigFile, 'ViewLyrics', 'ChordColor', '0x000000') <> '0x000000') Then $tmpString = $tmpString & IniRead($ConfigFile, 'ViewLyrics', 'ChordColor', '0x000000')
				If (StringLen($tmpString) > 0) Then JT_RichEditSetStyle($InputLyrics, 'ChordStyle', $tmpString)
				$tmpString = ''
				If (IniRead($ConfigFile, 'ViewLyrics', 'LyricsBold', 0) == 1) Then $tmpString = '+bo'
				If (IniRead($ConfigFile, 'ViewLyrics', 'LyricsItalic', 0) == 1) Then $tmpString = $tmpString & '+it'
				If (IniRead($ConfigFile, 'ViewLyrics', 'LyricsUnderline', 0) == 1) Then $tmpString = $tmpString & '+un'
				If (IniRead($ConfigFile, 'ViewLyrics', 'LyricsColor', '0x000000') <> '0x000000') Then $tmpString = $tmpString & IniRead($ConfigFile, 'ViewLyrics', 'LyricsColor', '0x000000')
				If (StringLen($tmpString) > 0) Then JT_RichEditSetStyle($InputLyrics, 'LyricsStyle', $tmpString)
				
				GUICtrlSetStyle($InputTitle, $ES_READONLY)
				GUICtrlSetStyle($InputAuthor, $ES_READONLY)
				GUICtrlSetStyle($InputCadence, $ES_READONLY)
				_GUICtrlRichEdit_SetReadOnly($InputLyrics, True)
				ControlFocus($FormMain, '', $InputSearch)
			EndIf
		EndIf
	EndIf
ENDFUNC ; <== JT_GetDataFromTree

FUNC JT_PrintLyrics($iData)
	GUISetState(@SW_DISABLE, $FormMain)
	SplashTextOn(JT_TranSlate('Print'), JT_TranSlate('Working'), 300, 100, -1, -1, 0, 'Tahoma')
	
	If (Not IsArray($iData)) Then
		SplashOff()
		GUISetState(@SW_ENABLE, $FormMain)
		Return 0
	EndIf
	
	If (StringLen($iData[0]) > 0 And StringLen($iData[3]) > 0) Then
		Local $PrintPath = @HomeDrive & '\' & StringReplace(_JT_ToLatin(StringStripWS($iData[0], 7)), ' ', '-') & '.html'
		Local $tmp = JT_GetTemplate('Print')
		If ($tmp == '') Then
			SplashOff()
			GUISetState(@SW_ENABLE, $FormMain)
			Return 0
		EndIf
		Local $tmp2 = FileOpen($PrintPath, 2+8+256)
		$tmp = StringReplace($tmp, '{_ARTIST_}', JT_TranSlate('Artist'), 1)
		$tmp = StringReplace($tmp, '{_AUTHOR_}', JT_TranSlate('Author'), 1)
		$tmp = StringReplace($tmp, '{_RHYTHM_}', JT_TranSlate('Rhythm'), 1)
		$tmp = StringReplace($tmp, '{_CADENCE_}', JT_TranSlate('Cadence'), 1)
		$tmp = StringReplace($tmp, '<!-- SONG TITLE - BEGIN --><!-- SONG TITLE - END -->', StringStripWS($iData[0], 7), 1)
		$tmp = StringReplace($tmp, '<!-- SONG AUTHOR - BEGIN --><!-- SONG AUTHOR - END -->', StringStripWS($iData[1], 7), 1)
		$tmp = StringReplace($tmp, '<!-- SONG CADENCE - BEGIN --><!-- SONG CADENCE - END -->', StringStripWS($iData[2], 7), 1)
		$tmp = StringReplace($tmp, '<!-- SONG LYRICS - BEGIN --><!-- SONG LYRICS - END -->', StringStripWS($iData[3], 3), 1)
		If (FileWrite($tmp2, $tmp)) Then
			Local $OldRegH = RegRead('HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup', 'header')
			Local $OldRegF = RegRead('HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup', 'footer')
			RegWrite('HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup', 'header', 'REG_SZ', '')
			RegWrite('HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup', 'footer', 'REG_SZ', '')
			RunWait(@ComSpec & ' /c rundll32.exe ' & @SystemDir & '\mshtml.dll,PrintHTML "' & $PrintPath, @TempDir, @SW_HIDE)
			RegWrite('HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup', 'header', 'REG_SZ', $OldRegH)
			RegWrite('HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\PageSetup', 'footer', 'REG_SZ', $OldRegF)
			FileClose($tmp2)
			FileDelete($PrintPath)
			SplashOff()
			GUISetState(@SW_ENABLE, $FormMain)
			Return 1
		Else
			FileClose($tmp2)
			SplashOff()
			GUISetState(@SW_ENABLE, $FormMain)
			Return 0
		EndIf
	EndIf
	SplashOff()
	GUISetState(@SW_ENABLE, $FormMain)
	Return 0
ENDFUNC ; <== JT_PrintLyrics

FUNC JT_ExportLIB_JDB($_Path, $_Version)
	If (StringLen($_Path) == 0) Then
		Return 0
	EndIf
	
	GUISetState(@SW_DISABLE, $FormMain)
	SplashTextOn(JT_TranSlate('Export libraries (JDB)'), JT_TranSlate('Working'), 300, 100, -1, -1, 0, 'Tahoma')
	
	Local $SongDB = $DB_Object.OpenDatabase($DB_Location, 0, 1)
	Local $RecordPointer = $SongDB.OpenRecordSet('SELECT * FROM ' & $SongDB_Table)
	Local $FileOpen, $FileWrite = '<!-- DATABASE VERSION - BEGIN -->' & $_Version & '<!-- DATABASE VERSION - END -->' & @CRLF
	Local $FirstSong = 1

	If ($RecordPointer.EOF <> -1 Or $RecordPointer.BOF <> -1) Then
		$RecordPointer.MoveFirst
		While ($RecordPointer.EOF <> -1)
			If ($FirstSong == 1) Then
				$FirstSong = 0
			Else
				$FileWrite = $FileWrite & @CRLF & '<!-- ===================== SONGBOOKS - BREAKER ===================== -->'
			EndIf
			$FileWrite = $FileWrite & @CRLF & '<!-- SONG TITLE - BEGIN -->' & $RecordPointer.Fields(1).Value & '<!-- SONG TITLE - END -->'
			$FileWrite = $FileWrite & @CRLF & '<!-- SONG AUTHOR - BEGIN -->' & $RecordPointer.Fields(2).Value & '<!-- SONG AUTHOR - END -->'
			$FileWrite = $FileWrite & @CRLF & '<!-- SONG ARTIST - BEGIN --><!-- SONG ARTIST - END -->'
			$FileWrite = $FileWrite & @CRLF & '<!-- SONG CADENCE - BEGIN -->' & $RecordPointer.Fields(3).Value & '<!-- SONG CADENCE - END -->'
			$FileWrite = $FileWrite & @CRLF & '<!-- SONG RHYTHM - BEGIN --><!-- SONG RHYTHM - END -->'
			$FileWrite = $FileWrite & @CRLF & '<!-- SONG LYRICS - BEGIN -->' & $RecordPointer.Fields(4).Value & '<!-- SONG LYRICS - END -->'
			$RecordPointer.MoveNext
		WEnd
	Else
		$SongDB.Close
		SplashOff()
		GUISetState(@SW_ENABLE, $FormMain)
		Return 0
	EndIf
	
	$SongDB.Close
	
	$FileOpen = FileOpen($_Path, 2+8+256)
	If ($FileOpen == -1) Then
		FileClose($FileOpen)
		SplashOff()
		GUISetState(@SW_ENABLE, $FormMain)
		Return 0
	EndIf
	If (FileWrite($FileOpen, $FileWrite) == 0) Then
		FileClose($FileOpen)
		SplashOff()
		GUISetState(@SW_ENABLE, $FormMain)
		Return 0
	EndIf
	FileClose($FileOpen)
	
	SplashOff()
	GUISetState(@SW_ENABLE, $FormMain)
	Return 1
ENDFUNC ; <== JT_ExportLIB_JDB

FUNC JT_ExportLIB_TXT($_Path)
	If ($_Path == '') Then
		Return 0
	EndIf
	
	GUISetState(@SW_DISABLE, $FormMain)
	SplashTextOn(JT_TranSlate('Export libraries (TXT)'), JT_TranSlate('Working'), 300, 100, -1, -1, 0, 'Tahoma')
	
	$_Path = $_Path & '\My Songbooks'
	
	; Start export
	If (Not FileExists($_Path)) Then DirCreate($_Path)
	Local $SongDB = $DB_Object.OpenDatabase($DB_Location, 0, 1)
	Local $RecordPointer = $SongDB.OpenRecordSet('SELECT * FROM ' & $SongDB_Table)
	Local $FileOpen, $FileWrite, $error = 0
	
	If ($RecordPointer.EOF <> -1 Or $RecordPointer.BOF <> -1) Then
		$RecordPointer.MoveFirst
		While ($RecordPointer.EOF <> -1)
			$FileOpen = FileOpen($_Path & '\' & $RecordPointer.Fields(1).Value & '.txt', 2+8+256)
			$FileWrite = ''
			$FileWrite = $FileWrite & JT_TranSlate('Author') & ': ' & $RecordPointer.Fields(2).Value & @CRLF
			$FileWrite = $FileWrite & JT_TranSlate('Cadence') & ': ' & $RecordPointer.Fields(3).Value & @CRLF & @CRLF
			$FileWrite = $FileWrite & $RecordPointer.Fields(4).Value
			If (FileWrite($FileOpen, $FileWrite) == 0) Then
				$error = 1
			EndIf
			FileClose($FileOpen)
			$RecordPointer.MoveNext
		WEnd
	Else
		$error = 1
	EndIf
	
	$SongDB.Close
	
	If ($error == 0) Then
		SplashOff()
		GUISetState(@SW_ENABLE, $FormMain)
		Return 1
	Else
		If (FileExists($_Path)) Then
			DirRemove($_Path, 1)
		EndIf
		SplashOff()
		GUISetState(@SW_ENABLE, $FormMain)
		Return 0
	EndIf
ENDFUNC ; <== JT_ExportLIB_TXT

FUNC JT_GetSongInfo($dbFilter)
	Local $Re[4] = ['', '', '', ''], $Title, $Author, $Cadence, $Lyrics
	Local $SongDB = $DB_Object.OpenDatabase($DB_Location, 0, 1)
	Local $RecordPointer = $SongDB.OpenRecordSet('SELECT * FROM ' & $SongDB_Table & ' WHERE ' & $SongDB_ColTitle & ' = "' & $dbFilter & '"')
	
	If $RecordPointer.EOF <> -1 Or $RecordPointer.BOF <> -1 Then
		$RecordPointer.MoveFirst
		$Re[0] = StringStripWS($RecordPointer.Fields(1).Value, 7)
		$Re[1] = StringStripWS($RecordPointer.Fields(2).Value, 7)
		$Re[2] = StringStripWS($RecordPointer.Fields(3).Value, 7)
		$Re[3] = StringStripWS($RecordPointer.Fields(4).Value, 3)
	EndIf
	
	$SongDB.Close
	
	Return $Re
ENDFUNC ; <== JT_GetSongInfo

FUNC JT_GetTemplate($mode='')
	Switch ($mode)
		Case ''
			Return ''
		Case 'Print'
			Return '<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">' & @CRLF & _
					'<html xmlns="http://www.w3.org/1999/xhtml">' & @CRLF & _
						'<head>' & @CRLF & _
							'<meta http-equiv="content-type" content="text/html; charset=UTF-8" />' & @CRLF & _
							'<title></title>' & @CRLF & _
							'<style type="text/css">' & @CRLF & _
							'/* <![CDATA[ */' & @CRLF & _
							'@media print, screen' & @CRLF & _
							'{' & @CRLF & _
								'body { margin: 20px; padding: 0; font-size: 1em; }' & @CRLF & _
								'#SongInfo { font-family: "Times New Roman", Tahoma; text-align: center; margin: 0 auto; }' & @CRLF & _
								'#SongInfo h1 { margin: 0; padding: 0; }' & @CRLF & _
								'#SongInfo h2, h3 { margin: 0; padding: 0; color: #696969; font-style: italic; }' & @CRLF & _
								'#LyricsPane { padding: 10px; font-family: "Courier New", Tahoma; white-space: pre-wrap; white-space: -moz-pre-wrap; white-space: -pre-wrap; white-space: -o-pre-wrap; word-wrap: break-word; }' & @CRLF & _
							'}' & @CRLF & _
							'/* ]]> */' & @CRLF & _
							'</style>' & @CRLF & _
						'</head>' & @CRLF & _
						'<body>' & @CRLF & _
							'<div id="SongInfo">' & @CRLF & _
								'<h1><!-- SONG TITLE - BEGIN --><!-- SONG TITLE - END --></h1>' & @CRLF & _
								'<h2>({_AUTHOR_}: <!-- SONG AUTHOR - BEGIN --><!-- SONG AUTHOR - END -->)</h2>' & @CRLF & _
								'<h3>({_CADENCE_}: <!-- SONG CADENCE - BEGIN --><!-- SONG CADENCE - END -->)</h3>' & @CRLF & _
							'</div>' & @CRLF & _
							'<div style="clear:both; height: 10px;"></div>' & @CRLF & _
							'<pre id="LyricsPane"><!-- SONG LYRICS - BEGIN --><!-- SONG LYRICS - END --></pre>' & @CRLF & _
						'</body>' & @CRLF & _
					'</html>'
	EndSwitch
	Return ''
ENDFUNC ; <== JT_GetTemplate

FUNC JT_SearchSong($keywords)
	If ($keywords == '' Or StringLen(StringStripWS($keywords, 8)) < 4) Then Return
	Local $Search = StringUpper(_JT_ToLatin(StringStripWS($keywords, 8)))
	Local $Node = '#26', $NodeCount, $i, $a, $b
	
	If (Asc(StringLeft($Search, 1)) >= 65 And Asc(StringLeft($Search, 1)) <= 90) Then
		$Node = '#' & (Asc(StringLeft($Search, 1)) - 65)
	EndIf
	
	$NodeCount = ControlTreeView($FormMain, '', $Tree, 'GetItemCount', $Node)
	
	For $i = 0 To $NodeCount - 1
		$a = ControlTreeView($FormMain, '', $Tree, 'GetText', $Node & '|#' & $i)
		$b = StringUpper(_JT_ToLatin(StringStripWS($a, 8)))
		If (StringLeft($b, StringLen($Search)) == $Search) Then
			ControlTreeView($FormMain, '', $Tree, 'Select', $Node & '|#' & $i)
			ExitLoop(1)
		EndIf
	Next
ENDFUNC ; <== JT_SearchSong

FUNC JT_TranSlate($string='', $subString1='', $subString2='', $subString3='', $subString4='')
	If ($string == '') Then Return $string
	
	Local $Lang = $string, $i
	Local $LangFile = _JT_GetFileContent($LanguageDir & '\' & IniRead($ConfigFile, 'Common', 'CustomLanguage', 'English') & '.ini')
	Local $LangMap = StringSplit($LangFile, @CRLF, 2)
	If (IniRead($ConfigFile, 'Common', 'CustomLanguage', 'English') <> 'English') Then
		For $i In $LangMap
			If (StringLeft($i, StringInStr($i, '=')-1) == $Lang) Then
				$Lang = StringMid($i, StringInStr($i, '=') + 1)
				ExitLoop(1)
			EndIf
		Next
	EndIf

	; Not found in $LangFile... return original string
	If ($Lang == '' And $string <> '') Then
		$Lang = $string
	EndIf

	Select
		Case ($subString4 <> '')
			Return StringFormat($Lang, $subString1, $subString2, $subString3, $subString4)
		Case ($subString3 <> '')
			Return StringFormat($Lang, $subString1, $subString2, $subString3)
		Case ($subString2 <> '')
			Return StringFormat($Lang, $subString1, $subString2)
		Case ($subString1 <> '')
			Return StringFormat($Lang, $subString1)
		Case Else
			Return $Lang
	EndSelect
ENDFUNC ; <== JT_TranSlate

; $gui				Return message to which form?
; $type				'Confirm', 'Error', 'Info', 'Warning'
; $string			Message text
FUNC JT_MessageBox($gui, $type='', $string='', $subString1='', $subString2='', $subString3='', $subString4='')
	Local $Code = 0, $Re = 0
	GUISetState(@SW_DISABLE, $gui)
	Switch ($type)
		Case 'Confirm'
			$Code = 32 + 262144 + 4
		Case 'Error'
			$Code = 16 + 262144
		Case 'Info'
			$Code = 64 + 262144
		Case 'Warning'
			$Code = 48 + 262144 + 4
	EndSwitch
	Select
		Case ($subString4 <> '')
			$Re = MsgBox($Code, JT_TranSlate($type), JT_TranSlate($string, $subString1, $subString2, $subString3, $subString4), 0, $gui)
		Case ($subString3 <> '')
			$Re = MsgBox($Code, JT_TranSlate($type), JT_TranSlate($string, $subString1, $subString2, $subString3), 0, $gui)
		Case ($subString2 <> '')
			$Re = MsgBox($Code, JT_TranSlate($type), JT_TranSlate($string, $subString1, $subString2), 0, $gui)
		Case ($subString1 <> '')
			$Re = MsgBox($Code, JT_TranSlate($type), JT_TranSlate($string, $subString1), 0, $gui)
		Case Else
			$Re = MsgBox($Code, JT_TranSlate($type), JT_TranSlate($string), 0, $gui)
	EndSelect
	
	GUISetState(@SW_ENABLE, $gui)
	Return $Re
ENDFUNC ; <== JT_MessageBox

FUNC JT_ViewChordPre()
	If (WinActive($FormMain) <> 0) Then
		JT_ViewChord(_GUICtrlRichEdit_GetSelText($InputLyrics))
	EndIf
ENDFUNC ; <== JT_ViewChordPre

FUNC JT_AutoBracketPre()
	If (WinActive($FormMain) <> 0) Then _JT_AutoBracket(_GUICtrlRichEdit_GetSelText($InputLyrics))
ENDFUNC ; <== JT_ViewChordPre

FUNC JT_ExitApp()
	_GUICtrlRichEdit_Destroy($InputLyrics)
	GuiDelete()
	Exit
ENDFUNC ; <== JT_ExitApp
#ENDREGION

#REGION ### FUNCTIONS BLOCK - CHILD GUI ###
FUNC JT_ViewChord($iChord = '')
	If ($iChord == -1) Then
		JT_MessageBox($FormMain, 'Error', 'Please mark chord name first')
		Return 0
	EndIf
	
	Local $InputChord = StringStripWS($iChord, 7)
	
	If ($InputChord == '') Then
		JT_MessageBox($FormMain, 'Error', 'Please mark chord name first')
		Return 0
	EndIf
	
	; Make form
	Local $FormViewChord = GUICreate(JT_TranSlate('View Chord'), 335, 185, 500, 200, $GUI_SS_DEFAULT_GUI, $WS_EX_MDICHILD, $FormMain)
	GUISetFont(9, 400, 0, 'Tahoma', $FormViewChord)
	Local $CurrentChord = GUICtrlCreateInput($InputChord, 10, 10, 315, 25, $SS_CENTER)
	GUICtrlSetFont(-1, 15, 800, 0, 'Times New Roman')
	Local $ModChord = GUICtrlCreateButton(JT_TranSlate('Add') & '/' & JT_TranSlate('Edit'), 10, 45, 155, 20)
	Local $Go = GUICtrlCreateButton(JT_TranSlate('Show'), 170, 45, 155, 20)
	Local $Screen = GUICtrlCreateEdit('', 10, 75, 315, 100, BitOr($ES_MULTILINE, $ES_CENTER, $ES_READONLY))
	GUICtrlSetFont(-1, 9, 400, 0, 'Courier New')
	; Analyze chord
	Local $AnalyzedChord = _JT_ChordAnalyzer($InputChord)
	Local $ChordName = $AnalyzedChord[0], $ChordExtend = $AnalyzedChord[1], $ChordSlash = $AnalyzedChord[2]
	
	If (StringLen($ChordSlash) > 0) Then
		GUICtrlSetData($CurrentChord, $ChordName & $ChordExtend & '/' & $ChordSlash)
		GUICtrlSetData($Screen, _JT_ChordGetMap($ChordName & $ChordExtend & '/' & $ChordSlash, $ChordMapFile, 0))
	Else
		GUICtrlSetData($CurrentChord, $ChordName & $ChordExtend)
		GUICtrlSetData($Screen, _JT_ChordGetMap($ChordName & $ChordExtend, $ChordMapFile, 0))
	EndIf
	
	Local $TmpArr[6] = [0,0,0,0,0,0]
	
	GUISetState(@SW_SHOW, $FormViewChord)
	GUISetState(@SW_DISABLE, $FormMain)
	
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				ExitLoop(1)
			Case $ModChord
				$InputChord = StringStripWS(GUICtrlRead($CurrentChord), 7)
				; Run('cmd.exe /c start "" "' & $ChordMapFile & '"')
				Local $FormModChord = GUICreate(JT_TranSlate('Add') & '/' & JT_TranSlate('Edit'), 320, 260, -1, -1, $GUI_SS_DEFAULT_GUI, $WS_EX_MDICHILD, $FormViewChord)
				GUISetFont(9, 400, 0, 'Tahoma', $FormModChord)
				Local $ModChord_Name = GuiCtrlCreateInput($InputChord, 15, 10, 290, 25, $SS_CENTER)
				GUICtrlSetFont(-1, 15, 800, 0, 'Times New Roman')
				GuiCtrlCreateLabel(JT_TranSlate('Fret'), 25, 40, 45, 20)
				GuiCtrlCreateLabel('0   1   2   3   4   5   6   7', 85, 40, 235, 20)
				GUICtrlSetFont(-1, 9, 400, 0, 'Courier New')
				GuiCtrlCreateGroup('', 5, 55, 310, 170)
				GuiCtrlCreateLabel('[E]>', 25, 70, 30, 20)
				Local $Str_E1c = GUICtrlCreateCheckbox('', 60, 70, 20, 20)
				GUICtrlSetTip(-1, JT_TranSlate('Mute'))
				Local $Str_E1 = GUICtrlCreateSlider(80, 70, 215, 20)
				GUICtrlSetLimit(-1, 7, 0)
				GuiCtrlCreateLabel('[B]>', 25, 95, 30, 20)
				Local $Str_Bc = GUICtrlCreateCheckbox('', 60, 95, 20, 20)
				GUICtrlSetTip(-1, JT_TranSlate('Mute'))
				Local $Str_B = GUICtrlCreateSlider(80, 95, 215, 20)
				GUICtrlSetLimit(-1, 7, 0)
				GuiCtrlCreateLabel('[G]>', 25, 120, 30, 20)
				Local $Str_Gc = GUICtrlCreateCheckbox('', 60, 120, 20, 20)
				GUICtrlSetTip(-1, JT_TranSlate('Mute'))
				Local $Str_G = GUICtrlCreateSlider(80, 120, 215, 20)
				GUICtrlSetLimit(-1, 7, 0)
				GuiCtrlCreateLabel('[D]>', 25, 145, 30, 20)
				Local $Str_Dc = GUICtrlCreateCheckbox('', 60, 145, 20, 20)
				GUICtrlSetTip(-1, JT_TranSlate('Mute'))
				Local $Str_D = GUICtrlCreateSlider(80, 145, 215, 20)
				GUICtrlSetLimit(-1, 7, 0)
				GuiCtrlCreateLabel('[A]>', 25, 170, 30, 20)
				Local $Str_Ac = GUICtrlCreateCheckbox('', 60, 170, 20, 20)
				GUICtrlSetTip(-1, JT_TranSlate('Mute'))
				Local $Str_A = GUICtrlCreateSlider(80, 170, 215, 20)
				GUICtrlSetLimit(-1, 7, 0)
				GuiCtrlCreateLabel('[E]>', 25, 195, 30, 20)
				Local $Str_E2c = GUICtrlCreateCheckbox('', 60, 195, 20, 20)
				GUICtrlSetTip(-1, JT_TranSlate('Mute'))
				Local $Str_E2 = GUICtrlCreateSlider(80, 195, 215, 20)
				GUICtrlSetLimit(-1, 7, 0)
				GuiCtrlCreateGroup('', -99, -99, 1, 1)
				Local $ModChord_Ok = GUICtrlCreateButton(JT_TranSlate('Apply'), 5, 230, 150, 25)
				Local $ModChord_Del = GUICtrlCreateButton(JT_TranSlate('Delete'), 165, 230, 150, 25)
				
				If (StringLen($InputChord) > 0) Then
					$TmpArr = _JT_ChordGetMap($InputChord, $ChordMapFile, 1)
					If (UBound($TmpArr) == 6) Then
						If ($TmpArr[0] > 15) Then
							GuiCtrlSetState($Str_E2c, $GUI_CHECKED)
						Else
							GuiCtrlSetData($Str_E2, $TmpArr[0])
						EndIf
						If ($TmpArr[1] > 15) Then
							GuiCtrlSetState($Str_Ac, $GUI_CHECKED)
						Else
							GuiCtrlSetData($Str_A, $TmpArr[1])
						EndIf
						If ($TmpArr[2] > 15) Then
							GuiCtrlSetState($Str_Dc, $GUI_CHECKED)
						Else
							GuiCtrlSetData($Str_D, $TmpArr[2])
						EndIf
						If ($TmpArr[3] > 15) Then
							GuiCtrlSetState($Str_Gc, $GUI_CHECKED)
						Else
							GuiCtrlSetData($Str_G, $TmpArr[3])
						EndIf
						If ($TmpArr[4] > 15) Then
							GuiCtrlSetState($Str_Bc, $GUI_CHECKED)
						Else
							GuiCtrlSetData($Str_B, $TmpArr[4])
						EndIf
						If ($TmpArr[5] > 15) Then
							GuiCtrlSetState($Str_E1c, $GUI_CHECKED)
						Else
							GuiCtrlSetData($Str_E1, $TmpArr[5])
						EndIf
					EndIf
				EndIf
				
				GUISetState(@SW_SHOW, $FormModChord)
				GUISetState(@SW_DISABLE, $FormViewChord)
				While 1
					Switch GUIGetMsg()
						Case $GUI_EVENT_CLOSE
							ExitLoop(1)
						Case $ModChord_Del
							$InputChord = StringStripWS(GUICtrlRead($ModChord_Name), 7)
							; Analyze chord
							$AnalyzedChord = _JT_ChordAnalyzer($InputChord)
							$ChordName = $AnalyzedChord[0]
							$ChordExtend = $AnalyzedChord[1]
							$ChordSlash = $AnalyzedChord[2]
							
							If (StringLen($ChordSlash) > 0) Then
								If (JT_MessageBox($FormModChord, 'Warning', 'Do you really want to DELETE this chord: "%s"', $ChordName & $ChordExtend & '/' & $ChordSlash) == 6) Then
									If (IniDelete ($ChordMapFile, $ChordName, $ChordName & $ChordExtend & '/' & $ChordSlash) == 1) Then
										JT_MessageBox($FormModChord, 'Info', 'Deleted')
									EndIf
								EndIf
							Else
								If (JT_MessageBox($FormModChord, 'Warning', 'Do you really want to DELETE this chord: "%s"', $ChordName & $ChordExtend) == 6) Then
									If (IniDelete ($ChordMapFile, $ChordName, $ChordName & $ChordExtend) == 1) Then
										JT_MessageBox($FormModChord, 'Info', 'Deleted')
									EndIf
								EndIf
							EndIf
						Case $ModChord_Ok
							$InputChord = StringStripWS(GuiCtrlRead($ModChord_Name), 7)
							; Analyze chord
							$AnalyzedChord = _JT_ChordAnalyzer($InputChord)
							$ChordName = $AnalyzedChord[0]
							$ChordExtend = $AnalyzedChord[1]
							$ChordSlash = $AnalyzedChord[2]
							
							If (StringLen($ChordSlash) > 0) Then
								GUICtrlSetData($ModChord_Name, $ChordName & $ChordExtend & '/' & $ChordSlash)
							Else
								GUICtrlSetData($ModChord_Name, $ChordName & $ChordExtend)
							EndIf
							
							$InputChord = GuiCtrlRead($ModChord_Name)
							
							If (GuiCtrlRead($Str_E1c) == $GUI_CHECKED) Then
								$TmpArr[5] = 99
							Else
								$TmpArr[5] = GuiCtrlRead($Str_E1)
							EndIf
							If (GuiCtrlRead($Str_Bc) == $GUI_CHECKED) Then
								$TmpArr[4] = 99
							Else
								$TmpArr[4] = GuiCtrlRead($Str_B)
							EndIf
							If (GuiCtrlRead($Str_Gc) == $GUI_CHECKED) Then
								$TmpArr[3] = 99
							Else
								$TmpArr[3] = GuiCtrlRead($Str_G)
							EndIf
							If (GuiCtrlRead($Str_Dc) == $GUI_CHECKED) Then
								$TmpArr[2] = 99
							Else
								$TmpArr[2] = GuiCtrlRead($Str_D)
							EndIf
							If (GuiCtrlRead($Str_Ac) == $GUI_CHECKED) Then
								$TmpArr[1] = 99
							Else
								$TmpArr[1] = GuiCtrlRead($Str_A)
							EndIf
							If (GuiCtrlRead($Str_E2c) == $GUI_CHECKED) Then
								$TmpArr[0] = 99
							Else
								$TmpArr[0] = GuiCtrlRead($Str_E2)
							EndIf
							
							If (IniWrite($ChordMapFile, $ChordName, $InputChord, $TmpArr[0] & ',' & $TmpArr[1] & ',' & $TmpArr[2] & ',' & $TmpArr[3] & ',' & $TmpArr[4] & ',' & $TmpArr[5])) Then							
								If (StringLen($ChordSlash) > 0) Then
									GUICtrlSetData($CurrentChord, $ChordName & $ChordExtend & '/' & $ChordSlash)
									GUICtrlSetData($Screen, _JT_ChordGetMap($ChordName & $ChordExtend & '/' & $ChordSlash, $ChordMapFile, 0))
								Else
									GUICtrlSetData($CurrentChord, $ChordName & $ChordExtend)
									GUICtrlSetData($Screen, _JT_ChordGetMap($ChordName & $ChordExtend, $ChordMapFile, 0))
								EndIf
								JT_MessageBox($FormModChord, 'Info', 'Add/Edit successful')
							Else
								JT_MessageBox($FormModChord, 'Error', 'Error')
							EndIf
					EndSwitch
				WEnd
				GUISetState(@SW_ENABLE, $FormViewChord)
				GuiDelete($FormModChord)
			Case $Go
				$InputChord = StringStripWS(GUICtrlRead($CurrentChord), 7)
				; Analyze chord
				$AnalyzedChord = _JT_ChordAnalyzer($InputChord)
				$ChordName = $AnalyzedChord[0]
				$ChordExtend = $AnalyzedChord[1]
				$ChordSlash = $AnalyzedChord[2]
				
				If (StringLen($ChordSlash) > 0) Then
					GUICtrlSetData($CurrentChord, $ChordName & $ChordExtend & '/' & $ChordSlash)
					GUICtrlSetData($Screen, _JT_ChordGetMap($ChordName & $ChordExtend & '/' & $ChordSlash, $ChordMapFile, 0))
				Else
					GUICtrlSetData($CurrentChord, $ChordName & $ChordExtend)
					GUICtrlSetData($Screen, _JT_ChordGetMap($ChordName & $ChordExtend, $ChordMapFile, 0))
				EndIf
		EndSwitch
	WEnd
	GUISetState(@SW_ENABLE, $FormMain)
	GuiDelete($FormViewChord)
ENDFUNC ; <== JT_ViewChord

FUNC JT_Update()
	Local $FormUpdate = GUICreate(JT_TranSlate('Update libraries'), 390, 180, 350, 200, $GUI_SS_DEFAULT_GUI, $WS_EX_MDICHILD, $FormMain)
	GUISetFont(9, 400, 0, 'Tahoma', $FormUpdate)
	GUICtrlCreateGroup(JT_TranSlate('Update from'), 5, 5, 380, 75)
	Local $FromFileMDB = GUICtrlCreateRadio(JT_TranSlate('File (MDB)'), 20, 25, 110, 20)
	GUICtrlSetState(-1, $GUI_CHECKED)
	Local $FromFileJDB = GUICtrlCreateRadio(JT_TranSlate('File (JDB)'), 140, 25, 110, 20)
	Local $FromWeb = GUICtrlCreateRadio(JT_TranSlate('Web'), 260, 25, 110, 20)
	Local $FromWebUrl = GUICtrlCreateInput($JdbUrl, 20, 50, 355, 20)
	GUICtrlCreateGroup('', -99, -99, 1, 1)
	GUICtrlCreateGroup(JT_TranSlate('If the song already exists'), 5, 90, 380, 45)
	Local $ModeSkip = GUICtrlCreateRadio(JT_TranSlate('Skip'), 20, 110, 110, 20)
	GUICtrlSetState(-1, $GUI_CHECKED)
	Local $ModeReplace = GUICtrlCreateRadio(JT_TranSlate('Overwrite'), 140, 110, 110, 20)
	Local $ModeKeep = GUICtrlCreateRadio(JT_TranSlate('Keep both'), 260, 110, 110, 20)
	GUICtrlCreateGroup('', -99, -99, 1, 1)
	Local $Start = GUICtrlCreateButton(JT_TranSlate('Start update'), 5, 145, 380, 30)
	GUISetState(@SW_SHOW, $FormUpdate)
	GUISetState(@SW_DISABLE, $FormMain)
	
	Local $UpdateData = '', $hDownload, $Mode
	Local $GUIMessage
	Local $Song[4] = ['', '', '', '']
	Local $error = 0, $skip = 0, $new = 0, $overWrite = 0
	
	While 1
		$GUIMessage = GUIGetMsg()
		Switch $GUIMessage
			Case $GUI_EVENT_CLOSE
				ExitLoop(1)
			Case $Start
				$error = 0
				$skip = 0
				$new = 0
				$overWrite = 0
				
				If (GUICtrlRead($ModeSkip) == $GUI_CHECKED) Then
					$Mode = 0 ; Skip
				Else
					If (GUICtrlRead($ModeReplace) == $GUI_CHECKED) Then
						$Mode = 1 ; Replace
					Else
						$Mode = 2 ; Keep both
					EndIf
				EndIf
				
				If (GUICtrlRead($FromFileMDB) == $GUI_CHECKED) Then
					$UpdateData = FileOpenDialog(JT_TranSlate('Choose file'), @ScriptDir, '[JT] Songbooks DB (*.mdb)', 2+1, 'SongbooksDB.mdb', $FormUpdate)
				Else
					If (GUICtrlRead($FromFileJDB) == $GUI_CHECKED) Then
						$UpdateData = FileOpenDialog(JT_TranSlate('Choose file'), @ScriptDir, '[JT] Songbooks DB (*.jdb)', 2+1, 'SongbooksDB.jdb', $FormUpdate)
					Else
						GUICtrlSetData($Start, JT_TranSlate('Stop update'))
						If (StringRight($FromWebUrl, 4) == '.mdb') Then
							$hDownload = InetGet($FromWebUrl, @TempDir & '\SongbooksDB.mdb', 1, 0)
						Else
							If (StringRight($FromWebUrl, 4) == '.jdb') Then
								$hDownload = InetGet($FromWebUrl, @TempDir & '\SongbooksDB.jdb', 1, 0)
							EndIf
						EndIf
						Do
							If ($GUIMessage == $Start) Then ExitLoop(1)
						Until InetGetInfo($hDownload, 2)
						If (StringRight($FromWebUrl, 4) == '.mdb') Then
							$UpdateData = @TempDir & '\SongbooksDB.mdb'
						Else
							$UpdateData = @TempDir & '\SongbooksDB.jdb'
						EndIf
						InetClose($hDownload)
					EndIf
				EndIf
				
				GUICtrlSetState($Start, $GUI_DISABLE)
				GUICtrlSetData($Start, JT_TranSlate('Working'))
				
				If (FileExists($UpdateData)) Then
					If (StringLower(StringRight($UpdateData, 4)) == '.mdb') Then
						Local $SongDB_New = $DB_Object.OpenDatabase($UpdateData, 0, 1)
						Local $Record_New = $SongDB_New.OpenRecordSet('SELECT * FROM ' & $SongDB_Table)
						
						If ($Record_New.EOF <> -1 Or $Record_New.BOF <> -1) Then
							$Record_New.MoveFirst
							
							While ($Record_New.EOF <> -1)
								$Song[0] = StringReplace(StringStripWS($Record_New.Fields(1).Value, 7), '"', "'")
								$Song[1] = StringReplace(StringStripWS($Record_New.Fields(2).Value, 7), '"', "'")
								$Song[2] = StringReplace(StringStripWS($Record_New.Fields(3).Value, 7), '"', "'")
								$Song[3] = StringAddCR(StringStripCR(StringReplace(StringStripWS($Record_New.Fields(4).Value, 3), '"', "'")))
								
								If (StringLen($Song[2]) > 5) Then $Song[2] = '0/0'
								
								If (StringLen($Song[0]) > 0) Then
									If (_JT_CheckRecord($DB_Object, $DB_Location, $SongDB_Table, $SongDB_ColTitle, $Song[0]) == 1) Then
										Switch ($Mode)
											Case 0 ; Skip
												$skip = $skip + 1
											Case 1 ; Replace
												; $Song[0] = $Song[0] & Random() & ' (New)'
												If (JT_AddSong($Song, $Song[0]) == '') Then
													$overWrite = $overWrite + 1
												Else
													$error = $error + 1
												EndIf
											Case 2 ; Keep both
												$Song[0] = $Song[0] & ' - ' & $Song[1] & ' (' & Random() & ')'
												; $Song[0] = $Song[0] & ' (New)'
												If (JT_AddSong($Song) == '') Then
													$new = $new + 1
												Else
													$error = $error + 1
												EndIf
										EndSwitch
									Else
										; $Song[0] = $Song[0] & Random() & ' (New)'
										If (JT_AddSong($Song) == '') Then
											$new = $new + 1
										Else
											$error = $error + 1
										EndIf
									EndIf
								Else
									$error = $error + 1
								EndIf
								
								$Record_New.MoveNext
							WEnd
						Else
							JT_MessageBox($FormUpdate, 'Error', JT_TranSlate('Nothing to update'))
						EndIf
						
						$SongDB_New.Close
					Else
						If (StringLower(StringRight($UpdateData, 4)) == '.jdb') Then
							Local $DBVersion, $SongArr, $Fetcher, $tmp = 0
							Local $Title, $Author, $Cadence, $Lyrics
							
							$UpdateData = _JT_GetFileContent($UpdateData)
							If (StringLen($UpdateData) > 0) Then
								$DBVersion = StringRegExp($UpdateData, '<!-- DATABASE VERSION - BEGIN -->(.*?)<!-- DATABASE VERSION - END -->' , 1)
								If (IsArray($DBVersion)) Then
									If (JT_MessageBox($FormUpdate, 'Warning', 'Version of the database is "%s", do you want to continue?', $DBVersion[0]) == 6) Then
										$SongArr = StringSplit($UpdateData, '<!-- ===================== SONGBOOKS - BREAKER ===================== -->', 1)
										If ($SongArr[0] > 1) Then
											For $Fetcher In $SongArr
												If ($tmp == 0) Then
													$tmp = 1
													ContinueLoop(1)
												EndIf
												
												$Title = StringRegExp($Fetcher, '<!-- SONG TITLE - BEGIN -->(.*?)<!-- SONG TITLE - END -->' , 1)
												$Author = StringRegExp($Fetcher, '<!-- SONG AUTHOR - BEGIN -->(.*?)<!-- SONG AUTHOR - END -->' , 1)
												$Cadence = StringRegExp($Fetcher, '<!-- SONG CADENCE - BEGIN -->(.*?)<!-- SONG CADENCE - END -->' , 1)
												$Lyrics = StringRegExp($Fetcher, '<!-- SONG LYRICS - BEGIN -->(?s)(.*?)<!-- SONG LYRICS - END -->' , 1)
												
												If (IsArray($Title)) Then
													$Song[0] = StringReplace(StringStripWS($Title[0], 7), '"', "'")
												Else
													$Song[0] = ''
												EndIf
												If (IsArray($Author)) Then
													$Song[1] = StringReplace(StringStripWS($Author[0], 7), '"', "'")
												Else
													$Song[1] = ''
												EndIf
												If (IsArray($Cadence)) Then
													$Song[2] = StringReplace(StringStripWS($Cadence[0], 7), '"', "'")
												Else
													$Song[2] = ''
												EndIf
												If (IsArray($Lyrics)) Then
													$Song[3] = StringAddCR(StringStripCR(StringReplace(StringStripWS($Lyrics[0], 3), '"', "'")))
												Else
													$Song[3] = ''
												EndIf
												
												If (StringLen($Song[2]) > 5) Then $Song[2] = '0/0'
												
												If (StringLen($Song[0]) > 0) Then
													If (_JT_CheckRecord($DB_Object, $DB_Location, $SongDB_Table, $SongDB_ColTitle, $Song[0]) == 1) Then
														Switch ($Mode)
															Case 0 ; Skip
																$skip = $skip + 1
															Case 1 ; Replace
																; $Song[0] = $Song[0] & Random() & ' (New)'
																If (JT_AddSong($Song, $Song[0]) == '') Then
																	$overWrite = $overWrite + 1
																Else
																	$error = $error + 1
																EndIf
															Case 2 ; Keep both
																$Song[0] = $Song[0] & ' - ' & $Song[1] & ' (' & Random() & ')'
																; $Song[0] = $Song[0] & ' (New)'
																If (JT_AddSong($Song) == '') Then
																	$new = $new + 1
																Else
																	$error = $error + 1
																EndIf
														EndSwitch
													Else
														; $Song[0] = $Song[0] & Random() & ' (New)'
														If (JT_AddSong($Song) == '') Then
															$new = $new + 1
														Else
															$error = $error + 1
														EndIf
													EndIf
												Else
													$error = $error + 1
												EndIf
											Next
										Else
											JT_MessageBox($FormUpdate, 'Error', JT_TranSlate('Nothing to update'))
										EndIf
									EndIf
								EndIf
							Else
								JT_MessageBox($FormUpdate, 'Error', JT_TranSlate('Nothing to update'))
							EndIf
						EndIf
					EndIf
				Else
					JT_MessageBox($FormUpdate, 'Error', JT_TranSlate('Nothing to update'))
				EndIf
				
				JT_MessageBox($FormUpdate, 'Info', 'STATISTICS: New song (%s) - Replaced (%s) - Skipped (%s) - Error (%s)', String($new), String($overWrite), String($skip), String($error))
				GUICtrlSetState($Start, $GUI_ENABLE)
				GUICtrlSetData($Start, JT_TranSlate('Start update'))
	
				If (FileExists(@TempDir & '\SongbooksDB.mdb')) Then FileDelete(@TempDir & '\SongbooksDB.mdb')
				If (FileExists(@TempDir & '\SongbooksDB.jdb')) Then FileDelete(@TempDir & '\SongbooksDB.jdb')
				If ($new > 0) Then $RebuildTreeNeeded = 1
			Case Else
				If (GUICtrlRead($FromWeb) == $GUI_CHECKED) Then
					If (BitAnd(GUICtrlGetState($FromWebUrl), $GUI_DISABLE)) Then
						GUICtrlSetState($FromWebUrl, $GUI_ENABLE)
					EndIf
				Else
					If (BitAnd(GUICtrlGetState($FromWebUrl), $GUI_ENABLE)) Then
						GUICtrlSetState($FromWebUrl, $GUI_DISABLE)
					EndIf
				EndIf
		EndSwitch
	WEnd
	
	GUISetState(@SW_ENABLE, $FormMain)
	GuiDelete($FormUpdate)
ENDFUNC ; <== JT_Update

FUNC JT_Settings()
	Local $FormConfig = GUICreate(JT_TranSlate('Settings'), 330, 275, 350, 150, $GUI_SS_DEFAULT_GUI, $WS_EX_MDICHILD, $FormMain)
	GuiSetFont(9, 400, 0, 'Tahoma', $FormConfig)
	; Group 1
	GuiCtrlCreateGroup(JT_TranSlate('Languages/Screen'), 5, 5, 320, 50)
	Local $LangName = GuiCtrlCreateCombo(IniRead($ConfigFile, 'Common', 'CustomLanguage', 'English'), 15, 25, 150, 20)
	GuiCtrlSetData(-1, 'English|Vietnamese')
	Local $Screen800 = GuiCtrlCreateRadio(JT_TranSlate('Small'), 200, 25, 50, 20)
	Local $Screen1024 = GuiCtrlCreateRadio(JT_TranSlate('Large'), 265, 25, 50, 20)
	GuiCtrlCreateGroup('', -99, -99, 1, 1)
	; Group 2
	GuiCtrlCreateGroup(JT_TranSlate('Font'), 5, 55, 320, 50)
	Local $FontName = GuiCtrlCreateCombo('', 15, 75, 250, 20)
	GuiCtrlSetData(-1, 'Arial|Comic Sans MS|Courier New|Lucida Console|Microsoft Serif|Monotype Corsiva|Palatino Linotype|Tahoma|Times New Roman', IniRead($ConfigFile, 'ViewLyrics', 'FontName', 'Courier New'))
	Local $FontSize = GuiCtrlCreateInput('', 270, 75, 45, 20, BitOR($ES_CENTER, $ES_NUMBER))
	GuiCtrlCreateUpdown(-1)
	GuiCtrlSetLimit(-1, 15, 6)
	GuiCtrlCreateGroup('', -99, -99, 1, 1)
	; Group 3
	GuiCtrlCreateGroup(JT_TranSlate('Chord'), 5, 115, 320, 50)
	Local $ChordBold = GuiCtrlCreateCheckbox('B', 15, 135, 30, 20)
	GuiCtrlSetFont(-1, 9, 800, 0)
	Local $ChordItalic = GuiCtrlCreateCheckbox('I', 50, 135, 30, 20)
	GuiCtrlSetFont(-1, 9, 800, 2)
	Local $ChordUnderline = GuiCtrlCreateCheckbox('U', 85, 135, 30, 20)
	GuiCtrlSetFont(-1, 9, 800, 4)
	Local $ChordBlack = GuiCtrlCreateButton('', 120, 135, 20, 20)
	GuiCtrlSetBkColor(-1, 0x000000)
	Local $ChordWhite = GuiCtrlCreateButton('', 145, 135, 20, 20)
	GuiCtrlSetBkColor(-1, 0xFFFFFF)
	Local $ChordRed = GuiCtrlCreateButton('', 170, 135, 20, 20)
	GuiCtrlSetBkColor(-1, 0xFF0000)
	Local $ChordGreen = GuiCtrlCreateButton('', 195, 135, 20, 20)
	GuiCtrlSetBkColor(-1, 0x00FF00)
	Local $ChordBlue = GuiCtrlCreateButton('', 220, 135, 20, 20)
	GuiCtrlSetBkColor(-1, 0x0000FF)
	Local $ChordColor = GuiCtrlCreateInput('', 245, 135, 70, 20)
	GuiCtrlSetTip(-1, JT_TranSlate('Chords color'))
	GuiCtrlCreateGroup('', -99, -99, 1, 1)
	; Group 4
	GuiCtrlCreateGroup(JT_TranSlate('Lyrics'), 5, 175, 320, 50)
	Local $LyricsBold = GuiCtrlCreateCheckbox('B', 15, 195, 30, 20)
	GuiCtrlSetFont(-1, 9, 800, 0)
	Local $LyricsItalic = GuiCtrlCreateCheckbox('I', 50, 195, 30, 20)
	GuiCtrlSetFont(-1, 9, 800, 2)
	Local $LyricsUnderline = GuiCtrlCreateCheckbox('U', 85, 195, 30, 20)
	GuiCtrlSetFont(-1, 9, 800, 4)
	Local $LyricsBlack = GuiCtrlCreateButton('', 120, 195, 20, 20)
	GuiCtrlSetBkColor(-1, 0x000000)
	Local $LyricsWhite = GuiCtrlCreateButton('', 145, 195, 20, 20)
	GuiCtrlSetBkColor(-1, 0xFFFFFF)
	Local $LyricsRed = GuiCtrlCreateButton('', 170, 195, 20, 20)
	GuiCtrlSetBkColor(-1, 0xFF0000)
	Local $LyricsGreen = GuiCtrlCreateButton('', 195, 195, 20, 20)
	GuiCtrlSetBkColor(-1, 0x00FF00)
	Local $LyricsBlue = GuiCtrlCreateButton('', 220, 195, 20, 20)
	GuiCtrlSetBkColor(-1, 0x0000FF)
	Local $LyricsColor = GuiCtrlCreateInput('', 245, 195, 70, 20, $ES_CENTER)
	GuiCtrlSetTip(-1, JT_TranSlate('Lyrics color'))
	GuiCtrlCreateGroup('', -99, -99, 1, 1)
	Local $ButtonDefault = GuiCtrlCreateButton(JT_TranSlate('Default'), 5, 240, 155, 30)
	Local $ButtonApply = GuiCtrlCreateButton(JT_TranSlate('Apply'), 170, 240, 155, 30)
	
	If (IniRead($ConfigFile, 'Common', 'Screen', 1024) == 1024) Then
		GuiCtrlSetState($Screen1024, $GUI_CHECKED)
	Else
		GuiCtrlSetState($Screen800, $GUI_CHECKED)
	EndIf
	GuiCtrlSetData($FontSize, IniRead($ConfigFile, 'ViewLyrics', 'FontSize', 10))
	If (IniRead($ConfigFile, 'ViewLyrics', 'ChordBold', 0) == 1) Then GuiCtrlSetState($ChordBold, $GUI_CHECKED)
	If (IniRead($ConfigFile, 'ViewLyrics', 'ChordItalic', 0) == 1) Then GuiCtrlSetState($ChordItalic, $GUI_CHECKED)
	If (IniRead($ConfigFile, 'ViewLyrics', 'ChordUnderline', 0) == 1) Then GuiCtrlSetState($ChordUnderline, $GUI_CHECKED)
	GuiCtrlSetData($ChordColor, IniRead($ConfigFile, 'ViewLyrics', 'ChordColor', '0x000000'))
	If (IniRead($ConfigFile, 'ViewLyrics', 'LyricsBold', 0) == 1) Then GuiCtrlSetState($LyricsBold, $GUI_CHECKED)
	If (IniRead($ConfigFile, 'ViewLyrics', 'LyricsItalic', 0) == 1) Then GuiCtrlSetState($LyricsItalic, $GUI_CHECKED)
	If (IniRead($ConfigFile, 'ViewLyrics', 'LyricsUnderline', 0) == 1) Then GuiCtrlSetState($LyricsUnderline, $GUI_CHECKED)
	GuiCtrlSetData($LyricsColor, IniRead($ConfigFile, 'ViewLyrics', 'LyricsColor', '0x000000'))
	GuiCtrlSetState($ButtonApply, $GUI_FOCUS)
	GuiSetState(@SW_SHOW, $FormConfig)
	GuiSetState(@SW_DISABLE, $FormMain)
	While 1
		Switch GuiGetMsg()
			Case $GUI_EVENT_CLOSE
				ExitLoop(1)
			Case $ButtonDefault
				GuiCtrlSetData($LangName, 'English')
				GuiCtrlSetState($Screen1024, $GUI_CHECKED)
				GuiCtrlSetData($FontName, 'Courier New')
				GuiCtrlSetData($FontSize, '10')
				GuiCtrlSetState($ChordBold, $GUI_UNCHECKED)
				GuiCtrlSetState($ChordItalic, $GUI_UNCHECKED)
				GuiCtrlSetState($ChordUnderline, $GUI_UNCHECKED)
				GuiCtrlSetData($ChordColor, '0x000000')
				GuiCtrlSetState($LyricsBold, $GUI_UNCHECKED)
				GuiCtrlSetState($LyricsItalic, $GUI_UNCHECKED)
				GuiCtrlSetState($LyricsUnderline, $GUI_UNCHECKED)
				GuiCtrlSetData($LyricsColor, '0x000000')
			; Chord
			Case $ChordBlack
				GuiCtrlSetData($ChordColor, '0x000000')
			Case $ChordWhite
				GuiCtrlSetData($ChordColor, '0xFFFFFF')
			Case $ChordBlue
				GuiCtrlSetData($ChordColor, '0xFF0000')
			Case $ChordGreen
				GuiCtrlSetData($ChordColor, '0x00FF00')
			Case $ChordRed
				GuiCtrlSetData($ChordColor, '0x0000FF')
			; Lyrics
			Case $LyricsBlack
				GuiCtrlSetData($LyricsColor, '0x000000')
			Case $LyricsWhite
				GuiCtrlSetData($LyricsColor, '0xFFFFFF')
			Case $LyricsBlue
				GuiCtrlSetData($LyricsColor, '0xFF0000')
			Case $LyricsGreen
				GuiCtrlSetData($LyricsColor, '0x00FF00')
			Case $LyricsRed
				GuiCtrlSetData($LyricsColor, '0x0000FF')
			; Apply
			Case $ButtonApply
				If (GuiCtrlRead($Screen1024) == $GUI_CHECKED) Then
					IniWrite($ConfigFile, 'Common', 'Screen', '1024')
				Else
					IniWrite($ConfigFile, 'Common', 'Screen', '800')
				EndIf
			
				IniWrite($ConfigFile, 'Common', 'CustomLanguage', GuiCtrlRead($LangName))
				IniWrite($ConfigFile, 'ViewLyrics', 'FontName', GuiCtrlRead($FontName))
				IniWrite($ConfigFile, 'ViewLyrics', 'FontSize', GuiCtrlRead($FontSize))
				
				IniWrite($ConfigFile, 'ViewLyrics', 'ChordBold', GuiCtrlRead($ChordBold))
				IniWrite($ConfigFile, 'ViewLyrics', 'ChordItalic', GuiCtrlRead($ChordItalic))
				IniWrite($ConfigFile, 'ViewLyrics', 'ChordUnderline', GuiCtrlRead($ChordUnderline))
				IniWrite($ConfigFile, 'ViewLyrics', 'ChordColor', GuiCtrlRead($ChordColor))

				IniWrite($ConfigFile, 'ViewLyrics', 'LyricsBold', GuiCtrlRead($LyricsBold))
				IniWrite($ConfigFile, 'ViewLyrics', 'LyricsItalic', GuiCtrlRead($LyricsItalic))
				IniWrite($ConfigFile, 'ViewLyrics', 'LyricsUnderline', GuiCtrlRead($LyricsUnderline))
				IniWrite($ConfigFile, 'ViewLyrics', 'LyricsColor', GuiCtrlRead($LyricsColor))
				JT_MessageBox($FormConfig, 'Info', 'All settings are saved')
				ExitLoop(1)
		EndSwitch
	WEnd
	GuiSetState(@SW_ENABLE, $FormMain)
	GuiDelete($FormConfig)
ENDFUNC ; <== JT_Settings

FUNC JT_About()
	Local $FormAbout = GUICreate(JT_TranSlate('About'), 370, 200, 350, 200, $GUI_SS_DEFAULT_GUI, $WS_EX_MDICHILD, $FormMain)
	Local $TotalSongs = _JT_GetRecordLists($DB_Object, $DB_Location, $SongDB_Table, $SongDB_ColTitle, 1)

	GUISetFont(9, 400, 0, 'Tahoma', $FormAbout)
	GUICtrlCreateGroup('', 5, 5, 360, 175)
	GUICtrlCreatePic($AboutFile, 15, 20, 150, 150)
	GUICtrlCreateLabel(JT_TranSlate('Product name') & ': [JT] Songbooks', 175, 20, 180, 20)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	GUICtrlCreateLabel(JT_TranSlate('Version') & ': ' & $AppVersion, 175, 38, 180, 20)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	GUICtrlCreateLabel(JT_TranSlate('Author') & ': o0johntam0o', 175, 56, 180, 20)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	GUICtrlCreateLabel(JT_TranSlate('Copyright') & ': [JT] Knowledge', 175, 74, 180, 20)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	GUICtrlCreateLabel(JT_TranSlate('Homepage') & ': ', 175, 92, 70, 20)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	Local $HomepageLabel = GUICtrlCreateLabel('GitHub', 250, 92, 105, 20)
	GUICtrlSetColor(-1, 0x0000ff)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	GUICtrlCreateLabel(JT_TranSlate('Contact') & ': o0johntam0o@gmail.com', 175, 110, 180, 20)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	GUICtrlCreateLabel(JT_TranSlate('Source code') & ': ', 175, 128, 70, 20)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	Local $SourceLabel = GUICtrlCreateLabel('GitHub', 250, 128, 105, 20)
	GUICtrlSetColor(-1, 0x0000ff)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	Local $StatLabel = GUICtrlCreateLabel(JT_TranSlate('Statistics') & ': ' & $TotalSongs & ' ' & JT_TranSlate('Song(s)'), 175, 146, 180, 20)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	GUICtrlCreateGroup('', -99, -99, 1, 1)
	Local $HelpLabel = GUICtrlCreateLabel(JT_TranSlate('Help'), 295, 183, 60, 15, $SS_RIGHT)
	GUICtrlSetColor(-1, 0x0000ff)
	GUICtrlSetFont(-1, 8, 400, 0, 'Tahoma')
	GUICtrlSetCursor(-1, 4)
	
	GUISetState(@SW_SHOW, $FormAbout)
	GUISetState(@SW_DISABLE, $FormMain)
	
	GuiCtrlSetData($StatLabel, JT_TranSlate('Statistics') & ': ' & $TotalSongs & ' ' & JT_TranSlate('Song(s)'))
	
	While 1
		Switch GUIGetMsg()
			Case $GUI_EVENT_CLOSE
				ExitLoop(1)
			Case $HelpLabel
				Run('notepad.exe "' & $HelpFile & '"')
			Case $HomepageLabel
				Run('cmd.exe /c start "" "https://github.com/o0johntam0o/JT-Songbooks/"')
			Case $SourceLabel
				Run('cmd.exe /c start "" "https://github.com/o0johntam0o/JT-Songbooks/"')
		EndSwitch
	WEnd
	GUISetState(@SW_ENABLE, $FormMain)
	GuiDelete($FormAbout)
ENDFUNC ; <== JT_About
#ENDREGION
