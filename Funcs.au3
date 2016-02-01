; $dbObj   => ObjCreate('DAO.DBEngine.36')
; $dbPath  => Fullpath to *.mdb
; $dbTable => Table name
; $dbCol   => Column name (put to list)
; $reCount => 0 (No), 1 (Yes)
; $dbCond   => Condition (WHERE <Condition>)
; Return   => List of records (Array) or Number of records
FUNC _JT_GetRecordLists($dbObj, $dbPath, $dbTable, $dbCol, $reCount = 0, $dbCond = '')
	Local $DbOpen = $dbObj.OpenDatabase($dbPath, 0, 1)
	Local $RecordPointer
	
	If ($dbCond <> '') Then $dbCond = ' WHERE ' & $dbCond
	$RecordPointer = $DbOpen.OpenRecordSet('SELECT ' & $dbCol & ' FROM ' & $dbTable & $dbCond)
	If ($RecordPointer.EOF <> -1 Or $RecordPointer.BOF <> -1) Then
		$RecordPointer.MoveLast
	Else
		$DbOpen.Close
		Return 0
	EndIf
	
	Local $tmpInt = $RecordPointer.RecordCount
	
	If ($reCount == 1) Then
		$DbOpen.Close
		Return $tmpInt
	Else
		$RecordPointer.MoveFirst
	EndIf
	
	Local $Re[$tmpInt], $i = 0
	
	For $i = 0 To $tmpInt - 1
		If ($RecordPointer.EOF <> -1 Or $RecordPointer.BOF <> -1) Then
			$Re[$i] = $RecordPointer.Fields(0).Value
			$RecordPointer.MoveNext
		EndIf
	Next
	
	$DbOpen.Close

	;Return _JT_ArraySortString($Re)
	Return $Re
ENDFUNC ; <== _JT_GetRecordLists

; $dbObj    => ObjCreate('DAO.DBEngine.36')
; $dbPath   => Fullpath to *.mdb
; $dbTable  => Table name
; $dbCol    => Column name (to search for)
; $dbFilter => A string (to search for)
; Return    => 0 (No - Record not found), 1 (Yes - Record exists)
FUNC _JT_CheckRecord($dbObj, $dbPath, $dbTable, $dbCol, $dbFilter)
	Local $DbOpen = $dbObj.OpenDatabase($dbPath, 0, 1)
	Local $RecordPointer = $DbOpen.OpenRecordSet('SELECT ' & $dbCol & ' FROM ' & $dbTable & ' WHERE ' & $dbCol & ' = "' & $dbFilter & '"')
	If ($RecordPointer.EOF <> -1 Or $RecordPointer.BOF <> -1) Then
		$DbOpen.Close
		Return 1
	Else
		$DbOpen.Close
		Return 0
	EndIf
ENDFUNC ; <== _JT_CheckRecord

; $dbObj    => ObjCreate('DAO.DBEngine.36')
; $dbPath   => Fullpath to *.mdb
; $dbTable  => Table name
; $dbFilter => A string (to search for)
; Return    => 0 (No - Field not found), 1 (Yes - Field exists)
FUNC _JT_CheckField($dbObj, $dbPath, $dbTable, $dbFilter)
	Local $DbOpen = $dbObj.OpenDatabase($dbPath, 0, 1)
	Local $RecordPointer = $DbOpen.OpenRecordSet($dbTable)
	Local $i
	
	If ($RecordPointer.Fields.Count > 0) Then
		For $i = 0 To $RecordPointer.Fields.Count - 1
			If $RecordPointer.Fields($i).Name = $dbFilter Then
				$DbOpen.Close
				Return 1
			EndIf
		Next
	EndIf
	
	$DbOpen.Close
	Return 0
ENDFUNC ; <== _JT_CheckField

; $array => An array of strings
; Return => Sorted array
FUNC _JT_ArraySortString($array)
	If (Not IsArray($array)) Then Return 0
	Local $i = 0, $j = 0
	
	For $i = 0 To UBound($array) - 1
		$array[$i] = _JT_ToLatin($array[$i]) & $array[$i]
	Next
	
	__ArrayQuickSort1D($array, 0, UBound($array) - 1)
	
	For $i  = 0 To UBound($array) - 1
		$array[$i] = StringRight($array[$i], StringLen($array[$i])/2)
	Next
	
	Return $array
ENDFUNC ; <== _JT_ArraySortString

; #INTERNAL_USE_ONLY# ===========================================================================================================
; Name...........: __ArrayQuickSort1D
; Description ...: Helper function for sorting 1D arrays
; Syntax.........: __ArrayQuickSort1D(ByRef $avArray, ByRef $iStart, ByRef $iEnd)
; Parameters ....: $avArray - Array to sort
;                  $iStart  - Index of array to start sorting at
;                  $iEnd    - Index of array to stop sorting at
; Return values .: None
; Author ........: Jos van der Zande, LazyCoder, Tylo, Ultima
; Modified.......:
; Remarks .......: For Internal Use Only
; Related .......:
; Link ..........:
; Example .......:
; ===============================================================================================================================
Func __ArrayQuickSort1D(ByRef $avArray, ByRef $iStart, ByRef $iEnd)
	If $iEnd <= $iStart Then Return

	Local $vTmp

	; InsertionSort (faster for smaller segments)
	If ($iEnd - $iStart) < 15 Then
		Local $vCur
		For $i = $iStart + 1 To $iEnd
			$vTmp = $avArray[$i]

			If IsNumber($vTmp) Then
				For $j = $i - 1 To $iStart Step -1
					$vCur = $avArray[$j]
					; If $vTmp >= $vCur Then ExitLoop
					If ($vTmp >= $vCur And IsNumber($vCur)) Or (Not IsNumber($vCur) And StringCompare($vTmp, $vCur) >= 0) Then ExitLoop
					$avArray[$j + 1] = $vCur
				Next
			Else
				For $j = $i - 1 To $iStart Step -1
					If (StringCompare($vTmp, $avArray[$j]) >= 0) Then ExitLoop
					$avArray[$j + 1] = $avArray[$j]
				Next
			EndIf

			$avArray[$j + 1] = $vTmp
		Next
		Return
	EndIf

	; QuickSort
	Local $L = $iStart, $R = $iEnd, $vPivot = $avArray[Int(($iStart + $iEnd) / 2)], $fNum = IsNumber($vPivot)
	Do
		If $fNum Then
			; While $avArray[$L] < $vPivot
			While ($avArray[$L] < $vPivot And IsNumber($avArray[$L])) Or (Not IsNumber($avArray[$L]) And StringCompare($avArray[$L], $vPivot) < 0)
				$L += 1
			WEnd
			; While $avArray[$R] > $vPivot
			While ($avArray[$R] > $vPivot And IsNumber($avArray[$R])) Or (Not IsNumber($avArray[$R]) And StringCompare($avArray[$R], $vPivot) > 0)
				$R -= 1
			WEnd
		Else
			While (StringCompare($avArray[$L], $vPivot) < 0)
				$L += 1
			WEnd
			While (StringCompare($avArray[$R], $vPivot) > 0)
				$R -= 1
			WEnd
		EndIf

		; Swap
		If $L <= $R Then
			$vTmp = $avArray[$L]
			$avArray[$L] = $avArray[$R]
			$avArray[$R] = $vTmp
			$L += 1
			$R -= 1
		EndIf
	Until $L > $R

	__ArrayQuickSort1D($avArray, $iStart, $R)
	__ArrayQuickSort1D($avArray, $L, $iEnd)
EndFunc   ;==>__ArrayQuickSort1D

; $chord => 'C', 'Dm', 'E/F', etc...
; $code  => '#', 'b', 1, 2, 3, etc..
; Return => Transposed chord
FUNC _JT_ChordTuner($chord, $code)
	Local $Sharp[34] =  ['A' , 'A#', 'B' , 'B' , 'C' , 'C#', 'D' , 'D' , 'D#', 'E' , 'E' , 'F' , 'F#', 'G' , 'G' , 'Ab', 'A' , 'LA' , 'LA#', 'SI' , 'SI' , 'DO' , 'DO#', 'RE' , 'RE' , 'RE#', 'MI' , 'MI' , 'FA' , 'FA#', 'SOL', 'SOL' , 'LAb', 'LA'  ]
	Local $Source[34] = ['Ab', 'A' , 'A#', 'Bb', 'B' , 'C' , 'C#', 'Db', 'D' , 'D#', 'Eb', 'E' , 'F' , 'F#', 'Gb', 'G' , 'G#', 'LAb', 'LA' , 'LA#', 'SIb', 'SI' , 'DO' , 'DO#', 'REb', 'RE' , 'RE#', 'MIb', 'MI' , 'FA' , 'FA#', 'SOLb', 'SOL', 'SOL#']
	Local $Flat[34] =   ['G' , 'Ab', 'A' , 'A' , 'Bb', 'B' , 'C' , 'C' , 'C#', 'D' , 'D' , 'Eb', 'E' , 'F' , 'F' , 'F#', 'G' , 'SOL', 'LAb', 'LA' , 'LA' , 'SIb', 'SI' , 'DO' , 'DO' , 'DO#', 'RE' , 'RE' , 'MIb', 'MI' , 'FA' , 'FA'  , 'FA#', 'SOL' ]
	Local $i, $Step, $ChordName = '', $ChordExtend = ''
	
	If (StringInStr($chord, '/') > 0) Then
		Local $pre = StringSplit($chord, '/', 2)
		Return _JT_ChordTuner($pre[0], $code) & '/' & _JT_ChordTuner($pre[1], $code)
	EndIf
	
	Switch (StringUpper($chord))
		Case 'DO', 'DO#', 'REB', 'RE', 'RE#', 'MIB', 'MI', 'FA', 'FA#', 'SOLB', 'SOL', 'SOL#', 'LAB', 'LA', 'LA#', 'SIB', 'SI' ; Note
			$ChordName = $chord
			$ChordExtend = ''
		Case Else ; Chord
			If (StringMid($chord, 2, 1) == '#' Or StringMid($chord, 2, 1) == 'b') Then 
				$ChordName = StringLeft($chord, 2)
				$ChordExtend = StringMid($chord, 3)
			Else
				$ChordName = StringLeft($chord, 1)
				$ChordExtend = StringMid($chord, 2)
			EndIf
	EndSwitch
	
	$Step = $code
	If (IsInt($Step)) Then
		If ($Step >= 12) Then
			$Step = Abs($Step) - 12*(Floor(Abs($Step)/12))
		EndIf
	EndIf
	
	$i = 0
	While $i < 34
		If (StringUpper($Source[$i]) == StringUpper($ChordName)) Then
			Select
				Case ($code == '#')
					Return $Sharp[$i] & $ChordExtend
					ExitLoop(1)
				Case ($code == 'b')
					Return $Flat[$i] & $ChordExtend
					ExitLoop(1)
				Case ($Step == 0)
					Return $ChordName & $ChordExtend
					ExitLoop(1)
				Case ($Step < 0)
					$Step = $Step + 1
					$ChordName = $Sharp[$i]
					$i = -1
				Case ($Step > 0)
					$Step = $Step - 1
					$ChordName = $Flat[$i]
					$i = -1
			EndSelect
		EndIf
		$i = $i + 1
	WEnd
	
	Return $chord
ENDFUNC ; <== _JT_ChordTuner

; $chord   => 'C', 'Dm', 'E/F', etc...
; $iniFile => Fullpath to 'ChordMap.ini' file
; Return   => _JT_ChordMapper
FUNC _JT_ChordGetMap($chord, $iniFile, $reArr)
	$chord = StringStripWS($chord, 7)
	If FileExists($iniFile) Then
		If ($reArr == 1) Then
			Return StringSplit(StringStripWS(IniRead($iniFile, StringLeft($chord, 1), $chord, '0,0,0,0,0,0'), 8), ',', 2)
		Else
			Return _JT_ChordMapper(StringSplit(StringStripWS(IniRead($iniFile, StringLeft($chord, 1), $chord, '0,0,0,0,0,0'), 8), ',', 2))
		EndIf
	Else
		; A
		IniWrite($iniFile, "A", "A", "99,0,2,2,2,0")
		IniWrite($iniFile, "A", "Am", "99,0,2,2,1,0")
		IniWrite($iniFile, "A", "A5", "99,0,2,2,5,5")
		IniWrite($iniFile, "A", "A6", "99,0,2,2,2,2")
		IniWrite($iniFile, "A", "A7", "99,0,2,0,2,0")
		IniWrite($iniFile, "A", "AAug", "99,0,3,2,2,1")
		IniWrite($iniFile, "A", "ADim", "99,0,1,5,4,5")
		IniWrite($iniFile, "A", "ASus2", "99,0,2,2,0,0")
		IniWrite($iniFile, "A", "ASus4", "99,0,0,2,3,0")
		IniWrite($iniFile, "A", "A#", "99,1,3,3,3,1")
		IniWrite($iniFile, "A", "A#m", "99,1,3,3,2,1")
		IniWrite($iniFile, "A", "A#5", "99,1,3,3,99,99")
		IniWrite($iniFile, "A", "A#6", "99,1,3,0,3,3")
		IniWrite($iniFile, "A", "A#7", "99,1,3,1,3,1")
		IniWrite($iniFile, "A", "A#Aug", "99,1,0,3,3,2")
		IniWrite($iniFile, "A", "A#Dim", "99,1,2,3,2,0")
		IniWrite($iniFile, "A", "A#Sus2", "99,1,3,3,1,1")
		IniWrite($iniFile, "A", "A#Sus4", "99,1,1,3,4,1")
		; B
		IniWrite($iniFile, "B", "B", "99,2,4,4,4,2")
		IniWrite($iniFile, "B", "Bm", "99,2,4,4,3,2")
		IniWrite($iniFile, "B", "B5", "99,2,4,4,99,99")
		IniWrite($iniFile, "B", "B6", "99,2,1,1,0,2")
		IniWrite($iniFile, "B", "B7", "99,2,1,2,0,2")
		IniWrite($iniFile, "B", "BAug", "99,2,1,0,0,3")
		IniWrite($iniFile, "B", "BDim", "99,2,0,4,3,1")
		IniWrite($iniFile, "B", "BSus2", "99,2,4,4,2,2")
		IniWrite($iniFile, "B", "BSus4", "99,2,2,4,5,2")
		; C
		IniWrite($iniFile, "C", "C", "99,3,2,0,1,0")
		IniWrite($iniFile, "C", "Cm", "99,3,5,5,4,3")
		IniWrite($iniFile, "C", "C5", "99,3,5,5,99,99")
		IniWrite($iniFile, "C", "C6", "99,3,5,5,5,5")
		IniWrite($iniFile, "C", "C7", "99,3,5,3,5,3")
		IniWrite($iniFile, "C", "CAug", "99,3,2,1,1,0")
		IniWrite($iniFile, "C", "CDim", "99,3,4,2,4,2")
		IniWrite($iniFile, "C", "CSus2", "99,3,0,0,3,3")
		IniWrite($iniFile, "C", "CSus4", "99,3,3,0,1,1")
		IniWrite($iniFile, "C", "C#", "99,4,3,1,2,1")
		IniWrite($iniFile, "C", "C#m", "99,4,2,1,2,0")
		IniWrite($iniFile, "C", "C#5", "99,4,6,6,99,99")
		IniWrite($iniFile, "C", "C#6", "99,4,3,3,99,4")
		IniWrite($iniFile, "C", "C#7", "99,4,3,1,0,4")
		IniWrite($iniFile, "C", "C#Aug", "99,4,3,2,2,5")
		IniWrite($iniFile, "C", "C#Dim", "99,4,2,0,2,0")
		IniWrite($iniFile, "C", "C#Sus2", "99,4,1,1,4,4")
		IniWrite($iniFile, "C", "C#Sus4", "99,4,4,1,2,2")
		; D
		IniWrite($iniFile, "D", "D", "99,99,0,2,3,2")
		IniWrite($iniFile, "D", "Dm", "99,99,0,2,3,1")
		IniWrite($iniFile, "D", "D5", "99,99,0,2,3,5")
		IniWrite($iniFile, "D", "D6", "99,99,0,2,0,2")
		IniWrite($iniFile, "D", "D7", "99,99,0,2,1,2")
		IniWrite($iniFile, "D", "DAug", "99,99,0,3,3,2")
		IniWrite($iniFile, "D", "DDim", "99,99,0,1,3,1")
		IniWrite($iniFile, "D", "DSus2", "99,99,0,2,3,0")
		IniWrite($iniFile, "D", "DSus4", "99,99,0,2,3,3")
		IniWrite($iniFile, "D", "D#", "99,99,1,3,4,3")
		IniWrite($iniFile, "D", "D#m", "99,99,1,3,4,2")
		IniWrite($iniFile, "D", "D#5", "99,99,1,3,4,99")
		IniWrite($iniFile, "D", "D#6", "99,99,1,3,1,3")
		IniWrite($iniFile, "D", "D#7", "99,99,1,3,2,3")
		IniWrite($iniFile, "D", "D#Aug", "99,99,1,0,0,3")
		IniWrite($iniFile, "D", "D#Dim", "99,99,1,2,4,2")
		IniWrite($iniFile, "D", "D#Sus2", "99,99,1,3,4,1")
		IniWrite($iniFile, "D", "D#Sus4", "99,99,1,3,4,4")
		; E
		IniWrite($iniFile, "E", "E", "0,2,2,1,0,0")
		IniWrite($iniFile, "E", "Em", "0,2,2,0,0,0")
		IniWrite($iniFile, "E", "E5", "0,2,2,99,99,99")
		IniWrite($iniFile, "E", "E6", "0,2,2,1,2,0")
		IniWrite($iniFile, "E", "E7", "0,2,2,1,3,0")
		IniWrite($iniFile, "E", "EAug", "0,3,2,1,1,0")
		IniWrite($iniFile, "E", "EDim", "99,99,2,3,5,2")
		IniWrite($iniFile, "E", "ESus2", "0,2,2,4,5,2")
		IniWrite($iniFile, "E", "ESus4", "0,0,2,2,0,0")
		; F
		IniWrite($iniFile, "F", "F", "99,99,3,2,1,1")
		IniWrite($iniFile, "F", "Fm", "99,99,3,1,1,1")
		IniWrite($iniFile, "F", "F5", "1,3,3,99,99,99")
		IniWrite($iniFile, "F", "F6", "99,99,3,5,3,5")
		IniWrite($iniFile, "F", "F7", "99,99,3,5,4,5")
		IniWrite($iniFile, "F", "FAug", "99,99,3,2,2,1")
		IniWrite($iniFile, "F", "FDim", "99,99,3,2,0,2")
		IniWrite($iniFile, "F", "FSus2", "99,99,3,0,1,1")
		IniWrite($iniFile, "F", "FSus4", "99,99,3,3,1,1")
		IniWrite($iniFile, "F", "F#", "99,99,4,3,2,2")
		IniWrite($iniFile, "F", "F#m", "99,99,4,2,2,2")
		IniWrite($iniFile, "F", "F#5", "2,4,4,99,99,99")
		IniWrite($iniFile, "F", "F#6", "2,1,1,3,2,99")
		IniWrite($iniFile, "F", "F#7", "99,99,4,3,2,0")
		IniWrite($iniFile, "F", "F#Aug", "99,99,4,3,3,2")
		IniWrite($iniFile, "F", "F#Dim", "99,99,4,2,1,2")
		IniWrite($iniFile, "F", "F#Sus2", "99,99,4,1,2,4")
		IniWrite($iniFile, "F", "F#Sus4", "99,99,4,4,2,2")
		; G
		IniWrite($iniFile, "G", "G", "3,2,0,0,0,3")
		IniWrite($iniFile, "G", "Gm", "3,1,0,0,3,3")
		IniWrite($iniFile, "G", "G5", "99,99,99,0,3,3")
		IniWrite($iniFile, "G", "G6", "3,2,0,0,0,0")
		IniWrite($iniFile, "G", "G7", "3,2,0,0,0,1")
		IniWrite($iniFile, "G", "GAug", "3,2,1,0,0,3")
		IniWrite($iniFile, "G", "GDim", "99,99,5,3,2,3")
		IniWrite($iniFile, "G", "GSus2", "3,0,0,0,3,3")
		IniWrite($iniFile, "G", "GSus4", "3,3,0,0,3,3")
		IniWrite($iniFile, "G", "G#", "4,3,1,1,1,4")
		IniWrite($iniFile, "G", "G#m", "4,2,1,1,4,4")
		IniWrite($iniFile, "G", "G#5", "99,99,99,1,4,4")
		IniWrite($iniFile, "G", "G#6", "4,3,1,1,1,1")
		IniWrite($iniFile, "G", "G#7", "4,3,1,1,1,2")
		IniWrite($iniFile, "G", "G#Aug", "99,99,99,1,1,0")
		IniWrite($iniFile, "G", "G#Dim", "4,2,0,1,0,4")
		IniWrite($iniFile, "G", "G#Sus2", "4,1,1,1,4,4")
		IniWrite($iniFile, "G", "G#Sus4", "4,4,1,1,4,4")
		
		If ($reArr == 1) Then
			Return StringSplit(StringStripWS(IniRead($iniFile, StringLeft($chord, 1), $chord, '0,0,0,0,0,0'), 8), ',', 2)
		Else
			Return _JT_ChordMapper(StringSplit(StringStripWS(IniRead($iniFile, StringLeft($chord, 1), $chord, '0,0,0,0,0,0'), 8), ',', 2))
		EndIf
	EndIf
ENDFUNC ; <== _JT_ChordGetMap

; $map   => Array[6] = [6, 5, 4, 3, 2, 1]
; Return => String that show how to play a chord
FUNC _JT_ChordMapper($map)
	If (Not IsArray($map)) Then
		Return ''
	Else
		If (UBound($map) <> 6) Then
			Return ''
		EndIf
	EndIf
	
	Local $Auto[6] = ['[E]>', '[A]>', '[D]>', '[G]>', '[B]>', '[E]>']
	Local $Re = ''
	Local $i = 0 ; String
	Local $j = 0 ; Fret

	For $i = 0 To 5
		For $j = 0 To 7
			If ($j == $map[$i] And $j <> 0) Then
				Switch ($j)
					Case 1
						Select
							Case ($i == 0 Or $i == 5)
								$Auto[$i] = $Auto[$i] & "-F--|"
							Case ($i == 1)
								$Auto[$i] = $Auto[$i] & "-A#-|"
							Case ($i == 2)
								$Auto[$i] = $Auto[$i] & "-D#-|"
							Case ($i == 3)
								$Auto[$i] = $Auto[$i] & "-G#-|"
							Case ($i == 4)
								$Auto[$i] = $Auto[$i] & "-C--|"
						EndSelect
					Case 2
						Select
							Case ($i==0 Or $i==5)
								$Auto[$i] = $Auto[$i] & "-F#-|"
							Case ($i == 1)
								$Auto[$i] = $Auto[$i] & "-B--|"
							Case ($i == 2)
								$Auto[$i] = $Auto[$i] & "-E--|"
							Case ($i == 3)
								$Auto[$i] = $Auto[$i] & "-A--|"
							Case ($i == 4)
								$Auto[$i] = $Auto[$i] & "-C#-|"
						EndSelect
					Case 3
						Select
							Case ($i==0 Or $i==5)
								$Auto[$i] = $Auto[$i] & "-G--|"
							Case ($i == 1)
								$Auto[$i] = $Auto[$i] & "-C--|"
							Case ($i == 2)
								$Auto[$i] = $Auto[$i] & "-F--|"
							Case ($i == 3)
								$Auto[$i] = $Auto[$i] & "-A#-|"
							Case ($i == 4)
								$Auto[$i] = $Auto[$i] & "-D--|"
						EndSelect
					Case 4
						Select
							Case ($i==0 Or $i==5)
								$Auto[$i] = $Auto[$i] & "-G#-|"
							Case ($i == 1)
								$Auto[$i] = $Auto[$i] & "-C#-|"
							Case ($i == 2)
								$Auto[$i] = $Auto[$i] & "-F#-|"
							Case ($i == 3)
								$Auto[$i] = $Auto[$i] & "-B--|"
							Case ($i == 4)
								$Auto[$i] = $Auto[$i] & "-D#-|"
						EndSelect
					Case 5
						Select
							Case ($i==0 Or $i==5)
								$Auto[$i] = $Auto[$i] & "-A--|"
							Case ($i == 1)
								$Auto[$i] = $Auto[$i] & "-D--|"
							Case ($i == 2)
								$Auto[$i] = $Auto[$i] & "-G--|"
							Case ($i == 3)
								$Auto[$i] = $Auto[$i] & "-C--|"
							Case ($i == 4)
								$Auto[$i] = $Auto[$i] & "-E--|"
						EndSelect
					Case 6
						Select
							Case ($i==0 Or $i==5)
								$Auto[$i] = $Auto[$i] & "-A#-|"
							Case ($i == 1)
								$Auto[$i] = $Auto[$i] & "-D#-|"
							Case ($i == 2)
								$Auto[$i] = $Auto[$i] & "-G#-|"
							Case ($i == 3)
								$Auto[$i] = $Auto[$i] & "-C#-|"
							Case ($i == 4)
								$Auto[$i] = $Auto[$i] & "-F--|"
						EndSelect
					Case 7
						Select
							Case ($i==0 Or $i==5)
								$Auto[$i] = $Auto[$i] & "-B--|"
							Case ($i == 1)
								$Auto[$i] = $Auto[$i] & "-E--|"
							Case ($i == 2)
								$Auto[$i] = $Auto[$i] & "-A--|"
							Case ($i == 3)
								$Auto[$i] = $Auto[$i] & "-D--|"
							Case ($i == 4)
								$Auto[$i] = $Auto[$i] & "-F#-|"
						EndSelect
				EndSwitch
			Else
				If ($j == 0) Then
					If ($map[$i] == 99) Then
						$Auto[$i] = $Auto[$i] & "-XX"
					Else
						$Auto[$i] = $Auto[$i] & "-||"
					EndIf
				Else
					$Auto[$i] = $Auto[$i] & "----|"
				EndIf
			EndIf
		Next
	Next

	For $i = 5 To 0 Step -1
		$Re = $Re & $Auto[$i] & @CRLF
	Next

	Return $Re
ENDFUNC ; <== _JT_ChordMapper

; $chord   => 'C', 'Dm', 'E/F', etc...
; Return   => Array['ChordName', 'ChordExtend', 'ChordSlash']
FUNC _JT_ChordAnalyzer($chord)
	Local $InputChord = StringStripWS($chord, 7)
	Local $ChordName = '', $ChordExtend = '', $ChordSlash = ''
	Local $Re[3] = ['', '', '']
	
	If (StringMid($InputChord, 2, 1) == '#' Or StringMid($InputChord, 2, 1) == 'b') Then
		$ChordName = StringLeft($InputChord, 2)
		$ChordExtend = StringMid($InputChord, 3)
	Else
		$ChordName = StringLeft($InputChord, 1)
		$ChordExtend = StringMid($InputChord, 2)
	EndIf
	
	If (StringInStr($ChordExtend, '/') > 0 And StringInStr($ChordExtend, '/') < StringLen($ChordExtend)) Then
		$ChordSlash = StringMid($ChordExtend, StringInStr($ChordExtend, '/') + 1)
		$ChordExtend = StringLeft($ChordExtend, StringLen($ChordExtend) - StringLen($ChordSlash) - 1)
	EndIf
	$ChordExtend = StringReplace(StringReplace($ChordExtend, 'major', ''), 'minor', 'm')
	$ChordExtend = StringReplace(StringReplace($ChordExtend, 'maj', ''), 'min', 'm')
	
	Switch ($ChordName)
		Case 'E#', 'F'
			$ChordName = 'F'
		Case 'B#', 'C'
			$ChordName = 'C'
		Case 'Ab', 'G#'
			$ChordName = 'G#'
		Case 'Bb', 'A#'
			$ChordName = 'A#'
		Case 'Cb', 'B'
			$ChordName = 'B'
		Case 'Db', 'C#'
			$ChordName = 'C#'
		Case 'Eb', 'D#'
			$ChordName = 'D#'
		Case 'Fb', 'E'
			$ChordName = 'E'
		Case 'Gb', 'F#'
			$ChordName = 'F#'
		Case 'A'
			$ChordName = 'A'
		Case 'D'
			$ChordName = 'D'
		Case 'G'
			$ChordName = 'G'
		Case Else
			$ChordName = ''
	EndSwitch
	Switch ($ChordSlash)
		Case 'E#', 'F'
			$ChordSlash = 'F'
		Case 'B#', 'C'
			$ChordSlash = 'C'
		Case 'Ab', 'G#'
			$ChordSlash = 'G#'
		Case 'Bb', 'A#'
			$ChordSlash = 'A#'
		Case 'Cb', 'B'
			$ChordSlash = 'B'
		Case 'Db', 'C#'
			$ChordSlash = 'C#'
		Case 'Eb', 'D#'
			$ChordSlash = 'D#'
		Case 'Fb', 'E'
			$ChordSlash = 'E'
		Case 'Gb', 'F#'
			$ChordSlash = 'F#'
		Case 'A'
			$ChordSlash = 'A'
		Case 'D'
			$ChordSlash = 'D'
		Case 'G'
			$ChordSlash = 'G'
		Case Else
			$ChordSlash = ''
	EndSwitch
	If ($ChordSlash == $ChordName) Then $ChordSlash = ''
	
	$Re[0] = $ChordName
	$Re[1] = $ChordExtend
	$Re[2] = $ChordSlash
	
	Return $Re
ENDFUNC ; <== _JT_ChordAnalyzer

; $filenameFull => Fullpath to a file
; Return        => Content of the file
FUNC _JT_GetFileContent($filenameFull)
	Local $Re, $FileOpen = FileOpen($filenameFull, 256)
	
	If ($FileOpen == -1) Then
		Return ''
	EndIf
	
	$Re = FileRead($FileOpen)
	FileClose($FileOpen)
	Return $Re
ENDFUNC ; <== _JT_GetFileContent

; $string => A string
; Return  => A String in Latin
FUNC _JT_ToLatin($string = '')
	; Don't waste time :)
	If (StringIsSpace($string)) Then
		Return $string
	EndIf
	
	Local $CharA[18] = ['a', 'á', 'à', 'ả', 'ã', 'ạ', 'ă', 'ắ', 'ằ', 'ẳ', 'ẵ', 'ặ', 'â', 'ấ', 'ầ', 'ẩ', 'ẫ', 'ậ']
	Local $CharI[6] = ['i', 'í', 'ì', 'ỉ', 'ĩ', 'ị']
	Local $CharY[6] = ['y', 'ý', 'ỳ', 'ỷ', 'ỹ', 'ỵ']
	Local $CharE[12] = ['e', 'é', 'è', 'ẻ', 'ẽ', 'ẹ', 'ê', 'ế', 'ề', 'ể', 'ễ', 'ệ']
	Local $CharO[18] = ['o', 'ó', 'ò', 'ỏ', 'õ', 'ọ', 'ô', 'ố', 'ồ', 'ổ', 'ỗ', 'ộ', 'ơ', 'ớ', 'ờ', 'ở', 'ỡ', 'ợ']
	Local $CharU[12] = ['u', 'ú', 'ù', 'ủ', 'ũ', 'ụ', 'ư', 'ứ', 'ừ', 'ữ', 'ử', 'ự']
	Local $CharD[2] = ['d', 'đ']
	Local $Re = '', $i, $Ok = False
	
	If (StringLen($string) == 1) Then
		$Re = StringLower($string)
		; Start replacement
		For $i In $CharA
			If ($Ok) Then
				ExitLoop(1)
			EndIf
			If $i == $Re Then
				$Re = $CharA[0]
				$Ok = True
				ExitLoop(1)
			EndIf
		Next
		For $i In $CharI
			If ($Ok) Then
				ExitLoop(1)
			EndIf
			If $i == $Re Then
				$Re = $CharI[0]
				$Ok = True
				ExitLoop(1)
			EndIf
		Next
		For $i In $CharY
			If ($Ok) Then
				ExitLoop(1)
			EndIf
			If $i == $Re Then
				$Re = $CharY[0]
				$Ok = True
				ExitLoop(1)
			EndIf
		Next
		For $i In $CharE
			If ($Ok) Then
				ExitLoop(1)
			EndIf
			If $i == $Re Then
				$Re = $CharE[0]
				$Ok = True
				ExitLoop(1)
			EndIf
		Next
		For $i In $CharO
			If ($Ok) Then
				ExitLoop(1)
			EndIf
			If $i == $Re Then
				$Re = $CharO[0]
				$Ok = True
				ExitLoop(1)
			EndIf
		Next
		For $i In $CharU
			If ($Ok) Then
				ExitLoop(1)
			EndIf
			If $i == $Re Then
				$Re = $CharU[0]
				$Ok = True
				ExitLoop(1)
			EndIf
		Next
		For $i In $CharD
			If ($Ok) Then
				ExitLoop(1)
			EndIf
			If $i == $Re Then
				$Re = $CharD[0]
				$Ok = True
				ExitLoop(1)
			EndIf
		Next
		
		If (StringIsUpper($string)) Then
			Return StringUpper($Re)
		Else
			Return $Re
		EndIf
	Else ; If $string has more than 1 word
		For $i = 1 To StringLen($string)
			$Re = $Re & _JT_ToLatin(StringMid($string, $i, 1))
		Next
		Return $Re
	EndIf
ENDFUNC ; <== _JT_ToLatin

; $string  => A string
; Return => Put every word into a pare of bracket
FUNC _JT_AutoBracket($string)
	Local $i, $Re = '', $Curr = '', $Left = ''
	$string = StringReplace($string, '[', ' ')
	$string = StringReplace($string, ']', ' ')
	$string = StringReplace($string, '<', ' ')
	$string = StringReplace($string, '>', ' ')
	$string = StringReplace($string, '(', ' ')
	$string = StringReplace($string, ')', ' ')
	$string = StringReplace($string, '–', ' ')
	$string = StringReplace($string, '-', ' ')
	$string = StringReplace($string, '_', ' ')
	$string = StringReplace($string, '.', ' ')
	$string = StringReplace($string, ',', ' ')
	$string = StringReplace($string, '…', ' ')
	$string = StringReplace($string, '|', ' ')
	$string = StringReplace($string, '=', ' ')
	$string = StringReplace($string, '*', ' ')
	$string = StringReplace($string, @TAB, ' ')
	$string = StringStripWS($string, 7)
	
	$Re = '['
	
	For $i = 1 To StringLen($string)
		$Left = StringMid($string, $i-1, 1)
		$Curr = StringMid($string, $i, 1)
		If ($Curr == ' ' Or $Curr == @CR Or $Curr == @LF) Then
			$Re = $Re & '] '
		Else
			If ($Left == ' ' Or $Left == @CR Or $Left == @LF) Then
				$Re = $Re & '[' & $Curr
			Else
				$Re = $Re & $Curr
			EndIf
		EndIf
	Next
	$Re = $Re & ']'
	;Send($Re)
	ClipPut($Re)
	;MsgBox(0, 'TEMP', $Re)
ENDFUNC ; <== JT_AutoBracket
