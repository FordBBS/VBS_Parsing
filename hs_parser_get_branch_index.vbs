Option Explicit

Function hs_parser_get_branch_index(ByVal objValue)
	'*** History ***********************************************************************************
	' 2020/12/12, BBS:	- First Release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Parser helper, Return Branch Index instruction object based on 'objValue'
	' e.g. 	objValue = {"CVS"}, 						RetVal = 0
	'		objValue = {{"SULEV", "CONT_BAG"}}			RetVal = 0%;%1
	'		objValue = {{"Bag", "Tunnel", "Diluted"}}	RetVal = 0%;%2
	'		objValue = {{"SULEV", "CONT_BAG"}, "", {"CONT_BAG_THC"}, {{"A", "B"}}}
	'		RetVal	 = 3%;%1;-1;0;0%;%-1;-1;-1;-1;1%;%-1;-1;-1;-1;-1;-1
	'  
	'***********************************************************************************************

	On Error Resume Next
	hs_parser_get_branch_index = -1

	'*** Pre-Validation ****************************************************************************
	If Not IsArray(objValue) Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim delim_lvl, cnt1, cnt2, cnt3, cnt_step, cnt_pos, arrTmp, arrBrnIdx, arrPrevIdx, arrCurrIdx
	Dim thisBrnIdx, thisStep

	delim_lvl = "%;%"
	arrBrnIdx = Array(UBound(objValue))

	'*** Operations ********************************************************************************
	'--- Create Branch Index Array -----------------------------------------------------------------
	Call hs_arr_append(arrBrnIdx, "")

	For cnt1 = 0 to UBound(objValue)
		If IsArray(objValue(cnt1)) Then
			thisBrnIdx = hs_parser_get_branch_index(objValue(cnt1)) 	'Recursives itself
			arrTmp 	   = Split(thisBrnIdx, delim_lvl)

			For cnt2 = 0 to UBound(arrTmp)
				If UBound(arrBrnIdx) < (cnt2 + 1) Then
					Call hs_arr_append(arrBrnIdx, "")
				End If

				If arrBrnIdx(cnt2 + 1) = "" Then
					arrBrnIdx(cnt2 + 1) = arrTmp(cnt2)
				Else
					arrBrnIdx(cnt2 + 1) = arrBrnIdx(cnt2 + 1) & ";" & arrTmp(cnt2)
				End If
			Next

		ElseIf arrBrnIdx(1) = "" Then
			arrBrnIdx(1) = "-1"
		Else	
			arrBrnIdx(1) = arrBrnIdx(1) & ";-1"
		End If
	Next

	'--- Branch correction for Non-Array datatype to have a correct reference position -------------
	For cnt1 = 1 to UBound(arrBrnIdx)
		arrPrevIdx = Split(arrBrnIdx(cnt1 - 1), ";")
		arrCurrIdx = Split(arrBrnIdx(cnt1), ";")
		thisBrnIdx = ""
		cnt_pos	   = 0
		
		For cnt2 = 0 to UBound(arrPrevIdx)
			thisStep = CInt(arrPrevIdx(cnt2))

			If thisStep = -1 Then 			'Non-Array Datatype
				thisBrnIdx = thisBrnIdx & ";-1"
			Else
				For cnt3 = cnt_pos to (cnt_pos + thisStep)
					thisBrnIdx = thisBrnIdx & ";" & arrCurrIdx(cnt3)
				Next
				cnt_pos = cnt_pos + thisStep + 1
			End If
		Next

		arrBrnIdx(cnt1) = Mid(thisBrnIdx, 2) 		'Push corrected Branch Index to result array
	Next

	'--- Assemble Branch Index results into String -------------------------------------------------
	hs_parser_get_branch_index = Join(arrBrnIdx, delim_lvl)

	'--- Release -----------------------------------------------------------------------------------
	If Err.Number <> 0 Then
		print("Error number " & Err.Number)
		Err.Clear
	End If
End Function
