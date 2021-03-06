Option Explicit

Function IUser_translate_json_strContent(ByVal strContent)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First Release
	' 2020/08/27, BBS:	- Bug fixed
	' 					1) Array tag is not removed when it's only one tag left
	'					2) Parameter path ends with "." for any value that has array value line
	' 2020/09/21, BBS:	- Bug fixed, Array branching is not closed correctly
	' 2020/10/07, BBS:	- Bug fixed, Empty parameter is not translated correctly
	' 2020/12/11, BBS:	- Implemented conditioner before release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return Parameters-Values array of provided JSON content in 'strContent' String type
	'	e.g. (("ecs.devicename", "ecs.activate"), ("CVS_SL", "Yes"))
	' 		 (("gmd.devicename"), ("%tag%0%tag%SULEV%;%%tag%1%tag%CONT_BAG"))
	'
	'***********************************************************************************************
	
	On Error Resume Next
	IUser_translate_json_strContent = Array(Array(), Array())

	'*** Pre-Validation ****************************************************************************
	If TypeName(strContent) <> "String" Then Exit Function
	If len(strContent) < 3 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, cnt_row, flg_append, tagBranch, tagLabel, tagLatest, thisParam, thisValue, existValue
	Dim curRoot, curPath, strTagArray, strTagValue, strTagRemove, strParamEx, strBrnTag, strBrnIdx
	Dim flg_clr_rt, arrContent, arrThisInfo, arrRoot, arrBrnIdx, arrParam(), arrValue()
	Redim Preserve arrParam(0), arrValue(0)

	arrContent  = Split(strContent, vbCrLf)
	cnt_row     = 0
	tagBranch   = "%tag%"
	curRoot 	= ""
	strTagArray = ""				' Storage: Parameter that has branch
	strTagValue = "" 				' Storage: Parameter that has value on its line
	strParamEx  = "" 				' Storage: Appended Parameter path
	strBrnTag   = "" 				' Storage: Branch, Parameter owner of each position index
	strBrnIdx   = "" 				' Storage: Branch, Position index
	
	'*** Operations ********************************************************************************
	'--- Parsing -----------------------------------------------------------------------------------
	while cnt_row < UBound(arrContent)
		arrThisInfo = IBase_get_value_from_strLine(arrContent(cnt_row), ":")
		thisParam   = arrThisInfo(0)
		thisValue	= CStr(arrThisInfo(1))
		flg_append	= False
		flg_clr_rt  = False

		If thisParam <> "" Then 	' Case: Parameter does exist
			' Case: Empty Parameter, No value, No SubGroup
			If InStr(arrContent(cnt_row), "{}") > 0 or InStr(arrContent(cnt_row), "[]") > 0 Then
				flg_append = True

			' Case: Value doesn't exist but SubGroup's or Array's symbol
			ElseIf InStr(arrContent(cnt_row), "{") > 0 or InStr(arrContent(cnt_row), "[") > 0 Then
				If curRoot <> "" Then
					curRoot = Join(Array(curRoot, thisParam), ".")
				Else
					curRoot = thisParam
				End If
				
				If strTagValue <> "" Then
					strTagValue = Join(Array(strTagValue, "%" & thisParam & "%"), ";")
				Else
					strTagValue = "%" & thisParam & "%"
				End If
				
				If InStr(arrContent(cnt_row), "[") > 0 Then
					If strTagArray <> "" Then
						strTagArray = Join(Array(strTagArray, "%" & thisParam & "%"), ";")
					Else
						strTagArray = "%" & thisParam & "%"
					End If

					If strBrnTag = "" Then
						strBrnTag = "%" & thisParam & "%"
						strBrnIdx = "0"
					Else
						strBrnTag = Join(Array(strBrnTag, "%" & thisParam & "%"), ";")
						strBrnIdx = Join(Array(strBrnIdx, "0"), ";")
					End If
				End If
			
			' Case: Value does exist
			Else
				flg_append = True
			End If
		ElseIf thisParam = "" Then 	' Case: Parameter doesn't exist
			' Case: End of current sub-tags group '{}', clear Root, and stored tags string
			If InStr(arrContent(cnt_row), "}") > 0 and len(curRoot) > 0 Then
				flg_clr_rt = True

			' Case: End of latest branch, clear all memo info of latest branch
			ElseIf InStr(arrContent(cnt_row), "]") > 0 Then
				arrRoot 	 = Split(curRoot, ".")
				arrBrn 		 = Split(strTagArray, ";")
				strTagRemove = Replace(arrBrn(UBound(arrBrn)), "%", "")

				If arrRoot(UBound(arrRoot)) = strTagRemove Then
					flg_clr_rt = True
				End If
	
				If InStrRev(strTagArray, ";") > 0 Then
					strTagArray = Mid(strTagArray, 1, InStrRev(strTagArray, ";") - 1)
				Else
					strTagArray = ""
				End If

				If InStrRev(strBrnTag, ";") > 0 Then
					strBrnTag = Mid(strBrnTag, 1, InStrRev(strBrnTag, ";") - 1)
					strBrnIdx = Mid(strBrnIdx, 1, InStrRev(strBrnIdx, ";") - 1)
				Else
					strBrnTag = ""
					strBrnIdx = ""
				End If

			' Case: Value line (e.g. Parameter that has array value will break its value into lines)
			ElseIf InStr(arrContent(cnt_row), "{") = 0 and InStr(arrContent(cnt_row), "[") = 0 Then
				flg_append  = True
			End If
		End If

		' Root Removal
		If flg_clr_rt Then
			strTagRemove = Mid(curRoot, InStrRev(curRoot, ".") + 1, len(curRoot))

			If InStr(strTagArray, "%" & strTagRemove & "%") = 0 and _
			 	InStr(strTagValue, "%" & strTagRemove & "%") > 0 Then

			 	If InStr(curRoot, ".") > 0 Then
			 		curRoot = Mid(curRoot, 1, InStrRev(curRoot, ".") - 1)
			 	Else
			 		curRoot = ""
			 	End If

			 	If InStr(strTagValue, ";") > 0 Then
			 		strTagValue = Mid(strTagValue, 1, InStrRev(strTagValue, ";") - 1)
			 	Else
			 		strTagValue = ""
			 	End If
			End If
		End If

		' Appending
		If flg_append Then
			If thisParam = "" Then
				curPath = curRoot
			ElseIf curRoot <> "" Then
				curPath = Join(Array(curRoot, thisParam), ".")
			Else
				curPath = thisParam
			End If

			' Case: Current Parameter path already exist
			If InStr(strParamEx, "%" & curPath & "%") > 0 Then
				' Get index of this Parameter path and existing value
				For cnt1 = 0 to UBound(arrParam)
					If arrParam(cnt1) = curPath Then
						existValue = arrValue(cnt1)
						Exit For
					End If
				Next

				'Appending - Branch check then append
				If InStrRev(existValue, "%;%") > 0 Then
					tagLatest = Mid(existValue, InStrRev(existValue, "%;%") + 3, len(existValue))
				Else
					tagLatest = existValue
				End If
				
				tagLatest = Mid(tagLatest, len(tagBranch) + 1, _
					 						InStrRev(tagLatest, tagBranch) - len(tagBranch) - 1)

				If tagLatest = strBrnIdx Then
					arrBrnIdx = Split(strBrnIdx, ";")
					arrBrnIdx(UBound(arrBrnIdx)) = CStr(CInt(arrBrnIdx(UBound(arrBrnIdx))) + 1)
					strBrnIdx = Join(arrBrnIdx, ";")
				End If

				arrValue(cnt1) = existValue & "%;%" & tagBranch & strBrnIdx & tagBranch & thisValue
				
			' Case: Current Parameter path has its first time appending
			Else
				' Store current parameter path
				If strParamEx <> "" Then
					strParamEx = Join(Array(strParamEx, "%" & curPath & "%"), ";")
				Else
					strParamEx = "%" & curPath & "%"
				End If

				' Prepare proper size for result arrays
				If Not (UBound(arrParam) = 0 and len(arrParam(0)) = 0) Then
					Redim Preserve arrParam(UBound(arrParam) + 1), arrValue(UBound(arrValue) + 1)
				End If

				' Create branch for 'thisValue' if it's necessary
				If strBrnIdx <> "" Then thisValue = tagBranch & strBrnIdx & tagBranch & thisValue	

				' Append Parameter path and its Value
				arrParam(UBound(arrParam)) = curPath
				arrValue(UBound(arrValue)) = thisValue
			End If
		End If

		' Release current line
		If cnt_row < 0 Then
			cnt_row = UBound(arrContent)
		Else
			cnt_row = cnt_row + 1
		End If
	wend
	
	'--- Translate 'arrValue' ----------------------------------------------------------------------
	For cnt1 = 0 to UBound(arrValue)
		If InStr(arrValue(cnt1), "%tag%") > 0 Then
			existValue 	   = IUser_translate_resParamValue(arrValue(cnt1), "")
			arrValue(cnt1) = existValue
		End If
	Next

	'--- Result's Conditioning ---------------------------------------------------------------------
	Call hs_parser_remove_redundant_params(arrParam, arrValue)
	
	'--- Release -----------------------------------------------------------------------------------
	IUser_translate_json_strContent = Array(arrParam, arrValue)

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
