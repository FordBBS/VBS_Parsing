Function IUser_clean_ParamPath(ByVal arrParamPath, ByVal arrReplaceTag)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First Release
	' 					- Improved handler of 'arrReplaceTag', If only one replace guide is needed,
	' 					it can be provided as Array("xxxxxx", "yyyyy")
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Replace target Parameter tag in 'Parameter Path' with specific string
	' 	e.g. "config_params.ecs" -> "ecs"
	' 		 "gmd.table.range" 	 -> "gmd.analyzerrange"
	'
	'***********************************************************************************************
	On Error Resume Next
	IUser_clean_ParamPath = arrParamPath

	'*** Pre-Validation ****************************************************************************
	If InStr(LCase(TypeName(arrReplaceTag)), "variant") = 0 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, cnt2, thisParam, thisReplaceGuide, flg_not_arr

	flg_not_arr = 0

	'*** Operations ********************************************************************************
	'--- Prepare 'arrReplaceTag' -------------------------------------------------------------------
	For cnt1 = 0 to UBound(arrReplaceTag)
		If InStr(LCase(TypeName(arrReplaceTag(cnt1))), "variant") = 0 Then
			flg_not_arr = flg_not_arr + 1
		End If
	Next

	If UBound(arrReplaceTag) = 1 and flg_not_arr > 0 Then
		arrReplaceTag = Array(arrReplaceTag)
	End If

	'--- Cleaning ----------------------------------------------------------------------------------
	For cnt1 = 0 to UBound(arrParamPath)
		thisParam = "." & arrParamPath(cnt1) & "."

		For cnt2 = 0 to UBound(arrReplaceTag)
			thisReplaceGuide = arrReplaceTag(cnt2)

			If InStr(LCase(TypeName(thisReplaceGuide)), "variant") > 0 Then
				If InStr(thisParam, "." & thisReplaceGuide(0) & ".") > 0 Then
					thisParam = Replace(thisParam, thisReplaceGuide(0), thisReplaceGuide(1))
				End If
			End If
		Next

		thisParam = Replace(thisParam, "..", ".")
		
		If Left(thisParam, 1) = "." Then thisParam = Mid(thisParam, 2, len(thisParam))
		If Right(thisParam, 1) = "." Then thisParam = Left(thisParam, len(thisParam) - 1)

		arrParamPath(cnt1) = thisParam
	Next

	IUser_clean_ParamPath = arrParamPath

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function