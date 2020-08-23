Function IUser_get_value_of_param(ByVal arrParamValue, ByVal strTargetParam, ByVal flg_case)
	'*** History ***********************************************************************************
	' 2020/08/23, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return the value of target Parameter path
	' 	Return "NotExist" if target Parameter path doesn't exist
	'	
	'	Argument(s)
	'	<Array> 	arrParamValue,  Parameter-Value Array
	'	<String> 	strTargetParam, Target parameter path
	' 	<Long> 		flg_case, 		0: Character's case doesn't matter, 1: Vice versa
	'
	'***********************************************************************************************
	On Error Resume Next
	IUser_get_value_of_param = "NotExist"

	'*** Pre-Validation ****************************************************************************
	If InStr(LCase(TypeName(arrParamValue)), "variant") = 0 Then Exit Function
	If UBound(arrParamValue) < 1 Then Exit Function
	If len(CStr(strTargetParam)) = 0 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, arrParam, arrValue, thisParam

	strTargetParam = CStr(strTargetParam)
	arrParam 	   = arrParamValue(0)
	arrValue 	   = arrParamValue(1)

	If LCase(TypeName(flg_case)) <> "integer" Then flg_case = 1
	If flg_case < 0 or flg_case > 1 Then flg_case = 1
	If flg_case = 0 Then strTargetParam = LCase(strTargetParam)

	'*** Operations ********************************************************************************
	For cnt1 = 0 to UBound(arrParam)
		thisParam = arrParam(cnt1)
		
		If flg_case = 0 Then thisParam = LCase(thisParam)
		If strTargetParam = thisParam Then
			IUser_get_value_of_param = arrValue(cnt1)
			Exit For
		End If
	Next

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
