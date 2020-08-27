Function IUser_translate_resParamValue(ByVal strValue, ByVal strIdxList)
	'*** History ***********************************************************************************
	' 2020/08/27, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' 	Return translated value of 'strValue' based on requested index in 'strIdxList'
	' 	e.g.  strValue 	   = "%tag%0;0%tag%SULEV%;%%tag%0;1%tag%CONT_BAG%;%%tag%1;0%tag%SULEV"
	'		  strIdxList = "0"	-> return ("SULEV", "CONT_BAG")
	' 		  strIdxList = "1" 	-> return ("SULEV")
	'' 		  strIdxList = "0;1"  -> return ("CONT_BAG")
	'	
	'	Argument(s)
	'	<String> strValue,   A string value created by 'IUser_translate_json_strContent'
	'	<String> strIdxList, A string list of desire index/indices
	'
	'***********************************************************************************************

	On Error Resume Next
	IUser_translate_resParamValue = Array()

	'*** Pre-Validation ****************************************************************************
	strValue   = CStr(strValue)
	strIdxList = CStr(strIdxList)
	If Not (len(strValue) > 0 and len(strIdxList) > 0) Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, tagBrnIdx, tmpValue
	Dim arrRes, arrValue, arrTarIdx, arrSplit

	tagBrnIdx = "%tag%"
	arrSplit  = Split(strValue, "%;%")
	arrTarIdx = Split(strIdxList, ";")
	arrValue  = IBase_create_resParamValue(strValue)

	'*** Operations ********************************************************************************
	'--- Post-Validation ---------------------------------------------------------------------------
	If UBound(arrSplit) < 0 Then Exit Function

	'--- Collect data based on target index list ---------------------------------------------------
	For cnt1 = 0 to UBound(arrTarIdx)
		If UBound(arrValue) < CInt(arrTarIdx(cnt1)) Then
			Exit Function
		Else
			tmpValue = arrValue(CInt(arrTarIdx(cnt1)))
			arrValue = tmpValue
		End If
	Next
	
	'--- Release -----------------------------------------------------------------------------------
	If InStr(LCase(TypeName(arrValue)), "variant") > 0 Then
		IUser_translate_resParamValue = arrValue
	Else
		IUser_translate_resParamValue = Array(arrValue)
	End If

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
