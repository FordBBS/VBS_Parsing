Function IBase_getinfo_resParamValue(ByVal strValue)
	'*** History ***********************************************************************************
	' 2020/08/26, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	'	Return Value and TagIndex information of 'strValue'
	'	e.g. strValue = "%tag%0;1;4%tag%SULEV"
	'		 return ("SULEV", "0;1;4")
	'
	'	Argument(s)
	'	<String> strValue, A string of single value created by 'IUser_translate_json_strContent'
	' 					   If more than one value exist, only first value will be manipulated
	'
	'***********************************************************************************************

	On Error Resume Next
	IBase_getinfo_resParamValue = Array("", "")

	'*** Pre-Validation ****************************************************************************
	strValue   = CStr(strValue)
	If len(strValue) = 0 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim arrValue, tagBrnIdx, thisValue, thisTagIdx

	tagBrnIdx  = "%tag%"
	arrValue   = Split(strValue, "%;%")

	'*** Operations ********************************************************************************
	If InStr(arrValue(0), tagBrnIdx) > 0 Then 
		thisValue  = Mid(arrValue(0), len(tagBrnIdx) + 1, len(arrValue(0)))
		thisTagIdx = Mid(thisValue, 1, InStr(thisValue, tagBrnIdx) - 1)
		thisValue  = Mid(thisValue, InStr(thisValue, tagBrnIdx) + len(tagBrnIdx), len(thisValue))
		IBase_getinfo_resParamValue = Array(thisValue, thisTagIdx)
	End If

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function