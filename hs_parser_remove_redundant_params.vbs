Option Explicit

Sub hs_parser_remove_redundant_params(ByRef arrOrder, ByRef arrValue)
	'*** History ***********************************************************************************
	' 2020/12/12, BBS:	- First Release
	'
	'***********************************************************************************************
	
	'*** Documentation *****************************************************************************
	' Parser helper, Remove redundant parameters
	'
	'***********************************************************************************************
	
	On Error Resume Next

	'*** Pre-Validation ****************************************************************************
	If Not (IsArray(arrParam) and IsArray(arrValue)) Then
		Exit Sub
	End If

	'*** Initialization ****************************************************************************
	Dim idx1, arrCondParam, arrCondValue, objIdx, flg_val

	'*** Operations ********************************************************************************
	For idx1 = 0 to UBound(arrOrder)
		flg_val = False

		If idx1 < UBound(arrOrder) Then
			objIdx = hs_arr_val_exist_ex(hs_arr_slice(arrOrder, idx1 + 1, -1), arrOrder(idx1), 1, False)
		Else
			objIdx = hs_arr_val_exist_ex(hs_arr_slice(arrOrder, -2, 0), arrOrder(idx1), 1, False)
		End If

		If IsArray(objIdx) Then
			If UBound(objIdx) = 0 and IsNumeric(CStr(objIdx(0))) Then
				If CInt(objIdx(0)) < 0 Then
					flg_val = True
				End If
			End If
		ElseIf IsNumeric(CStr(objIdx)) Then
			If CInt(objIdx) < 0 Then
				flg_val = True
			End If
		End If

		If flg_val Then
			Call hs_arr_append(arrCondParam, arrOrder(idx1))
			Call hs_arr_append(arrCondValue, arrValue(idx1))
		End If
	Next

	arrOrder = arrCondParam
	arrValue = arrCondValue

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Sub
