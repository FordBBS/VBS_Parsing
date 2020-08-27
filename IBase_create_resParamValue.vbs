Function IBase_create_resParamValue(ByVal strValue)
	'*** History ***********************************************************************************
	' 2020/08/26, BBS:	- First Release
	' 2020/08/27, BBS:	- bug fixed when 'strValue' has only one level
	'					- bug fixed, invalid If-Else condition for creating result for new branch
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' 	Return a general form Value of 'strValue'
	' 	e.g.  strValue = "%tag%0;0%tag%SULEV%;%%tag%0;1%tag%CONT_BAG%;%%tag%1;0%tag%CONT_BAG_THC"
	' 		  Return (("SULEV", "CONT_BAG"), ("CONT_BAG_THC"))
	'
	' 		  strValue = "%tag%0;0;0%tag%Modal%;%%tag%0;0;1%tag%Bag%;%%tag%0;1;0%tag%THC"
	' 		  Return ((("Modal", "Bag"), ("THC")))
	'	
	'	Argument(s)
	'	<String> strValue, A string value created by 'IUser_translate_json_strContent'
	'
	'***********************************************************************************************

	On Error Resume Next
	IBase_create_resParamValue = Array()

	'*** Pre-Validation ****************************************************************************
	strValue = CStr(strValue)
	If len(strValue) = 0 Then Exit Function

	'*** Initialization ****************************************************************************
	Dim cnt1, cnt2, cnt_level, thisInfo, thisValue, flg_create, tmpValue
	Dim arrValue, arrTagIdx, arrThis, arrTmpSnap, arrSnapObj, arrSnapIdx, arrBase(), arrRet()
	Redim Preserve  arrBase(0), arrRet(0)

	arrValue = Split(strValue, "%;%")

	'*** Operations ********************************************************************************
	For cnt1 = 0 to UBound(arrValue)
		thisInfo   = IBase_getinfo_resParamValue(arrValue(cnt1))
		thisValue  = thisInfo(0) 
		arrTagIdx  = Split(thisInfo(1), ";")
		flg_create = False

		If cnt1 = 0 Then
			flg_create = True
			cnt_level  = UBound(arrTagIdx) - 1
		Else
			arrThis    = arrRet
			arrSnapObj = Array()
			arrSnapIdx = Array()

			For cnt2 = 0 to UBound(arrTagIdx)
				Call hs_arr_append(arrSnapIdx, CInt(arrTagIdx(cnt2))) 	'Snaps target index

				If UBound(arrThis) < CInt(arrTagIdx(cnt2)) Then 		'Branch is needed now
					cnt_level = UBound(arrTagIdx) - cnt2

					If cnt_level > 0 Then
						flg_create = True
					Else
						Call hs_arr_append(arrThis, thisValue)
					End If

					Exit For
				Else
					' Snapshot of thisLevel's object(s)
 					Call hs_arr_append(arrSnapObj, arrThis)

 					' Update 'arrThis' for next level validation
 					tmpValue = arrThis(CInt(arrTagIdx(cnt2)))
 					arrThis  = tmpValue
				End If
			Next
		End If

		' Create Array of thisValue for the new branch case
		If flg_create Then
			' Create base array
			Erase arrBase
			Redim Preserve arrBase(0)

			If UBound(arrTagIdx) < 1 Then
				arrThis = thisValue
			Else
				For cnt2 = 0 to (CInt(arrTagIdx(UBound(arrTagIdx))) - 1)
					Call hs_arr_append(arrBase, "")
				Next

				Call hs_arr_append(arrBase, thisValue)
				arrThis = arrBase
			End If
		End If

		' Compelete Result appending
		If cnt1 = 0 Then
			For cnt2 = 0 to (CInt(arrTagIdx(0)) - 1)
				Call hs_arr_append(arrRet, "")
			Next

			Call hs_arr_append(arrRet, arrThis)
			Call hs_arr_stack(arrThis, cnt_level)
		Else
			If UBound(arrSnapObj) > -1 Then
				' New branch was created, arrSnapObj and arrSnapIdx are needed to be prepared first
				If cnt_level > 0 Then
					arrTmpSnap = arrSnapObj(UBound(arrSnapObj))(arrSnapIdx(UBound(arrSnapIdx) - 1))
					Call hs_arr_append(arrTmpSnap, arrThis)
					Call hs_arr_append(arrSnapObj, arrTmpSnap)
				End If

				' Complete Result array inside out
				For cnt2 = UBound(arrSnapObj) to 0 Step -1
					arrTmpSnap = arrSnapObj(cnt2)

					If cnt2 = UBound(arrSnapObj) Then
						arrTmpSnap(arrSnapIdx(cnt2)) = arrThis
					Else
						arrTmpSnap(arrSnapIdx(cnt2)) = arrSnapObj(cnt2 + 1)
					End If

					arrSnapObj(cnt2) = arrTmpSnap
				Next

				' Load Result array which are prepared in 'arrSnapObj' level 0 to 'arrRet'
				For cnt2 = 0 to UBound(arrSnapObj(0))
					arrRet(cnt2) = arrSnapObj(0)(cnt2)
				Next
			ElseIf UBound(arrTagIdx) = 0 Then
				Erase arrRet
				Redim Preserve arrRet(UBound(arrThis))
				
				For cnt2 = 0 to UBound(arrRet)
					arrRet(cnt2) = arrThis(cnt2)
				Next
			Else
				Call hs_arr_stack(arrThis, cnt_level - 1)
				Call hs_arr_append(arrRet, arrThis)
			End If
		End If
	Next

	IBase_create_resParamValue = arrRet

	'*** Error handler *****************************************************************************
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function