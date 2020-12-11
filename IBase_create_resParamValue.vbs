Option Explicit

Function IBase_create_resParamValue(ByVal strValue)
	'*** History ***********************************************************************************
	' 2020/08/26, BBS:	- First Release
	' 2020/08/27, BBS:	- Bug fixed, when 'strValue' has only one level
	'					- Bug fixed, invalid If-Else condition for creating result for new branch
	' 2020/12/11, BBS:	- Overhaul mechanism
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
	'		  strValue = "%tag%0;0%tag%SULEV%;%%tag%0;1%tag%CONT_BAG%;%%tag%2;0%tag%CONT_BAG_THC"
	' 		  Return (("SULEV", "CONT_BAG"), (), ("CONT_BAG_THC"))
	'	
	'	Argument(s)
	'	<String> strValue, A string value created by 'IUser_translate_json_strContent'
	'
	'***********************************************************************************************

	On Error Resume Next
	IBase_create_resParamValue = Array()

	'*** Pre-Validation ****************************************************************************
	strValue = CStr(strValue)
	If Len(strValue) = 0 Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim cnt1, cnt2, cnt3, stack_lvl, corr_lvl, thisInfo, thisValue, flg_create
	Dim arrRet, arrValue, arrTagIdx, arrPrep, arrTmp, arrSnapIdx, arrSnapObj, arrSnapTmp
	Dim thisTag, tarSnapObj, tarSnapIdx
	
	arrValue = Split(strValue, "%;%")
	arrRet 	 = Array()

	'*** Operations ********************************************************************************
	For cnt1 = 0 to UBound(arrValue)
		thisInfo   = IBase_getinfo_resParamValue(arrValue(cnt1))
		thisValue  = thisInfo(0) 
		arrTagIdx  = Split(thisInfo(1), ";")
		flg_create = False
		flg_snap   = False
		arrSnapIdx = Array()
		arrSnapObj = Array()

		'--- Analysis of 'thisValue' information ---------------------------------------------------
		' Case: First value, skips snapshot analysis
		If cnt1 = 0 Then
			flg_create 	= True
			stack_lvl	= UBound(arrTagIdx) - 1		'Set needed stack level
			corr_lvl 	= 0							'Set start level of correction
		
		' Case: General, performs snapshot analysis
		Else
			arrTmp = arrRet

			For cnt2 = 0 to UBound(arrTagIdx)
				thisTag = CInt(arrTagIdx(cnt2))
				Call hs_arr_append(arrSnapIdx, thisTag)

				' Case: Result array has no target position yet
				If UBound(arrTmp) < thisTag Then
					stack_lvl  = UBound(arrTagIdx) - cnt2 - 1	'Set needed stack level
					corr_lvl   = cnt2 							'Set start level of correction
					flg_create = True
					Exit For

				' Case: Result array covers target position
				Else
					Call hs_arr_append(arrSnapObj, arrTmp) 	'Snapshot

					' Prepared next iteration
					arrSnapTmp = arrTmp(thisTag)
					arrTmp     = arrSnapTmp
				End If
			Next
		End If

		'--- Create new base array for 'thisValue' -------------------------------------------------
		If flg_create Then
			If UBound(arrTagIdx) > 0 and UBound(arrSnapObj) < 0 Then
				arrPrep = Array()

				For cnt2 = corr_lvl to (CInt(arrTagIdx(UBound(arrTagIdx))) - 1)
					Call hs_arr_append(arrPrep, "")
				Next
				
				Call hs_arr_append(arrPrep, thisValue)
				Call hs_arr_stack(arrPrep, stack_lvl)
			Else
				arrPrep = thisValue
			End If
		End If

		'--- Manipulate return array ---------------------------------------------------------------
		' Method: Snapshot Restoration
		If UBound(arrSnapObj) >= 0 Then
			Call hs_arr_append(arrSnapObj, arrPrep) 		'Append prepared value as last Snapshot

			For cnt2 = UBound(arrSnapObj) to 1 Step -1 		'Snapshot Restoration process
				tarSnapObj = arrSnapObj(cnt2 - 1)
				tarSnapIdx = arrSnapIdx(cnt2 - 1)
				arrTmp 	   = tarSnapObj(tarSnapIdx)

				For cnt3 = UBound(arrTmp) to (CInt(arrSnapIdx(cnt2)) - 2)
					Call hs_arr_append(arrTmp, "")
				Next
				Call hs_arr_append(arrTmp, arrSnapObj(cnt2))

				tarSnapObj(tarSnapIdx) = arrTmp
				arrSnapObj(cnt2 - 1)   = tarSnapObj
			Next

			arrRet = arrSnapObj(0) 			'Set top level of Snapshot as current Result

		' Method: Direct Appending
		Else
			For cnt2 = UBound(arrRet) to (CInt(arrTagIdx(0)) - 2)
				Call hs_arr_append(arrRet, "")
			Next
			Call hs_arr_append(arrRet, arrPrep)
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	IBase_create_resParamValue = arrRet

	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
