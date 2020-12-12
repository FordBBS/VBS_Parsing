Option Explicit

Function hs_parser_get_root_of_params(ByVal paramA, ByVal paramB)
	'*** History ***********************************************************************************
	' 2020/12/12, BBS:	- First Release
	'
	'***********************************************************************************************

	'*** Documentation *****************************************************************************
	' Parser helper, Return the deepest common parameter among 'paramA' and 'paramB'
	'  
	'***********************************************************************************************

	On Error Resume Next
	hs_parser_get_root_of_params = ""

	'*** Pre-Validation ****************************************************************************
	paramA = CStr(paramA)
	paramB = CStr(paramB)

	If Len(paramA) = 0 or Len(paramB) = 0 Then
		Exit Function
	End If

	'*** Initialization ****************************************************************************
	Dim idx, n_size, arrParamA, arrParamB, arrRet()

	'*** Operations ********************************************************************************
	'--- Exception, Base cases ---------------------------------------------------------------------
	' ParamA is a subset of ParamB or ParamA is the same as ParamB
	If InStr(LCase(paramB), LCase(paramA)) > 0 Then
		hs_parser_get_root_of_params = paramA
		Exit Function
	End If

	' ParamB is a subset of ParamA
	If InStr(LCase(paramA), LCase(paramB)) > 0 Then
		hs_parser_get_root_of_params = paramB
		Exit Function
	End If

	'--- Create Parameter Arrays ------------------------------------------------------------------- 
	arrParamA = Split(paramA, ".")
	arrParamB = Split(paramB, ".")
	
	'--- Get minimum size --------------------------------------------------------------------------
	n_size = UBound(arrParamA)
	
	If n_size > UBound(arrParamB) Then
		n_size = UBound(arrParamB)
	End If

	'--- Get largest index position ----------------------------------------------------------------
	For idx = 0 to n_size
		If arrParamA(idx) <> arrParamB(idx) Then
			Exit For
		ElseIf hs_parser_get_root_of_params = "" Then
			hs_parser_get_root_of_params = arrParamA(idx)
		Else
			hs_parser_get_root_of_params = hs_parser_get_root_of_params & "." & arrParamA(idx)
		End If
	Next

	'--- Release -----------------------------------------------------------------------------------
	If Err.Number <> 0 Then
		Err.Clear
	End If
End Function
