Option Explicit

Class Vector
	Private dblEpsilon

	Public _
		conRewriteError, _
		conDimensionMismatchError, _
		conNonNumericError, _
		conDivideByZeroError, _
		conLengthIsZeroError

	Private avarValues()
	Private lngLength, varNorm

	Private boolReadOnly

	Private Sub Class_Initialize()
		dblEpsilon = 1e-7
		boolReadOnly = False

		conRewriteError = vbObjectError
		conDimensionMismatchError = vbObjectError + 1
		conNonNumericError = vbObjectError + 2
		conDivideByZeroError = vbObjectError + 3
		conLengthIsZeroError = vbObjectError + 4
	End Sub

	Public Property Let Values(ByRef avarRawVector)
		If Not boolReadOnly Then
			If Not IsArray(avarRawVector) Then
				Err.Raise conNonNumericError, "Vector", "The Input must be an array of numbers."
			End If
			On Error Resume Next
			Call UBound(avarRawVector, 1)
			If Err.Number <> 0 Then
				On Error GoTo 0
				Err.Raise _
					conLengthIsZeroError, _
					"Vector", _
					"Input Array is empty."
			End If
			On Error GoTo 0
			
			lngLength = _
				UBound(avarRawVector) - _
				LBound(avarRawVector) + 1
			
			If lngLength = 0 Then
				Err.Raise _
					conLengthIsZeroError, _
					"Vector", _
					"Input Array is empty."
			End If

			Dim lngIndex
			ReDim avarValues(UBound(avarRawVector))
			
			varNorm = 0
			For lngIndex = LBound(avarRawVector) To UBound(avarRawVector)
				If Not IsNumeric(avarRawVector(lngIndex)) Then
					Err.Raise _
						conNonNumericError, _
						"Vector", _
						"Array contains non-numeric value(s)."
				End If
				avarValues(lngIndex) = CDbl(avarRawVector(lngIndex))
				varNorm = varNorm + avarValues(lngIndex) ^ 2
			Next
			varNorm = Sqr(varNorm)

			boolReadOnly = True
		Else
			Err.Raise conRewriteError, _
				"Vector", _
				"Values property is read only."
		End If
	End Property
	
	Public Property Get Values()
		Values = avarValues
	End Property
	
	Public Property Get Value(ByRef lngIndex)
		Value = avarValues(lngIndex)
	End Property
	
	Public Property Get Length()
		Length = lngLength
	End Property

	Public Property Get Stringify()
		Stringify = "[" + Join(avarValues, " ") + "]"
	End Property

	Public Function Add(ByRef objAnotherVector)
		Dim avarResult()
		If lngLength = objAnotherVector.Length Then
			ReDim avarResult(lngLength - 1)
			Dim lngIndex
			For lngIndex = LBound(avarResult) To UBound(avarResult)
				avarResult(lngIndex) = _
					avarValues(lngIndex) + objAnotherVector.Value(lngIndex)
			Next
			Set Add = New Vector
			Add.Values = avarResult
		Else
			Err.Raise _
				conDimensionMismatchError, _
				"Vector", _
				"Two Vector's length is not equal."
		End If
	End Function

	Public Function Negate()
		Dim avarResult()
		ReDim avarResult(lngLength - 1)
		Dim lngIndex
		For lngIndex = LBound(avarResult) To UBound(avarResult)
			avarResult(lngIndex) = -avarValues(lngIndex)
		Next
		Set Negate = New Vector
		Negate.Values = avarResult
	End Function

	Public Function Subtract(ByRef objAnotherVector)
		If lngLength = objAnotherVector.Length Then
			Set Subtract = Add(objAnotherVector.Negate())
		Else
			Err.Raise _
				conDimensionMismatchError, _
				"Vector", _
				"Two Vector's length is not equal."
		End If
	End Function

	Public Property Get Norm()
		Norm = varNorm
	End Property

	Private Function IsZero(ByRef varValue)
		IsZero = Abs(varValue) < dblEpsilon
	End Function

	Public Function Normalize()
		If IsZero(varNorm) Then
			Err.Raise _
				conDivideByZeroError, _
				"Vector", _
				"Cannot normalize zero vector."
		End If
		Dim avarResult()
		ReDim avarResult(lngLength - 1)
		Dim lngIndex
		For lngIndex = LBound(avarResult) To UBound(avarResult)
			avarResult(lngIndex) = avarValues(lngIndex) / varNorm
		Next
		Set Normalize = New Vector
		Normalize.Values = avarResult
	End Function

	Public Function DotProduct(ByRef objAnotherVector)
		If lngLength = objAnotherVector.Length Then
			Dim lngIndex
			Dim varResult
			For lngIndex = LBound(avarValues) To UBound(avarValues)
				varResult = _
					varResult + _
					avarValues(lngIndex) * objAnotherVector.Value(lngIndex)
			Next
			DotProduct = varResult
		Else
			Err.Raise _
				conDimensionMismatchError, _
				"Vector", _
				"Two Vector's length is not equal."
		End If
	End Function

	Public Function Multiply(ByRef varAnotherNumber)
		If IsNumeric(varAnotherNumber) Then
			Dim avarResult()
			ReDim avarResult(lngLength - 1)
			Dim lngIndex
			For lngIndex = LBound(avarResult) To UBound(avarResult)
				avarResult(lngIndex) = _
					avarValues(lngIndex) * varAnotherNumber
			Next
			Set Multiply = New Vector
			Multiply.Values = avarResult
		Else
			Err.Raise _
				conNonNumericError, _
				"Vector", _
				"Multiply operand is not numeric."
		End If
	End Function

	Public Function Divide(ByRef varAnotherNumber)
		If IsNumeric(varAnotherNumber) Then
			If Not IsZero(varAnotherNumber) Then
				Set Divide = Multiply(1 / varAnotherNumber)
			Else
				Err.Raise _
					conDivideByZeroError, _
					"Vector", _
					"Cannot divide by zero."
			End If
		Else
			Err.Raise _
				conNonNumericError, _
				"Vector", _
				"Divide operand is not numeric."
		End If
	End Function
End Class