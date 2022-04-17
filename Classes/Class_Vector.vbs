Option Explicit

Class Vector
	Private avarValues()
	Private lngLength, varNorm

	Private boolReadOnly

	Private objVectorGenerator

	Private lngErrorNumber

	Private Sub Class_Initialize()
		boolReadOnly = False

		Set objVectorGenerator = New VectorGenerator
	End Sub

	Private Sub Assert(ByVal boolCondition, ByRef strMessage)
		If Not boolCondition Then
			Err.Raise vbObjectError, "Vector", strMessage
		End If
	End Sub

	Public Property Let Values(ByRef avarRawVector)
		Assert Not boolReadOnly, "Values property is read only."
		Assert IsArray(avarRawVector), "The Input must be an array of numbers."
		
		On Error Resume Next
			Call UBound(avarRawVector, 1)
			lngErrorNumber = Err.Number
		On Error GoTo 0
		Assert lngErrorNumber = 0, "Input Array is empty."
		
		lngLength = _
			UBound(avarRawVector) - _
			LBound(avarRawVector) + 1
		Assert lngLength > 0, "Input Array is empty."

		Dim lngIndex
		ReDim avarValues(UBound(avarRawVector))

		varNorm = 0
		For lngIndex = LBound(avarRawVector) To UBound(avarRawVector)
			Assert IsNumeric(avarRawVector(lngIndex)), _
				"Input Array contains non-numeric values."
			avarValues(lngIndex) = CDbl(avarRawVector(lngIndex))
			varNorm = varNorm + avarValues(lngIndex) ^ 2
		Next
		varNorm = Sqr(varNorm)

		boolReadOnly = True
	End Property
	
	Public Property Get Values()
		Values = avarValues
	End Property
	
	Private Function IsInteger(ByRef varValue)
		IsInteger = IsNumeric(varValue) And Fix(varValue) = varValue
	End Function

	Public Property Get Value(ByRef lngIndex)
		Assert IsInteger(lngIndex), "Index must be an integer."

		Value = avarValues(lngIndex)
	End Property
	
	Public Property Get Length()
		Length = lngLength
	End Property

	Public Property Get Stringify()
		Stringify = "[" + Join(avarValues, " ") + "]"
	End Property

	Public Function Add(ByRef objAnotherVector)
		Assert TypeName(objAnotherVector) = "Vector", _
			"Input must be a Vector."
		Assert lngLength = objAnotherVector.Length, _
			"Two Vector's length must be the same."
		
		Dim avarResult()
		ReDim avarResult(lngLength - 1)
		Dim lngIndex
		For lngIndex = LBound(avarResult) To UBound(avarResult)
			avarResult(lngIndex) = _
				avarValues(lngIndex) + objAnotherVector.Value(lngIndex)
		Next
		Set Add = objVectorGenerator.Init(avarResult)
	End Function

	Public Function Negate()
		Dim avarResult()
		ReDim avarResult(lngLength - 1)
		Dim lngIndex
		For lngIndex = LBound(avarResult) To UBound(avarResult)
			avarResult(lngIndex) = -avarValues(lngIndex)
		Next
		Set Negate = objVectorGenerator.Init(avarResult)
	End Function

	Public Function Subtract(ByRef objAnotherVector)
		Set Subtract = Add(objAnotherVector.Negate())
	End Function

	Public Property Get Norm()
		Norm = varNorm
	End Property

	Private Function IsZero(ByRef varValue)
		IsZero = Abs(varValue) < 1E-7
	End Function

	Public Function Normalize()
		Assert Not IsZero(varNorm), "Cannot normalize a zero vector."

		Dim avarResult()
		ReDim avarResult(lngLength - 1)
		Dim lngIndex
		For lngIndex = LBound(avarResult) To UBound(avarResult)
			avarResult(lngIndex) = avarValues(lngIndex) / varNorm
		Next
		Set Normalize = objVectorGenerator.Init(avarResult)
	End Function

	Public Function DotProduct(ByRef objAnotherVector)
		Assert TypeName(objAnotherVector) = "Vector", _
			"Input must be a Vector."
		Assert lngLength = objAnotherVector.Length, _
			"Two Vector's length must be the same."
		
		Dim lngIndex
		Dim varResult
		For lngIndex = LBound(avarValues) To UBound(avarValues)
			varResult = _
				varResult + _
				avarValues(lngIndex) * objAnotherVector.Value(lngIndex)
		Next
		DotProduct = varResult
	End Function

	Public Function Multiply(ByRef varAnotherNumber)
		Assert IsNumeric(varAnotherNumber), "Multiply operand is not numeric."
		
		Dim avarResult()
		ReDim avarResult(lngLength - 1)
		Dim lngIndex
		For lngIndex = LBound(avarResult) To UBound(avarResult)
			avarResult(lngIndex) = _
				avarValues(lngIndex) * varAnotherNumber
		Next
		Set Multiply = objVectorGenerator.Init(avarResult)
	End Function

	Public Function Divide(ByRef varAnotherNumber)
		Assert IsNumeric(varAnotherNumber), "Divide operand is not numeric."
		Assert Not IsZero(varAnotherNumber), "Cannot divide by zero."
		
		Set Divide = Multiply(1 / varAnotherNumber)
	End Function
End Class