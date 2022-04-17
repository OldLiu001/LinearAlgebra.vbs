Option Explicit

Class Matrix
	Private lngRow, lngColumn
	Private avarValues()

	Private boolReadOnly

	Public _
		conRewriteError, _
		conDimensionMismatchError, _
		conNonNumericError, _
		conDivideByZeroError, _
		conLengthIsZeroError, _
		conIndexOutOfRangeError, _
		conTypeMismatchError
	
	Private dblEpsilon

	Private Sub Class_Initialize()
		dblEpsilon = 1e-7
		boolReadOnly = False

		conRewriteError = vbObjectError
		conDimensionMismatchError = vbObjectError + 1
		conNonNumericError = vbObjectError + 2
		conDivideByZeroError = vbObjectError + 3
		conLengthIsZeroError = vbObjectError + 4
		conIndexOutOfRangeError = vbObjectError + 5
		conTypeMismatchError = vbObjectError + 6
	End Sub

	Public Property Let Values(ByRef avarRaw)
		Rem Input Type: Vector or Array(Array()) or Array2D().
		
		If boolReadOnly Then
			Err.Raise conRewriteError, "Matrix", "Matrix is read-only."
		End If

		If TypeName(avarRaw) = "Matrix" Then
			boolReadOnly = True
			Exit Property
		End If
		
		Rem Turn any Input into Array2D(Number).
		If TypeName(avarRaw) = "Vector" Then
			ReDim avarValues(0, avarRaw.Length - 1)
			Dim lngIndex
			For lngIndex = LBound(avarRaw.Values()) To UBound(avarRaw.Values())
				avarValues(0, lngIndex) = avarRaw.Value(lngIndex)
			Next
		ElseIf IsArray(avarRaw) Then
			On Error Resume Next
			Call UBound(avarRaw, 1)
			If Err.Number <> 0 Then
				On Error GoTo 0
				Err.Raise conLengthIsZeroError, "Matrix", "Input array is empty."
			End If
			On Error GoTo 0
			If UBound(avarRaw, 1) = -1 Then
				Err.Raise conLengthIsZeroError, "Matrix", "Input array is empty."
			Else
				Dim lngRowIndex
				Dim lngColumnIndex
				Dim avarValue
				For Each avarValue In avarRaw
					If TypeName(avarValue) = "Variant()" Then
						Rem Turn Array(Array(...)) into Array2d(...).
						ReDim avarValues(UBound(avarRaw), UBound(avarValue))
						For lngRowIndex = LBound(avarRaw) To UBound(avarRaw)
							For lngColumnIndex = LBound(avarValue) To UBound(avarValue)
								If IsArray(avarRaw(lngRowIndex)) Then
									If UBound(avarValue) = UBound(avarRaw(lngRowIndex)) Then
										If Not IsNumeric(avarRaw(lngRowIndex)(lngColumnIndex)) Then
											Err.Raise _
												conNonNumericError, _
												"Matrix", _
												"Array contains non-numeric value(s)."
										Else
											avarValues(lngRowIndex, lngColumnIndex) = CDbl(avarRaw(lngRowIndex)(lngColumnIndex))
										End If
									Else
										Err.Raise _
											conDimensionMismatchError, _
											"Matrix", _
											"Input array is not a rectangular matrix."
									End If
								Else
									Err.Raise _
										conTypeMismatchError, _
										"Matrix", _
										"Input array is not a rectangular matrix."
								End If
							Next
						Next
					ElseIf IsNumeric(avarValue) Then
						ReDim avarValues(UBound(avarRaw, 1), UBound(avarRaw, 2))
						For lngRowIndex = LBound(avarRaw) To UBound(avarRaw)
							For lngColumnIndex = LBound(avarRaw, 2) To UBound(avarRaw, 2)
								If Not IsNumeric(avarRaw(lngRowIndex, lngColumnIndex)) Then
									Err.Raise _
										conNonNumericError, _
										"Matrix", _
										"Array contains non-numeric value(s)."
								Else
									avarValues(lngRowIndex, lngColumnIndex) = CDbl(avarRaw(lngRowIndex, lngColumnIndex))
								End If
							Next
						Next
					End If
					Exit For
				Next
			End If
		Else
			Err.Raise conTypeMismatchError, "Matrix", "Input is not a vector or array."
		End If

		lngRow = UBound(avarValues, 1) + 1
		lngColumn = UBound(avarValues, 2) + 1
		boolReadOnly = True
	End Property

	Public Property Get Stringify()
		Dim lngRowIndex
		Dim lngColumnIndex
		Stringify = "[" & vbNewLine
		For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
			Stringify = Stringify & "	[ "
			For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
				Stringify = Stringify & avarValues(lngRowIndex, lngColumnIndex) & " "
			Next
			Stringify = Stringify & "]" & vbNewLine
		Next
		Stringify = Stringify & "]"
	End Property

	Public Property Get RowCount()
		RowCount = lngRow
	End Property

	Public Property Get ColumnCount()
		ColumnCount = lngColumn
	End Property
	
	Public Property Get Length()
		Length = lngRow * lngColumn
	End Property

	Public Property Get Row(ByVal lngRowIndex)
		Dim adblRow
		
		If lngRowIndex >= lngRow Then
			Err.Raise conIndexOutOfRangeError, "Matrix", "Row index out of range."
		Else
			ReDim adblRow(lngColumn - 1)
			Dim lngColumnIndex
			For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
				adblRow(lngColumnIndex) = avarValues(lngRowIndex, lngColumnIndex)
			Next
		End If
		Row = adblRow
	End Property

	Public Property Get RowVector(ByVal lngRowIndex)
		Set RowVector = New VectorGenerator
		Set RowVector = RowVector.Init(Row(lngRowIndex))
	End Property


	Public Property Get Column(ByVal lngColumnIndex)
		Dim adblColumn

		If lngColumnIndex >= lngColumn Then
			Err.Raise conIndexOutOfRangeError, "Matrix", "Column index out of range."
		Else
			ReDim adblColumn(lngRow - 1)
			Dim lngRowIndex
			For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
				adblColumn(lngRowIndex) = avarValues(lngRowIndex, lngColumnIndex)
			Next
		End If
		Column = adblColumn
	End Property

	Public Property Get ColumnVector(ByVal lngColumnIndex)
		Set ColumnVector = New VectorGenerator
		Set ColumnVector = ColumnVector.Init(Column(lngColumnIndex))
	End Property

	Public Property Get Value(ByVal lngRowIndex, ByVal lngColumnIndex)
		If lngRowIndex >= lngRow Then
			Err.Raise conIndexOutOfRangeError, "Matrix", "Row index out of range."
		ElseIf lngColumnIndex >= lngColumn Then
			Err.Raise conIndexOutOfRangeError, "Matrix", "Column index out of range."
		Else
			Value = avarValues(lngRowIndex, lngColumnIndex)
		End If
	End Property

	Public Property Get Values()
		Values = avarValues
	End Property

	Public Function Transpose()
		Dim adblTransposed()
		Dim lngRowIndex
		Dim lngColumnIndex
		ReDim adblTransposed(lngColumn - 1, lngRow - 1)
		For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
			For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
				adblTransposed(lngColumnIndex, lngRowIndex) = avarValues(lngRowIndex, lngColumnIndex)
			Next
		Next
		Set Transpose = New Matrix
		Transpose.Values = adblTransposed
	End Function

	Public Function Add(ByVal objAnotherMatrix)
		If TypeName(objAnotherMatrix) = "Matrix" Then
			Dim adblAdded()
			Dim lngRowIndex
			Dim lngColumnIndex
			If lngRow = objAnotherMatrix.RowCount And lngColumn = objAnotherMatrix.ColumnCount Then
				ReDim adblAdded(lngRow - 1, lngColumn - 1)
				For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
					For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
						adblAdded(lngRowIndex, lngColumnIndex) = _
							avarValues(lngRowIndex, lngColumnIndex) + _
							objAnotherMatrix.Value(lngRowIndex, lngColumnIndex)
					Next
				Next
				Set Add = New Matrix
				Add.Values = adblAdded
			Else
				Err.Raise conDimensionMismatchError, "Matrix", "Dimension mismatch."
			End If
		Else
			Err.Raise conTypeMismatchError, "Matrix", "Type mismatch."
		End If
	End Function

	Public Function Negate()
		Dim adblNegated()
		Dim lngRowIndex
		Dim lngColumnIndex
		ReDim adblNegated(lngRow - 1, lngColumn - 1)
		For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
			For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
				adblNegated(lngRowIndex, lngColumnIndex) = -avarValues(lngRowIndex, lngColumnIndex)
			Next
		Next
		Set Negate = New Matrix
		Negate.Values = adblNegated
	End Function

	Public Function Subtract(ByVal objAnotherMatrix)
		Set Subtract = Add(objAnotherMatrix.Negate)
	End Function

	Public Function Multiply(ByVal objAnother)
		Dim adblMultiplied()
		If IsNumeric(objAnother) Then
			ReDim adblMultiplied(lngRow - 1, lngColumn - 1)
			For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
				For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
					adblMultiplied(lngRowIndex, lngColumnIndex) = _
						avarValues(lngRowIndex, lngColumnIndex) * objAnother
				Next
			Next
			Set Multiply = New Matrix
			Multiply.Values = adblMultiplied
		ElseIf TypeName(objAnother) = "Matrix" Then
			If lngColumn = objAnother.RowCount Then
				ReDim adblMultiplied(lngRow - 1, objAnother.ColumnCount - 1)
				Dim lngRowIndex
				Dim lngColumnIndex
				Dim lngAnotherColumnIndex
				For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
					For lngAnotherColumnIndex = LBound(objAnother.Values, 2) To UBound(objAnother.Values, 2)
						adblMultiplied(lngRowIndex, lngAnotherColumnIndex) = 0
						For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
							adblMultiplied(lngRowIndex, lngAnotherColumnIndex) = _
								adblMultiplied(lngRowIndex, lngAnotherColumnIndex) + _
								avarValues(lngRowIndex, lngColumnIndex) * _
								objAnother.Value(lngColumnIndex, lngAnotherColumnIndex)
						Next
					Next
				Next
				Set Multiply = New Matrix
				Multiply.Values = adblMultiplied
			Else
				Err.Raise conDimensionMismatchError, "Matrix", "Dimension mismatch."
			End If
		ElseIf TypeName(objAnother) = "Vector" Then
			Rem Assume that the vector is a column vector.
			If lngColumn = objAnother.Length Then
				ReDim adblMultiplied(lngRow - 1, 0)
				For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
					adblMultiplied(lngRowIndex, 0) = 0
					For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
						adblMultiplied(lngRowIndex, 0) = _
							adblMultiplied(lngRowIndex, 0) + _
							avarValues(lngRowIndex, lngColumnIndex) * _
							objAnother.Value(lngColumnIndex)
					Next
				Next
				Set Multiply = New Matrix
				Multiply.Values = adblMultiplied
			Else
				Err.Raise conDimensionMismatchError, "Matrix", "Dimension mismatch."
			End If
		Else
			Err.Raise conTypeMismatchError, "Matrix", "Type mismatch."
		End If
	End Function

	Public Function Divide(ByVal varAnotherNumber)
		If IsNumeric(varAnotherNumber) Then
			Set Divide = Multiply(1 / varAnotherNumber)
		Else
			Err.Raise conTypeMismatchError, "Matrix", "Type mismatch."
		End If
	End Function

	Public Function Append(ByVal objAnotherMatrix)
		Rem Append the matrix to the bottom of the current matrix.
		Dim adblAppended()
		If TypeName(objAnotherMatrix) = "Matrix" Then
			If lngColumn = objAnotherMatrix.ColumnCount Then
				ReDim adblAppended(lngRow + objAnotherMatrix.RowCount - 1, lngColumn - 1)
				Dim lngRowIndex
				Dim lngColumnIndex
				For lngRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
					For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
						adblAppended(lngRowIndex, lngColumnIndex) = avarValues(lngRowIndex, lngColumnIndex)
					Next
				Next
				For lngRowIndex = LBound(objAnotherMatrix.Values, 1) To UBound(objAnotherMatrix.Values, 1)
					For lngColumnIndex = LBound(objAnotherMatrix.Values, 2) To UBound(objAnotherMatrix.Values, 2)
						adblAppended(lngRowIndex + lngRow, lngColumnIndex) = _
							objAnotherMatrix.Value(lngRowIndex, lngColumnIndex)
					Next
				Next
				Set Append = New Matrix
				Append.Values = adblAppended
			Else
				Err.Raise conDimensionMismatchError, "Matrix", "Dimension mismatch."
			End If
		Else
			Err.Raise conTypeMismatchError, "Matrix", "Type mismatch, expected Matrix."
		End If
	End Function

	Public Function Combine(ByVal objAnotherMatrix)
		Rem Combine the matrix to the right of the current matrix.
		Set Combine = Transpose().Append(objAnotherMatrix.Transpose()).Transpose()
	End Function

	Public Property Get Determinant()
		If lngRow = lngColumn Then
			If lngRow = 1 Then
				Determinant = avarValues(0, 0)
			ElseIf lngRow <= 5 Then
				Dim lngColumnIndex
				Determinant = 0
				For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
					Determinant = Determinant + _
						avarValues(0, lngColumnIndex) * _
						AlgebraicCofactor(0, lngColumnIndex)
				Next
			Else
			End If
		Else
			Err.Raise conDimensionMismatchError, "Matrix", "Dimension mismatch, expected square matrix."
		End If
	End Property

	Public Function RemoveRow(lngRowIndex)
		Rem Remove the specified row from the matrix.
		Dim adblRemoved()
		If lngRowIndex >= LBound(avarValues, 1) And lngRowIndex <= UBound(avarValues, 1) Then
			ReDim adblRemoved(lngRow - 2, lngColumn - 1)
			Dim lngTemporaryRowIndex
			Dim lngColumnIndex
			For lngTemporaryRowIndex = LBound(avarValues, 1) To UBound(avarValues, 1)
				For lngColumnIndex = LBound(avarValues, 2) To UBound(avarValues, 2)
					adblRemoved(
						lngTemporaryRowIndex + (Sgn(lngTemporaryRowIndex - lngRowIndex) + 1)/2, _
						lngColumnIndex) = _
						avarValues(lngTemporaryRowIndex, lngColumnIndex)
				Next
			Next
			Set RemoveRow = New Matrix
			RemoveRow.Values = adblRemoved
		Else
			Err.Raise conDimensionMismatchError, "Matrix", "Dimension mismatch."
		End If
	End Function

	Public Function RemoveColumn(lngColumnIndex)
		Rem Remove the specified column from the matrix.
		Set RemoveColumn = Transpose().RemoveRow(lngColumnIndex).Transpose()
	End Function

	Public Property Get Cofactor(ByVal lngRowIndex, ByVal lngColumnIndex)
		If lngRow = lngColumn Then
			Dim objCofactor
			Set objCofactor = New Matrix
			objCofactor.Values = RemoveRow(lngRowIndex).RemoveColumn(lngColumnIndex).Values
			Cofactor = objCofactor.Determinant
		Else
			Err.Raise conDimensionMismatchError, "Matrix", "Dimension mismatch, expected square matrix."
		End If
	End Property

	Public Property Get AlgebraicCofactor(ByVal lngRowIndex, ByVal lngColumnIndex)
		AlgebraicCofactor = (-1) ^ (lngRowIndex + lngColumnIndex) * Cofactor(lngRowIndex, lngColumnIndex)
	End Property
End Class