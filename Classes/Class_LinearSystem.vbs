Option Explicit

Class LinearSystem
	Private objCoefficientMatrix, objConstantMatrix, objAugmentedMatrix
	Private avarCoefficientMatrix, avarConstantMatrix, avarAugmentedMatrix

	Private objMatrixGenerator

	Private Sub Class_Initialize()
		Set objMatrixGenerator = New MatrixGenerator
	End Sub

	Public Sub Init(ByRef varCoefficientMatrix, ByRef varConstantMatrix)
		Set objCoefficientMatrix = objMatrixGenerator.Init(varCoefficientMatrix)
		If varConstantMatrix = Empty Then
			Set objConstantMatrix = Nothing
			avarConstantMatrix = Empty
			Set objAugmentedMatrix = objCoefficientMatrix
		Else
			Set objConstantMatrix = objMatrixGenerator.Init(varConstantMatrix)
			avarConstantMatrix = objConstantMatrix.Values
			Set objAugmentedMatrix = objCoefficientMatrix.Combine(objConstantMatrix)
		End If

		avarCoefficientMatrix = objCoefficientMatrix.Values
		avarAugmentedMatrix = objAugmentedMatrix.Values
	End Sub

	Public Sub Elimination(ByVal strName)
		Select Case strName
			Case "Gauss"
				GaussElimination()
			Case "Jordan"
				JordanElimination()
			Case "GaussJordan"
				GaussElimination()
				JordanElimination()
			Case "GaussSeidel"
				'not implemented yet
				GaussSeidelElimination()
		End Select
	End Sub

	Private Sub GaussElimination()
		'Colunm-Pivoting Method
		Dim lngRow, lngColumn, lngMaxRow
		lngRow = 0
		lngColumn = 0

		While lngRow < objCoefficientMatrix.RowCount And lngColumn < objCoefficientMatrix.ColumnCount
			lngMaxRow = MaxRow(lngRow, lngColumn)
		WEnd
	End Sub

	Private Function MaxColumnPivot(ByRef lngRow, ByRef lngColumn)
		Dim dblMax, lngMaxRow
		
		dblMax = avarAugmentedMatrix(lngRow, lngColumn)
		Dim i
		For i = lngColumn To objCoefficientMatrix.ColumnCount
			If avarAugmentedMatrix(lngRow, i) > dblMax Then
				dblMax = avarAugmentedMatrix(lngRow, i)
				lngMaxRow = i
			End If
		Next

	End Function

	'LDU LU PLU PLUP'

	Public Function factorization(ByVal strType)
		Select Case strType
			Case "LDU"
				Return LDUDecomposition()
			Case "LU"
				Return LUDecomposition()
			Case "PLU"
				Return PLUDecomposition()
			Case "PLUP"
				Return PLUPDecomposition()
		End Select
	End Function

	
	' Public Function Factorization(strType)
	' End Function

	' Private boolGaussEliminationed
	' Private adblPermutationLeft, adblPermutationRight
	' Private adblLowerTriangular, adblDiagonal, adblUpperTriangular
	' Private Function GaussElimination()
	' 	If Not boolGaussEliminationed Then
	' 		Dim lngRowIndex, lngColumnIndex
	' 		Set objPermutationLeft = objMatrixGenerator.Indentity(RowCount)
	' 		Set objPermutationRight = objMatrixGenerator.Indentity(ColumnCount)

	' 		adblUpperTriangular = Values
			
	' 		Dim adblPivot(ColumnCount - 1)

	' 		Dim lngPivotRowIndex, lngPivotColumnIndex
	' 		lngPivotColumnIndex = 0
	' 		For lngPivotRowIndex = 0 To RowCount - 1
	' 			If Not IsZero(adblUpperTriangular(lngPivotRowIndex, lngPivotColumnIndex)) Then



	' 		Next
			
	' 	End If
	' End Function

End Class