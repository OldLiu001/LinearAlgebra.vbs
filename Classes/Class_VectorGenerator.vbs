Option Explicit

Class VectorGenerator
	Public Function Init(ByRef avarRawVector)
		Set Init = New Vector
		Init.Values = avarRawVector
	End Function

	Public Function Zero(ByRef lngLength)
		Dim avarZero()
		ReDim avarZero(lngLength - 1)
		Fill avarZero, 0
		Set Zero = New Vector
		Zero.Values = avarZero
	End Function

	Private Sub Fill(ByRef avarArray, ByRef lngValue)
		Dim lngIndex
		For lngIndex = LBound(avarArray) To UBound(avarArray)
			avarArray(lngIndex) = lngValue
		Next
	End Sub
End Class