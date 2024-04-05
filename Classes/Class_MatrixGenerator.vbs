Option Explicit

Class MatrixGenerator
	Public Function Init(ByRef avarRawMatrix)
		Set Init = New Matrix
		Init.Values = avarRawMatrix
	End Function
	
	Public Function Zero(ByRef lngRowLength, ByRef lngColumnLength)
		Dim avarZero()
		ReDim avarZero(lngRowLength - 1, lngColumnLength - 1)
		Dim lngRowIndex
		Dim lngColumnIndex
		For lngRowIndex = LBound(avarZero, 1) To UBound(avarZero, 1)
			For lngColumnIndex = LBound(avarZero, 2) To UBound(avarZero, 2)
				avarZero(lngRowIndex, lngColumnIndex) = 0
			Next
		Next

		Set Zero = New Matrix
		Zero.Values = avarZero
	End Function
	
	Public Function Identity(ByRef lngDimension)
		Dim avarIdentity()
		ReDim avarIdentity(lngDimension - 1, lngDimension - 1)
		Dim lngRowIndex
		Dim lngColumnIndex
		For lngRowIndex = LBound(avarIdentity, 1) To UBound(avarIdentity, 1)
			For lngColumnIndex = LBound(avarIdentity, 2) To UBound(avarIdentity, 2)
				If lngRowIndex = lngColumnIndex Then
					avarIdentity(lngRowIndex, lngColumnIndex) = 1
				Else
					avarIdentity(lngRowIndex, lngColumnIndex) = 0
				End If
			Next
		Next

		Set Identity = New Matrix
		Identity.Values = avarIdentity
	End Function
End Class