If IsEmpty(WSH) Then Set WSH = WScript
Set objVectorGenerator = CreateObject("LinearAlgebra.VectorGenerator")
Set objMatrixGenerator = CreateObject("LinearAlgebra.MatrixGenerator")

'1. Create a matrix from array(array())
WSH.Echo objMatrixGenerator.Init(Array(Array(1,2),Array(3,4))).Stringify()
'2. Create a matrix from array2d()
Dim arr2d(1,1)
arr2d(0,0) = 1 : arr2d(0,1) = 2 : arr2d(1,0) = 3 : arr2d(1,1) = 4
WSH.Echo objMatrixGenerator.Init(arr2d).Stringify()
'3. Create a matrix from Vector()
WSH.Echo objMatrixGenerator.Init(objVectorGenerator.Init(Array(1,2,3,4))).Stringify()
WSH.Echo objMatrixGenerator.Init(objVectorGenerator.Init(Array(1,2,3,4))).Transpose().Stringify()
WSH.Echo ""

WSH.Echo objMatrixGenerator.Zero(2,3).Stringify()
WSH.Echo objMatrixGenerator.Identity(4).Stringify()
WSH.Echo ""

Set objMatrix = objMatrixGenerator.Init(Array(Array(1,2,3),Array(4,5,6)))
WSH.Echo "matrix: " & objMatrix.Stringify()
WSH.Echo "row: " & objMatrix.RowCount() & " col: " & objMatrix.ColumnCount()
WSH.Echo "length: " & objMatrix.Length()
WSH.Echo "RowVector(0): " & objMatrix.RowVector(0).Stringify()
WSH.Echo "ColumnVector(1): " & objMatrix.ColumnVector(1).Stringify()
WSH.Echo "value(0,1): " & objMatrix.Value(0,1)
WSH.Echo "values()(1,0): " & objMatrix.Values()(1,0)
WSH.Echo "Transpose(): " & objMatrix.Transpose().Stringify()
WSH.Echo "Nagate(): " & objMatrix.Negate().Stringify()
WSH.Echo ""

Set A = objMatrixGenerator.Init(Array(Array(1,2),Array(3,4)))
Set B = objMatrixGenerator.Init(Array(Array(5,6),Array(7,8)))
Set C = objVectorGenerator.Init(Array(5,7))
Set D = objMatrixGenerator.Init(Array(Array(5),Array(7)))
WSH.Echo "A: " & A.Stringify()
WSH.Echo "B: " & B.Stringify()
WSH.Echo "C: " & C.Stringify()
WSH.Echo "D: " & D.Stringify()
WSH.Echo "A + B: " & A.Add(B).Stringify()
WSH.Echo "A - B: " & A.Subtract(B).Stringify()
WSH.Echo "A * B: " & A.Multiply(B).Stringify()
WSH.Echo "B * A: " & B.Multiply(A).Stringify()
WSH.Echo "Matrix * Vector : assume Vector as column vector"
WSH.Echo "A * C: " & A.Multiply(C).Stringify()
WSH.Echo "A * D: " & A.Multiply(D).Stringify()
WSH.Echo "A * 2: " & A.Multiply(2).Stringify()
WSH.Echo "A / 2: " & A.Divide(2).Stringify()
WSH.Echo ""

Set X = objMatrixGenerator.Init(Array(Array(1,2)))
Set Y = objMatrixGenerator.Init(Array(Array(3),Array(4)))
Set I = objMatrixGenerator.Identity(2)
WSH.Echo "X: " & X.Stringify()
WSH.Echo "Y: " & Y.Stringify()
WSH.Echo "I: " & I.Stringify()
WSH.Echo "X * Y: " & X.Multiply(Y).Stringify()
WSH.Echo "Y * X: " & Y.Multiply(X).Stringify()
WSH.Echo "X * I: " & X.Multiply(I).Stringify()
WSH.Echo "I * Y: " & I.Multiply(Y).Stringify()
WSH.Echo ""

Set A = objMatrixGenerator.Init(Array(Array(1,2),Array(3,4)))
Set B = objMatrixGenerator.Init(Array(Array(5,6),Array(7,8)))
WSH.Echo "A: " & A.Stringify()
WSH.Echo "B: " & B.Stringify()
WSH.Echo "A.Append(B): " & A.Append(B).Stringify()
WSH.Echo "B.Append(A): " & B.Append(A).Stringify()
WSH.Echo "A.Combine(B): " & A.Combine(B).Stringify()
WSH.Echo "B.Combine(A): " & B.Combine(A).Stringify()
WSH.Echo ""


Set A = objMatrixGenerator.Init(Array(Array(-1)))
WSH.Echo "A: " & A.Stringify()
WSH.Echo "A.Determinant(): " & A.Determinant()
Set B = objMatrixGenerator.Init(Array(Array(1,2),Array(3,4)))
WSH.Echo "B: " & B.Stringify()
WSH.Echo "B.Determinant(): " & B.Determinant()
Set C = objMatrixGenerator.Init(Array(Array(1,2,3),Array(4,5,6),Array(7,8,9)))
WSH.Echo "C: " & C.Stringify()
WSH.Echo "C.Determinant(): " & C.Determinant()
