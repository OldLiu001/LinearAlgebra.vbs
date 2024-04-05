If IsEmpty(WSH) Then Set WSH = WScript

Set objVectorGenerator = CreateObject("LinearAlgebra.VectorGenerator")

Set objVector = objVectorGenerator.Init(Array(4,3))
WSH.Echo objVector.Stringify() + ".Length()", "=", CStr(objVector.Length())
WSH.Echo objVector.Stringify() + ".Value(0)", "=", CStr(objVector.Value(0))
WSH.Echo objVector.Stringify() + ".Value(1)", "=", CStr(objVector.Value(1))
WSH.Echo objVector.Stringify() + ".Values()", "=", "Array(" + Join(objVector.Values(), ",") + ")"
WSH.Echo ""

Set objVector2 = objVectorGenerator.Init(Array(3,1))
WSH.Echo objVector.Stringify(), "+", objVector2.Stringify(), "=", objVector.Add(objVector2).Stringify()
WSH.Echo objVector.Stringify(), "-", objVector2.Stringify(), "=", objVector.Subtract(objVector2).Stringify()
WSH.Echo ""

WSH.Echo objVector.Stringify(), "*", "2", "=", objVector.Multiply(2).Stringify()
WSH.Echo objVector.Stringify(), "/", "2", "=", objVector.Divide(2).Stringify()
WSH.Echo ""

WSH.Echo objVector.Stringify() + ".DotProduct(" + objVector2.Stringify() + ") =", CStr(objVector.DotProduct(objVector2))
WSH.Echo ""

WSH.Echo objVector.Stringify() + ".Negate() =", objVector.Negate().Stringify()
WSH.Echo objVector.Stringify() + ".Norm() =", CStr(objVector.Norm())
WSH.Echo objVector.Stringify() + ".Normalize() =", objVector.Normalize().Stringify()
WSH.Echo objVector.Stringify() + ".Normalize().Norm() =", CStr(objVector.Normalize().Norm())
WSH.Echo ""

Set objZeroVector = objVectorGenerator.Zero(2)
WSH.Echo objVector.Stringify(), "+", objZeroVector.Stringify(), "=", objVector.Add(objZeroVector).Stringify()
WSH.Echo objVector.Stringify(), "-", objZeroVector.Stringify(), "=", objVector.Subtract(objZeroVector).Stringify()
WSH.Echo objZeroVector.Stringify(), "+", objVector.Stringify(), "=", objZeroVector.Add(objVector).Stringify()
WSH.Echo objZeroVector.Stringify(), "-", objVector.Stringify(), "=", objZeroVector.Subtract(objVector).Stringify()
WSH.Echo ""

On Error Resume Next
WSH.Echo objZeroVector.Stringify() + ".Normalize()"
WSH.Echo objZeroVector.Normalize()
WSH.Echo Err.Description
On Error GoTo 0
WSH.Echo ""