---
title: TypeName Function
keywords: vblr6.chm1010100
f1_keywords:
- vblr6.chm1010100
ms.prod: office
ms.assetid: 9353f1d5-5b64-9cad-5cc3-e1487bdd3afd
ms.date: 04/27/2018
---


# TypeName Function

Returns a  **String** that provides information about a [variable](vbe-glossary.md).</br></br>
## Syntax
**TypeName(**_varname_**)**</br>
The required _varname_ [argument](vbe-glossary.md) is a [Variant](vbe-glossary.md) containing any variable except a variable of a [user-defined type](vbe-glossary.md).
## Remarks
The string returned by **TypeName** can be any one of the following:</br>

|**String returned**|**Variable**|
|:-----|:-----|
|[object type](vbe-glossary.md)|An object whose type is  _objecttype_|
|[Byte](vbe-glossary.md)|Byte value|
|[Integer](vbe-glossary.md)|Integer|
|[Long](vbe-glossary.md)|Long integer|
|[Single](vbe-glossary.md)|Single-precision floating-point number|
|[Double](vbe-glossary.md)|Double-precision floating-point number|
|[Currency](vbe-glossary.md)|Currency value|
|[Decimal](vbe-glossary.md)|Decimal value|
|[Date](vbe-glossary.md)|Date value|
|[String](vbe-glossary.md)|String|
|[Boolean](vbe-glossary.md)|Boolean value|
|**Error**|An error value|
|[Empty](vbe-glossary.md)|Uninitialized|
|[Null](vbe-glossary.md)|No valid data|
|[Object](vbe-glossary.md)|An object|
|Unknown|An object whose type is unknown|
|**Nothing**|Object variable that doesn't refer to an object|
<br>
If  _varname_ is an [array](vbe-glossary.md), the returned string can be any one of the possible returned strings (or  **Variant** ) with empty parentheses appended. For example, if _varname_ is an array of integers, **TypeName** returns `"Integer()`".

## Example

This example uses the **TypeName** function to return information about a variable.


```vb
' Declare variables.
Dim NullVar, MyType, StrVar As String, IntVar As Integer, CurVar As Currency
Dim ArrayVar (1 To 5) As Integer
NullVar = Null    ' Assign Null value.
MyType = TypeName(StrVar)    ' Returns "String".
MyType = TypeName(IntVar)    ' Returns "Integer".
MyType = TypeName(CurVar)    ' Returns "Currency".
MyType = TypeName(NullVar)    ' Returns "Null".
MyType = TypeName(ArrayVar)    ' Returns "Integer()".

```


