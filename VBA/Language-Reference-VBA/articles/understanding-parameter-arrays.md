---
title: Understanding Parameter Arrays
keywords: vbcn6.chm1076759
f1_keywords:
- vbcn6.chm1076759
ms.prod: office
ms.assetid: 42438a68-37a8-85d0-6404-1df4266fe33d
ms.date: 06/08/2017
---


# Understanding Parameter Arrays

A [parameter](vbe-glossary.md)[array](vbe-glossary.md) can be used to pass an array of [arguments](vbe-glossary.md) to a [procedure](vbe-glossary.md). You don't have to know the number of elements in the array when you define the procedure.

You use the  **ParamArray** keyword to denote a parameter array. The array must be declared as an array of type **Variant**, and it must be the last argument in the procedure definition.

The following example shows how you might define a procedure with a parameter array.




```vb
Sub AnyNumberArgs(strName As String, ParamArray intScores() As Variant) 
 Dim intI As Integer 
 
 Debug.Print strName; " Scores" 
 ' Use UBound function to determine upper limit of array. 
 For intI = 0 To UBound(intScores()) 
 Debug.Print " "; intScores(intI) 
 Next intI 
End Sub
```

The following examples show how you can call this procedure.



```vb
AnyNumberArgs "Jamie", 10, 26, 32, 15, 22, 24, 16 
 
AnyNumberArgs "Kelly", "High", "Low", "Average", "High" 
```


