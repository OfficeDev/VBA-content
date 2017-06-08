---
title: CVErr Function
keywords: vblr6.chm1008821
f1_keywords:
- vblr6.chm1008821
ms.prod: office
ms.assetid: 244ab040-3816-a744-7afb-06675a4b076d
ms.date: 06/08/2017
---


# CVErr Function



Returns a [Variant](vbe-glossary.md) of subtype **Error** containing an[error number](vbe-glossary.md) specified by the user.
 **Syntax**
 **CVErr(**_errornumber_**)**
The required  _errornumber_[argument](vbe-glossary.md) is any valid error number.
 **Remarks**
Use the  **CVErr** function to create user-defined errors in user-created[procedures](vbe-glossary.md). For example, if you create a function that accepts several arguments and normally returns a string, you can have your function evaluate the input arguments to ensure they are within acceptable range. If they are not, it is likely your function will not return what you expect. In this event,  **CVErr** allows you to return an error number that tells you what action to take.
Note that implicit conversion of an  **Error** is not allowed. For example, you can't directly assign the return value of **CVErr** to a[variable](vbe-glossary.md) that is not a **Variant**. However, you can perform an explicit conversion (using **CInt**, **CDbl**, and so on) of the value returned by **CVErr** and assign that to a variable of the appropriate[data type](vbe-glossary.md).

## Example

This example uses the  **CVErr** function to return a **Variant** whose **VarType** is **vbError** (10). The user-defined function `CalculateDouble` returns an error if the argument passed to it isn't a number. You can use **CVErr** to return user-defined errors from user-defined procedures or to defer handling of a run-time error. Use the **IsError** function to test if the value represents an error.


```vb
' Call CalculateDouble with an error-producing argument.
Sub Test()
    Debug.Print CalculateDouble("345.45robert")
End Sub
' Define CalculateDouble Function procedure.
Function CalculateDouble(Number)
    If IsNumeric(Number) Then
        CalculateDouble = Number * 2    ' Return result.
    Else
        CalculateDouble = CVErr(2001)    ' Return a user-defined error 
    End If    ' number.
End Function
```


