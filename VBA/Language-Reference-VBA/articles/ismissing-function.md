---
title: IsMissing Function
keywords: vblr6.chm1010101
f1_keywords:
- vblr6.chm1010101
ms.prod: office
ms.assetid: 63193fae-038a-b95c-1776-ac820f62fbb2
ms.date: 06/08/2017
---


# IsMissing Function



Returns a  **Boolean** value indicating whether an optional **Variant**[argument](vbe-glossary.md) has been passed to a[procedure](vbe-glossary.md).
 **Syntax**
 **IsMissing(**_argname_**)**
The required  _argname_ argument contains the name of an optional **Variant** procedure argument.
 **Remarks**
Use the  **IsMissing** function to detect whether or not optional **Variant** arguments have been provided in calling a procedure. **IsMissing** returns **True** if no value has been passed for the specified argument; otherwise, it returns **False**. If **IsMissing** returns **True** for an argument, use of the missing argument in other code may cause a user-defined error. If **IsMissing** is used on a **ParamArray** argument, it always returns **False**. To detect an empty **ParamArray**, test to see if the[array's](vbe-glossary.md) upper bound is less than its lower bound.

 **Note**   **IsMissing** does not work on simple data types (such as **Integer** or **Double** ) because, unlike **Variants**, they don't have a provision for a "missing" flag bit. Because of this, the syntax for typed optional arguments allows you to specify a default value. If the argument is omitted when the procedure is called, then the argument will have this default value, as in the example below:




```vb
Sub MySub(Optional MyVar As String = "specialvalue")
    If MyVar = "specialvalue" Then
        ' MyVar was omitted.
    Else
    ...
End Sub
```

In many cases you can omit the  `If MyVar` test entirely by making the default value equal to the value you want `MyVar` to contain if the user omits it from the function call. This makes your code more concise and efficient.

## Example

This example uses the  **IsMissing** function to check if an optional argument has been passed to a user-defined procedure. Note that **Optional** arguments can now have default values and types other than **Variant**.


```vb
Dim ReturnValue
' The following statements call the user-defined function procedure.
ReturnValue = ReturnTwice()    ' Returns Null.
ReturnValue = ReturnTwice(2)    ' Returns 4.

' Function procedure definition.
Function ReturnTwice(Optional A)
    If IsMissing(A) Then
        ' If argument is missing, return a Null.
        ReturnTwice = Null
    Else
        ' If argument is present, return twice the value.
        ReturnTwice = A * 2
    End If
End Function
```


