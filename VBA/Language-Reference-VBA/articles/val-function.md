---
title: Val Function
keywords: vblr6.chm1009055
f1_keywords:
- vblr6.chm1009055
ms.prod: office
ms.assetid: 49909fd8-07bf-3b94-4c6c-85ad595d6fcb
ms.date: 06/08/2017
---


# Val Function



Returns the numbers contained in a string as a numeric value of appropriate type.
 **Syntax**
 **Val(**_string_**)**
The required  _string_[argument](vbe-glossary.md) is any valid[string expression](vbe-glossary.md).
 **Remarks**
The  **Val** function stops reading the string at the first character it can't recognize as part of a number. Symbols and characters that are often considered parts of numeric values, such as dollar signs and commas, are not recognized. However, the function recognizes the radix prefixes `&;O` (for octal) and (for octal) and `&;H` (for hexadecimal). Blanks, tabs, and linefeed characters are stripped from the argument.
The following returns the value 1615198:



```vb
Val("    1615 198th Street N.E.")
```

In the code below,  **Val** returns the decimal value -1 for the hexadecimal value shown:



```vb
Val("&;HFFFF")
```


 **Note**  The  **Val** function recognizes only the period ( **.** ) as a valid decimal separator. When different decimal separators are used, as in international applications, use **CDbl** instead to convert a string to a number.


## Example

This example uses the  **Val** function to return the numbers contained in a string.


```vb
Dim MyValue
MyValue = Val("2457")    ' Returns 2457.
MyValue = Val(" 2 45 7")    ' Returns 2457.
MyValue = Val("24 and 57")    ' Returns 24.
```


