---
title: Type Conversion Functions
keywords: vblr6.chm1008820
f1_keywords:
- vblr6.chm1008820
ms.prod: MULTIPLEPRODUCTS
ms.assetid: fd602e34-9de2-1e8b-46fe-6a2873d6a785
---


# Type Conversion Functions

Each function coerces an +AFs-expression+AF0-(vbe-glossary.md) to a specific+AFs-data type+AF0-(vbe-glossary.md).

 +ACoAKg-Syntax+ACoAKg-

 +ACoAKg-CBool(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-

 +ACoAKg-CByte(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CCur(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CDate(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CDbl(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CDec(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CInt(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CLng(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CLngLng(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg- (Valid on 64-bit platforms only.)
 +ACoAKg-CLngPtr(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CSng(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CStr(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
 +ACoAKg-CVar(+ACoAKgBf-expression+AF8AKgAq-)+ACoAKg-
The required  +AF8-expression+AF8AWw-argument+AF0-(vbe-glossary.md) is any+AFs-string expression+AF0-(vbe-glossary.md) or+AFs-numeric expression+AF0-(vbe-glossary.md).
 +ACoAKg-Return Types+ACoAKg-
The function name determines the return type as shown in the following:


+AHwAKgAq-Function+ACoAKgB8ACoAKg-Return Type+ACoAKgB8ACoAKg-Range for  +AF8-expression+AF8- argument+ACoAKgB8-
+AHw-:-----+AHw-:-----+AHw-:-----+AHw-
+AHwAKgAq-CBool+ACoAKgB8AFs-Boolean+AF0-(vbe-glossary.md)+AHw-Any valid  +ACoAKg-string+ACoAKg- or numeric expression.+AHw-
+AHwAKgAq-CByte+ACoAKgB8AFs-Byte+AF0-(vbe-glossary.md)+AHw-0 to 255.+AHw-
+AHwAKgAq-CCur+ACoAKgB8AFs-Currency+AF0-(vbe-glossary.md)+AHw--922,337,203,685,477.5808 to 922,337,203,685,477.5807.+AHw-
+AHwAKgAq-CDate+ACoAKgB8AFs-Date+AF0-(vbe-glossary.md)+AHw-Any valid +AFs-date expression+AF0-(vbe-glossary.md).+AHw-
+AHwAKgAq-CDbl+ACoAKgB8AFs-Double+AF0-(vbe-glossary.md)+AHw--1.79769313486231E308 to -4.94065645841247E-324 for negative values+ADs- 4.94065645841247E-324 to 1.79769313486232E308 for positive values.+AHw-
+AHwAKgAq-CDec+ACoAKgB8AFs-Decimal+AF0-(vbe-glossary.md)+AHw-79,228,162,514,264,337,593,543,950,335 for zero-scaled numbers, that is, numbers with no decimal places. For numbers with 28 decimal places, the range is 7.9228162514264337593543950335. The smallest possible non-zero number is 0.0000000000000000000000000001.+AHw-
+AHwAKgAq-CInt+ACoAKgB8AFs-Integer+AF0-(vbe-glossary.md)+AHw--32,768 to 32,767+ADs- fractions are rounded.+AHw-
+AHwAKgAq-CLng+ACoAKgB8AFs-Long+AF0-(vbe-glossary.md)+AHw--2,147,483,648 to 2,147,483,647+ADs- fractions are rounded.+AHw-
+AHwAKgAq-CLngLng+ACoAKgB8AFs-LongLong+AF0-(longlong-data-type.md)+AHw--9,223,372,036,854,775,808 to 9,223,372,036,854,775,807+ADs- fractions are rounded. (Valid on 64-bit platforms only.)+AHw-
+AHwAKgAq-CLngPtr+ACoAKgB8AFs-LongPtr+AF0-(longptr-data-type.md)+AHw--2,147,483,648 to 2,147,483,647 on 32-bit systems, -9,223,372,036,854,775,808 to 9,223,372,036,854,775,807 on 64-bit systems+ADs- fractions are rounded for 32-bit and 64-bit systems.+AHw-
+AHwAKgAq-CSng+ACoAKgB8AFs-Single+AF0-(vbe-glossary.md)+AHw--3.402823E38 to -1.401298E-45 for negative values+ADs- 1.401298E-45 to 3.402823E38 for positive values.+AHw-
+AHwAKgAq-CStr+ACoAKgB8AFs-String+AF0-(vbe-glossary.md)+AHwAWw-Returns for CStr+AF0-(returns-for-cstr.md) depend on the +AF8-expression+AF8- argument.+AHw-
+AHwAKgAq-CVar+ACoAKgB8AFs-Variant+AF0-(vbe-glossary.md)+AHw-Same range as  +ACoAKg-Double+ACoAKg- for numerics. Same range as +ACoAKg-String+ACoAKg- for non-numerics.+AHw-
 +ACoAKg-Remarks+ACoAKg-
If the  +AF8-expression+AF8- passed to the function is outside the range of the data type being converted to, an error occurs.

 +ACoAKg-Note+ACoAKg-  Conversion functions must be used to explicitly assign  +ACoAKg-LongLong+ACoAKg- (including +ACoAKg-LongPtr+ACoAKg- on 64-bit platforms) to smaller integral types. Implicit conversions of +ACoAKg-LongLong+ACoAKg- to smaller integrals are not allowed.

In general, you can document your code using the data-type conversion functions to show that the result of some operation should be expressed as a particular data type rather than the default data type. For example, use  +ACoAKg-CCur+ACoAKg- to force currency arithmetic in cases where single-precision, double-precision, or integer arithmetic normally would occur.
You should use the data-type conversion functions instead of  +ACoAKg-Val+ACoAKg- to provide internationally aware conversions from one data type to another. For example, when you use +ACoAKg-CCur+ACoAKg-, different decimal separators, different thousand separators, and various currency options are properly recognized depending on the+AFs-locale+AF0-(vbe-glossary.md) setting of your computer.
When the fractional part is exactly 0.5,  +ACoAKg-CInt+ACoAKg- and +ACoAKg-CLng+ACoAKg- always round it to the nearest even number. For example, 0.5 rounds to 0, and 1.5 rounds to 2. +ACoAKg-CInt+ACoAKg- and +ACoAKg-CLng+ACoAKg- differ from the +ACoAKg-Fix+ACoAKg- and +ACoAKg-Int+ACoAKg- functions, which truncate, rather than round, the fractional part of a number. Also, +ACoAKg-Fix+ACoAKg- and +ACoAKg-Int+ACoAKg- always return a value of the same type as is passed in.
Use the  +ACoAKg-IsDate+ACoAKg- function to determine if +AF8-date+AF8- can be converted to a date or time. +ACoAKg-CDate+ACoAKg- recognizes+AFs-date literals+AF0-(vbe-glossary.md) and time literals as well as some numbers that fall within the range of acceptable dates. When converting a number to a date, the whole number portion is converted to a date. Any fractional part of the number is converted to a time of day, starting at midnight.
 +ACoAKg-CDate+ACoAKg- recognizes date formats according to the locale setting of your system. The correct order of day, month, and year may not be determined if it is provided in a format other than one of the recognized date settings. In addition, a long date format is not recognized if it also contains the day-of-the-week string.
A  +ACoAKg-CVDate+ACoAKg- function is also provided for compatibility with previous versions of Visual Basic. The syntax of the +ACoAKg-CVDate+ACoAKg- function is identical to the +ACoAKg-CDate+ACoAKg- function, however, +ACoAKg-CVDate+ACoAKg- returns a +ACoAKg-Variant+ACoAKg- whose subtype is +ACoAKg-Date+ACoAKg- instead of an actual +ACoAKg-Date+ACoAKg- type. Since there is now an intrinsic +ACoAKg-Date+ACoAKg- type, there is no further need for +ACoAKg-CVDate+ACoAKg-. The same effect can be achieved by converting an expression to a +ACoAKg-Date,+ACoAKg- and then assigning it to a +ACoAKg-Variant+ACoAKg-. This technique is consistent with the conversion of all other intrinsic types to their equivalent +ACoAKg-Variant+ACoAKg- subtypes.

 +ACoAKg-Note+ACoAKg-  The  +ACoAKg-CDec+ACoAKg- function does not return a discrete data type+ADs- instead, it always returns a +ACoAKg-Variant+ACoAKg- whose value has been converted to a +ACoAKg-Decimal+ACoAKg- subtype.


+ACMAIw- CBool Function Example

This example uses the  +ACoAKg-CBool+ACoAKg- function to convert an expression to a +ACoAKg-Boolean+ACoAKg-. If the expression evaluates to a nonzero value, +ACoAKg-CBool+ACoAKg- returns +ACoAKg-True+ACoAKgA7- otherwise, it returns +ACoAKg-False+ACoAKg-.


+AGAAYABg-vb
Dim A, B, Check 
A +AD0- 5: B +AD0- 5 ' Initialize variables. 
Check +AD0- CBool(A +AD0- B) ' Check contains True. 
 
A +AD0- 0 ' Define variable. 
Check +AD0- CBool(A) ' Check contains False. 

+AGAAYABg-


+ACMAIw- CByte Function Example

This example uses the  +ACoAKg-CByte+ACoAKg- function to convert an expression to a +ACoAKg-Byte+ACoAKg-.


+AGAAYABg-vb
Dim MyDouble, MyByte 
MyDouble +AD0- 125.5678 ' MyDouble is a Double. 
MyByte +AD0- CByte(MyDouble) ' MyByte contains 126. 

+AGAAYABg-


+ACMAIw- CCur Function Example

This example uses the  +ACoAKg-CCur+ACoAKg- function to convert an expression to a +ACoAKg-Currency+ACoAKg-.


+AGAAYABg-vb
Dim MyDouble, MyCurr 
MyDouble +AD0- 543.214588 ' MyDouble is a Double. 
MyCurr +AD0- CCur(MyDouble +ACo- 2) ' Convert result of MyDouble +ACo- 2 
 ' (1086.429176) to a 
 ' Currency (1086.4292). 

+AGAAYABg-


+ACMAIw- CDate Function Example

This example uses the  +ACoAKg-CDate+ACoAKg- function to convert a string to a +ACoAKg-Date+ACoAKg-. In general, hard-coding dates and times as strings (as shown in this example) is not recommended. Use date literals and time literals, such as +ACM-2/12/1969+ACM- and +ACM-4:45:23 PM+ACM-, instead.


+AGAAYABg-vb
Dim MyDate, MyShortDate, MyTime, MyShortTime 
MyDate +AD0- +ACI-February 12, 1969+ACI- ' Define date. 
MyShortDate +AD0- CDate(MyDate) ' Convert to Date data type. 
 
MyTime +AD0- +ACI-4:35:47 PM+ACI- ' Define time. 
MyShortTime +AD0- CDate(MyTime) ' Convert to Date data type. 

+AGAAYABg-


+ACMAIw- CDbl Function Example

This example uses the  +ACoAKg-CDbl+ACoAKg- function to convert an expression to a +ACoAKg-Double+ACoAKg-.


+AGAAYABg-vb
Dim MyCurr, MyDouble 
MyCurr +AD0- CCur(234.456784) ' MyCurr is a Currency. 
MyDouble +AD0- CDbl(MyCurr +ACo- 8.2 +ACo- 0.01) ' Convert result to a Double. 

+AGAAYABg-


+ACMAIw- CDec Function Example

This example uses the  +ACoAKg-CDec+ACoAKg- function to convert a numeric value to a +ACoAKg-Decimal+ACoAKg-.


+AGAAYABg-vb
Dim MyDecimal, MyCurr 
MyCurr +AD0- 10000000.0587 ' MyCurr is a Currency. 
MyDecimal +AD0- CDec(MyCurr) ' MyDecimal is a Decimal. 

+AGAAYABg-


+ACMAIw- CInt Function Example

This example uses the  +ACoAKg-CInt+ACoAKg- function to convert a value to an +ACoAKg-Integer+ACoAKg-.


+AGAAYABg-vb
Dim MyDouble, MyInt 
MyDouble +AD0- 2345.5678 ' MyDouble is a Double. 
MyInt +AD0- CInt(MyDouble) ' MyInt contains 2346. 

+AGAAYABg-


+ACMAIw- CLng Function Example

This example uses the  +ACoAKg-CLng+ACoAKg- function to convert a value to a +ACoAKg-Long+ACoAKg-.


+AGAAYABg-vb
Dim MyVal1, MyVal2, MyLong1, MyLong2 
MyVal1 +AD0- 25427.45: MyVal2 +AD0- 25427.55 ' MyVal1, MyVal2 are Doubles. 
MyLong1 +AD0- CLng(MyVal1) ' MyLong1 contains 25427. 
MyLong2 +AD0- CLng(MyVal2) ' MyLong2 contains 25428. 

+AGAAYABg-


+ACMAIw- CSng Function Example

This example uses the  +ACoAKg-CSng+ACoAKg- function to convert a value to a +ACoAKg-Single+ACoAKg-.


+AGAAYABg-vb
Dim MyDouble1, MyDouble2, MySingle1, MySingle2 
' MyDouble1, MyDouble2 are Doubles. 
MyDouble1 +AD0- 75.3421115: MyDouble2 +AD0- 75.3421555 
MySingle1 +AD0- CSng(MyDouble1) ' MySingle1 contains 75.34211. 
MySingle2 +AD0- CSng(MyDouble2) ' MySingle2 contains 75.34216. 

+AGAAYABg-


+ACMAIw- CStr Function Example

This example uses the  +ACoAKg-CStr+ACoAKg- function to convert a numeric value to a +ACoAKg-String+ACoAKg-.


+AGAAYABg-vb
Dim MyDouble, MyString 
MyDouble +AD0- 437.324 ' MyDouble is a Double. 
MyString +AD0- CStr(MyDouble) ' MyString contains +ACI-437.324+ACI-. 

+AGAAYABg-


+ACMAIw- CVar Function Example

This example uses the  +ACoAKg-CVar+ACoAKg- function to convert an expression to a +ACoAKg-Variant+ACoAKg-.


+AGAAYABg-vb
Dim MyInt, MyVar 
MyInt +AD0- 4534 ' MyInt is an Integer. 
MyVar +AD0- CVar(MyInt +ACYAOw- +ACI-000+ACI-) ' MyVar contains the string 
 ' 4534000. 

+AGAAYABg-


