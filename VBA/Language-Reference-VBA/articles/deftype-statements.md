---
title: Deftype Statements
keywords: vblr6.chm1008787
f1_keywords:
- vblr6.chm1008787
ms.prod: office
ms.assetid: 14396fc2-494a-9025-d8a5-86174fcc8a74
ms.date: 06/08/2017
---


# Deftype Statements



Used at [module level](vbe-glossary.md) to set the default[data type](vbe-glossary.md) for[variables](vbe-glossary.md), [arguments](vbe-glossary.md) passed to[procedures](vbe-glossary.md), and the return type for  **Function** and **Property** **Get** procedures whose names start with the specified characters.
 **Syntax**
 **DefBool**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefByte**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefInt**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefLng**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefLngLng**_letterrange_ [ **,**_letterrange_ ] **. . .** (Valid on 64-bit platforms only.)
 **DefLngPtr**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefCur**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefSng**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefDbl**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefDec**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefDate**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefStr**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefObj**_letterrange_ [ **,**_letterrange_ ] **. . .**
 **DefVar**_letterrange_ [ **,**_letterrange_ ] **. . .**
The required  _letterrange_ argument has the following syntax:
 _letter1_ [ **-**_letter2_ ]
The  _letter1_ and _letter2_ arguments specify the name range for which you can set a default data type. Each argument represents the first letter of the variable, argument, **Function** procedure, or **Property Get** procedure name and can be any letter of the alphabet. The case of letters in _letterrange_ isn't significant.
 **Remarks**
The statement name determines the data type:


|**Statement**|**Data Type**|
|:-----|:-----|
|**DefBool**|[Boolean](vbe-glossary.md)|
|**DefByte**|[Byte](vbe-glossary.md)|
|**DefInt**|[Integer](vbe-glossary.md)|
|**DefLng**|[Long](vbe-glossary.md)|
|**DefLngLng**|[LongLong](longlong-data-type.md) (Valid on 64-bit platforms only.)|
|**DefLngPtr**|[LongPtr](longptr-data-type.md)|
|**DefCur**|[Currency](vbe-glossary.md)|
|**DefSng**|[Single](vbe-glossary.md)|
|**DefDbl**|[Double](vbe-glossary.md)|
|**DefDec**|[Decimal](vbe-glossary.md) (not currently supported)|
|**DefDate**|[Date](vbe-glossary.md)|
|**DefStr**|[String](vbe-glossary.md)|
|**DefObj**|[Object](vbe-glossary.md)|
|**DefVar**|[Variant](vbe-glossary.md)|
For example, in the following program fragment, is a string variable:
A  **Def**_type_ statement affects only the[module](vbe-glossary.md) where it is used. For example, a **DefInt** statement in one module affects only the default data type of variables, arguments passed to procedures, and the return type for **Function** and **Property** **Get** procedures declared in that module; the default data type of variables, arguments, and return types in other modules is unaffected. If not explicitly declared with a **Def**_type_ statement, the default data type for all variables, all arguments, all **Function** procedures, and all **Property** **Get** procedures is **Variant**.
When you specify a letter range, it usually defines the data type for variables that begin with letters in the first 128 characters of the character set. However, when you specify the letter range A-Z, you set the default to the specified data type for all variables, including variables that begin with international characters from the extended part of the character set (128-255).
Once the range A-Z has been specified, you can't further redefine any subranges of variables using  **Def**_type_ statements. Once a range has been specified, if you include a previously defined letter in another **Def**_type_ statement, an error occurs. However, you can explicitly specify the data type of any variable, defined or not, using a **Dim** statement with an **As**_type_ clause. For example, you can use the following code at module level to define a variable as a **Double** even though the default data type is **Integer**:
 **Def**_type_ statements don't affect elements of[user-defined types](vbe-glossary.md) because the elements must be explicitly declared.

## Example

This example shows various uses of the  **Def**_type_ statements to set default data types of variables and function procedures whose names start with specified characters. The default data type can be overridden only by explicit assignment using the **Dim** statement. **Def**_type_ statements can only be used at the module level (that is, not within procedures).


