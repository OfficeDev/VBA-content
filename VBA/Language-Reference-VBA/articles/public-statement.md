---
title: Public Statement
keywords: vblr6.chm1008831
f1_keywords:
- vblr6.chm1008831
ms.prod: office
ms.assetid: c8c8771b-d4cf-d5dc-4160-110472e252b4
ms.date: 06/08/2017
---


# Public Statement

Used at [module level](vbe-glossary.md) to declare public [variables](vbe-glossary.md) and allocate storage space.

 **Syntax**

 **Public** [ **WithEvents** ] _varname_ [ **(** [ _subscripts_ ] **)** ] [ **As** [ **New** ] _type_ ] [ **,** [ **WithEvents** ] _varname_ [ **(** [ _subscripts_ ] **)** ] [ **As** [ **New** ] _type_ ]] **. . .**

The  **Public** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**WithEvents**|Optional. [Keyword](vbe-glossary.md) specifying that _varname_ is an [object variable](vbe-glossary.md) used to respond to events triggered by an [ActiveX object](vbe-glossary.md).  **WithEvents** is valid only in [class modules](vbe-glossary.md). You can declare as many individual variables as you like using  **WithEvents**, but you can't create [arrays](vbe-glossary.md) with **WithEvents**. You can't use **New** with **WithEvents**.|
| _varname_|Required. Name of the variable; follows standard naming conventions.|
| _subscripts_|Optional. Dimensions of an array variable; up to 60 multiple dimensions may be declared. The  _subscripts_ [argument](vbe-glossary.md) uses the following syntax: [ _lower_**To** ] _upper_ [ **,** [ _lower_**To** ] _upper_ ] **. . .** When not explicitly stated in _lower_, the lower bound of an array is controlled by the **Option** **Base** statement. The lower bound is zero if no **Option** **Base** statement is present.|
|**New**|Optional. Keyword that enables implicit creation of an object. If you use  **New** when declaring the object variable, a new instance of the object is created on first reference to it, so you don't have to use the **Set** statement to assign the object reference. The **New** keyword can't be used to declare variables of any intrinsic[data type](vbe-glossary.md), can't be used to declare instances of dependent objects, and can't be used with  **WithEvents**.|
| _type_|Optional. Data type of the variable; may be [Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported), [Date](vbe-glossary.md), [String](vbe-glossary.md), (for variable-length strings),  **String** * _length_ (for fixed-length strings), [Object](vbe-glossary.md), [Variant](vbe-glossary.md), a [user-defined type](vbe-glossary.md), or an [object type](vbe-glossary.md). Use a separate  **As**_type_ clause for each variable being defined.|
 **Remarks**
Variables declared using the  **Public** statement are available to all procedures in all modules in all applications unless **Option** **Private Module** is in effect; in which case, the variables are public only within the [project](vbe-glossary.md) in which they reside.
The  **Public** statement can't be used in a class module to declare a fixed-length string variable.
Use the  **Public** statement to declare the data type of a variable. For example, the following statement declares a variable as an **Integer**:



```vb
Public NumberOfEmployees As Integer 

```

Also use a  **Public** statement to declare the object type of a variable. The following statement declares a variable for a new instance of a worksheet.



```vb
Public X As New Worksheet 

```

If the  **New** keyword is not used when declaring an object variable, the variable that refers to the object must be assigned an existing object using the **Set** statement before it can be used. Until it is assigned an object, the declared object variable has the special value **Nothing**, which indicates that it doesn't refer to any particular instance of an object.
You can also use the  **Public** statement with empty parentheses to declare a dynamic array. After declaring a dynamic array, use the **ReDim** statement within a procedure to define the number of dimensions and elements in the array. If you try to redeclare a dimension for an array variable whose size was explicitly specified in a **Private**, **Public**, or **Dim** statement, an error occurs.
If you don't specify a data type or object type and there is no  **Def**_type_ statement in the module, the variable is **Variant** by default.
When variables are initialized, a numeric variable is initialized to 0, a variable-length string is initialized to a zero-length string (""), and a fixed-length string is filled with zeros.  **Variant** variables are initialized to[Empty](vbe-glossary.md). Each element of a user-defined type variable is initialized as if it were a separate variable.

## Example

This example uses the  **Public** statement at the module level (General section) of a standard module to explicitly declare variables as public; that is, they are available to all procedures in all modules in all applications unless **Option Private Module** is in effect.


```vb
Public Number As Integer ' Public Integer variable. 
Public NameArray(1 To 5) As String ' Public array variable. 
' Multiple declarations, two Variants and one Integer, all Public. 
Public MyVar, YourVar, ThisVar As Integer 

```


