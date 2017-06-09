---
title: ReDim Statement
keywords: vblr6.chm1008999
f1_keywords:
- vblr6.chm1008999
ms.prod: office
ms.assetid: 5044cb55-6cdc-16a7-6558-dcff7ab4b933
ms.date: 06/08/2017
---


# ReDim Statement

Used at [procedure level](vbe-glossary.md) to reallocate storage space for dynamic array[variables](vbe-glossary.md).

 **Syntax**

 **ReDim** [ **Preserve** ] _varname_**(**_subscripts_**)** [ **As**_type_ ] [ **,**_varname_**(**_subscripts_**)** [ **As**_type_ ]] **. . .**

The  **ReDim** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Preserve**|Optional. [Keyword](vbe-glossary.md) used to preserve the data in an existing[array](vbe-glossary.md) when you change the size of the last dimension.|
| _varname_|Required. Name of the variable; follows standard variable naming conventions.|
| _subscripts_|Required. Dimensions of an array variable; up to 60 multiple dimensions may be declared. The  _subscripts_[argument](vbe-glossary.md) uses the following syntax: [ _lower_**To** ] _upper_ [ **,** [ _lower_**To** ] _upper_ ] **. . .** When not explicitly stated in _lower_, the lower bound of an array is controlled by the **Option** **Base** statement. The lower bound is zero if no **Option** **Base** statement is present.|
| _type_|Optional. [Data type](vbe-glossary.md) of the variable; may be[Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported),[Date](vbe-glossary.md), [String](vbe-glossary.md) (for variable-length strings), **String** * _length_ (for fixed-length strings),[Object](vbe-glossary.md), [Variant](vbe-glossary.md), a [user-defined type](vbe-glossary.md), or an [object type](vbe-glossary.md). Use a separate  **As**_type_ clause for each variable being defined. For a **Variant** containing an array, _type_ describes the type of each element of the array, but doesn't change the **Variant** to some other type.|
 **Remarks**
The  **ReDim**[statement](vbe-glossary.md) is used to size or resize a dynamic array that has already been formally declared using a **Private**, **Public**, or **Dim** statement with empty parentheses (without dimension subscripts).
You can use the  **ReDim** statement repeatedly to change the number of elements and dimensions in an array. However, you can't declare an array of one data type and later use **ReDim** to change the array to another data type, unless the array is contained in a **Variant**. If the array is contained in a **Variant**, the type of the elements can be changed using an **As**_type_ clause, unless you're using the **Preserve** keyword, in which case, no changes of data type are permitted.
If you use the  **Preserve** keyword, you can resize only the last array dimension and you can't change the number of dimensions at all. For example, if your array has only one dimension, you can resize that dimension because it is the last and only dimension. However, if your array has two or more dimensions, you can change the size of only the last dimension and still preserve the contents of the array. The following example shows how you can increase the size of the last dimension of a dynamic array without erasing any existing data contained in the array.



```
ReDim X(10, 10, 10) 
. . . 
ReDim Preserve X(10, 10, 15) 

```

Similarly, when you use  **Preserve**, you can change the size of the array only by changing the upper bound; changing the lower bound causes an error.
If you make an array smaller than it was, data in the eliminated elements will be lost. If you pass an array to a procedure by reference, you can't redimension the array within the procedure.
When variables are initialized, a numeric variable is initialized to 0, a variable-length string is initialized to a zero-length string (""), and a fixed-length string is filled with zeros.  **Variant** variables are initialized to[Empty](vbe-glossary.md). Each element of a user-defined type variable is initialized as if it were a separate variable. A variable that refers to an object must be assigned an existing object using the  **Set** statement before it can be used. Until it is assigned an object, the declared[object variable](vbe-glossary.md) has the special value **Nothing**, which indicates that it doesn't refer to any particular instance of an object.
The  **ReDim** statement acts as a declarative statement if the variable it declares doesn't exist at[module level](vbe-glossary.md) or[procedure level](vbe-glossary.md). If another variable with the same name is created later, even in a wider [scope](vbe-glossary.md),  **ReDim** will refer to the later variable and won't necessarily cause a compilation error, even if **Option Explicit** is in effect. To avoid such conflicts, **ReDim** should not be used as a declarative statement, but simply for redimensioning arrays.

 **Note**  To resize an array contained in a  **Variant**, you must explicitly declare the **Variant** variable before attempting to resize its array.


## Example

This example uses the  **ReDim** statement to allocate and reallocate storage space for dynamic-array variables. It assumes the **Option Base** is **1**.


```vb
Dim MyArray() As Integer ' Declare dynamic array. 
Redim MyArray(5) ' Allocate 5 elements. 
For I = 1 To 5 ' Loop 5 times. 
 MyArray(I) = I ' Initialize array. 
Next I 

```

The next statement resizes the array and erases the elements.




```
Redim MyArray(10) ' Resize to 10 elements. 
For I = 1 To 10 ' Loop 10 times. 
 MyArray(I) = I ' Initialize array. 
Next I 

```

The following statement resizes the array but does not erase elements.




```
Redim Preserve MyArray(15) ' Resize to 15 elements. 

```


