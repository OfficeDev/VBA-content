---
title: Type Statement
keywords: vblr6.chm1009049
f1_keywords:
- vblr6.chm1009049
ms.prod: office
ms.assetid: e253420f-2074-6c2a-49c3-6474d2439d5f
ms.date: 06/08/2017
---


# Type Statement

Used at [module level](vbe-glossary.md) to define a user-defined[data type](vbe-glossary.md) containing one or more elements.

 **Syntax**

[ **Private** | **Public** ] **Type** _varname_  
 _elementname_ [ **(** [ _subscripts_ ] **)** ] **As** _type_  
 [ _elementname_ [ **(** [ _subscripts_ ] **)** ] **As** _type_ ]  
 **. . .**  

 **End Type**  
The  **Type** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Public**|Optional. Used to declare [user-defined types](vbe-glossary.md) that are available to all[procedures](vbe-glossary.md) in all[modules](vbe-glossary.md) in all[projects](vbe-glossary.md).|
|**Private**|Optional. Used to declare user-defined types that are available only within the module where the [declaration](vbe-glossary.md) is made.|
| _varname_|Required. Name of the user-defined type; follows standard [variable](vbe-glossary.md) naming conventions.|
| _elementname_|Required. Name of an element of the user-defined type. Element names also follow standard variable naming conventions, except that [keyword](vbe-glossary.md)s can be used.|
| _subscripts_|When not explicitly stated in  _lower_, the lower bound of an array is controlled by the **Option** **Base** statement. The lower bound is zero if no **Option** **Base** statement is present.|
| _type_|Required. Data type of the element; may be [Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported),[Date](vbe-glossary.md), [String](vbe-glossary.md) (for variable-length strings), **String** * _length_ (for fixed-length strings),[Object](vbe-glossary.md), [Variant](vbe-glossary.md), another user-defined type, or an [object type](vbe-glossary.md).|

**Remarks**

The **Type** statement can be used only at module level. Once you have declared a user-defined type using the **Type** statement, you can declare a variable of that type anywhere within the[scope](vbe-glossary.md) of the declaration. Use **Dim**, **Private**, **Public**, **ReDim**, or **Static** to declare a variable of a user-defined type.
In [standard modules](vbe-glossary.md) and[class modules](vbe-glossary.md), user-defined types are public by default. This visibility can be changed using the  **Private** keyword.
[Line numbers](vbe-glossary.md) and[line labels](vbe-glossary.md) aren't allowed in **Type...End Type** blocks.
User-defined types are often used with data records, which frequently consist of a number of related elements of different data types.
The following example shows the use of fixed-size arrays in a user-defined type:



```vb
Type StateData 
    CityCode (1 To 100) As Integer    ' Declare a static array. 
    County As String * 30 
End Type 
 
Dim Washington(1 To 100) As StateData 

```

In the preceding example,  `StateData` includes the `CityCode` static array, and the record `Washington` has the same structure as `StateData`.
When you declare a fixed-size array within a user-defined type, its dimensions must be declared with numeric literals or [constants](vbe-glossary.md) rather than variables.

## Example

This example uses the  **Type** statement to define a user-defined data type. The **Type** statement is used at the module level only. If it appears in a class module, a **Type** statement must be preceded by the keyword **Private**.


```vb
Type EmployeeRecord    ' Create user-defined type. 
    ID As Integer    ' Define elements of data type. 
    Name As String * 20 
    Address As String * 30 
    Phone As Long 
    HireDate As Date 
End Type 
Sub CreateRecord() 
    Dim MyRecord As EmployeeRecord    ' Declare variable. 
 
    ' Assignment to EmployeeRecord variable must occur in a procedure. 
    MyRecord.ID = 12003    ' Assign a value to an element. 
End Sub
```


