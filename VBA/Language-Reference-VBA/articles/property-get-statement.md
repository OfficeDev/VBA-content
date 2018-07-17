---
title: Property Get Statement
keywords: vblr6.chm1009538
f1_keywords:
- vblr6.chm1009538
ms.prod: office
ms.assetid: 39d1fb20-653e-a174-7a98-e2b33f260d39
ms.date: 06/08/2017
---


# Property Get Statement

Declares the name, [arguments](vbe-glossary.md), and code that form the body of a  **Property**[procedure](vbe-glossary.md), which gets the value of a [property](vbe-glossary.md).

 **Syntax**

[ **Public** |**Private** |**Friend** ] [ **Static** ] **Property** **Get**_name_ [ **(**_arglist_**)** ] [ **As**_type_ ]
[ _statements_ ]
[ _name_**=**_expression_ ]
[ **Exit Property** ]
[ _statements_ ]
[ _name_**=**_expression_ ]

 **End Property**
The  **Property Get** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
|**Public**|Optional. Indicates that the  **Property** **Get** procedure is accessible to all other procedures in all[modules](vbe-glossary.md). If used in a module that contains an  **Option Private** statement, the procedure is not available outside the[project](vbe-glossary.md).|
|**Private**|Optional. Indicates that the  **Property** **Get** procedure is accessible only to other procedures in the module where it is declared.|
|**Friend**|Optional. Used only in a [class module](vbe-glossary.md). Indicates that the  **Property Get** procedure is visible throughout the project, but not visible to a controller of an instance of an object.|
|**Static**|Optional. Indicates that the  **Property** **Get** procedure's local[variables](vbe-glossary.md) are preserved between calls. The **Static** attribute doesn't affect variables that are declared outside the **Property Get** procedure, even if they are used in the procedure.|
| _name_|Required. Name of the  **Property** **Get** procedure; follows standard variable naming conventions, except that the name can be the same as a **Property** **Let** or **Property Set** procedure in the same module.|
| _arglist_|Optional. List of variables representing arguments that are passed to the  **Property** **Get** procedure when it is called. Multiple arguments are separated by commas. The name and[data type](vbe-glossary.md) of each argument in a **Property** **Get** procedure must be the same as the corresponding argument in a **Property** **Let** procedure (if one exists).|
| _type_|Optional. Data type of the value returned by the  **Property** **Get** procedure; may be[Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported),[Date](vbe-glossary.md), [String](vbe-glossary.md) (except fixed length),[Object](vbe-glossary.md), [Variant](vbe-glossary.md), [user-defined type](vbe-glossary.md), and [Arrays](vbe-glossary.md). The return  _type_ of a **Property** **Get** procedure must be the same data type as the last (or sometimes the only) argument in a corresponding **Property** **Let** procedure (if one exists) that defines the value assigned to the property on the right side of an[expression](vbe-glossary.md).|
| _statements_|Optional. Any group of statements to be executed within the body of the  **Property** **Get** procedure.|
| _expression_|Optional. Value of the property returned by the procedure defined by the  **Property Get** statement.|
The  _arglist_ argument has the following syntax and parts:
[ **Optional** ] [ **ByVal** |**ByRef** ] [ **ParamArray** ] _varname_ [ **( )** ] [ **As**_type_ ] [ **=**_defaultvalue_ ]


|**Part**|**Description**|
|:-----|:-----|
|**Optional**|Optional. Indicates that an argument is not required. If used, all subsequent arguments in  _arglist_ must also be optional and declared using the **Optional** keyword.|
|**ByVal**|Optional. Indicates that the argument is passed [by value](vbe-glossary.md).|
|**ByRef**|Optional. Indicates that the argument is passed [by reference](vbe-glossary.md).  **ByRef** is the default in Visual Basic.|
|**ParamArray**|Optional. Used only as the last argument in  _arglist_ to indicate that the final argument is an **Optional** array of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. It may not be used with **ByVal**, **ByRef**, or **Optional**.|
| _varname_|Required. Name of the variable representing the argument; follows standard variable naming conventions.|
| _type_|Optional. Data type of the argument passed to the procedure; may be  **Byte**, **Boolean**, **Integer**, **Long**, **Currency**, **Single**, **Double**, **Decimal** (not currently supported), **Date**, **String** (variable length only), **Object**, **Variant**, or a specific[object type](vbe-glossary.md). If the parameter is not  **Optional**, a user-defined type may also be specified.|
| _defaultvalue_|Optional. Any [constant](vbe-glossary.md) or constant expression. Valid for **Optional** parameters only. If the type is an **Object**, an explicit default value can only be **Nothing**.|
 **Remarks**
If not explicitly specified using  **Public**, **Private**, or **Friend**, **Property** procedures are public by default. If **Static** is not used, the value of local variables is not preserved between calls. The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure doesn't appear in the[type library](vbe-glossary.md) of its parent class, nor can a **Friend** procedure be late bound.
All executable code must be in procedures. You can't define a  **Property** **Get** procedure inside another **Property**, **Sub**, or **Function** procedure.
The  **Exit Property** statement causes an immediate exit from a **Property Get** procedure. Program execution continues with the statement following the statement that called the **Property** **Get** procedure. Any number of **Exit Property** statements can appear anywhere in a **Property** **Get** procedure.
Like a  **Sub** and **Property Let** procedure, a **Property Get** procedure is a separate procedure that can take arguments, perform a series of statements, and change the values of its arguments. However, unlike a **Sub** or **Property Let** procedure, you can use a **Property Get** procedure on the right side of an expression in the same way you use a **Function** or a property name when you want to return the value of a property.

## Example

This example uses the  **Property Get** statement to define a property procedure that gets the value of a property. The property identifies the current color of a pen as a string.


```vb
Dim CurrentColor As Integer 
Const BLACK = 0, RED = 1, GREEN = 2, BLUE = 3 
 
' Returns the current color of the pen as a string. 
Property Get PenColor() As String 
 Select Case CurrentColor 
 Case RED 
 PenColor = "Red" 
 Case GREEN 
 PenColor = "Green" 
 Case BLUE 
 PenColor = "Blue" 
 End Select 
End Property 
 
' The following code gets the color of the pen 
' calling the Property Get procedure. 
ColorName = PenColor 

```


