---
title: Writing a Property Procedure
keywords: vbcn6.chm1101346
f1_keywords:
- vbcn6.chm1101346
ms.prod: office
ms.assetid: 7ec62de7-4628-423e-54af-a49b0aef9d3c
ms.date: 06/08/2017
---


# Writing a Property Procedure

A  **Property** procedure is a series of Visual Basic[statements](vbe-glossary.md) that allow a programmer to create and manipulate custom properties.



-  **Property** procedures can be used to create read-only properties for[forms](vbe-glossary.md), [standard modules](vbe-glossary.md), and [class modules](vbe-glossary.md).

-  **Property** procedures should be used instead of **Public** variables in code that must be executed when the property value is set.

- Unlike  **Public** variables, **Property** procedures can have Help strings assigned to them in the[Object Browser](vbe-glossary.md).


When you create a  **Property** procedure, it becomes a property of the[module](vbe-glossary.md) containing the procedure. Visual Basic provides the following three types of **Property** procedures:


| <strong>Procedure</strong>    | <strong>Description</strong>                    |
|:------------------------------|:------------------------------------------------|
| <strong>Property Let</strong> | A procedure that sets the value of a property.  |
| <strong>Property Get</strong> | A procedure that returns the value a property.  |
| <strong>Property Set</strong> | A procedure that sets a reference to an object. |

The syntax for declaring a  <strong>Property</strong> procedure is:
[ <strong>Public</strong> |<strong>Private</strong> ] [ <strong>Static</strong> ] <strong>Property</strong> { <strong>Get</strong> |<strong>Let</strong> |<strong>Set</strong> } <em>propertyname__ [( _arguments</em> )] [ <strong>As</strong><em>type</em> ]
 
<em>statements</em>
 
<strong>End Property</strong>
 
<strong>Property</strong> procedures are usually used in pairs: <strong>Property Let</strong> with <strong>Property Get</strong> and <strong>Property Set</strong> with <strong>Property Get</strong>. Declaring a <strong>Property Get</strong> procedure alone is like declaring a read-only property. Using all three <strong>Property</strong> procedure types together is only useful for <strong>Variant</strong> variables, since only a <strong>Variant</strong> can contain either an object or other data type information. <strong>Property Set</strong> is intended for use with objects; <strong>Property Let</strong> isn't.
The required arguments in  
<strong>Property</strong> procedure declarations are shown in the following table:


| <strong>Procedure</strong>    | <strong>Declaration Syntax</strong>                                                                 |
|:------------------------------|:----------------------------------------------------------------------------------------------------|
| <strong>Property Get</strong> | <strong>Property Get</strong><em>propname</em> (1, …, <em>n</em> ) <strong>As</strong><em>type</em> |
| <strong>Property Let</strong> | <strong>Property Let</strong><em>propname</em> (1, …,,,, <em>n</em>, <em>n</em> +1)                 |
| <strong>Property Set</strong> | <strong>Property Set</strong><em>propname</em> (1, …, <em>n</em>, <em>n</em> +1)                    |

The first argument through the next to last argument (1, …,  _n_ ) must share the same names and data types in all **Property** procedures with the same name.
A  **Property Get** procedure declaration takes one less argument than the related **Property Let** and **Property Set** declarations. The data type of the **Property Get** procedure must be the same as the data type as the data type of the last argument ( _n_ +1) in the related **Property Let** and **Property Set** declarations. For example, if you declare the following **Property Let** procedure, the **Property Get** declaration must use arguments with the same name and data type as the arguments in the **Property Let** procedure.



```
Property Let Names(intX As Integer, intY As Integer, varZ As Variant) 
 ' Statement here. 
End Property 

Property Get Names(intX As Integer, intY As Integer) As Variant 
 ' Statement here. 
End Property 
```

The data type of the final argument in a  **Property Set** declaration must be either an[object type](vbe-glossary.md) or a **Variant**.

