---
title: Sub Statement
keywords: vblr6.chm1009038
f1_keywords:
- vblr6.chm1009038
ms.prod: office
ms.assetid: 7931d739-a61a-78ba-5b33-960c1bf908ce
ms.date: 06/08/2017
---


# Sub Statement

Declares the name, [arguments](vbe-glossary.md), and code that form the body of a  **Sub** [procedure](vbe-glossary.md).

 **Syntax**

[ **Private** |**Public** |**Friend** ] [ **Static** ] **Sub** _name_ [ **(**_arglist_**)** ]
[ _statements_ ]
[ **Exit Sub** ]
[ _statements_ ]

 **End Sub**
The  **Sub** statement syntax has these parts:


| <strong>Part</strong>    | <strong>Description</strong>                                                                                                                                                                                                                                                                  |
|:-------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Public</strong>  | Optional. Indicates that the  <strong>Sub</strong> procedure is accessible to all other procedures in all [modules](vbe-glossary.md). If used in a module that contains an  <strong>Option Private</strong> statement, the procedure is not available outside the [project](vbe-glossary.md). |
| <strong>Private</strong> | Optional. Indicates that the  <strong>Sub</strong> procedure is accessible only to other procedures in the module where it is declared.                                                                                                                                                       |
| <strong>Friend</strong>  | Optional. Used only in a [class module](vbe-glossary.md). Indicates that the  <strong>Sub</strong> procedure is visible throughout the [project](vbe-glossary.md), but not visible to a controller of an instance of an object.                                                               |
| <strong>Static</strong>  | Optional. Indicates that the  <strong>Sub</strong> procedure's local [variables](vbe-glossary.md) are preserved between calls. The <strong>Static</strong> attribute doesn't affect variables that are declared outside the <strong>Sub</strong>, even if they are used in the procedure.     |
| <em>name</em>            | Required. Name of the  <strong>Sub</strong>; follows standard [variable](vbe-glossary.md) naming conventions.                                                                                                                                                                                 |
| <em>arglist</em>         | Optional. List of variables representing arguments that are passed to the  <strong>Sub</strong> procedure when it is called. Multiple variables are separated by commas.                                                                                                                      |
| <em>statements</em>      | Optional. Any group of [statements](vbe-glossary.md) to be executed within the <strong>Sub</strong> procedure.                                                                                                                                                                                |

The  _arglist_ argument has the following syntax and parts:
[ **Optional** ] [ **ByVal** |**ByRef** ] [ **ParamArray** ] _varname_ [ **( )** ] [ **As**_type_ ] [ **=**_defaultvalue_ ]


| <strong>Part</strong>       | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
|:----------------------------|:-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Optional</strong>   | Optional. [Keyword](vbe-glossary.md) indicating that an argument is not required. If used, all subsequent arguments in <em>arglist</em> must also be optional and declared using the <strong>Optional</strong> keyword. <strong>Optional</strong> can't be used for any argument if <strong>ParamArray</strong> is used.                                                                                                                                                                                                                                                                                                                   |
| <strong>ByVal</strong>      | Optional. Indicates that the argument is passed [by value](vbe-glossary.md).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                               |
| <strong>ByRef</strong>      | Optional. Indicates that the argument is passed [by reference](vbe-glossary.md).  <strong>ByRef</strong> is the default in Visual Basic.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| <strong>ParamArray</strong> | Optional. Used only as the last argument in  <em>arglist</em> to indicate that the final argument is an <strong>Optional</strong> [array](vbe-glossary.md) of <strong>Variant</strong> elements. The <strong>ParamArray</strong> keyword allows you to provide an arbitrary number of arguments. <strong>ParamArray</strong> can't be used with <strong>ByVal</strong>, <strong>ByRef</strong>, or <strong>Optional</strong>.                                                                                                                                                                                                              |
| <em>varname</em>            | Required. Name of the variable representing the argument; follows standard variable naming conventions.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| <em>type</em>               | Optional. [Data type](vbe-glossary.md) of the argument passed to the procedure; may be[Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported), [Date](vbe-glossary.md), [String](vbe-glossary.md) (variable-length only), [Object](vbe-glossary.md), [Variant](vbe-glossary.md), or a specific [object type](vbe-glossary.md). If the parameter is not  <strong>Optional</strong>, a [user-defined type](vbe-glossary.md) may also be specified. |
| <em>defaultvalue</em>       | Optional. Any [constant](vbe-glossary.md) or constant [expression](vbe-glossary.md). Valid for  <strong>Optional</strong> parameters only. If the type is an <strong>Object</strong>, an explicit default value can only be <strong>Nothing</strong>.                                                                                                                                                                                                                                                                                                                                                                                      |

**Remarks**

If not explicitly specified using  **Public**, **Private**, or **Friend**, **Sub** procedures are public by default.

If **Static** isn't used, the value of local variables is not preserved between calls.

The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure doesn't appear in the [type library](vbe-glossary.md) of its parent class, nor can a **Friend** procedure be late bound.

**Sub** procedures can be recursive; that is, they can call themselves to perform a given task. However, recursion can lead to stack overflow. The **Static** keyword usually is not used with recursive **Sub** procedures.

All executable code must be in [procedures](vbe-glossary.md). You can't define a  **Sub** procedure inside another **Sub**, **Function**, or **Property** procedure.

The **Exit Sub** keywords cause an immediate exit from a **Sub** procedure. Program execution continues with the statement following the statement that called the **Sub** procedure. Any number of **Exit Sub** statements can appear anywhere in a **Sub** procedure.

Like a **Function** procedure, a **Sub** procedure is a separate procedure that can take arguments, perform a series of statements, and change the value of its arguments. However, unlike a **Function** procedure, which returns a value, a **Sub** procedure can't be used in an expression.

You call a **Sub** procedure using the procedure name followed by the argument list. See the **Call** statement for specific information on how to call **Sub** procedures.

Variables used in **Sub** procedures fall into two categories: those that are explicitly declared within the procedure and those that are not. Variables that are explicitly declared in a procedure (using **Dim** or the equivalent) are always local to the procedure. Variables that are used but not explicitly declared in a procedure are also local unless they are explicitly declared at some higher level outside the procedure.

A procedure can use a variable that is not explicitly declared in the procedure, but a naming conflict can occur if anything you defined at the [module level](vbe-glossary.md) has the same name. If your procedure refers to an undeclared variable that has the same name as another procedure, constant or variable, it is assumed that your procedure is referring to that module-level name. To avoid this kind of conflict, explicitly declare variables. You can use an **Option Explicit** statement to force explicit declaration of variables.

 **Note** You can't use **GoSub**, **GoTo**, or **Return** to enter or exit a **Sub** procedure.

## Example

This example uses the **Sub** statement to define the name, arguments, and code that form the body of a **Sub** procedure.

```vb
' Sub procedure definition. 
' Sub procedure with two arguments. 
Sub SubComputeArea(Length, TheWidth) 

   Dim Area As Double ' Declare local variable. 

   If Length = 0 Or TheWidth = 0 Then 
      ' If either argument = 0. 
      Exit Sub ' Exit Sub immediately. 
   End If 

   Area = Length * TheWidth ' Calculate area of rectangle. 
   Debug.Print Area ' Print Area to Debug window. 

End Sub
```

