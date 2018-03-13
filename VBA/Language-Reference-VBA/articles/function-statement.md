---
title: Function Statement
keywords: vblr6.chm1008927
f1_keywords:
- vblr6.chm1008927
ms.prod: office
ms.assetid: 407a6e70-b3e4-f13a-bda9-59296b288287
ms.date: 06/08/2017
---


# Function Statement

Declares the name, [arguments](vbe-glossary.md), and code that form the body of a  **Function**[procedure](vbe-glossary.md).

 **Syntax**

[ **Public** |**Private | Friend** ] [ **Static** ] **Function**_name_ [ **(**_arglist_**)** ] [ **As**_type_ ]
[ _statements_ ]
[ _name_**=**_expression_ ]
[ **Exit Function** ]
[ _statements_ ]
[ _name_**=**_expression_ ]

 **End Function**
The  **Function** statement syntax has these parts:


| <strong>Part</strong>    | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
|:-------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Public</strong>  | Optional. Indicates that the  <strong>Function</strong> procedure is accessible to all other procedures in all[modules](vbe-glossary.md). If used in a module that contains an  <strong>Option Private</strong>, the procedure is not available outside the[project](vbe-glossary.md).                                                                                                                                                                                                                                                             |
| <strong>Private</strong> | Optional. Indicates that the  <strong>Function</strong> procedure is accessible only to other procedures in the module where it is declared.                                                                                                                                                                                                                                                                                                                                                                                                       |
| <strong>Friend</strong>  | Optional. Used only in a [class module](vbe-glossary.md). Indicates that the  <strong>Function</strong> procedure is visible throughout the project, but not visible to a controller of an instance of an object.                                                                                                                                                                                                                                                                                                                                  |
| <strong>Static</strong>  | Optional. Indicates that the  <strong>Function</strong> procedure's local[variables](vbe-glossary.md) are preserved between calls. The <strong>Static</strong> attribute doesn't affect variables that are declared outside the <strong>Function</strong>, even if they are used in the procedure.                                                                                                                                                                                                                                                 |
| <em>name</em>            | Required. Name of the  <strong>Function</strong>; follows standard variable naming conventions.                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| <em>arglist</em>         | Optional. List of variables representing arguments that are passed to the  <strong>Function</strong> procedure when it is called. Multiple variables are separated by commas.                                                                                                                                                                                                                                                                                                                                                                      |
| <em>type</em>            | Optional. [Data type](vbe-glossary.md) of the value returned by the <strong>Function</strong> procedure; may be[Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported),[Date](vbe-glossary.md), [String](vbe-glossary.md), or (except fixed length), [Object](vbe-glossary.md), [Variant](vbe-glossary.md), or any [user-defined type](vbe-glossary.md). |
| <em>statements</em>      | Optional. Any group of statements to be executed within the  <strong>Function</strong> procedure.                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| <em>expression</em>      | Optional. Return value of the  <strong>Function</strong>.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |

The  _arglist_ argument has the following syntax and parts:
[ **Optional** ] [ **ByVal** |**ByRef** ] [ **ParamArray** ] _varname_ [ **( )** ] [ **As**_type_ ] [ **=**_defaultvalue_ ]


| <strong>Part</strong>       | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
|:----------------------------|:-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Optional</strong>   | Optional. Indicates that an argument is not required. If used, all subsequent arguments in  <em>arglist</em> must also be optional and declared using the <strong>Optional</strong> keyword. <strong>Optional</strong> can't be used for any argument if <strong>ParamArray</strong> is used.                                                                                                                                                                                                                                                                                |
| <strong>ByVal</strong>      | Optional. Indicates that the argument is passed [by value](vbe-glossary.md).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| <strong>ByRef</strong>      | Optional. Indicates that the argument is passed [by reference](vbe-glossary.md).  <strong>ByRef</strong> is the default in Visual Basic.                                                                                                                                                                                                                                                                                                                                                                                                                                     |
| <strong>ParamArray</strong> | Optional. Used only as the last argument in  <em>arglist</em> to indicate that the final argument is an <strong>Optional</strong> array of <strong>Variant</strong> elements. The <strong>ParamArray</strong> keyword allows you to provide an arbitrary number of arguments. It may not be used with <strong>ByVal</strong>, <strong>ByRef</strong>, or <strong>Optional</strong>.                                                                                                                                                                                          |
| <em>varname</em>            | Required. Name of the variable representing the argument; follows standard variable naming conventions.                                                                                                                                                                                                                                                                                                                                                                                                                                                                      |
| <em>type</em>               | Optional. Data type of the argument passed to the procedure; may be  <strong>Byte</strong>, <strong>Boolean</strong>, <strong>Integer</strong>, <strong>Long</strong>, <strong>Currency</strong>, <strong>Single</strong>, <strong>Double</strong>, <strong>Decimal</strong> (not currently supported) <strong>Date</strong>, <strong>String</strong> (variable length only), <strong>Object</strong>, <strong>Variant</strong>, or a specific[object type](vbe-glossary.md). If the parameter is not  <strong>Optional</strong>, a user-defined type may also be specified. |
| <em>defaultvalue</em>       | Optional. Any [constant](vbe-glossary.md) or constant expression. Valid for <strong>Optional</strong> parameters only. If the type is an <strong>Object</strong>, an explicit default value can only be <strong>Nothing</strong>.                                                                                                                                                                                                                                                                                                                                            |

 **Remarks**
If not explicitly specified using  **Public**, **Private**, or **Friend**, **Function** procedures are public by default. If **Static** isn't used, the value of local variables is not preserved between calls. The **Friend** keyword can only be used in class modules. However, **Friend** procedures can be accessed by procedures in any module of a project. A **Friend** procedure does't appear in the[type library](vbe-glossary.md) of its parent class, nor can a **Friend** procedure be late bound.
 **Function** procedures can be recursive; that is, they can call themselves to perform a given task. However, recursion can lead to stack overflow. The **Static** keyword usually isn't used with recursive **Function** procedures.
All executable code must be in procedures. You can't define a  **Function** procedure inside another **Function**, **Sub**, or **Property** procedure.
The  **Exit Function** statement causes an immediate exit from a **Function** procedure. Program execution continues with the statement following the statement that called the **Function** procedure. Any number of **Exit Function** statements can appear anywhere in a **Function** procedure.
Like a  **Sub** procedure, a **Function** procedure is a separate procedure that can take arguments, perform a series of statements, and change the values of its arguments. However, unlike a **Sub** procedure, you can use a **Function** procedure on the right side of an[expression](vbe-glossary.md) in the same way you use any intrinsic function, such as **Sqr**, **Cos**, or **Chr**, when you want to use the value returned by the function.
You call a  **Function** procedure using the function name, followed by the argument list in parentheses, in an expression. See the **Call** statement for specific information on how to call **Function** procedures.
To return a value from a function, assign the value to the function name. Any number of such assignments can appear anywhere within the procedure. If no value is assigned to  _name_, the procedure returns a default value: a numeric function returns 0, a string function returns a zero-length string (""), and a **Variant** function returns[Empty](vbe-glossary.md). A function that returns an object reference returns  **Nothing** if no object reference is assigned to _name_ (using **Set** ) within the **Function**.
The following example shows how to assign a return value to a function named . In this case,  **False** is assigned to the name to indicate that some value was not found.



```vb
Function BinarySearch(. . .) As Boolean 
'. . . 
 ' Value not found. Return a value of False. 
 If lower > upper Then 
 BinarySearch = False 
 Exit Function 
 End If 
'. . . 
End Function
```

Variables used in  **Function** procedures fall into two categories: those that are explicitly declared within the procedure and those that are not. Variables that are explicitly declared in a procedure (using **Dim** or the equivalent) are always local to the procedure. Variables that are used but not explicitly declared in a procedure are also local unless they are explicitly declared at some higher level outside the procedure.
A procedure can use a variable that is not explicitly declared in the procedure, but a naming conflict can occur if anything you defined at the [module level](vbe-glossary.md) has the same name. If your procedure refers to an undeclared variable that has the same name as another procedure, constant, or variable, it is assumed that your procedure refers to that module-level name. Explicitly declare variables to avoid this kind of conflict. You can use an **Option Explicit** statement to force explicit declaration of variables.
Visual Basic may rearrange arithmetic expressions to increase internal efficiency. Avoid using a  **Function** procedure in an arithmetic expression when the function changes the value of variables in the same expression.

## Example

This example uses the  **Function** statement to declare the name, arguments, and code that form the body of a **Function** procedure. The last example uses hard-typed, initialized **Optional** arguments.


```vb
' The following user-defined function returns the square root of the 
' argument passed to it. 
Function CalculateSquareRoot(NumberArg As Double)As Double 
 If NumberArg < 0 Then ' Evaluate argument. 
 Exit Function ' Exit to calling procedure. 
 Else 
 CalculateSquareRoot = Sqr(NumberArg) ' Return square root. 
 End If 
End Function
```

Using the  **ParamArray** keyword enables a function to accept a variable number of arguments. In the following definition, is passed by value.




```vb
Function CalcSum(ByVal FirstArg As Integer,ParamArray OtherArgs()) 
Dim ReturnValue 
' If the function is invoked as follows: 
ReturnValue = CalcSum(4, 3 ,2 ,1) 
' Local variables are assigned the following values: FirstArg = 4, 
' OtherArgs(1) = 3, OtherArgs(2) = 2, and so on, assuming default 
' lower bound for arrays = 1. 
```

 **Optional** arguments can have default values and types other than **Variant**.




```vb
' If a function's arguments are defined as follows: 
Function MyFunc(MyStr As String,Optional MyArg1 As _ Integer = 5,Optional MyArg2 = "Dolly") 
Dim RetVal 
' The function can be invoked as follows: 
RetVal = MyFunc("Hello", 2, "World") ' All 3 arguments supplied. 
RetVal = MyFunc("Test", , 5) ' Second argument omitted. 
' Arguments one and three using named-arguments. 
RetVal = MyFunc(MyStr:="Hello ", MyArg1:=7) 
```


