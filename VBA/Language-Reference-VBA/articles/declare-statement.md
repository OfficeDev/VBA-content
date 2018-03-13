---
title: Declare Statement
keywords: vblr6.chm1008781
f1_keywords:
- vblr6.chm1008781
ms.prod: office
ms.assetid: 82f68f6b-76c6-2efd-72d2-652000b3a083
ms.date: 06/08/2017
---


# Declare Statement

Used at [module level](vbe-glossary.md) to declare references to external[procedures](vbe-glossary.md) in a[dynamic-link library](vbe-glossary.md) (DLL).


 **Note**  Declare statements with the [PtrSafe](ptrsafe-keyword.md) keyword is the recommended syntax. Declare statements that include **PtrSafe** work correctly in the VBA7 development environment on both 32-bit and 64-bit platforms only after all data types in the **Declare** statement (parameters and return values) that need to store 64-bit quantities are updated to use[LongLong](longlong-data-type.md) for 64-bit integrals or[LongPtr](longptr-data-type.md) for pointers and handles. To ensure backwards compatibility with VBA version 6 and earlier use the following construct:


```vb
#If VBA7 Then 
Declare PtrSafe Sub... 
#Else 
Declare Sub... 
#EndIf
```

 **Syntax 1**
[ **Public** |**Private** ] **Declare** **Sub**_name_**Lib** **"**_libname_**"** [ **Alias** **"**_aliasname_**"** ] [ **(** [ _arglist_ ] **)** ]
 **Syntax 2**
[ **Public** |**Private** ] **Declare** **Function**_name_**Lib** **"**_libname_**"** [ **Alias** **"**_aliasname_**"** ] [ **(** [ _arglist_ ] **)** ] [ **As**_type_ ]
VBA7 Declare Statement Syntax

 **Note**  For code to run in 64-bit versions of Microsoft Office all Declare statements must include the  **PtrSafe** keyword, and all data types in the **Declare** statement (parameters and return values) that need to store 64-bit quantities must be updated to use[LongLong](longlong-data-type.md) for 64-bit integrals or[LongPtr](longptr-data-type.md) for pointers and handles.

 <strong>Syntax 1 (Sub)</strong>

[ <strong>Public</strong> |<strong>Private</strong> ] <strong>Declare</strong> <strong>PtrSafe</strong> ** Sub<strong><em>name</em></strong>Lib** <strong>"</strong><em>libname</em><strong>"</strong> [ <strong>Alias</strong> <strong>"</strong><em>aliasname</em><strong>"</strong> ] [ <strong>(</strong> [ <em>arglist</em> ] <strong>)</strong> ]
 
<strong>Syntax 2 (Function)</strong>

[ <strong>Public</strong> |<strong>Private</strong> ] <strong>Declare</strong> <strong>PtrSafe</strong> <strong>Function</strong><em>name</em><strong>Lib</strong> <strong>"</strong><em>libname</em><strong>"</strong> [ <strong>Alias</strong> <strong>"</strong><em>aliasname</em><strong>"</strong> ] [ <strong>(</strong> [ <em>arglist</em> ] <strong>)</strong> ] [ <strong>As</strong><em>type</em> ]


| <strong>Part</strong>     | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              |
|:--------------------------|:----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Public</strong>   | Optional. Used to declare procedures that are available to all other procedures in all [modules](vbe-glossary.md).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| <strong>Private</strong>  | Optional. Used to declare procedures that are available only within the module where the [declaration](vbe-glossary.md) is made.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| <strong>PtrSafe</strong>  | Required on 64-bit. The  <strong><a href="ptrsafe-keyword.md" data-raw-source="[PtrSafe](ptrsafe-keyword.md)">PtrSafe</a></strong> keyword asserts that a <strong>Declare</strong> statement is safe to run in 64-bit versions of Microsoft Office                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| <strong>Sub</strong>      | Optional (either  <strong>Sub</strong> or <strong>Function</strong> must appear). Indicates that the procedure doesn't return a value.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| <strong>Function</strong> | Optional (either  <strong>Sub</strong> or <strong>Function</strong> must appear). Indicates that the procedure returns a value that can be used in an[expression](vbe-glossary.md).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       |
| <em>name</em>             | Required. Any valid procedure name. Note that DLL entry points are case sensitive.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| <strong>Lib</strong>      | Required. Indicates that a DLL or code resource contains the procedure being declared. The  <strong>Lib</strong> clause is required for all declarations.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| <em>libname</em>          | Required. Name of the DLL or code resource that contains the declared procedure.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| <strong>Alias</strong>    | Optional. Indicates that the procedure being called has another name in the DLL. This is useful when the external procedure name is the same as a keyword. You can also use  <strong>Alias</strong> when a DLL procedure has the same name as a public[variable](vbe-glossary.md), [constant](vbe-glossary.md), or any other procedure in the same [scope](vbe-glossary.md).  <strong>Alias</strong> is also useful if any characters in the DLL procedure name aren't allowed by the DLL naming convention.                                                                                                                                                                                              |
| <em>aliasname</em>        | Optional. Name of the procedure in the DLL or code resource. If the first character is not a number sign ( <strong>#</strong> ), <em>aliasname</em> is the name of the procedure's entry point in the DLL. If ( <strong>#</strong> ) is the first character, all characters that follow must indicate the ordinal number of the procedure's entry point.                                                                                                                                                                                                                                                                                                                                                  |
| <em>arglist</em>          | Optional. List of variables representing [arguments](vbe-glossary.md) that are passed to the procedure when it is called.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| <em>type</em>             | Optional. [Data type](vbe-glossary.md) of the value returned by a <strong>Function</strong> procedure; may be[Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [LongLong](longlong-data-type.md), [LongPtr](longptr-data-type.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported),[Date](vbe-glossary.md), [String](vbe-glossary.md) (variable length only), or[Variant](vbe-glossary.md), a [user-defined type](vbe-glossary.md), or an [object type](vbe-glossary.md). ( <strong>LongLong</strong> is a valid declared type only on 64-bit platforms.) |

The  _arglist_ argument has the following syntax and parts:
[ **Optional** ] [ **ByVal** |**ByRef** ] [ **ParamArray** ] _varname_ [ **( )** ] [ **As**_type_ ]


| <strong>Part</strong>       | <strong>Description</strong>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
|:----------------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <strong>Optional</strong>   | Optional. Indicates that an argument is not required. If used, all subsequent arguments in  <em>arglist</em> must also be optional and declared using the <strong>Optional</strong> keyword. <strong>Optional</strong> can't be used for any argument if <strong>ParamArray</strong> is used.                                                                                                                                                                                                                                                                                                                  |
| <strong>ByVal</strong>      | Optional. Indicates that the argument is passed [by value](vbe-glossary.md).                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   |
| <strong>ByRef</strong>      | Indicates that the argument is passed [by reference](vbe-glossary.md).  <strong>ByRef</strong> is the default in Visual Basic.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 |
| <strong>ParamArray</strong> | Optional. Used only as the last argument in  <em>arglist</em> to indicate that the final argument is an <strong>Optional</strong>[array](vbe-glossary.md) of <strong>Variant</strong> elements. The <strong>ParamArray</strong> keyword allows you to provide an arbitrary number of arguments. The <strong>ParamArray</strong> keyword can't be used with <strong>ByVal</strong>, <strong>ByRef</strong>, or <strong>Optional</strong>.                                                                                                                                                                       |
| <em>varname</em>            | Required. Name of the variable representing the argument being passed to the procedure; follows standard variable naming conventions.                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| <strong>( )</strong>        | Required for array variables. Indicates that  <em>varname</em> is an array.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| <em>type</em>               | Optional. Data type of the argument passed to the procedure; may be  <strong>Byte</strong>, <strong>Boolean</strong>, <strong>Integer</strong>, <strong>Long</strong>, <strong>LongLong</strong>, <strong>LongPtr</strong>, <strong>Currency</strong>, <strong>Single</strong>, <strong>Double</strong>, <strong>Decimal</strong> (not currently supported), <strong>Date</strong>, <strong>String</strong> (variable length only), <strong>Object</strong>, <strong>Variant</strong>, a user-defined type, or an object type. ( <strong>LongLong</strong> is a valid declared type only on 64-bit platforms.) |

 **Remarks**
For  **Function** procedures, the data type of the procedure determines the data type it returns. You can use an **As** clause following _arglist_ to specify the return type of the function. Within _arglist_, you can use an **As** clause to specify the data type of any of the arguments passed to the procedure. In addition to specifying any of the standard data types, you can specify **As Any** in _arglist_ to inhibit type checking and allow any data type to be passed to the procedure.
Empty parentheses indicate that the  **Sub** or **Function** procedure has no arguments and that Visual Basic should ensure that none are passed. In the following example, `First` takes no arguments. If you use arguments in a call to takes no arguments. If you use arguments in a call to `First`, an error occurs:



```vb
Declare Sub First Lib "MyLib" () 
```

If you include an argument list, the number and type of arguments are checked each time the procedure is called. In the following example, takes one  **Long** argument:



```vb
Declare Sub First Lib "MyLib" (X As Long) 
```


 **Note**  You can't have fixed-length strings in the argument list of a  **Declare** statement; only variable-length strings can be passed to procedures. Fixed-length strings can appear as procedure arguments, but they are converted to variable-length strings before being passed.


 **Note**  The  **vbNullString** constant is used when calling external procedures, where the external procedure requires a string whose value is zero. This is not the same thing as a zero-length string ("").


## Example

This example shows how the  **Declare** statement is used at the module level of a standard module to declare a reference to an external procedure in a dynamic-link library (DLL). You can place the **Declare** statements in class modules if the **Declare** statements are **Private**.


 **Note**  


```vb
' In Microsoft Windows (16-bit): 
Declare Sub MessageBeep Lib "User" (ByVal N As Integer) 
' Assume SomeBeep is an alias for the procedure name. 
Declare Sub MessageBeep Lib "User" Alias "SomeBeep"(ByVal N As Integer) 
' Use an ordinal in the Alias clause to call GetWinFlags. 
Declare Function GetWinFlags Lib "Kernel" Alias "#132"()As Long 

' In 32-bit Microsoft Windows systems, specify the library USER32.DLL, 
' rather than USER.DLL. You can use conditional compilation to write 
' code that can run on either Win32 or Win16. 
#If Win32 Then 
    Declare Sub MessageBeep Lib "User32" (ByVal N As Long) 
#Else 
    Declare Sub MessageBeep Lib "User" (ByVal N As Integer) 
#End If 


' 64-bit Declare statement example: 
Declare PtrSafe Function GetActiveWindow Lib "User32" () As LongPtr 

' Conditional Compilation Example 
#If Vba7 Then 
     ' Code is running in  32-bit or 64-bit VBA7. 
     #If Win64 Then 
          ' Code is running in 64-bit VBA7. 
     #Else 
          ' Code is not running in 64-bit VBA7. 
     #End If 
#Else 
     ' Code is NOT running in 32-bit or 64-bit VBA7. 
#End If 
```


