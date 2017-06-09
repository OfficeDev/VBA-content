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

 **Syntax 1 (Sub)**
[ **Public** |**Private** ] **Declare** **PtrSafe** ** Sub**_name_**Lib** **"**_libname_**"** [ **Alias** **"**_aliasname_**"** ] [ **(** [ _arglist_ ] **)** ]
 **Syntax 2 (Function)**
[ **Public** |**Private** ] **Declare** **PtrSafe** **Function**_name_**Lib** **"**_libname_**"** [ **Alias** **"**_aliasname_**"** ] [ **(** [ _arglist_ ] **)** ] [ **As**_type_ ]


|**Part**|**Description**|
|:-----|:-----|
|**Public**|Optional. Used to declare procedures that are available to all other procedures in all [modules](vbe-glossary.md).|
|**Private**|Optional. Used to declare procedures that are available only within the module where the [declaration](vbe-glossary.md) is made.|
|**PtrSafe**|Required on 64-bit. The  **[PtrSafe](ptrsafe-keyword.md)** keyword asserts that a **Declare** statement is safe to run in 64-bit versions of Microsoft Office|
|**Sub**|Optional (either  **Sub** or **Function** must appear). Indicates that the procedure doesn't return a value.|
|**Function**|Optional (either  **Sub** or **Function** must appear). Indicates that the procedure returns a value that can be used in an[expression](vbe-glossary.md).|
| _name_|Required. Any valid procedure name. Note that DLL entry points are case sensitive.|
|**Lib**|Required. Indicates that a DLL or code resource contains the procedure being declared. The  **Lib** clause is required for all declarations.|
| _libname_|Required. Name of the DLL or code resource that contains the declared procedure.|
|**Alias**|Optional. Indicates that the procedure being called has another name in the DLL. This is useful when the external procedure name is the same as a keyword. You can also use  **Alias** when a DLL procedure has the same name as a public[variable](vbe-glossary.md), [constant](vbe-glossary.md), or any other procedure in the same [scope](vbe-glossary.md).  **Alias** is also useful if any characters in the DLL procedure name aren't allowed by the DLL naming convention.|
| _aliasname_|Optional. Name of the procedure in the DLL or code resource. If the first character is not a number sign ( **#** ), _aliasname_ is the name of the procedure's entry point in the DLL. If ( **#** ) is the first character, all characters that follow must indicate the ordinal number of the procedure's entry point.|
| _arglist_|Optional. List of variables representing [arguments](vbe-glossary.md) that are passed to the procedure when it is called.|
| _type_|Optional. [Data type](vbe-glossary.md) of the value returned by a **Function** procedure; may be[Byte](vbe-glossary.md), [Boolean](vbe-glossary.md), [Integer](vbe-glossary.md), [Long](vbe-glossary.md), [LongLong](longlong-data-type.md), [LongPtr](longptr-data-type.md), [Currency](vbe-glossary.md), [Single](vbe-glossary.md), [Double](vbe-glossary.md), [Decimal](vbe-glossary.md) (not currently supported),[Date](vbe-glossary.md), [String](vbe-glossary.md) (variable length only), or[Variant](vbe-glossary.md), a [user-defined type](vbe-glossary.md), or an [object type](vbe-glossary.md). ( **LongLong** is a valid declared type only on 64-bit platforms.)|
The  _arglist_ argument has the following syntax and parts:
[ **Optional** ] [ **ByVal** |**ByRef** ] [ **ParamArray** ] _varname_ [ **( )** ] [ **As**_type_ ]


|**Part**|**Description**|
|:-----|:-----|
|**Optional**|Optional. Indicates that an argument is not required. If used, all subsequent arguments in  _arglist_ must also be optional and declared using the **Optional** keyword. **Optional** can't be used for any argument if **ParamArray** is used.|
|**ByVal**|Optional. Indicates that the argument is passed [by value](vbe-glossary.md).|
|**ByRef**|Indicates that the argument is passed [by reference](vbe-glossary.md).  **ByRef** is the default in Visual Basic.|
|**ParamArray**|Optional. Used only as the last argument in  _arglist_ to indicate that the final argument is an **Optional**[array](vbe-glossary.md) of **Variant** elements. The **ParamArray** keyword allows you to provide an arbitrary number of arguments. The **ParamArray** keyword can't be used with **ByVal**, **ByRef**, or **Optional**.|
| _varname_|Required. Name of the variable representing the argument being passed to the procedure; follows standard variable naming conventions.|
|**( )**|Required for array variables. Indicates that  _varname_ is an array.|
| _type_|Optional. Data type of the argument passed to the procedure; may be  **Byte**, **Boolean**, **Integer**, **Long**, **LongLong**, **LongPtr**, **Currency**, **Single**, **Double**, **Decimal** (not currently supported), **Date**, **String** (variable length only), **Object**, **Variant**, a user-defined type, or an object type. ( **LongLong** is a valid declared type only on 64-bit platforms.)|
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


