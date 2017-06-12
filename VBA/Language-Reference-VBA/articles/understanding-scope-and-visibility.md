---
title: Understanding Scope and Visibility
keywords: vbcn6.chm1076771
f1_keywords:
- vbcn6.chm1076771
ms.prod: office
ms.assetid: 38f2ffcc-1bb6-3e0b-2ea5-3ca2355c8b92
ms.date: 06/08/2017
---


# Understanding Scope and Visibility

Scope refers to the availability of a [variable](vbe-glossary.md), [constant](vbe-glossary.md), or [procedure](vbe-glossary.md) for use by another procedure. There are three scoping levels: [procedure-level](vbe-glossary.md), private  [module-level](vbe-glossary.md), and public module-level.

You determine the scope of a variable when you declare it. It's a good idea to declare all variables explicitly to avoid naming-conflict errors between variables with different scopes.

## Defining Procedure-Level Scope

A variable or constant defined within a procedure is not visible outside that procedure. Only the procedure that contains the variable declaration can use it. In the following example, the first procedure displays a message box that contains a string. The second procedure displays a blank message box because the variable is local to the first procedure.


```vb
Sub LocalVariable() 
 Dim strMsg As String 
 strMsg = "This variable can't be used outside this procedure." 
 MsgBox strMsg 
End Sub 
 
Sub OutsideScope() 
 MsgBox strMsg 
End Sub
```


## Defining Private Module-Level Scope

You can define module-level variables and constants in the Declarations section of a module. Module-level variables can be either public or private. Public variables are available to all procedures in all modules in a  [project](vbe-glossary.md); private variables are available only to procedures in that module. By default, variables declared with the  **Dim** statement in the Declarations section are scoped as private. However, by preceding the variable with the **Private** keyword, the scope is obvious in your code.

In the following example, the string variable  `strMsg` is available to any procedures defined in the module. When the second procedure is called, it displays the contents of the string variable is available to any procedures defined in the module. When the second procedure is called, it displays the contents of the string variable `strMsg` in a dialog box.




```vb
' Add following to Declarations section of module. 
Private strMsg sAs String 
 
Sub InitializePrivateVariable() 
 strMsg = "This variable can't be used outside this module." 
End Sub 
 
Sub UsePrivateVariable() 
 MsgBox strMsg 
End Sub
```


 **Note**  Public procedures in a [standard module](vbe-glossary.md) or [class module](vbe-glossary.md) are available to any [referencing project](vbe-glossary.md). To limit the scope of all procedures in a module to the current project, add an  **Option Private Module** statement to the Declarations section of the module. Public variables and procedures will still be available to other procedures in the current project, but not to referencing projects.


## Defining Public Module-Level Scope

If you declare a module-level variable as public, it's available to all procedures in the project. In the following example, the string variable can be used by any procedure in any module in the project.


```vb
' Include in Declarations section of module. 
Public strMsg As String 

```

All procedures are public by default, except for event procedures. When Visual Basic creates an event procedure, the  **Private** [keyword](vbe-glossary.md) is automatically inserted before the procedure declaration. For all other procedures, you must explicitly declare the procedure with the **Private** keyword if you do not want it to be public.

You can use public procedures, variables, and constants defined in standard modules or class modules from referencing projects. However, you must first set a reference to the project in which they are defined.

Public procedures, variables, and constants defined in other than standard or class modules, such as [form modules](vbe-glossary.md) or report modules, are not available to referencing projects, because these modules are private to the project in which they reside.


