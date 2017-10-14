---
title: Option Private Statement
keywords: vblr6.chm1011061
f1_keywords:
- vblr6.chm1011061
ms.prod: office
ms.assetid: bd4d8b8b-d513-62a0-7c78-45c15b462bdc
ms.date: 06/08/2017
---


# Option Private Statement

When used in host applications that allow references across multiple [projects](vbe-glossary.md),  **Option Private Module** prevents a[module's](vbe-glossary.md) contents from being referenced outside its project. In host applications that don't permit such references, for example, standalone versions of Visual Basic, **Option Private** has no effect.

 **Syntax**

 **Option Private Module**

 **Remarks**
If used, the  **Option** **Private** statement must appear at[module level](vbe-glossary.md), before any [procedures](vbe-glossary.md).
When a module contains  **Option Private Module**, the public parts, for example,[variables](vbe-glossary.md), [objects](vbe-glossary.md), and [user-defined types](vbe-glossary.md) declared at module level, are still available within the[project](vbe-glossary.md) containing the module, but they are not available to other applications or projects.

 **Note**   **Option Private** is only useful for[host applications](vbe-glossary.md) that support simultaneous loading of multiple projects and permit references between the loaded projects. For example, Microsoft Excel permits loading of multiple projects and **Option Private Module** can be used to restrict cross-project visibility. Although Visual Basic permits loading of multiple projects, references between projects are never permitted in Visual Basic.


## Example

This example demonstrates the  **Option Private** statement, which is used at module level to indicate that the entire module is private. With **Option Private Module**, module-level parts not declared **Private** are available to other modules in the project, but not to other projects or applications.


```vb
Option Private Module ' Indicates that module is private. 

```


