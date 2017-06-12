---
title: Set Reference to a Type Library
keywords: vbhw6.chm1105230
f1_keywords:
- vbhw6.chm1105230
ms.prod: office
ms.assetid: 5b695bf5-5ab3-977e-b037-13aea3097b9c
ms.date: 06/08/2017
---


# Set Reference to a Type Library

Automation (formerly OLE Automation) enables you to use objects from other applications in Visual Basic code. An application that provides its objects for use by other applications also provides information about those objects in a [type library](vbe-glossary.md). To achieve the best possible performance when using another application's objects, you should set a reference to that application's type library.

 **To set a reference to an application's type library**




1. Click  **References** on the **Tools** menu.
    
2. Select the check boxes for the applications with type libraries you want to reference.
    

If you are writing code that manipulates objects in another application, you should set a reference to that application's type library for best possible access to those objects. You don't have to set a reference to use another application's objects, but doing so provides several advantages for your application.
Your code will run faster if you set a reference to another application's type library before you work with its objects. If you set a reference, you can declare an [object variable](vbe-glossary.md) representing an object in the other application as its most specific type. For example, if you are writing code to work with Microsoft Excel objects, you can declare an object variable of type **Excel.Application** if you created a reference to the Microsoft Excel type library. The following code is the fastest way to create a variable to represent the Microsoft Excel **Application** object.



```vb
Dim appXL As Excel.Application 

```

If you haven't set a reference to the Microsoft Excel type library, you must declare the [variable](vbe-glossary.md) as a generic variable of type[Object](vbe-glossary.md). The following code runs more slowly.



```vb
Dim appXL As Object 

```

If you set a reference to an application's type library, all of its objects and their [methods](vbe-glossary.md) and[properties](vbe-glossary.md) are listed in the **Object Browser**. This makes it easy to determine what properties and methods are available to each object.
For Microsoft applications that can also serve as Automation servers, you can set references to their type libraries from another application, and control their objects from that application.

