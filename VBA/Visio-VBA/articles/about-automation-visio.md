---
title: About Automation (Visio)
keywords: vis_sdr.chm0
f1_keywords:
- vis_sdr.chm0
ms.prod: visio
ms.assetid: d34dd6a0-7f11-d8ce-65d2-2a9654cdb06d
ms.date: 06/08/2017
---


# About Automation (Visio)

You can write programs to control Visio in Visual Basic for Applications (VBA), Visual Basic, Visual C++, any of the Visual Studio .NET languages, or in any programming language that supports Automation.

A program can use Automation to incorporate Visio drawing and diagramming capabilities or to automate simple repetitive tasks in Visio. For example, a program might generate an organization chart from a list of names and positions or print all of the masters on a stencil.

## How a program uses automation to control Visio

A program controls Visio by accessing its objects and then using their properties, methods, and events.


-  _Objects_ represent items you work with in the Visio application, such as documents, drawing pages, shapes, and cells containing formulas.
    
-  _Properties_ are attributes that determine the appearance or behavior of objects. For example, a **Shape** object has a **Name** property, which represents the name of that shape.
    
-  _Methods_ are actions provided by an object. For instance, a program can perform the **Add** method on a **Page** object. This is the same as adding a page to a document by clicking **Blank Page** on the **Insert** tab.
    
-  _Events_ trigger code or entire programs. For example, an event can programmatically trigger code when a document is opened or trigger a program when a shape is double-clicked.
    

## The VBA programming environment in Visio

 Visio includes the Visual Basic for Applications (VBA) programming environment. To create, view, debug, and run programs in this environment, use the Visual Basic Editor:


- Create VBA programs by inserting modules, class modules, and user forms into your VBA project and by writing code.
    
- View VBA project items by choosing the project of an open Visio document in the  **Project Explorer**. To view the  **Code** window for individual items, open the appropriate folder in the **Project Explorer** and double-click the project item, or right-click the item and click **View Code** on the shortcut menu.
    
- Debug VBA programs by adding breakpoints, including watch expressions, and stepping through code as it runs.
    
- Run VBA macros in the following ways:
    
In the  _Visual Basic Editor:_ On the **Run** menu, click **Run Macro.**

In  _Visio:_ In the **Code** group on the [Developer](run-visio-in-developer-mode.md) tab, click **Macros**.


