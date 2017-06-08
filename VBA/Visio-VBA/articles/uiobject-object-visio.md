---
title: UIObject Object (Visio)
keywords: vis_sdr.chm10300
f1_keywords:
- vis_sdr.chm10300
ms.prod: visio
api_name:
- Visio.UIObject
ms.assetid: 2d842398-df53-0d59-6ee5-89d411440863
ms.date: 06/08/2017
---


# UIObject Object (Visio)

Represents a set of Microsoft Visio menus, toolbars, and accelerators, from either the built-in Visio user interface or a customized version of it. 


## Remarks


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

To retrieve a  **UIObject** object that contains




- Visio menus and accelerators, use the  **BuiltInMenus** property of an **Application** object and then the **MenuSets** or **AccelTables** collections of the **UIObject** object returned from the **BuiltInMenus** property.
    
- Visio toolbars, use the  **BuiltInToolbars** property of an **Application** object and then the **ToolbarSets** collection of the **UIObject** object returned from the **BuiltInToolbars** property.
    


If an  **Application** object or **Document** object has a customized user interface, use the **CustomMenus** or **CustomToolbars** properties to retrieve **UIObject** objects that represent these.

A  **UIObject** object can be stored in a file and loaded into Visio. Use the **SaveToFile** method to save the object and the **LoadFromFile** method to load it, or set the **CustomMenusFile** or **CustomToolbarsFile** property of an **Application** object or **Document** object to the name of the stored user interface file.

Beginning with Visio 2002, a program can manipulate menus and toolbars in the Visio user interface by manipulating the  **CommandBars** collection returned by the **CommandBars** property. The **CommandBars** collection has an interface identical to the **CommandBars** collection exposed by the suite of Microsoft System applications such as Microsoft Word and Microsoft Excel. Consequently, programs can manipulate the Visio menus and toolbars by using either the **CommandBars** collection or **UIObject** objects.


