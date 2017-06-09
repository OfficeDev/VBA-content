---
title: InvisibleApp Object (Visio)
keywords: vis_sdr.chm60185
f1_keywords:
- vis_sdr.chm60185
ms.prod: visio
api_name:
- Visio.InvisibleApp
ms.assetid: 70a30571-2017-af8b-eaa1-bf93c758a46a
ms.date: 06/08/2017
---


# InvisibleApp Object (Visio)

Represents an invisible instance of Microsoft Visio. The properties of the  **InvisibleApp** object are identical to those of the **Application** object.


## Remarks

You can use the  **InvisibleApp** object when you want to take advantage of Automation in Visio without end-user interaction or knowledge.

An external program typically creates or retrieves an  **Application** or **InvisibleApp** object before it can retrieve other Visio objects from that instance. Use the Microsoft Visual Basic **CreateObject** function or the **New** keyword to run a new instance. Set the value of the **InvisibleApp** object's **Visible** property to **True** to show it.


 **Note**  You cannot use the Visual Basic  **GetObject** function to retrieve an **InvisibleApp** object for an instance of Visio that is already running. Attempts to do so will fail.

Use the  **Documents** , **Windows** , and **Addons** properties of an **InvisibleApp** object to retrieve the **Documents** , **Windows** , and **Addons** collections of the instance respectively.

Use the  **ActiveDocument** , **ActivePage** , or **ActiveWindow** property to retrieve the currently active **Document** , **Page** , or **Window** object.


 **Note**  Starting with Visio, the Microsoft Office Fluent user interface (UI) replaces the previous system of layered menus, toolbars, and task panes. VBA objects and members that you used to customize the user interface in previous versions of Visio are still available in Visio, but they function differently.

 **ActiveDocument** is the default property of an **InvisibleApp** object.


 **Note**  Code in the Microsoft Visual Basic for Applications (VBA) project of a Visio document can use the Visio global object instead of a Visio  **Invisible App** object to retrieve other objects.


