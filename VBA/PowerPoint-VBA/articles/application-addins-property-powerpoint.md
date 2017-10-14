---
title: Application.AddIns Property (PowerPoint)
keywords: vbapp10.chm502019
f1_keywords:
- vbapp10.chm502019
ms.prod: powerpoint
api_name:
- PowerPoint.Application.AddIns
ms.assetid: 5a5a030f-45cd-3b82-f41a-eab53b1ed48f
ms.date: 06/08/2017
---


# Application.AddIns Property (PowerPoint)

Returns the program-specific  **AddIns** collection that represents all the add-ins listed in the **Add-Ins** dialog box (click the **Office** button, click **PowerPoint Options**, click  **Add-Ins**, click  **PowerPoint Add-Ins** on the **Manage** list). Read-only.


## Syntax

 _expression_. **AddIns**

 _expression_ A variable that represents an **Application** object.


## Remarks

Microsoft Office PowerPoint-specific add-ins are identified by a .ppa or .ppam file name extension. Component Object Model (COM) add-ins can be used universally across Microsoft programming products and have a .dll or .exe file name extension.

For information about returning a single member of a collection, see [Returning an Object from a Collection](return-objects-from-collections.md).


## Example

This example adds the add-in named "Myaddin.ppa" to the list in the  **Add-Ins** dialog box and loads the add-in automatically.


```vb
Set myAddIn = Application.AddIns.Add(FileName:="c:\myaddin.ppa")

myAddIn.Loaded = True

MsgBox myAddIn.Name &; " has been added to the list"


```


## See also


#### Concepts


[Application Object](application-object-powerpoint.md)

