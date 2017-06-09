---
title: COMAddIn.ProgId Property (Office)
keywords: vbaof11.chm219003
f1_keywords:
- vbaof11.chm219003
ms.prod: office
api_name:
- Office.COMAddIn.ProgId
ms.assetid: eb917d53-512e-35dd-ff70-ac7b976e6500
ms.date: 06/08/2017
---


# COMAddIn.ProgId Property (Office)

Gets the programmatic identifier (ProgID) for the specified  **COMAddIn** object. Read-only.


## Syntax

 _expression_. **ProgId**

 _expression_ A variable that represents a **COMAddIn** object.


## Example

The following example displays the ProgID and GUID for COM add-in one in a message box.


```
MsgBox "My ProgID is " &amp; _ 
 Application.COMAddIns(1).ProgID &amp; _ 
 " and my GUID is " &amp; _ 
 Application.COMAddIns(1).Guid
```


## See also


#### Concepts


[COMAddIn Object](comaddin-object-office.md)
#### Other resources


[COMAddIn Object Members](comaddin-members-office.md)

