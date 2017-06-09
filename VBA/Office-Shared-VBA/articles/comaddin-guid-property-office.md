---
title: COMAddIn.Guid Property (Office)
keywords: vbaof11.chm219004
f1_keywords:
- vbaof11.chm219004
ms.prod: office
api_name:
- Office.COMAddIn.Guid
ms.assetid: 1e3218d9-dce7-21e2-55a7-4435ca58bb35
ms.date: 06/08/2017
---


# COMAddIn.Guid Property (Office)

Gets the class identifier (CLSID) for the specified  **COMAddIn** object. Read-only.


## Syntax

 _expression_. **Guid**

 _expression_ A variable that represents a **COMAddIn** object.


## Example

The following example displays the ProgID and CLSID for the first COM add-in in the  **COMAddIns** collection in a message box.


```
MsgBox "My ProgID is " &amp; _ 
 Application.COMAddIns(1).ProgID &amp; _ 
 " and my CLSID is " &amp; _ 
 Application.COMAddIns(1).Guid
```


## See also


#### Concepts


[COMAddIn Object](comaddin-object-office.md)
#### Other resources


[COMAddIn Object Members](comaddin-members-office.md)

