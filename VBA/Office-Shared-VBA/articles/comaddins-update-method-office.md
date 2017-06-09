---
title: COMAddIns.Update Method (Office)
keywords: vbaof11.chm220004
f1_keywords:
- vbaof11.chm220004
ms.prod: office
api_name:
- Office.COMAddIns.Update
ms.assetid: 4cbaff64-10e8-d792-60b5-29f6de97dc8f
ms.date: 06/08/2017
---


# COMAddIns.Update Method (Office)

Updates the contents of the COMAddIns collection from the list of add-ins stored in the Windows registry.


## Syntax

 _expression_. **Update**

 _expression_ A variable that represents a **COMAddIns** object.


## Remarks

Before you can use a given COM add-in in a Microsoft Office application, that add-in must be registered in the Windows registry as a COM component with a corresponding Component Category ID. Normally the setup program for a COM add-in will add the necessary entries to the registry.


## Example

The following example updates the contents of the COMAddIns collection from the list of add-ins stored in the Windows registry.


```
Application.COMAddIns.Update
```


## See also


#### Concepts


[COMAddIns Object](comaddins-object-office.md)
#### Other resources


[COMAddIns Object Members](comaddins-members-office.md)

