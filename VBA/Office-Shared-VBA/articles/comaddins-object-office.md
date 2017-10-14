---
title: COMAddIns Object (Office)
keywords: vbaof11.chm220000
f1_keywords:
- vbaof11.chm220000
ms.prod: office
api_name:
- Office.COMAddIns
ms.assetid: f6efa1cc-8d30-27d5-8b07-7ddad22f16ef
ms.date: 06/08/2017
---


# COMAddIns Object (Office)

A collection of  **COMAddIn** objects that provide information about a COM add-in registered in the Windows registry.


## Example

Use the  **COMAddIns** property of the **Application** object to return the **COMAddIns** collection for a Microsoft Office host application. This collection contains all of the COM add-ins that are available to a given Office host application, and the **Count** property of the **COMAddins** collection returns the number of available COM add-ins, as in the following example.


```
MsgBox Application.COMAddIns.Count
```

Use the  **Update** method of the **COMAddins** collection to refresh the list of COM add-ins from the Windows registry, as in the following example.




```
Application.COMAddIns.Update
```

Use  **COMAddIns.Item(index)**, where _index_ is either an ordinal value that returns the COM add-in at that position in the **COMAddIns** collection, or a **String** value that represents the ProgID of the specified COM add-in. The following example displays a COM add-in's description text and ProgID (" **msodraa9.ShapeSelect** ") in a message box.




```
MsgBox Application.COMAddIns.Item("msodraa9.ShapeSelect").Description
```


## See also


#### Concepts


[Object Model Reference](reference-object-library-reference-for-office.md)
#### Other resources


[COMAddIns Object Members](comaddins-members-office.md)

