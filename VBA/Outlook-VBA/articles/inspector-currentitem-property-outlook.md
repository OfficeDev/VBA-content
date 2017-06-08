---
title: Inspector.CurrentItem Property (Outlook)
keywords: vbaol11.chm2962
f1_keywords:
- vbaol11.chm2962
ms.prod: outlook
api_name:
- Outlook.Inspector.CurrentItem
ms.assetid: eaaf0192-a169-c107-95a6-b8e759a3b873
ms.date: 06/08/2017
---


# Inspector.CurrentItem Property (Outlook)

Returns an  **Object** representing the current item being displayed in the inspector. Read-only.


## Syntax

 _expression_ . **CurrentItem**

 _expression_ A variable that represents an **Inspector** object.


## Remarks

If no item is currently open, an error message will be returned.


## Example

This Visual Basic for Applications (VBA) example uses the  **[CurrentItem](inspector-currentitem-property-outlook.md)** property to obtain the current item that the user is viewing and closes it. If no item is currently open, an error message will be returned.


```vb
Sub CloseItem() 
 
 Dim myItem As Object 
 
 
 
 Set myItem = Application.ActiveInspector.CurrentItem 
 
 myItem.Close olSave 
 
End Sub
```


## See also


#### Concepts


[Inspector Object](inspector-object-outlook.md)

