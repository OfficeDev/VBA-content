---
title: DistListItem.DLName Property (Outlook)
keywords: vbaol11.chm1148
f1_keywords:
- vbaol11.chm1148
ms.prod: outlook
api_name:
- Outlook.DistListItem.DLName
ms.assetid: 38d027b7-89f9-1659-84e0-35473b07c088
ms.date: 06/08/2017
---


# DistListItem.DLName Property (Outlook)

Returns or sets a  **String** representing the display name of a distribution list. Read/write.


## Syntax

 _expression_ . **DLName**

 _expression_ A variable that represents a **DistListItem** object.


## Example

This Microsoft Visual Basic for Applications (VBA) example creates a new distribution list and then prompts the user for a name.


```vb
Sub CreateDL() 
 
 Dim myDistList As Outlook.DistListItem 
 
 
 
 Set myDistList = Application.CreateItem(olDistributionListItem) 
 
 myDistList.DLName = InputBox("Type the name of the new distribution list.") 
 
 myDistList.Save 
 
 myDistList.Display 
 
End Sub
```


## See also


#### Concepts


[DistListItem Object](distlistitem-object-outlook.md)

