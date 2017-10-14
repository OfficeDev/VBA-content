---
title: MailItem.Close Method (Outlook)
keywords: vbaol11.chm1320
f1_keywords:
- vbaol11.chm1320
ms.prod: outlook
api_name:
- Outlook.MailItem.Close
ms.assetid: 00a8a4e8-9bdc-d1bc-cb61-c6d925fb754f
ms.date: 06/08/2017
---


# MailItem.Close Method (Outlook)

Closes and optionally saves changes to the Outlook item.


## Syntax

 _expression_ . **Close**( **_SaveMode_** )

 _expression_ A variable that represents a **MailItem** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveMode_|Required| **[OlInspectorClose](olinspectorclose-enumeration-outlook.md)**|The close behavior. If the item displayed within the inspector has not been changed, this argument has no effect.|

## Example

This Visual Basic for Applications (VBA) example saves and closes the item displayed in the active inspector without prompting the user. To run this example, you need to have an item displayed in an inspector window.


```vb
Sub CloseItem() 
 
 Dim myinspector As Outlook.Inspector 
 
 Dim myItem As Outlook.MailItem 
 
 
 
 Set myinspector = Application.ActiveInspector 
 
 Set myItem = myinspector.CurrentItem 
 
 myItem.Close olSave 
 
End Sub
```


## See also


#### Concepts


[MailItem Object](mailitem-object-outlook.md)

