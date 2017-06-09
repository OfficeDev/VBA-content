---
title: Inspector.Close Method (Outlook)
keywords: vbaol11.chm2965
f1_keywords:
- vbaol11.chm2965
ms.prod: outlook
api_name:
- Outlook.Inspector.Close
ms.assetid: de821cf4-72f8-ba62-3d8d-96548db0b4a0
ms.date: 06/08/2017
---


# Inspector.Close Method (Outlook)

Closes the  **[Inspector](inspector-object-outlook.md)** and optionally saves changes to the displayed Outlook item.


## Syntax

 _expression_ . **Close**( **_SaveMode_** )

 _expression_ A variable that represents an **Inspector** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _SaveMode_|Required| **[OlInspectorClose](olinspectorclose-enumeration-outlook.md)**|The close behavior. If the item displayed within the inspector has not been changed, this argument has no effect.|

## Remarks


 **Note**  Do not use this method from within the [Inspector.Activate Event (Outlook)](inspector-activate-event-outlook.md) event handler.


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


[Inspector Object](inspector-object-outlook.md)

