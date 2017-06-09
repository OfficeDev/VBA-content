---
title: OutlookBarPane.Contents Property (Outlook)
keywords: vbaol11.chm363
f1_keywords:
- vbaol11.chm363
ms.prod: outlook
api_name:
- Outlook.OutlookBarPane.Contents
ms.assetid: ec7b8c50-7bf5-50d5-6c0b-32091106350e
ms.date: 06/08/2017
---


# OutlookBarPane.Contents Property (Outlook)

Returns the  **[OutlookBarStorage](outlookbarstorage-object-outlook.md)** object for the specified Outlook Bar pane. Read-only.


## Syntax

 _expression_ . **Contents**

 _expression_ A variable that represents an **OutlookBarPane** object.


## Example

This Microsoft Visual Basic for Applications example displays a message listing the groups in the Outlook Bar.


```vb
Sub ListGroups() 
 
 Dim myOlBar As Outlook.OutlookBarPane 
 
 Dim myOlGroups As Outlook.OutlookBarGroups 
 
 
 
 myMsg = "The groups in the Outlook Bar are:" 
 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 
 Set myOlGroups = myOlBar.Contents.Groups 
 
 For x = 1 To myOlGroups.Count 
 
 myMsg = myMsg &; Chr(13) &; myOlGroups.Item(x) 
 
 Next x 
 
 MsgBox myMsg 
 
End Sub
```


## See also


#### Concepts


[OutlookBarPane Object](outlookbarpane-object-outlook.md)

