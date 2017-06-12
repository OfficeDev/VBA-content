---
title: OutlookBarShortcut.SetIcon Method (Outlook)
keywords: vbaol11.chm344
f1_keywords:
- vbaol11.chm344
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcut.SetIcon
ms.assetid: d54a60b5-e667-e030-0724-d61be3ad3745
ms.date: 06/08/2017
---


# OutlookBarShortcut.SetIcon Method (Outlook)

Sets the icon for the specified shortcut on the  **Shortcuts** pane.


## Syntax

 _expression_ . **SetIcon**( **_Icon_** )

 _expression_ A variable that represents an **OutlookBarShortcut** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Icon_|Required| **Variant**|The path of the icon.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a group called MicrosoftSites and adds a shortcut to the Microsoft Network Web page. Then it sets the icon of the shortcut to the icon image MSN.ico located on the user's computer. The example assumes that this icon exists in the specified location.


```vb
 Sub CreateMSNShortcutWithIcon() 
 
 Dim exp As Outlook.Explorer 
 
 Dim pans As Outlook.Panes 
 
 Dim bpan As Outlook.OutlookBarPane 
 
 Dim bgrps As Outlook.OutlookBarGroups 
 
 Dim bgrp As Outlook.OutlookBarGroup 
 
 Dim bscs As Outlook.OutlookBarShortcuts 
 
 Dim bsc As Outlook.OutlookBarShortcut 
 
 Dim bsc2 As Outlook.OutlookBarShortcut 
 
 
 
 Set exp = Application.ActiveExplorer 
 
 Set pans = exp.Panes 
 
 Set bpan = pans.Item("OutlookBar") 
 
 Set bgrps = bpan.Contents.Groups 
 
 Set bgrp = bgrps.Add("MicrosoftSites") 
 
 Set bscs = bgrp.Shortcuts 
 
 Set bsc = bscs.Add("http://www.msn.com", "MSN Home Page") 
 
 bsc.SetIcon "C:\MSN.ico" 
 
End Sub
```


## See also


#### Concepts


[OutlookBarShortcut Object](outlookbarshortcut-object-outlook.md)

