---
title: OutlookBarShortcuts.Add Method (Outlook)
keywords: vbaol11.chm335
f1_keywords:
- vbaol11.chm335
ms.prod: outlook
api_name:
- Outlook.OutlookBarShortcuts.Add
ms.assetid: 801d1a9e-f2b6-cbcd-8181-003eba1025b2
ms.date: 06/08/2017
---


# OutlookBarShortcuts.Add Method (Outlook)

Adds a new shortcut to a group in the  **Shortcuts** pane.


## Syntax

 _expression_ . **Add**( **_Target_** , **_Name_** , **_Index_** )

 _expression_ A variable that represents an **OutlookBarShortcuts** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Target_|Required| **Variant**|The target of the shortcut being created.|
| _Name_|Required| **String**|The name of the shortcut being created.|
| _Index_|Optional| **Long**|The position at which the new shortcut will be inserted in the  **Shortcuts** pane group. Position one is at the top of the group.The **Target** type depends on the shortcut type. If the type is **Folder** , the shortcut represents a Microsoft Outlook folder. If the type is a **String** , the shortcut represents a file-system path or a URL.|

### Return Value

An  **[OutlookBarShortcut](outlookbarshortcut-object-outlook.md)** object that represents the new shortcut.


## Example

The following Microsoft Visual Basic for Applications example adds a shortcut to the Microsoft home page on the Web.


```vb
Sub AddShortcut() 
 Dim myOlBar As Outlook.OutlookBarPane 
 Dim myolGroup As Outlook.OutlookBarGroup 
 Dim myOlShortcuts As Outlook.OutlookBarShortcuts 
 
 Set myOlBar = Application.ActiveExplorer.panes.Item("OutlookBar") 
 Set myolGroup = myOlBar.Contents.Groups.Item(1) 
 Set myOlShortcuts = myolGroup.Shortcuts 
 myOlShortcuts.Add "http://www.microsoft.com", _ 
 "Microsoft Home Page", 1 
End Sub
```


## See also


#### Concepts


[OutlookBarShortcuts Object](outlookbarshortcuts-object-outlook.md)

