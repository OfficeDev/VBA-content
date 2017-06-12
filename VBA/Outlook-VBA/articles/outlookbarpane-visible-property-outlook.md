---
title: OutlookBarPane.Visible Property (Outlook)
keywords: vbaol11.chm366
f1_keywords:
- vbaol11.chm366
ms.prod: outlook
api_name:
- Outlook.OutlookBarPane.Visible
ms.assetid: d9d00e7a-52ef-b481-7a56-729e1ac04534
ms.date: 06/08/2017
---


# OutlookBarPane.Visible Property (Outlook)

Returns or sets a  **Boolean** indicating the visible state of the specified object. Read/write.


## Syntax

 _expression_ . **Visible**

 _expression_ A variable that represents an **OutlookBarPane** object.


## Remarks

 **True** to display the object; **False** to hide the object.

You can also use the  **[ShowPane](explorer-showpane-method-outlook.md)** method or the **[IsPaneVisible](explorer-ispanevisible-method-outlook.md)** method of an **[Explorer](explorer-object-outlook.md)** object to set or retrieve this value.


## Example

This Microsoft Visual Basic for Applications (VBA) example toggles the visible state of the Shortcuts pane.


```vb
Sub ShowHideShortcutsBar() 
 
 Dim myOlBar As Outlook.OutlookBarPane 
 
 
 
 Set myOlBar = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 
 myOlBar.Visible = Not myOlBar.Visible 
 
End Sub
```


## See also


#### Concepts


[OutlookBarPane Object](outlookbarpane-object-outlook.md)

