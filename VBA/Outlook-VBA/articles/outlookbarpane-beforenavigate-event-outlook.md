---
title: OutlookBarPane.BeforeNavigate Event (Outlook)
keywords: vbaol11.chm374
f1_keywords:
- vbaol11.chm374
ms.prod: outlook
api_name:
- Outlook.OutlookBarPane.BeforeNavigate
ms.assetid: f632928b-01a9-b467-1cee-0a86e0023f4d
ms.date: 06/08/2017
---


# OutlookBarPane.BeforeNavigate Event (Outlook)

Occurs when the user clicks a shortcut in the  **Shortcuts** pane to navigate to a different folder.


## Syntax

 _expression_ . **BeforeNavigate**( **_Shortcut_** , **_Cancel_** )

 _expression_ A variable that represents an **[OutlookBarPane](outlookbarpane-object-outlook.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Shortcut_|Required| **[OutlookBarShortcut](outlookbarshortcut-object-outlook.md)**|The shortcut that the user clicked.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the current folder is not changed.|

## Remarks

This event is not available in Microsoft Visual Basic Scripting Edition (VBScript).


## Example

This Microsoft Visual Basic for Applications (VBA) example prevents the user from using the  **Shortcuts** pane to open the **Notes** folder. The sample code must be placed in a class module, and the `Initialize_handler` routine must be called before the event procedure can be called by Microsoft Outlook. If you do not have a shortcut to the **Notes** folder already, you need to create one to run this example.


```vb
Public WithEvents myOlPane As Outlook.OutlookBarPane 
 
 
 
Public Sub Initialize_handler() 
 
 Set myOlPane = Application.ActiveExplorer.Panes.Item("OutlookBar") 
 
End Sub 
 
 
 
Private Sub myOlPane_BeforeNavigate(ByVal Shortcut As Outlook.OutlookBarShortcut, Cancel As Boolean) 
 
 If Shortcut.Name = "Notes" Then 
 
 MsgBox "You cannot view the Notes folder." 
 
 Cancel = True 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[OutlookBarPane Object](outlookbarpane-object-outlook.md)

