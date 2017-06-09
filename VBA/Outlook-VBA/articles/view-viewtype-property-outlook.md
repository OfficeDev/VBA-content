---
title: View.ViewType Property (Outlook)
keywords: vbaol11.chm2494
f1_keywords:
- vbaol11.chm2494
ms.prod: outlook
api_name:
- Outlook.View.ViewType
ms.assetid: db44b9ec-cb55-c9f4-d621-32d2f46598dd
ms.date: 06/08/2017
---


# View.ViewType Property (Outlook)

Returns an  **[OlViewType](olviewtype-enumeration-outlook.md)** constant representing the view type of a **[View](view-object-outlook.md)** object. Read-only.


## Syntax

 _expression_ . **ViewType**

 _expression_ An expression that returns a **View** object.


## Remarks

This property does not have any effect on the icons displayed in the Shortcuts pane. Large icons have been removed and if this property is set to  **olLargeIcon** , it will not have any effect.


## Example

The following Visual Basic for Applicatons (VBA) example displays the name and type of all views in the user's  **Inbox**.


```vb
Sub DisplayViewMode() 
 
 'Displays the names and view modes for all views 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 Dim strTypes As String 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 'Collect names and view types for all views 
 
 For Each objView In objViews 
 
 strTypes = strTypes &; objView.Name &; vbTab &; vbTab &; objView.ViewType &; vbCr 
 
 Next objView 
 
 'Display message box 
 
 MsgBox "Current Inbox Views and Viewtypes:" &; vbCr &; _ 
 
 vbCr &; strTypes 
 
End Sub
```


## See also


#### Concepts


[View Object](view-object-outlook.md)

