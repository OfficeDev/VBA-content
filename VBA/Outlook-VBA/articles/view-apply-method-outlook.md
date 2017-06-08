---
title: View.Apply Method (Outlook)
keywords: vbaol11.chm2484
f1_keywords:
- vbaol11.chm2484
ms.prod: outlook
api_name:
- Outlook.View.Apply
ms.assetid: b121d1ce-24b7-4ace-8369-42e5c7becd0a
ms.date: 06/08/2017
---


# View.Apply Method (Outlook)

Applies the view to the Microsoft Outlook user interface.


## Syntax

 _expression_ . **Apply**

 _expression_ A variable that represents a **View** object.


## Remarks

To properly reset the current view, you must do a  **[View.Reset](view-reset-method-outlook.md)** and then a **View.Apply** . The code sample below illustrates the order of the calls:


```vb
Sub ResetView() 
 
 Dim v as Outlook.View 
 
 ' Save a reference to the current view object 
 
 Set v = Application.ActiveExplorer.CurrentView 
 
 ' Reset and then apply the current view 
 
 v.Reset 
 
 v.Apply 
 
End Sub
```


## Example

The following Visual Basic for Applications (VBA) example creates a new view called  **New Table** and applies it.


```vb
Sub CreateView() 
 
 'Creates a new view 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objViews As Outlook.Views 
 
 Dim objNewView As Outlook.View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 Set objNewView = objViews.Add(Name:="New Table", _ 
 
 ViewType:=olTableView) 
 
 objNewView.Save 
 
 objNewView.Apply 
 
End Sub
```


## See also


#### Concepts


[View Object](view-object-outlook.md)

