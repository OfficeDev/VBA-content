---
title: Views.ViewRemove Event (Outlook)
keywords: vbaol11.chm552
f1_keywords:
- vbaol11.chm552
ms.prod: outlook
api_name:
- Outlook.ViewRemove
ms.assetid: a0d405fd-aa57-c333-8e33-aa482019d9c8
ms.date: 06/08/2017
---


# Views.ViewRemove Event (Outlook)

Occurs when a view has been removed from the specified collection.


## Syntax

 _expression_ . **ViewRemove**( **_View_** )

 _expression_ A variable that represents a **Views** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _View_|Required| **[View](view-object-outlook.md)**|The view which was removed from the collection prior to this event.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the name of the view that has been removed from the collection when the  **ViewRemove** event is fired. The sample code must be placed in a class module such as `ThisOutlookSession`, and the  `DeleteView()` procedure should be called before the event procedure can be called by Microsoft Outlook.


```vb
Public WithEvents objViews As Outlook.Views 
 
Sub DeleteView() 
 Set objViews = Application.ActiveExplorer.CurrentFolder.Views 
 objViews.Item("New Table View").Delete 
End Sub 
 
Sub objViews_ViewRemove(ByVal View As View) 
 'Displays view name 
 MsgBox "The view: " &; View.Name &; " was removed programmatically." 
End Sub
```


