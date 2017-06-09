---
title: CalendarView.Standard Property (Outlook)
keywords: vbaol11.chm2620
f1_keywords:
- vbaol11.chm2620
ms.prod: outlook
api_name:
- Outlook.CalendarView.Standard
ms.assetid: 7d4ac52a-8a3d-25b1-6900-3799fe0fde70
ms.date: 06/08/2017
---


# CalendarView.Standard Property (Outlook)

Returns a  **Boolean** value that indicates whether the **[CalendarView](calendarview-object-outlook.md)** object is a built-in Outlook view. Read-only.


## Syntax

 _expression_ . **Standard**

 _expression_ An expression that returns a **CalendarView** object.


## Remarks

The  **[Reset](view-reset-method-outlook.md)** method can only be used on a view if the value of this property is set to **True** .


## Example

The following Visual Basic for Applications (VBA) example enumerates through the  **[Views](views-object-outlook.md)** collection of the current **[Folder](folder-object-outlook.md)** object using the **Standard** property to determine if a **View** object is a built-in Outlook view. If the **View** object is a built-in Outlook view, the sample calls the **Reset** method to reset the view to its default settings. Otherwise, the sample uses the **[Delete](view-delete-method-outlook.md)** method to delete the view.


```vb
Private Sub RemoveAllViewCustomization() 
 
 Dim objView As View 
 
 
 
 ' Enumerate each View object in the Views collection 
 
 ' of the current Folder object. 
 
 For Each objView In Application.ActiveExplorer.CurrentFolder.Views 
 
 ' If the View object is a built-in Outlook view, reset 
 
 ' the view to its default settings. If the View object 
 
 ' is a custom view, delete it. 
 
 If objView.Standard Then 
 
 objView.Reset 
 
 Else 
 
 objView.Delete 
 
 End If 
 
 Next 
 
End Sub
```


## See also


#### Concepts


[CalendarView Object](calendarview-object-outlook.md)

