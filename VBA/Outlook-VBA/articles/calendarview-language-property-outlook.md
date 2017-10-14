---
title: CalendarView.Language Property (Outlook)
keywords: vbaol11.chm2616
f1_keywords:
- vbaol11.chm2616
ms.prod: outlook
api_name:
- Outlook.CalendarView.Language
ms.assetid: e8d1a39b-c0f7-bd62-5831-d4ac02a0f2ee
ms.date: 06/08/2017
---


# CalendarView.Language Property (Outlook)

Returns or sets a  **String** value that represents the language setting for the view. Read/write.


## Syntax

 _expression_ . **Language**

 _expression_ A variable that represents a **CalendarView** object.


## Remarks

The  **Language** property uses a **String** to represent an ISO language tag. For example, the string "EN-US" represents the ISO code for "United States - English."

If a valid language code is specified, the object will only be available in the  **View** menu for the specified language type. If no value is specified, the object item is available for all language types. The default value for this property is an empty string.


## Example

The following Microsoft Visual Basic for Applications (VBA) example sets the language type of all  **[View](view-object-outlook.md)** objects of type **olTableView** to U.S. English.


```vb
Sub SetLanguage() 
 
 'Sets the language of all table views to U.S. English. 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 
 
 Set objViews = _ 
 
 Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Views 
 
 'Iterate through each view in the collection. 
 
 For Each objView In objViews 
 
 Debug.Print objView.Name 
 
 'If view is of type olTableVIew then set language. 
 
 If objView.ViewType = olTableView And objView.Standard = False Then 
 
 objView.Language = "EN-US" 
 
 End If 
 
 Next objView 
 
End Sub
```


## See also


#### Concepts


[CalendarView Object](calendarview-object-outlook.md)

