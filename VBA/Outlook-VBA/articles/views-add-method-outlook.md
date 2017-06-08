---
title: Views.Add Method (Outlook)
keywords: vbaol11.chm547
f1_keywords:
- vbaol11.chm547
ms.prod: outlook
api_name:
- Outlook.Views.Add
ms.assetid: 8005ca2e-8b28-1286-74d1-448f2a168c65
ms.date: 06/08/2017
---


# Views.Add Method (Outlook)

Creates a new view in the  **[Views](views-object-outlook.md)** collection.


## Syntax

 _expression_ . **Add**( **_Name_** , **_ViewType_** , **_SaveOption_** )

 _expression_ A variable that represents a **Views** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|The name of the new view.|
| _ViewType_|Required| **[OlViewType](olviewtype-enumeration-outlook.md)**|The type of the new view.|
| _SaveOption_|Optional| **[OlViewSaveOption](olviewsaveoption-enumeration-outlook.md)**|The save option that specifies the permissions of the new view. 
<ul xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:MSHelp="http://msdn.microsoft.com/mshelp" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><li><p><b>olViewSaveOptionAllFoldersOfType</b>  The view can be accessed in all folders of this type.</p></li><li><p><b>olViewSaveOptionThisFolderEveryOne</b>  The view can be accessed by all users in this folder only.</p></li><li><p><b>olViewSaveOptionThisFolderOnlyMe</b>  The view can be accessed in this folder only by the user.</p></li></ul>|

### Return Value

A  **[View](view-object-outlook.md)** object that represents the new view.


## Remarks

If you add a  **View** to a **Views** collection of a folder that is not the current folder, you must first save a copy of the **Views** collection object and then add the **View** to this collection object, as shown in the code sample below. This is a work-around for an existing problem which will otherwise cause a call to **[View.Apply](view-apply-method-outlook.md)** for the added **View** to fail.


```vb
Sub CalendarView() 
 Dim calView As Outlook.View 
 Dim vws As Views 
 
 Set Application.ActiveExplorer.CurrentFolder = Application.Session.GetDefaultFolder(olFolderInbox) 
 ' Current folder is Inbox; add a View to the Calendar folder which is not the current folder 
 ' Keep a copy of the object for the Views collection for the Calendar 
 Set vws = Application.Session.GetDefaultFolder(olFolderCalendar).Views 
 ' Add the View to this Views collection object 
 Set calView = vws.Add("New Calendar", olCalendarView, olViewSaveOptionThisFolderEveryone) 
 calView.Save 
 ' This Apply call will be fine 
 calView.Apply 
End Sub
```


## Example

The following Visual Basic for Applications (VBA) example creates a new view called New Table and stores it in a variable called  `objNewView`.


```vb
Sub CreateView() 
 'Creates a new view 
 Dim objName As Outlook.NameSpace 
 Dim objViews As Outlook.Views 
 Dim objNewView As Outlook.View 
 
 Set objName = Application.GetNamespace("MAPI") 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 Set objNewView = objViews.Add(Name:="New Table", _ 
 ViewType:=olTableView, SaveOption:=olViewSaveOptionThisFolderEveryone) 
End Sub
```


## See also


#### Concepts


[Views Object](views-object-outlook.md)

