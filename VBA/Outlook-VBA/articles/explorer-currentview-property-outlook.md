---
title: Explorer.CurrentView Property (Outlook)
keywords: vbaol11.chm2766
f1_keywords:
- vbaol11.chm2766
ms.prod: outlook
api_name:
- Outlook.Explorer.CurrentView
ms.assetid: 177e6387-9ccb-cb71-bbe5-332c25485848
ms.date: 06/08/2017
---


# Explorer.CurrentView Property (Outlook)

Returns or sets a  **Variant** representing the current view. Read/write.


## Syntax

 _expression_ . **CurrentView**

 _expression_ A variable that represents an **Explorer** object.


## Remarks

To obtain a  **[View](view-object-outlook.md)** object for the view of the current **[Explorer](explorer-object-outlook.md)** , use **Explorer.CurrentView** instead of the **[CurrentView](folder-currentview-property-outlook.md)** property of the current **[Folder](folder-object-outlook.md)** object returned by **[Explorer.CurrentFolder](explorer-currentfolder-property-outlook.md)** .

You must save a reference to the  **View** object returned by **CurrentView** before you proceed to use it for any purpose.

To properly reset the current view, you must do a  **[View.Reset](view-reset-method-outlook.md)** and then a **[View.Apply](view-apply-method-outlook.md)** . The code sample below illustrates the order of the calls:




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

When this property is set, two events occur:  **[BeforeViewSwitch](explorer-beforeviewswitch-event-outlook.md)** occurs before the actual view change takes place and can be used to cancel the change and **[ViewSwitch](explorer-viewswitch-event-outlook.md)** takes place after the change is effective.


## Example

The following Visual Basic for Applications (VBA) example sets the current view in the active explorer to messages if the  **Inbox** is displayed.


```vb
Sub ChangeCurrentView() 
 
 Dim myOlExp As Outlook.Explorer 
 
 
 
 Set myOlExp = Application.ActiveExplorer 
 
 If myOlExp.CurrentFolder = "Inbox" Then 
 
 myOlExp.CurrentView = "Messages" 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Explorer Object](explorer-object-outlook.md)

