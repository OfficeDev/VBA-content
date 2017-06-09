---
title: View.Reset Method (Outlook)
keywords: vbaol11.chm2487
f1_keywords:
- vbaol11.chm2487
ms.prod: outlook
api_name:
- Outlook.View.Reset
ms.assetid: fb909688-309d-0a70-0b67-0f1793f6a27d
ms.date: 06/08/2017
---


# View.Reset Method (Outlook)

Resets a built-in Microsoft Outlook view to its original settings.


## Syntax

 _expression_ . **Reset**

 _expression_ A variable that represents a **View** object.


## Remarks

This method works only on built-in Outlook views.

To properly reset the current view, you must do a  **View.Reset** and then a **[View.Apply](view-apply-method-outlook.md)** . The code sample below illustrates the order of the calls:




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

The following Microsoft Visual Basic for Applications (VBA) example resets all built-in views in the user's  **Inbox** to their original settings. The **[Standard](view-standard-property-outlook.md)** property is returned to determine if the view is a built-in Outlook view.


```vb
Sub ResetViews() 
 
 'Resets all standard views in the user's Inbox 
 
 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 For Each objView In objViews 
 
 If objView.Standard = True Then 
 
 objView.Reset 
 
 End If 
 
 Next objView 
 
End Sub
```


## See also


#### Concepts


[View Object](view-object-outlook.md)

