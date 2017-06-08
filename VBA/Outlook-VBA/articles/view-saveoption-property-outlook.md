---
title: View.SaveOption Property (Outlook)
keywords: vbaol11.chm2492
f1_keywords:
- vbaol11.chm2492
ms.prod: outlook
api_name:
- Outlook.View.SaveOption
ms.assetid: d7990708-5eb4-1b11-944e-127793bdb5b1
ms.date: 06/08/2017
---


# View.SaveOption Property (Outlook)

Returns an  **[OlViewSaveOption](olviewsaveoption-enumeration-outlook.md)** constant that specifies the folders in which the specified view is available and the read permissions attached to the view. Read-only.


## Syntax

 _expression_ . **SaveOption**

 _expression_ A variable that represents a **View** object.


## Remarks

The  **SaveOption** property is set when the **[View](view-object-outlook.md)** object is created by using the **[Views.Add](views-add-method-outlook.md)** method.


## Example

The following Microsoft Visual Basic for Applications (VBA) example displays the names of all views that can be accessed by all users in the  **Notes** folder.

The following example locks the user interface for all views that are available to all users. The subroutine  `LockView` accepts the **View** object and a **Boolean** value that indicates if the View interface will be locked. In this example the procedure is always called with the **Boolean** value set to **True** .




```vb
Sub LocksPublicViews() 
 
 'Locks the interface of all views that are available to 
 
 'all users of this folder. 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderNotes).Views 
 
 For Each objView In objViews 
 
 If objView.SaveOption = olViewSaveOptionThisFolderEveryone Then 
 
 Call LockView(objView, True) 
 
 End If 
 
 Next objView 
 
End Sub 
 
 
 
Sub LockView(ByRef objView As View, ByVal blnAns As Boolean) 
 
 'Locks the user interface of the view. 
 
 'Accepts and returns a View object and user response. 
 
 With objView 
 
 If blnAns = True Then 
 
 'if true lock UI 
 
 .LockUserChanges = True 
 
 .Save 
 
 Else 
 
 'if false don't lock UI 
 
 .LockUserChanges = False 
 
 End If 
 
 End With 
 
End Sub
```


## See also


#### Concepts


[View Object](view-object-outlook.md)

