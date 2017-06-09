---
title: CardView.LockUserChanges Property (Outlook)
keywords: vbaol11.chm2589
f1_keywords:
- vbaol11.chm2589
ms.prod: outlook
api_name:
- Outlook.CardView.LockUserChanges
ms.assetid: a9fea66c-36b3-6c72-8aee-dc77ca8ae215
ms.date: 06/08/2017
---


# CardView.LockUserChanges Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether a user can modify the settings of the view. Read/write.


## Syntax

 _expression_ . **LockUserChanges**

 _expression_ A variable that represents a **CardView** object.


## Remarks

If  **True** , the user can modify the settings of the specified view. However, changes made to the interface will not be saved. If **False** (the default), any changes will be saved.


## Example

The following Visual Basic for Applications (VBA) example locks the user interface for all views that are available to all users. The subroutine  `LockView` accepts the **[View](view-object-outlook.md)** object and a **Boolean** value that indicates if the **View** user interface will be locked. In this example, the procedure is always called with the **Boolean** value set to **True** .


```vb
Sub LockPublicViews() 
 
 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As View 
 
 
 
 ' Get the Views collection for the Contacts default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderContacts).Views 
 
 
 
 ' Enumerate the Views collection and lock the user 
 
 ' interface for any view that can be accessed by 
 
 ' all users who have access to the Notes default folder. 
 
 For Each objView In objViews 
 
 If objView.SaveOption = olViewSaveOptionThisFolderEveryone Then 
 
 Call LockView(objView, True) 
 
 End If 
 
 Next objView 
 
 
 
End Sub 
 
 
 
Sub LockView(ByRef objView As View, ByVal blnAns As Boolean) 
 
 
 
 ' Examine the view object. 
 
 With objView 
 
 If blnAns = True Then 
 
 ' Lock the user interface and 
 
 ' save the view 
 
 .LockUserChanges = True 
 
 .Save 
 
 Else 
 
 ' Unlock the user interface of the view. 
 
 .LockUserChanges = False 
 
 End If 
 
 End With 
 
 
 
End Sub
```


## See also


#### Concepts


[CardView Object](cardview-object-outlook.md)

