---
title: View.LockUserChanges Property (Outlook)
keywords: vbaol11.chm2490
f1_keywords:
- vbaol11.chm2490
ms.prod: outlook
api_name:
- Outlook.View.LockUserChanges
ms.assetid: f4347b6f-b00d-6508-09e3-35cf98da26b1
ms.date: 06/08/2017
---


# View.LockUserChanges Property (Outlook)

Returns or sets a  **Boolean** value that indicates whether a user can modify the settings of the current view. Read/write.


## Syntax

 _expression_ . **LockUserChanges**

 _expression_ A variable that represents a **View** object.


## Remarks

If  **True** , the user can modify the settings of the current view. However, changes made to the interface will not be saved. If **False** (the default), any changes will be saved.


## Example

The following example locks the user interface for all views that are available to all users. The subroutine  `LockView` accepts the **[View](view-object-outlook.md)** object and a **Boolean** value that indicates if the **View** interface will be locked. In this example the procedure is always called with the **Boolean** value set to **True** .


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

