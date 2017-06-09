---
title: IconView.SaveOption Property (Outlook)
keywords: vbaol11.chm2569
f1_keywords:
- vbaol11.chm2569
ms.prod: outlook
api_name:
- Outlook.IconView.SaveOption
ms.assetid: 23beb50f-94ab-c3dc-da06-7c8a51355c4a
ms.date: 06/08/2017
---


# IconView.SaveOption Property (Outlook)

Returns an  **[OlViewSaveOption](olviewsaveoption-enumeration-outlook.md)** constant that specifies the folders in which the specified view is available and the read permissions attached to the view. Read-only.


## Syntax

 _expression_ . **SaveOption**

 _expression_ A variable that represents an **IconView** object.


## Remarks

The value of the  **SaveOption** property is set when the **[IconView](iconview-object-outlook.md)** object is created by using the **[Add](views-add-method-outlook.md)** method of the **[Views](views-object-outlook.md)** collection.


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
 
 If objView.SaveOption = _ 
 
 olViewSaveOptionThisFolderEveryone Then 
 
 
 
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


[IconView Object](iconview-object-outlook.md)

