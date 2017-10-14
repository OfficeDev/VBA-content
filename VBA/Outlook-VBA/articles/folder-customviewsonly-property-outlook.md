---
title: Folder.CustomViewsOnly Property (Outlook)
keywords: vbaol11.chm2010
f1_keywords:
- vbaol11.chm2010
ms.prod: outlook
api_name:
- Outlook.Folder.CustomViewsOnly
ms.assetid: b94b60f3-acd8-a968-f333-cb6d4bae8683
ms.date: 06/08/2017
---


# Folder.CustomViewsOnly Property (Outlook)

Returns or sets a  **Boolean** that determines which views are displayed on the **View** menu for a given folder. Read/write.


## Syntax

 _expression_ . **CustomViewsOnly**

 _expression_ A variable that represents a **Folder** object.


## Remarks

If set to the  **True** , only user-created views will appear on the menu.

This property has an effect only on the  **View** menu. It does not affect the display of views in the Navigation Pane.


## Example

The following example prompts the user to select a view option. If the user chooses to view all views, the  **CustomViewsOnly** property is set to **False** . If the user chooses to view only custom views, the **CustomViewsOnly** property is set to **True** . Once the property is changed, the outcome of the change can be seen in the user interface.


```vb
Sub SetCusView() 
 
 'Sets the CustomViewsOnly property depending on the user's response 
 
 Dim nmsName As Outlook.NameSpace 
 
 Dim fldFolder As Outlook.Folder 
 
 Dim lngAns As Long 
 
 
 
 Set nmsName = Application.GetNamespace("MAPI") 
 
 Set fldFolder = nmsName.GetDefaultFolder(olFolderInbox) 
 
 'Prompt user for input 
 
 lngAns = MsgBox( _ 
 
 "Would you like to view only custom views in the View menu?", vbYesNo) 
 
 Call SetVal(fldFolder, lngAns) 
 
End Sub 
 
 
 
Sub SetVal(ByRef fldFolder As Folder, ByVal lngAns As Long) 
 
 'Modifies the CustomViewsOnly property to display views on the View menu 
 
 If lngAns = vbYes Then 
 
 fldFolder.CustomViewsOnly = True 
 
 Else 
 
 fldFolder.CustomViewsOnly = False 
 
 End If 
 
 'Display only custom views 
 
 If lngAns = vbYes Then 
 
 MsgBox "The View menu for the " _ 
 
 &; fldFolder.Name &; " folder will now display only custom views." 
 
 'Display all views 
 
 Else 
 
 MsgBox "The View menu for the " _ 
 
 &; fldFolder.Name &; " folder will now display all views." 
 
 End If 
 
End Sub
```


## See also


#### Concepts


[Folder Object](folder-object-outlook.md)

