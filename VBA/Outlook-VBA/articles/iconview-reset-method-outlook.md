---
title: IconView.Reset Method (Outlook)
keywords: vbaol11.chm2564
f1_keywords:
- vbaol11.chm2564
ms.prod: outlook
api_name:
- Outlook.IconView.Reset
ms.assetid: 0cb48612-db32-19e4-1478-694898775752
ms.date: 06/08/2017
---


# IconView.Reset Method (Outlook)

Resets a built-in Microsoft Outlook view to its original settings.


## Syntax

 _expression_ . **Reset**

 _expression_ A variable that represents an **IconView** object.


## Remarks

This method works only on built-in Outlook views.


## Example

The following Visual Basic for Applications (VBA) example resets all built-in views in the user's  **Inbox** default folder to their original settings. The **[Standard](view-standard-property-outlook.md)** property is returned to determine if the view is a built-in Outlook view.


```vb
Sub ResetInboxViews() 
 
 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As View 
 
 
 
 ' Get the Views collection of the Inbox default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderInbox).Views 
 
 
 
 ' Enumerate the Views collection, calling the Reset 
 
 ' method for each View object with its Standard 
 
 ' property value set to True. 
 
 For Each objView In objViews 
 
 If objView.Standard = True Then 
 
 objView.Reset 
 
 End If 
 
 Next objView 
 
 
 
End Sub
```


## See also


#### Concepts


[IconView Object](iconview-object-outlook.md)

