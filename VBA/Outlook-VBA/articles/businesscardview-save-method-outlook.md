---
title: BusinessCardView.Save Method (Outlook)
keywords: vbaol11.chm2925
f1_keywords:
- vbaol11.chm2925
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.Save
ms.assetid: 9d3d85b7-4ed1-fea3-abb1-7506a0851b50
ms.date: 06/08/2017
---


# BusinessCardView.Save Method (Outlook)

Saves the view, or saves the changes to a view.


## Syntax

 _expression_ . **Save**

 _expression_ An expression that returns a **BusinessCardView** object.


## Example

The following Visual Basic for Applications (VBA) example creates, saves, and applies a new  **BusinessCardView** object.


```vb
Sub CreateBusinessCardView() 
 
 
 
 Dim objName As NameSpace 
 
 Dim objViews As Views 
 
 Dim objView As BusinessCardView 
 
 
 
 ' Get the Views collection of the Contacts default folder. 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderContacts).Views 
 
 
 
 ' Create the new view. 
 
 Set objView = objViews.Add( _ 
 
 "Card View", _ 
 
 olBusinessCardView, _ 
 
 olViewSaveOptionAllFoldersOfType) 
 
 
 
 ' Save and apply the new view. 
 
 objView.Save 
 
 objView.Apply 
 
 
 
End Sub
```


## See also


#### Concepts


[BusinessCardView Object](businesscardview-object-outlook.md)

