---
title: BusinessCardView.Apply Method (Outlook)
keywords: vbaol11.chm2921
f1_keywords:
- vbaol11.chm2921
ms.prod: outlook
api_name:
- Outlook.BusinessCardView.Apply
ms.assetid: 4a64b59e-0d52-8439-30bb-32d0624cf28a
ms.date: 06/08/2017
---


# BusinessCardView.Apply Method (Outlook)

Applies the  **[BusinessCardView](businesscardview-object-outlook.md)** object to the current view.


## Syntax

 _expression_ . **Apply**

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

