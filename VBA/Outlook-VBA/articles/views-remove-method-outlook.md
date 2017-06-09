---
title: Views.Remove Method (Outlook)
keywords: vbaol11.chm548
f1_keywords:
- vbaol11.chm548
ms.prod: outlook
api_name:
- Outlook.Views.Remove
ms.assetid: 73a92be6-8dc4-6fb9-7f20-0ff678445737
ms.date: 06/08/2017
---


# Views.Remove Method (Outlook)

Removes an object from the collection.


## Syntax

 _expression_ . **Remove** **_Index_**

 _expression_ A variable that represents a **Views** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Required| **Variant**|The name or 1-based index value of an object within a collection.|

## Example

The following example removes a View object from the Views collection.


```vb
Sub DeleteView() 
 
 'Deletes a view from the collection 
 
 Dim objName As Outlook.NameSpace 
 
 Dim objViews As Outlook.Views 
 
 Dim objView As Outlook.View 
 
 Dim strName As String 
 
 
 
 strName = "New Icon View" 
 
 Set objName = Application.GetNamespace("MAPI") 
 
 Set objViews = objName.GetDefaultFolder(olFolderNotes).Views 
 
 For Each objView In objViews 
 
 If objView.Name = strName Then 
 
 objViews.Remove (strName) 
 
 End If 
 
 Next objView 
 
End Sub
```


## See also


#### Concepts


[Views Object](views-object-outlook.md)

