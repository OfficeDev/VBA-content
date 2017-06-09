---
title: View.Copy Method (Outlook)
keywords: vbaol11.chm2485
f1_keywords:
- vbaol11.chm2485
ms.prod: outlook
api_name:
- Outlook.View.Copy
ms.assetid: dfa82ef6-94f1-5c7d-eea5-600f992992d3
ms.date: 06/08/2017
---


# View.Copy Method (Outlook)

Creates a new instance of a  **[View](view-object-outlook.md)** object.


## Syntax

 _expression_ . **Copy**( **_Name_** , **_SaveOption_** )

 _expression_ A variable that represents a **View** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Name_|Required| **String**|Represents the name of the new  **View** object.|
| _SaveOption_|Optional| **[OlViewSaveOption](olviewsaveoption-enumeration-outlook.md)**|The save option that defines the permissions of the  **View** object.|

## Example

The following Microsoft Visual Basic for Applications (VBA) example creates a copy of a view called "New Table View" and saves it in the current folder. To run this example, you need to first create a view called 'New Table View' programmatically or by using the Outlook user interface.


```vb
Sub CopyView() 
 
 'Copies a view 
 
 Dim objViews As Outlook.Views 
 
 Dim objNewView As Outlook.View 
 
 
 
 Set objViews = _ 
 
 Application.GetNamespace("MAPI").GetDefaultFolder(olFolderInbox).Views 
 
 'Create copy of View object 
 
 Set objNewView = objViews("New Table View").Copy(Name:="Table View Copy", _ 
 
 SaveOption:=olViewSaveOptionThisFolderEveryone) 
 
End Sub
```


## See also


#### Concepts


[View Object](view-object-outlook.md)

