---
title: NewFile.Add Method (Office)
keywords: vbaof11.chm235001
f1_keywords:
- vbaof11.chm235001
ms.prod: office
api_name:
- Office.NewFile.Add
ms.assetid: 094e4093-fc2d-beaa-4a63-b3ad88557907
ms.date: 06/08/2017
---


# NewFile.Add Method (Office)

Adds a new item to the  **New Item** task pane. Returns a **Boolean** value to indicate whether the operation was successful.


## Syntax

 _expression_. **Add**( **_FileName_**, **_Section_**, **_DisplayName_**, **_Action_** )

 _expression_ Required. A variable that represents a **[NewFile](newfile-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file to add to the list of files on the task pane.|
| _Section_|Optional|**Variant**|The section to which to add the file. Can be any  **msoFileNew** constant.|
| _DisplayName_|Optional|**Variant**|The text to display in the task pane.|
| _Action_|Optional|**Variant**|The action to take when a user clicks the item. Can be any  **msoFileNew** constant.|

## See also


#### Concepts


[NewFile Object](newfile-object-office.md)
#### Other resources


[NewFile Object Members](newfile-members-office.md)

