---
title: NewFile.Remove Method (Office)
keywords: vbaof11.chm235002
f1_keywords:
- vbaof11.chm235002
ms.prod: office
api_name:
- Office.NewFile.Remove
ms.assetid: 1954580b-3c8b-3e4b-0884-8d32932fbf58
ms.date: 06/08/2017
---


# NewFile.Remove Method (Office)

Removes an item from the  **New Item** task pane. Returns a **Boolean** value to indicate whether the operation was successful.


## Syntax

 _expression_. **Remove**( **_FileName_**, **_Section_**, **_DisplayName_**, **_Action_** )

 _expression_ Required. A variable that represents a **[NewFile](newfile-object-office.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _FileName_|Required|**String**|The name of the file reference.|
| _Section_|Optional|**Variant**|The section of the task pane where the file reference exists. Can be any  **msoFileNew** constant.|
| _DisplayName_|Optional|**Variant**|The display text of the file reference.|

## Remarks

The arguments supplied to the  **Remove** method must match the arguments that were supplied to the **Add** method of the **NewFile** object, or the **Remove** method will fail. For example, if the **Action** argument was supplied when the **NewFile** object was added, then the same **Action** argument must be supplied to remove the **NewFile** object, or the **Remove** method will fail.


## See also


#### Concepts


[NewFile Object](newfile-object-office.md)
#### Other resources


[NewFile Object Members](newfile-members-office.md)

