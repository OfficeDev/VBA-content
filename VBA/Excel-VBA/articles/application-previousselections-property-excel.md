---
title: Application.PreviousSelections Property (Excel)
keywords: vbaxl10.chm133191
f1_keywords:
- vbaxl10.chm133191
ms.prod: excel
api_name:
- Excel.Application.PreviousSelections
ms.assetid: 967ba122-700c-dca5-1b95-aeaf59e9f19c
ms.date: 06/08/2017
---


# Application.PreviousSelections Property (Excel)

Returns an array of the last four ranges or names selected. Each element in the array is a  **[Range](range-object-excel.md)** object. Read-only **Variant** .


## Syntax

 _expression_ . **PreviousSelections**( **_Index_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Index_|Optional| **Variant**|The index number (from 1 to 4) of the previous range or name.|

## Remarks

Each time you go to a range or cell by using the  **Name** box or the **Go To** command ( **Edit** menu), or each time a macro calls the **[Goto](application-goto-method-excel.md)** method, the previous range is added to this array as element number 1, and the other items in the array are moved down.


## Example

This example displays the cell addresses of all items in the array of previous selections. If there are no previous selections, the  **LBound** function returns an error. This error is trapped, and a message box appears.


```vb
On Error GoTo noSelections 
For i = LBound(Application.PreviousSelections) To _ 
 UBound(Application.PreviousSelections) 
 MsgBox Application.PreviousSelections(i).Address 
Next i 
Exit Sub 
On Error GoTo 0 
 
noSelections: 
 MsgBox "There are no previous selections"
```


## See also


#### Concepts


[Application Object](application-object-excel.md)

