---
title: Workbook.AfterSave Event (Excel)
keywords: vbaxl10.chm503107
f1_keywords:
- vbaxl10.chm503107
ms.prod: excel
api_name:
- Excel.Workbook.AfterSave
ms.assetid: 97fee36a-f77c-29ab-de1d-b6069b2d74d8
ms.date: 06/08/2017
---


# Workbook.AfterSave Event (Excel)

Occurs after the workbook is saved.


## Syntax

 _expression_ . **AfterSave**( **_Success_** )

 _expression_ A variable that represents a **[Workbook](workbook-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Success_|Required| **Boolean**|Returns  **True** if the save operation was successful; otherwise **False** .|

### Return Value

Nothing


## Example

The following code example displays a message box if the workbook was successfully saved.


```vb
Private Sub Workbook_AfterSave(ByVal Success As Boolean) 
If Success Then 
 MsgBox ("The workbook was successfully saved.") 
End If 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

