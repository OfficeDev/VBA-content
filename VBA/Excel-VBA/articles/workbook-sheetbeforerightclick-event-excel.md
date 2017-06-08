---
title: Workbook.SheetBeforeRightClick Event (Excel)
keywords: vbaxl10.chm503087
f1_keywords:
- vbaxl10.chm503087
ms.prod: excel
api_name:
- Excel.Workbook.SheetBeforeRightClick
ms.assetid: d84dd9fd-85d3-009e-281b-cfc0d2874859
ms.date: 06/08/2017
---


# Workbook.SheetBeforeRightClick Event (Excel)

Occurs when any worksheet is right-clicked, before the default right-click action.


## Syntax

 _expression_ . **SheetBeforeRightClick**( **_Sh_** , **_Target_** , **_Cancel_** )

 _expression_ An expression that returns a **Workbook** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Sh_|Required| **Object**|A  **[Worksheet](worksheet-object-excel.md)** object that represents the sheet.|
| _Target_|Required| **Range**|The cell nearest to the mouse pointer when the right-click occurred.|
| _Cancel_|Required| **Boolean**| **False** when the event occurs. If the event procedure sets this argument to **True** , the default right-click action isn't performed when the procedure is finished.|

## Remarks

This event doesn't occur on chart sheets.


## Example

This example disables the default right-click action. For another example, see the [BeforeRightClick](workbook-sheetbeforerightclick-event-excel.md)event example.


```vb
Private Sub Workbook_SheetBeforeRightClick(ByVal Sh As Object, _ 
 ByVal Target As Range, ByVal Cancel As Boolean) 
 Cancel = True 
End Sub
```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

