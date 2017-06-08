---
title: Window.Caption Property (Excel)
keywords: vbaxl10.chm356080
f1_keywords:
- vbaxl10.chm356080
ms.prod: excel
api_name:
- Excel.Window.Caption
ms.assetid: d8a5ca13-90b8-d7ce-d041-2cdc544789e5
ms.date: 06/08/2017
---


# Window.Caption Property (Excel)

Returns or sets a  **Variant** value that represents the name that appears in the title bar of the document window.


## Syntax

 _expression_ . **Caption**

 _expression_ A variable that represents a **[Window](window-object-excel.md)** object.


## Remarks

When you set the name, you can use that name as the index to the  **[Windows](windows-object-excel.md)** collection (as demonstrated in the example.)


## Example

This example sets the name of the first window in the active workbook to be "Consolidated Balance Sheet." This name is then used as the index to that window in the  **Windows** collection.


```vb
ActiveWorkbook.Windows(1).Caption = "Consolidated Balance Sheet" 
ActiveWorkbook.Windows("Consolidated Balance Sheet") _ 
 .ActiveSheet.Calculate
```


## See also


#### Concepts


[Window Object](window-object-excel.md)

