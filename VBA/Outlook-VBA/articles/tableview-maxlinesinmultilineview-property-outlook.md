---
title: TableView.MaxLinesInMultiLineView Property (Outlook)
keywords: vbaol11.chm2520
f1_keywords:
- vbaol11.chm2520
ms.prod: outlook
api_name:
- Outlook.TableView.MaxLinesInMultiLineView
ms.assetid: e9001b61-bae4-72f2-4aa2-6d1c1e4fc086
ms.date: 06/08/2017
---


# TableView.MaxLinesInMultiLineView Property (Outlook)

Returns or sets a  **Long** value that determines the maximum number of lines displayed in multiline mode for the **[TableView](tableview-object-outlook.md)** object. Read/write.


## Syntax

 _expression_ . **MaxLinesInMultiLineView**

 _expression_ A variable that represents a **TableView** object.


## Remarks

This property can be set to a value between 2 and 20. If this property is set to a value less than 2, the property is set to 2. If this property is set to a value greater than 20, the property is set to 20. The default value for this property is 2.


## See also


#### Concepts


[TableView Object](tableview-object-outlook.md)

