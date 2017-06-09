---
title: TableStyle.Condition Method (Word)
keywords: vbawd10.chm244776976
f1_keywords:
- vbawd10.chm244776976
ms.prod: word
api_name:
- Word.TableStyle.Condition
ms.assetid: f0adb8b7-434d-3134-38d0-d21d221a27d3
ms.date: 06/08/2017
---


# TableStyle.Condition Method (Word)

Returns a  **[ConditionalStyle](conditionalstyle-object-word.md)** object that represents special style formatting for a portion of a table.


## Syntax

 _expression_ . **Condition**( **_ConditionCode_** )

 _expression_ Required. A variable that represents a **[TableStyle](tablestyle-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ConditionCode_|Required| [**WdConditionCode**](wdconditioncode-enumeration-word.md)|The area of the table to which to apply the formatting.|

## Example

This example selects the first table in the active document and adds a 20 percent shading to odd-numbered columns.


```vb
Sub TableStylesTest() 
 With ActiveDocument 
 
 'Select the table to which the conditional 
 'formatting will apply 
 .Tables(1).Select 
 
 'Specify the conditional formatting 
 .Styles("Table Grid").Table _ 
 .Condition(wdOddColumnBanding).Shading _ 
 .BackgroundPatternColor = wdColorGray20 
 End With 
End Sub
```


## See also


#### Concepts


[TableStyle Object](tablestyle-object-word.md)

