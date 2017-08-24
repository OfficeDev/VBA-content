---
title: Table.ApplyAutoFormat Method (Publisher)
keywords: vbapb10.chm4784137
f1_keywords:
- vbapb10.chm4784137
ms.prod: publisher
api_name:
- Publisher.Table.ApplyAutoFormat
ms.assetid: f792a5f3-0d1c-06de-a030-7a588ca372d2
ms.date: 06/08/2017
---


# Table.ApplyAutoFormat Method (Publisher)

Applies automatic built-in table formatting to a specified table.


## Syntax

 _expression_. **ApplyAutoFormat**( **_AutoFormat_**,  **_TextFormatting_**,  **_TextAlignment_**,  **_Fill_**,  **_Borders_**)

 _expression_A variable that represents a  **Table** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|AutoFormat|Required| **PbTableAutoFormatType**|The type of automatic formatting to apply to the specified table.|
|TextFormatting|Optional| **Boolean**| **True** to apply font formatting to the text in the table. Default value is **True**.|
|TextAlignment|Optional| **Boolean**| **True** to apply text alignment to the text in the table. Default value is **True**.|
|Fill|Optional| **Boolean**| **True** to apply fill formatting to cells in the table. Default value is **True**.|
|Borders|Optional| **Boolean**| **True** to apply borders to cells in the table. Default value is **True**.|

## Remarks

The AutoFormat parameter can be one of the  **[PbTableAutoFormatType](pbtableautoformattype-enumeration-publisher.md)** constants declared in the Microsoft Publisher type library.


## Example

This example applies the checkbook register automatic formatting, with fill and borders, to the specified table.


```vb
Sub ApplyAutomaticTableFormatting() 
 ActiveDocument.Pages(1).Shapes(1).Table.ApplyAutoFormat _ 
 AutoFormat:=pbTableAutoFormatCheckbookRegister, _ 
 Borders:=False 
End Sub
```


