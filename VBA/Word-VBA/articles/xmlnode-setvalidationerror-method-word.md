---
title: XMLNode.SetValidationError Method (Word)
keywords: vbawd10.chm37748841
f1_keywords:
- vbawd10.chm37748841
ms.prod: word
api_name:
- Word.XMLNode.SetValidationError
ms.assetid: 19e2cb53-5e57-4cfe-52d6-c1d42154bc46
ms.date: 06/08/2017
---


# XMLNode.SetValidationError Method (Word)

Changes the validation error text displayed to a user for a specified node and forces Word to report a node as invalid.


## Syntax

 _expression_ . **SetValidationError**( **_Status_** , **_ErrorText_** , **_ClearedAutomatically_** )

 _expression_ Required. A variable that represents a **[XMLNode](xmlnode-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Status_|Required| **WdXMLValidationStatus**|Specifies whether to set the validation status error text ( **wdXMLValidationStatusCustom** ) or to clear the validation status error text ( **wdXMLValidationStatusOK** ).|
| _ErrorText_|Optional| **Variant**|The text displayed to the user. Leave blank when the Status parameter is set to  **wdXMLValidationStatusOK** .|
| _ClearedAutomatically_|Optional| **Boolean**| **True** automatically clears the error message as soon as the next validation event occurs on the specified node. **False** requires running the **SetValidationError** method with a Status parameter of **wdXMLValidationStatusOK** to clear the custom error text.|

## Remarks

To set custom error text, use the  **wdXMLValidationStatusCustom** constant.


## Example

The following example specifies custom validation error text.


```vb
Dim objNode As XMLNode 
 
Set objNode = ActiveDocument.XMLNodes(1) 
objNode.SetValidationError wdXMLValidationStatusCustom, _ 
 "Error Text", True
```


## See also


#### Concepts


[XMLNode Object](xmlnode-object-word.md)

