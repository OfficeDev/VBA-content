---
title: Form.BeforeScreenTip Event (Access)
keywords: vbaac10.chm13678
f1_keywords:
- vbaac10.chm13678
ms.prod: access
api_name:
- Access.Form.BeforeScreenTip
ms.assetid: 08e67747-9023-e880-c246-1aa9e9c447ed
ms.date: 06/08/2017
---


# Form.BeforeScreenTip Event (Access)

Occurs before a ScreenTip is displayed for an element in a PivotChart view or PivotTable view.


## Syntax

 _expression_. **BeforeScreenTip**( ** _ScreenTipText_**, ** _SourceObject_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ScreenTipText_|Required|**Object**|Set the Value property of this object to the ScreenTip that you want to display. Changing this argument to an empty string effectively hides the ScreenTip.|
| _SourceObject_|Required|**Object**|The object that generates the ScreenTip.|

### Return Value

nothing


## See also


#### Concepts


[Form Object](form-object-access.md)

