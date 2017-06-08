---
title: ComboBox.Selected Property (Access)
keywords: vbaac10.chm11494
f1_keywords:
- vbaac10.chm11494
ms.prod: access
api_name:
- Access.ComboBox.Selected
ms.assetid: fc643ebc-084a-c11c-2489-7d1504d5b17b
ms.date: 06/08/2017
---


# ComboBox.Selected Property (Access)

You can use the  **Selected** property in Visual Basic to determine if an item in a combo box is selected. Read/write **Long**.


## Syntax

 _expression_. **Selected**( ** _lRow_** )

 _expression_ A variable that represents a **ComboBox** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _lRow_|Required|**Long**|The item in the combo box. The first item is represented by a zero (0), the second by a one (1), and so on.|

## Remarks

The  **Selected** property is a zero-based array that contains the selected state of each item in a combo box.



|**Setting**|**Description**|
|:-----|:-----|
|**True**|The combo box item is selected.|
|**False**|The combo box item is not selected.|
This property is available only at run time.

You can use the  **Selected** property to select items in a combo box by using Visual Basic. For example, the following expression selects the fifth item in the list:




```vb
Me!Combobox.Selected(4) = True
```


## See also


#### Concepts


[ComboBox Object](combobox-object-access.md)

