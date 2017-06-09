---
title: CalloutFormat.Border Property (Excel)
keywords: vbaxl10.chm104010
f1_keywords:
- vbaxl10.chm104010
ms.prod: excel
api_name:
- Excel.CalloutFormat.Border
ms.assetid: 6d0c78d9-b30a-c1ff-940a-e15b4decad42
ms.date: 06/08/2017
---


# CalloutFormat.Border Property (Excel)

Returns or sets a  **[MsoTriState](http://msdn.microsoft.com/library/2036cfc9-be7d-e05c-bec7-af05e3c3c515%28Office.15%29.aspx)** value that represents the visibility options for the border of the object.


## Syntax

 _expression_ . **Border**

 _expression_ A variable that represents a **CalloutFormat** object.


## Remarks

The value of this property can be set to one of the following  **MsoTriState** constants:



| **msoCTrue** Does not apply to this object.|
| **msoFalse** Sets the border invisible.|
| **msoTriStateMixed** Does not apply to this object.|
| **msoTriStateToggle** Allows the user to switch the border from visible to invisible and vice versa.|
| **msoTrue**_default_ . Sets the border visible.|

## See also


#### Concepts


[CalloutFormat Object](calloutformat-object-excel.md)

