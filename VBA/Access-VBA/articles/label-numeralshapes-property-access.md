---
title: Label.NumeralShapes Property (Access)
keywords: vbaac10.chm10232
f1_keywords:
- vbaac10.chm10232
ms.prod: access
api_name:
- Access.Label.NumeralShapes
ms.assetid: 3da2f917-a257-b9aa-3517-f4d65bc3af18
ms.date: 06/08/2017
---


# Label.NumeralShapes Property (Access)





## Syntax

 _expression_. **NumeralShapes**

 _expression_ A variable that represents a **Label** object.


## Remarks

The  **NumeralShapes** property uses the following settings.



|**Setting**|**Visual Basic**|**Description**|
|:-----|:-----|:-----|
|System|0|Numeral shapes determined by the  **Numeral Shapes** system setting.|
|Arabic|1|Arabic digit shapes will be used to display and print numerals.|
|National|2|National digit shapes will be used to display and print numerals.|
|Context|3|Numeral shapes determined by Unicode context rules for adjacent text.|

## Example

The following example changes the  **NumeralShapes** property for the selected control to 0 (numeral shapes will be determined by the **Numeral Shapes** system setting).


```vb
Public Sub ChangeNumeralShapes(ctl As Control) 
 ctl.NumeralShapes = 0 
End Sub
```


## See also


#### Concepts


[Label Object](label-object-access.md)

