---
title: IconCriterion.Icon Property (Excel)
keywords: vbaxl10.chm814077
f1_keywords:
- vbaxl10.chm814077
ms.prod: excel
api_name:
- Excel.IconCriterion.Icon
ms.assetid: bcf25274-2dbb-535d-404c-0eec0f312a15
ms.date: 06/08/2017
---


# IconCriterion.Icon Property (Excel)

Returns or specifies the icon for a criterion in an icon set conditional formatting rule. Read/write


## Syntax

 _expression_ . **Icon**

 _expression_ A variable that represents an **[IconCriterion](iconcriterion-object-excel.md)** object.


## Remarks

After you set the  **Icon** property for the icon criterion in an icon set conditional formatting rule, the **[IconSet](iconsetcondition-iconset-property-excel.md)** property is changed to **xlCustomSet** .


## Example

The following code example creates an icon set conditional formatting rule that displays four icons split across the specified percentages. The icon set is initially set to use the  **4 Arrows (Colored)** icon set, but the **Icon** property is used to override which icons are used for the first and third criteria. After running the code, the icon for the first criterion is the **Red Cross** icon, the icon for the second criterion is the second arrow from the **4 Arrows (Colored)** icon set, the icon for the third criterion is the **Yellow Traffic Light** icon, and the icon for the fourth criterion is the fourth arrow from **4 Arrows (Colored)** icon set.


```vb
Range("A1:A10").Select 
Selection.FormatConditions.AddIconSetCondition 
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority 
 
With Selection.FormatConditions(1) 
 .ReverseOrder = False 
 .ShowIconOnly = False 
 .IconSet = ActiveWorkbook.IconSets(xl4Arrows) 
End With 
 
With Selection.FormatConditions(1).IconCriteria(1) 
 .Icon = xlIconRedCross 
End With 
 
With Selection.FormatConditions(1).IconCriteria(2) 
 .Type = xlConditionValuePercent 
 .Value = 25 
 .Operator = 7 
End With 
 
With Selection.FormatConditions(1).IconCriteria(3) 
 .Type = xlConditionValuePercent 
 .Value = 50 
 .Operator = 7 
 .Icon = xlIconYellowTrafficLight 
End With 
 
With Selection.FormatConditions(1).IconCriteria(4) 
 .Type = xlConditionValuePercent 
 .Value = 75 
 .Operator = 7 
End With
```


## See also


#### Concepts


[IconCriterion Object](iconcriterion-object-excel.md)

