---
title: IconCriteria Object (Excel)
keywords: vbaxl10.chm813072
f1_keywords:
- vbaxl10.chm813072
ms.prod: excel
api_name:
- Excel.IconCriteria
ms.assetid: c3b0480a-6def-c315-32ed-137b64708810
ms.date: 06/08/2017
---


# IconCriteria Object (Excel)

Represents the collection of  **[IconCriterion](iconcriterion-object-excel.md)** objects. Each **IconCriterion** represents the values and threshold type for each icon in an icon set conditional formatting rule.


## Remarks

The  **IconCriteria** collection is returned from the **[IconCriteria](iconsetcondition-iconcriteria-property-excel.md)** property of the **[IconSetCondition](iconsetcondition-object-excel.md)** object. You can access each **IconCriterion** object in the collection by passing an index into the collection. See the example for details.


## Example

The following code example creates a range of numbers representing test scores and then applies an icon set conditional formatting rule to that range. The type of icon set is then changed from the default icons to a five-arrow icon set. Finally, the threshold type is modified from percentile to a hard-coded number.


```vb
Sub CreateIconSetCF() 
 
 Dim cfIconSet As IconSetCondition 
 
 'Fill cells with sample data from 1 to 10 
 With ActiveSheet 
 .Range("C1") = 55 
 .Range("C2") = 92 
 .Range("C3") = 88 
 .Range("C4") = 77 
 .Range("C5") = 66 
 .Range("C6") = 93 
 .Range("C7") = 76 
 .Range("C8") = 80 
 .Range("C9") = 79 
 .Range("C10") = 83 
 .Range("C11") = 66 
 .Range("C12") = 74 
 End With 
 
 Range("C1:C12").Select 
 
 'Create an icon set conditional format for the created sample data range 
 Set cfIconSet = Selection.FormatConditions.AddIconSetCondition 
 
 'Change the icon set to a five-arrow icon set 
 cfIconSet.IconSet = ActiveWorkbook.IconSets(xl5Arrows) 
 
 'The IconCriterion collection contains all IconCriteria 
 'By indexing into the collection you can modify each criterion 
 
 With cfIconSet.IconCriteria(2) 
 .Type = xlConditionValueNumber 
 .Value = 60 
 .Operator = 7 
 End With 
 With cfIconSet.IconCriteria(3) 
 .Type = xlConditionValueNumber 
 .Value = 70 
 .Operator = 7 
 End With 
 With cfIconSet.IconCriteria(4) 
 .Type = xlConditionValueNumber 
 .Value = 80 
 .Operator = 7 
 End With 
 With cfIconSet.IconCriteria(5) 
 .Type = xlConditionValueNumber 
 .Value = 90 
 .Operator = 7 
 End With 
 
End Sub
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

