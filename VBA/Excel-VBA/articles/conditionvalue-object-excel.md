---
title: ConditionValue Object (Excel)
keywords: vbaxl10.chm803072
f1_keywords:
- vbaxl10.chm803072
ms.prod: excel
api_name:
- Excel.ConditionValue
ms.assetid: a39335db-4e0a-66aa-393b-3aa7e5268c00
ms.date: 06/08/2017
---


# ConditionValue Object (Excel)

Represents how the shortest bar or longest bar is evaluated for a data bar conditional formatting rule.


## Remarks

The  **ConditionValue** object is returned using either the **[MaxPoint](databar-maxpoint-property-excel.md)** or **[MinPoint](databar-minpoint-property-excel.md)** property of the **[Databar](databar-object-excel.md)** object.

You can change the type of evaluation from the default setting (lowest value for the shortest bar and highest value for the longest bar) by using the  **[Modify](conditionvalue-modify-method-excel.md)** method.


## Example

The following example creates a range of data and then applies a data bar to the range. You will notice that because there is an extremely low and high value in the range, the middle values have data bars that are of similiar length. To disambiguate the middle values, the sample code uses the  **ConditionValue** object to change how the thresholds are evaluated to percentiles.


```vb
Sub CreateDataBarCF() 
 
 Dim cfDataBar As Databar 
 
 'Create a range of data with a couple of extreme values 
 With ActiveSheet 
 .Range("D1") = 1 
 .Range("D2") = 45 
 .Range("D3") = 50 
 .Range("D2:D3").AutoFill Destination:=Range("D2:D8") 
 .Range("D9") = 500 
 End With 
 
 Range("D1:D9").Select 
 
 'Create a data bar with default behavior 
 Set cfDataBar = Selection.FormatConditions.AddDatabar 
 MsgBox "Because of the extreme values, middle data bars are very similar" 
 
 'The MinPoint and MaxPoint properties return a ConditionValue object 
 'which you can use to change threshold parameters 
 cfDataBar.MinPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=5 
 cfDataBar.MaxPoint.Modify newtype:=xlConditionValuePercentile, _ 
 newvalue:=75 
 
End Sub
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


