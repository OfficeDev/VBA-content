---
title: Scenarios Object (Excel)
keywords: vbaxl10.chm361072
f1_keywords:
- vbaxl10.chm361072
ms.prod: excel
api_name:
- Excel.Scenarios
ms.assetid: 90d6ff4b-f329-a04c-040e-a39bb501a58b
ms.date: 06/08/2017
---


# Scenarios Object (Excel)

A collection of all the  **[Scenario](scenario-object-excel.md)** objects on the specified worksheet.


## Remarks

 A scenario is a group of input values (called changing cells) that's named and saved.


## Example

Use the  **[Scenarios](worksheet-scenarios-method-excel.md)** method to return the **Scenarios** collection. The following example creates a summary for the scenarios on the worksheet named "Options," using cells J10 and J20 as the result cells.


```vb
Worksheets("options").Scenarios.CreateSummary _ 
 resultCells:=Worksheets("options").Range("j10,j20")
```

Use the  **[Add](scenarios-add-method-excel.md)** method to create a new scenario and add it to the collection. The following example adds a new scenario named "Typical" to the worksheet named "Options." The new scenario has two changing cells, A2 and A12, with the respective values 55 and 60.




```vb
Worksheets("options").Scenarios.Add name:="Typical", _ 
 changingCells:=Worksheets("options").Range("A2,A12"), _ 
 values:=Array("55", "60")
```

Use  **Scenarios** ( _index_ ), where _index_ is the scenario name or index number, to return a single **Scenario** object. The following example shows the scenario named "Typical" on the worksheet named "Options."




```vb
Worksheets("options").Scenarios("typical").Show
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

