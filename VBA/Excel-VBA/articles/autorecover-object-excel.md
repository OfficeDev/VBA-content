---
title: AutoRecover Object (Excel)
keywords: vbaxl10.chm695072
f1_keywords:
- vbaxl10.chm695072
ms.prod: excel
api_name:
- Excel.AutoRecover
ms.assetid: 02fb24e7-4823-7e52-79d7-3d2726f31227
ms.date: 06/08/2017
---


# AutoRecover Object (Excel)

Represents the automatic recovery features of a workbook. 


## Remarks

Properties for the  **AutoRecover** object determine the path and time interval for backing up all files.

Use the  **[AutoRecover](application-autorecover-property-excel.md)** property of the **[Application](application-object-excel.md)** object to return an **AutoRecover** object.

Use the  **[Path](autorecover-path-property-excel.md)** property of the **AutoRecover** object to set the path for where the AutoRecover file will be saved.


## Example

The following example sets the path of the AutoRecover file to drive C.


```vb
Sub SetPath() 
 
 Application.AutoRecover.Path = "C:\" 
 
End Sub
```

Use the  **[Time](autorecover-time-property-excel.md)** property of the **AutoRecover** object to set the time interval for backing up all files.


 **Note**  Units for the  **Time** property are in minutes.




```vb
Sub SetTime() 
 
 Application.AutoRecover.Time = 5 
 
End Sub
```


## See also


#### Other resources



[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)

