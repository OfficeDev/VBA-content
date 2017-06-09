---
title: UsedObjects Object (Excel)
keywords: vbaxl10.chm677072
f1_keywords:
- vbaxl10.chm677072
ms.prod: excel
api_name:
- Excel.UsedObjects
ms.assetid: b94ad3d1-411f-acf6-19bb-8e6c4a484748
ms.date: 06/08/2017
---


# UsedObjects Object (Excel)

Represents objects that have been allocated in a workbook.


## Example

Use the  **[UsedObjects](application-usedobjects-property-excel.md)** property of the **[Application](application-object-excel.md)** object to return a **UsedObjects** collection.

Once a  **UsedObjects** collection is returned, you can determine the quantity of used objects in a Microsoft Excel application using the **[Count](usedobjects-count-property-excel.md)** property.



In this example, Microsoft Excel determines the quantity of objects that have been allocated and notifies the user. This example assumes a recalculation was performed in the application and was interrupted before finishing.






```vb
Sub CountUsedObjects() 
 
 MsgBox "The number of used objects in this application is: " &; _ 
 Application.UsedObjects.Count 
 
End Sub
```


## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)


