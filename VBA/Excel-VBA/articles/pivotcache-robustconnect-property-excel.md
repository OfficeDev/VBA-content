---
title: PivotCache.RobustConnect Property (Excel)
keywords: vbaxl10.chm227105
f1_keywords:
- vbaxl10.chm227105
ms.prod: excel
api_name:
- Excel.PivotCache.RobustConnect
ms.assetid: 354d0124-e178-342b-9565-fa74e9dae5d5
ms.date: 06/08/2017
---


# PivotCache.RobustConnect Property (Excel)

Returns or sets how the PivotTable cache connects to its data source. Read/write  **[XlRobustConnect](xlrobustconnect-enumeration-excel.md)** .


## Syntax

 _expression_ . **RobustConnect**

 _expression_ A variable that represents a **PivotCache** object.


## Remarks





| **XlRobustConnect** can be one of these **XlRobustConnect** constants.|
| **xlAlways** . The cache always uses external source information (as defined by the **[SourceConnectionFile](pivotcache-sourceconnectionfile-property-excel.md)** or **[SourceDataFile](pivotcache-sourcedatafile-property-excel.md)** property) to reconnect.|
| **xlAsRequired** . The cache uses external source info to reconnect using the **[Connection](pivotcache-connection-property-excel.md)** property.|
| **xlNever** . The cache never uses source info to reconnect.|

## Example

The following example determines the setting for the cache connection and notifies the user. The example assumes a PivotTable exists on the active worksheet.


```vb
Sub CheckRobustConnect() 
 
 Dim pvtCache As PivotCache 
 
 Set pvtCache = Application.ActiveWorkbook.PivotCaches.Item(1) 
 
 ' Determine the connection robustness and notify user. 
 Select Case pvtCache.RobustConnect 
 Case xlAlways 
 MsgBox "The PivotTable cache is always connected to its source." 
 Case xlAsRequired 
 MsgBox "The PivotTable cache is connected to its source as required." 
 Case xlNever 
 MsgBox "The PivotTable cache is never connected to its source." 
 End Select 
 
End Sub
```


## See also


#### Concepts


[PivotCache Object](pivotcache-object-excel.md)

