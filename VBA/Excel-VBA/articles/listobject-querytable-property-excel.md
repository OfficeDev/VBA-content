---
title: ListObject.QueryTable Property (Excel)
keywords: vbaxl10.chm734089
f1_keywords:
- vbaxl10.chm734089
ms.prod: excel
api_name:
- Excel.ListObject.QueryTable
ms.assetid: fe019d61-654a-9c87-0bf4-30590a1274ca
ms.date: 06/08/2017
---


# ListObject.QueryTable Property (Excel)

Returns the  **[QueryTable](querytable-object-excel.md)** object that provides a link for the **[ListObject](listobject-object-excel.md)** object to the list server. Read-only.


## Syntax

 _expression_ . **QueryTable**

 _expression_ A variable that represents a **ListObject** object.


## Example

The following example creates a connection to a SharePoint site and publishes the  **ListObject** object named `List1` to the server. A reference to the **QueryTable** object for the list object is created and the **MaintainConnection** property of the **QueryTable** object is set to **True** so that the connection to the SharePoint site is maintained between trips to the server.


```vb
Dim wrksht As Worksheet 
 Dim objListObj As ListObject 
 Dim objQryTbl As QueryTable 
 Dim prpQryProp As pro 
 Dim arTarget(4) As String 
 Dim strSTSConnection As String 
 
 Set wrksht = ActiveWorkbook.Worksheets("Sheet1") 
 Set objListObj = wrksht.ListObjects(1) 
 
 arTarget(0) = "0" 
 arTarget(1) = "http://myteam/project1" 
 arTarget(2) = "1" 
 arTarget(3) = "List1" 
 
 strSTSConnection = objListObj.Publish(arTarget, True) 
 
 Set objQryTbl = objListObj.QueryTable 
 
 objQryTbl.MaintainConnection = True
```


## See also


#### Concepts


[ListObject Object](listobject-object-excel.md)

