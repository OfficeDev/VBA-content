---
title: OLEDBConnection.RobustConnect Property (Excel)
keywords: vbaxl10.chm794088
f1_keywords:
- vbaxl10.chm794088
ms.prod: excel
api_name:
- Excel.OLEDBConnection.RobustConnect
ms.assetid: 47ca146c-54ba-b2d5-6d16-15571daf08f3
ms.date: 06/08/2017
---


# OLEDBConnection.RobustConnect Property (Excel)

 Returns or sets how OLE DB connection connects to its data source. Read/write **[XlRobustConnect](xlrobustconnect-enumeration-excel.md)** .


## Syntax

 _expression_ . **RobustConnect**

 _expression_ A variable that represents an **OLEDBConnection** object.


## Remarks

If you use robust connectivity, when Microsoft Excel is unable to establish a connection using the workbook connection information, Excel will check if the workbook connection has a path to an external connection file. If it does, Excel will open the external connection file and try to connect using the information contained in the external connection file. If a connection can be established after using the external connection file, Excel will then update the workbook's connection object. 

This provides robust connectivity in scenarios where an Information Technology Department maintains and updates connections in a central place, permitting a user's workbook to automatically fetch those updates whenever the previous version of the connection (cached within the workbook) fails. 




 **Note**  Robust connectivity isn't secure. If you create a connection on your computer and use it in a workbook and then send someone the workbook, that person will be able to see the path to the connection file on your computer. It is a good idea to remove the connection file information from the workbook before you send the workbook to someone else.


## See also


#### Concepts


[OLEDBConnection Object](oledbconnection-object-excel.md)

