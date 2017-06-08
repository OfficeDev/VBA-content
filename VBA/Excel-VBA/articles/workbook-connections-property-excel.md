---
title: Workbook.Connections Property (Excel)
keywords: vbaxl10.chm199235
f1_keywords:
- vbaxl10.chm199235
ms.prod: excel
api_name:
- Excel.Workbook.Connections
ms.assetid: 9c4f4ba7-dd4b-0bc2-65b7-16455014097f
ms.date: 06/08/2017
---


# Workbook.Connections Property (Excel)

The  **Connections** property establishes a connection between the workbook and an ODBC or an OLEDB data source and refreshes the data without prompting the user. Read-only.


## Syntax

 _expression_ . **Connections**

 _expression_ A variable that represents a **Workbook** object.


## Example

The following example rereshes the OBDC and OLEDB connections of the active workbook.

```vb
ActiveWorkbook.Connections(1).ODBCConnection.Refresh 
ActiveWorkbook.Connections(1).OLEDBConnection.Refresh 

```


## See also


#### Concepts


[Workbook Object](workbook-object-excel.md)

