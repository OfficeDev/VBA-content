---
title: Application Object
keywords: vbagr10.chm3077640
f1_keywords:
- vbagr10.chm3077640
ms.prod: excel
api_name:
- Excel.Application
ms.assetid: 553a0ee2-83da-6d32-f082-15e93e7b0e4d
ms.date: 06/08/2017
---


# Application Object

Represents the entire Microsoft Graph application. The  **Application** object represents the top level of the object hierarchy and contains all of the objects, properties, and methods for the application.


## Using the Application Object

Use the  **Application** property to return the **Application** object. The following example applies the **DataSheet** property to the **Application** object.


```
myChart.Application.DataSheet.Range("A1").Value = 32
```


