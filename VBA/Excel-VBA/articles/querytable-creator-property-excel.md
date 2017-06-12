---
title: QueryTable.Creator Property (Excel)
keywords: vbaxl10.chm517074
f1_keywords:
- vbaxl10.chm517074
ms.prod: excel
api_name:
- Excel.QueryTable.Creator
ms.assetid: 6384b8d4-295c-1566-9405-a7450551b4f1
ms.date: 06/08/2017
---


# QueryTable.Creator Property (Excel)

Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long** .


## Syntax

 _expression_ . **Creator**

 _expression_ A variable that represents a **QueryTable** object.


## Remarks

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The  **Creator** property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL.

Data from Web queries or text queries is imported as a  **[QueryTable](querytable-object-excel.md)** object, while all other external data is imported as a **[ListObject](listobject-object-excel.md)** object. You can use the **[QueryTable](listobject-querytable-property-excel.md)** property of the **ListObject** to access the **Creator** property.


## See also


#### Concepts


[QueryTable Object](querytable-object-excel.md)

