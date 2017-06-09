---
title: ServerViewableItems.DeleteAll Method (Excel)
keywords: vbaxl10.chm833076
f1_keywords:
- vbaxl10.chm833076
ms.prod: excel
api_name:
- Excel.ServerViewableItems.DeleteAll
ms.assetid: 8f2bf876-50ba-3b91-d353-6d73a35e9462
ms.date: 06/08/2017
---


# ServerViewableItems.DeleteAll Method (Excel)

Deletes references to all the objects in the  **[ServerViewableItems](serverviewableitems-object-excel.md)** collection in the workbook.


## Syntax

 _expression_ . **DeleteAll**

 _expression_ A variable that represents a **ServerViewableItems** object.


## Remarks

If you do not want any of the objects in the  **ServerViewableItems** collection to be viewable on the server, use this method to remove them all at once.


 **Note**  If the  **ServerViewableItems** collection does not contain at least one object, you will see the message "Unable to Display Specified Named Range or Item" when viewing the workbook in Excel Services.


## See also


#### Concepts


[ServerViewableItems Object](serverviewableitems-object-excel.md)

