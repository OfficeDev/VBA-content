---
title: AllFunctions Object (Access)
keywords: vbaac10.chm13250
f1_keywords:
- vbaac10.chm13250
ms.prod: access
api_name:
- Access.AllFunctions
ms.assetid: 1420cf24-906e-7b65-29f3-29a28cdf92cf
ms.date: 06/08/2017
---


# AllFunctions Object (Access)

The  **AllFunctions** collection contains an **[AccessObject](accessobject-object-access.md)** object for each function in the **[CurrentData](currentdata-object-access.md)** or **[CodeData](codedata-object-access.md)** object.


## Remarks

The  **CurrentData** or **CodeData** object has an **AllFunctions** collection containing **AccessObject** objects that describe instances of all functions specified by the **CurrentData** or **CodeData** objects. For example, you can enumerate the **AllFunctions** collection in Visual Basic to set or return the values of properties of individual **AccessObject** objects in the collection.

You can refer to an individual  **AccessObject** object in the **AllFunctions** collection either by referring to the object by name, or by referring to its index within the collection. If you want to refer to a specific object in the **AllFunctions** collection, it's better to refer to the function by name because a function's collection index may change.

The  **AllFunctions** collection is indexed beginning with zero. If you refer to a function by its index, the first function is AllFunctions(0), the second table is AllFunctions(1), and so on.

To list all open functions in the database, use the  **[IsLoaded](accessobject-isloaded-property-access.md)** property of each **AccessObject** object in the **AllFunctions** collection. You can then use the **Name** property of each individual **AccessObject** object to return the name of a function.

You can't add or delete an  **AccessObject** object from the **AllFunctions** collection.


## Properties



|**Name**|
|:-----|
|[Application](allfunctions-application-property-access.md)|
|[Count](allfunctions-count-property-access.md)|
|[Item](allfunctions-item-property-access.md)|
|[Parent](allfunctions-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
