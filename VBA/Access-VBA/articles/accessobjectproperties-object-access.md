---
title: AccessObjectProperties Object (Access)
keywords: vbaac10.chm12698
f1_keywords:
- vbaac10.chm12698
ms.prod: access
api_name:
- Access.AccessObjectProperties
ms.assetid: 2df86891-6038-d147-2a32-f1c77b841067
ms.date: 06/08/2017
---


# AccessObjectProperties Object (Access)

The  **AccessObjectProperties** collection contains all of the custom **[AccessObjectProperty](accessobjectproperty-object-access.md)** objects of a specific instance of an object. These **AccessObjectProperty** objects (which are often just called properties) uniquely characterize that instance of the object.


## Remarks

Use the  **AccessObjectProperties** collection in Visual Basic or in an expression to refer to properties of the **CurrentProject**, **CodeProject**, or **AccessObject** object. For example, you can enumerate the **AccessObjectProperties** collection to set or return the values of properties of an individual report.


 **Note**  The  **AccessObjectProperties** collection isn't accessible for objects derived from the **CurrentData** object (for example, `CurrentData.AllTables!Table1`). For objects derived in this manner, you can only access their built-in properties by direct calls to the desired property (for example,  `CurrentData.AllTables!Table1.Name`).

To add a user-defined property to an existing instance of an object, first define its characteristics and add it to the collection with the  **[Add](accessobjectproperties-add-method-access.md)** method. Referencing a user-defined **AccessObjectProperty** object that has not yet been appended to an **AccessObjectProperties** collection will cause an error, as will appending a user-defined AccessObjectProperty object to an **AccessObjectProperties** collection containing an **AccessObjectProperty** object of the same name.

You can use the  **[Remove](accessobjectproperties-remove-method-access.md)** method to remove user-defined properties from the **AccessObjectProperties** collection.


 **Note**  A built-in or user-defined  **AccessObjectProperty** object is associated only with the specific instance of an object. The property isn't defined for all instances of objects of the selected type.

To refer to a built-in or user-defined  **AccessObjectProperty** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms:




```
CurrentProject.AllForms("Form1").Properties(0) 
CurrentProject.AllForms("Form1").Properties("name") 
CurrentProject.AllForms("Form1").Properties![name]
```

With the same syntax forms, you can also refer to the Value property of a  **AccessObjectProperty** object. The context of the reference will determine whether you are referring to the **AccessObjectProperty** object itself or the **[Value](accessobjectproperty-value-property-access.md)** property of the **AccessObjectProperty** object.


 **Note**  Properties in the  **AccessObjectProperties** collection are not stored and can be lost when the object they are associated with is checked in or out using the Source Code Control add-in.


## Methods



|**Name**|
|:-----|
|[Add](accessobjectproperties-add-method-access.md)|
|[Remove](accessobjectproperties-remove-method-access.md)|

## Properties



|**Name**|
|:-----|
|[Application](accessobjectproperties-application-property-access.md)|
|[Count](accessobjectproperties-count-property-access.md)|
|[Item](accessobjectproperties-item-property-access.md)|
|[Parent](accessobjectproperties-parent-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
