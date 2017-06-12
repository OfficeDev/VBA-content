---
title: AccessObjectProperty Object (Access)
keywords: vbaac10.chm12693
f1_keywords:
- vbaac10.chm12693
ms.prod: access
api_name:
- Access.AccessObjectProperty
ms.assetid: b1a44d34-8ca1-af7d-1878-f2c14fb481f7
ms.date: 06/08/2017
---


# AccessObjectProperty Object (Access)

An  **AccessObjectProperty** object represents a built-in or user-defined characteristic of an **AccessObject** object.


## Remarks

Every  **AccessObject** object contains an **AccessObjectProperties** collection that has **AccessObjectProperty** objects corresponding to the properties of that **AccessObject** object. The user can also define **AccessObjectProperty** objects and append them to the **AccessObjectProperties** collection of some **AccessObject** objects.

You can create user-defined properties for the following objects:


-  **CodeData**, **CodeProject**, **CurrentProject**, and **CurrentData** objects
    
-  **AccessObject** objects in the following collections.
    

|**CurrentProject and CodeProject object collections**|**CodeData and CodeProject object collections**|
|:-----|:-----|
|**[AllForms](allforms-object-access.md)**|**[AllQueries](allqueries-object-access.md)**|
|**[AllReports](allreports-object-access.md)**|**[AllViews](allviews-object-access.md)**|
|**[AllMacros](allmacros-object-access.md)**|**[AllStoredProcedures](allstoredprocedures-object-access.md)**|
|**[AllModules](allmodules-object-access.md)**|**[AllDatabaseDiagrams](alldatabasediagrams-object-access.md)**|
|**[AllTables](alltables-object-access.md)**||

 **Note**  The  **AccessObjectProperties** collection isn't accessible for objects derived from the **CurrentData** object (for example, CurrentData.AllTables!Table1). For objects derived in this manner, you can only access their built-in properties by direct calls to the desired property (for example, CurrentData.AllTables!Table1.Name).

To add a user-defined property, use the  **Add** method to create and add an **AccessObjectProperty** object with a unique **Name** property setting and **Value** property of the new **AccessObjectProperty** object to the **AccessObjectProperties** collection of the appropriate object. The object to which you are adding the user-defined property must already be appended to a collection. Referencing a user-defined **AccessObjectProperty** object that has not yet been appended to an **AccessObjectProperties** collection will cause an error, as will appending a user-defined **AccessObjectProperty** object to an **AccessObjectProperties** collection containing an **AccessObjectProperty** object of the same name.

You can delete user-defined properties from the  **AccessObjectProperties** collection.


 **Note**  A user-defined  **AccessObjectProperty** object is associated only with the specific instance of an object. The property isn't defined for all instances of objects of the selected type.

The  **AccessObjectProperty** object has two built-in properties:


- The  **Name** property, a **String** that uniquely identifies the property.
    
- The  **Value** property, a **Variant** that contains the property setting.
    
To refer to a built-in or user-defined  **AccessObjectProperty** object in a collection by its ordinal number or by its **Name** property setting, use any of the following syntax forms:




```
CurrentProject.AllForms("Form1").Properties(0) 
CurrentProject.AllForms("Form1").Properties("name") 
CurrentProject.AllForms("Form1").Properties![name]
```

With the same syntax forms, you can also refer to the  **Value** property of a **AccessObjectProperty** object. The context of the reference will determine whether you are referring to the **AccessObjectProperty** object itself or the **Value** property of the **AccessObjectProperty** object.


 **Note**  Properties in the  **AccessObjectProperties** collection are not stored and can be lost when when the object they are associated with is checked in or out using the Source Code Control add-in.


## Properties



|**Name**|
|:-----|
|[Name](accessobjectproperty-name-property-access.md)|
|[Value](accessobjectproperty-value-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
