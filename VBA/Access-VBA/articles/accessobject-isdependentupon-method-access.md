---
title: AccessObject.IsDependentUpon Method (Access)
keywords: vbaac10.chm12755
f1_keywords:
- vbaac10.chm12755
ms.prod: access
api_name:
- Access.AccessObject.IsDependentUpon
ms.assetid: aba465c5-4176-c69a-8eb8-1a6737b6d8cf
ms.date: 06/08/2017
---


# AccessObject.IsDependentUpon Method (Access)

Returns a  **Boolean** value that indicates whether the specified object is dependent upon the database object specified in the _ObjectName_ argument.


## Syntax

 _expression_. **IsDependentUpon**( ** _ObjectType_**, ** _ObjectName_** )

 _expression_ A variable that represents an **AccessObject** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Required|**AcObjectType**|An  **[AcObjectType](acobjecttype-enumeration-access.md)** constant that represents the type of database object to check for dependency.|
| _ObjectName_|Required|**String**|The name of the database object to check for dependency.|

### Return Value

Boolean


## Remarks

This method will return a run-time error if any of the following conditions are true:


- The  **Track name AutoCorrect info** setting ( **Tools** menu, **Options** dialog box, **General** tab) is disabled. You can use the following code to enable the **Track name AutoCorrect info** setting and update the dependency information for all of the objects in the database: `Application.SetOption "Track Name AutoCorrect Info", 1`
    
- You have insufficient permissions to check the dependency information for the specified  **AccessObject** object.
    
- This method is being called from an Access project (.adp).
    


Access does not search Visual Basic for Applications (VBA) code, macros, or data access pages for dependencies.


## See also


#### Concepts


[AccessObject Object](accessobject-object-access.md)

