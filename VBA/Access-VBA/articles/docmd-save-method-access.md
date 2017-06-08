---
title: DoCmd.Save Method (Access)
keywords: vbaac10.chm4177
f1_keywords:
- vbaac10.chm4177
ms.prod: access
api_name:
- Access.DoCmd.Save
ms.assetid: 7e01f370-36c9-9f4d-b506-61bc8886ee18
ms.date: 06/08/2017
---


# DoCmd.Save Method (Access)

The  **Save** method carries out the Save action in Visual Basic.


## Syntax

 _expression_. **Save**( ** _ObjectType_**, ** _ObjectName_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Optional|**AcObjectType**|A  **[AcObjectType](acobjecttype-enumeration-access.md)** constant that specifies the type of object you want to save.|
| _ObjectName_|Optional|**Variant**|A string expression that's the valid name of an object of the type selected by the  _objecttype_ argument.|

## Remarks

The  **Save** method works on all database objects that the user can explicitly open and save. The specified object must be open for the **Save** method to have any effect on the object.

If you leave the  _objecttype_ and _objectname_ arguments blank (the default constant, **acDefault**, is assumed for the _objecttype_ argument), Microsoft Access saves the active object. If you leave the _objecttype_ argument blank, but enter a name in the _objectname_ argument, Microsoft Access saves the active object with the specified name. If you enter an object type in the _objecttype_ argument, you must enter an existing object's name in the _objectname_ argument.


 **Note**  You can't use the  **Save** method to save any of the following with a new name:


- A form in Form view or Datasheet view.
    
- A report in Print Preview.
    
- A module.
    
- A server view in Datasheet view or Print Preview.
    
- A table in Datasheet view or Print Preview.
    
- A query in Datasheet view or Print Preview.
    
- A stored procedure in Datasheet view or Print Preview.
    

## Example

The following example uses the  **Save** method to save the form named "New Employees Form". This form must be open when the code containing this method runs.


```vb
DoCmd.Save acForm, "New Employees Form"
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

