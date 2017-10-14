---
title: DoCmd.GoToRecord Method (Access)
keywords: vbaac10.chm4154
f1_keywords:
- vbaac10.chm4154
ms.prod: access
api_name:
- Access.DoCmd.GoToRecord
ms.assetid: 5494b6fc-112f-e944-9072-873b00271ab1
ms.date: 06/08/2017
---


# DoCmd.GoToRecord Method (Access)

The  **GoToRecord** method carries out the GoToRecord action in Visual Basic.


## Syntax

 _expression_. **GoToRecord**( ** _ObjectType_**, ** _ObjectName_**, ** _Record_**, ** _Offset_** )

 _expression_ A variable that represents a **DoCmd** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _ObjectType_|Optional|**AcDataObjectType**|A  **AcDataObjectType** constant that specifies the type of object that contains the record you want to make current.|
| _ObjectName_|Optional|**Variant**|A string expression that's the valid name of an object of the type selected by the  _objecttype_ argument.|
| _Record_|Optional|**AcRecord**|A [AcRecord](acrecord-enumeration-access.md) constant that specifies the record to make the current record. The default value is **acNext**.|
| _Offset_|Optional|**Variant**|A numeric expression that represents the number of records to move forward or backward if you specify  **acNext** or **acPrevious** for the _record_ argument, or the record to move to if you specify **acGoTo** for the _record_ argument. The expression must result in a valid record number.|

## Remarks

You can use the  **GoToRecord** method to make the specified record the current record in an open table, form, or query result set.

If you leave the  _objecttype_ and _objectname_ arguments blank (the default constant, **acActiveDataObject**, is assumed for _objecttype_), the active object is assumed.

You can use the  **GoToRecord** method to make a record on a hidden form the current record if you specify the hidden form in the _objecttype_ and _objectname_ arguments.


## Example

The following example uses the  **GoToRecord** method to make the seventh record in the form Employees current:


```vb
DoCmd.GoToRecord acDataForm, "Employees", acGoTo, 7
```


## See also


#### Concepts


[DoCmd Object](docmd-object-access.md)

