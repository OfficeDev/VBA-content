---
title: TimeScaleValues.Add Method (Project)
ms.prod: project-server
api_name:
- Project.TimeScaleValues.Add
ms.assetid: 083ef154-31ce-55ec-793a-0627c1eff211
ms.date: 06/08/2017
---


# TimeScaleValues.Add Method (Project)

Adds a  **TimeScaleValue** object to a **TimeScaleValues** collection.


## Syntax

 _expression_. **Add**( ** _Value_**, ** _Position_** )

 _expression_ A variable that represents a **TimeScaleValues** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Value_|Required|**Variant**|The value of the timescaled data.|
| _Position_|Optional|**Variant**|The position of the new value. The default value is  **n + 1**, where **n** is the number of items in the collection. If the value specified for Position is **n + 2** or greater, the intervening items are given a value of 0.|

### Return Value

 **TimeScaleValue**


## See also


#### Concepts


[TimeScaleValues Collection Object](timescalevalues-object-project.md)
