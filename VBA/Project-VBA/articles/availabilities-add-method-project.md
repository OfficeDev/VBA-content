---
title: Availabilities.Add Method (Project)
ms.prod: project-server
api_name:
- Project.Availabilities.Add
ms.assetid: 4506674e-947b-905b-93bd-73a58281d676
ms.date: 06/08/2017
---


# Availabilities.Add Method (Project)

Adds an  **Availability** object to an **Availabilities** collection.


## Syntax

 _expression_. **Add**( ** _AvailableFrom_**, ** _AvailableTo_**, ** _AvailableUnit_** )

 _expression_ A variable that represents an **Availabilities** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _AvailableFrom_|Required|**Variant**|The earliest date the resource is available for work on the project.|
| _AvailableTo_|Required|**Variant**| The latest date the resource is available for work on the project.|
| _AvailableUnit_|Required|**Double**|The unit value for the availability period.|

### Return Value

 **Availability**


## See also


#### Concepts


[Availabilities Collection Object](availabilities-object-project.md)
