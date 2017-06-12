---
title: Application.DurationValue Method (Project)
ms.prod: project-server
api_name:
- Project.Application.DurationValue
ms.assetid: 745acbd3-600c-1179-1d61-be0dab88cdf5
ms.date: 06/08/2017
---


# Application.DurationValue Method (Project)

Returns the number of minutes in a duration.


## Syntax

 _expression_. **DurationValue**( ** _Duration_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Duration_|Required|**String**|The duration to be expressed in minutes.|

### Return Value

 **Variant**


## Example

The following example adds the entered value to the duration of the selected task.


```vb
Sub DurationAdder() 
 
 Dim Temp As String 
 
 Temp = InputBox$("Enter amount by which to increase the duration:") 
 ActiveCell.Task.Duration = ActiveCell.Task.Duration + DurationValue(Temp) 
 
End Sub
```


