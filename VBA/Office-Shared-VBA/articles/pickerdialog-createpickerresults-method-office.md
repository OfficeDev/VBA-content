---
title: PickerDialog.CreatePickerResults Method (Office)
keywords: vbaof11.chm340004
f1_keywords:
- vbaof11.chm340004
ms.prod: office
api_name:
- Office.PickerDialog.CreatePickerResults
ms.assetid: 39954f3e-53ef-f33c-9e90-a2247fd7882a
ms.date: 06/08/2017
---


# PickerDialog.CreatePickerResults Method (Office)

Creates an empty  **PickerResults** object.


## Syntax

 _expression_. **CreatePickerResults**

 _expression_ An expression that returns a **PickerDialog** object.


### Return Value

PickerResults


## Remarks

 You can add the PickerResult to the returned object and specify it to the second parameter of the **Show** method as already existing results of the **PickerDialog** object.


## Example

The following code sets various properties of the Picker Dialog and adds the already existing PickerResults to the results.


```
Dim objPickerDialog As PickerDialog 
Dim objPickerExistingResults As PickerResults 
 
Set objPickerDialog = Application.PickerDialog 
objPickerDialog.DataHandlerId = "{000CDF0A-0000-0000-C000-000000000046}" 
objPickerDialog.Title = "Sample Picker Dialog" 
 
Set objPickerExistingResults = objPickerDialog.CreatePickerResults 
Set objPickerExistingResult = objPickerExistingResults.Add("johndoe@contoso.com", "John Doe", "User") 
Set objPickerResults = objPickerDialog.Show(True, objPickerExistingResult) 

```


## See also


#### Concepts


[PickerDialog Object](pickerdialog-object-office.md)
#### Other resources


[PickerDialog Object Members](pickerdialog-members-office.md)

