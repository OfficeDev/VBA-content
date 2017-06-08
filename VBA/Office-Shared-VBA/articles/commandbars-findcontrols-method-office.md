---
title: CommandBars.FindControls Method (Office)
keywords: vbaof11.chm2014
f1_keywords:
- vbaof11.chm2014
ms.prod: office
api_name:
- Office.CommandBars.FindControls
ms.assetid: 79c46884-816d-def6-2bff-85b59b0831ea
ms.date: 06/08/2017
---


# CommandBars.FindControls Method (Office)

Gets the  **CommandBarControls** collection that fits the specified criteria.


## 


 **Note**  The use of CommandBars in some Microsoft Office applications has been superseded by the new ribbon component of the Microsoft Office Fluent user interface. For more information, search Help for the keyword "ribbon."


## Syntax

 _expression_. **FindControls**( **_Type_**, **_Id_**, **_Tag_**, **_Visible_** )

 _expression_ A variable that represents a **CommandBars** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Optional|**Variant**|Is one of the  **MsoControlType** constants specfying the type of control.|
| _Id_|Optional|**Variant**|The control's identifier.|
| _Tag_|Optional|**Variant**|The control's tag value.|
| _Visible_|Optional|**Variant**|**True** to include only visible command bar controls in the search. The default value is False.|

### Return Value

CommandBarControls


## Remarks

If no controls that fits the criteria are found, the  **FindControls** method returns **Nothing**.


## Example

This example uses the FindControls method to return all members of the CommandBars collection that have an ID of 18 and displays (in a message box) the number of controls that meet the search criteria.


```
Dim myControls As CommandBarControls 
Set myControls = CommandBars.FindControls(Type:=msoControlButton, ID:=18) 
MsgBox "There are " &amp; myControls.Count &amp; _ 
    " controls that meet the search criteria."
```


## See also


#### Concepts


[CommandBars Object](commandbars-object-office.md)
#### Other resources


[CommandBars Object Members](commandbars-members-office.md)

