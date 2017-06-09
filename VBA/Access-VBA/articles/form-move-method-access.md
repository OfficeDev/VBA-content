---
title: Form.Move Method (Access)
keywords: vbaac10.chm13518
f1_keywords:
- vbaac10.chm13518
ms.prod: access
api_name:
- Access.Form.Move
ms.assetid: 21529c39-70be-45ab-fe8a-b54b4f78b4c8
ms.date: 06/08/2017
---


# Form.Move Method (Access)

Moves the specified object to the coordinates specified by the argument values.


## Syntax

 _expression_. **Move**( ** _Left_**, ** _Top_**, ** _Width_**, ** _Height_** )

 _expression_ A variable that represents a **Form** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Left_|Required|**Variant**|The screen position in twips for the left edge of the object relative to the left edge of the Microsoft Access window.|
| _Top_|Optional|**Variant**|The screen position in twips for the top edge of the object relative to the top edge of the Microsoft Access window.|
| _Width_|Optional|**Variant**|The desired width in twips of the object.|
| _Height_|Optional|**Variant**|The desired height in twips of the object.|

## Remarks

Only the  _Left_ argument is required. However, to specify any other arguments, you must specify all the arguments that precede it. For example, you cannot specify _Width_ without specifying _Left_ and _Top_. Any trailing arguments that are unspecified remain unchanged.

This method overrides the  **Moveable** property.

If a form is modal, it is still positioned relative to the Access window, but the values for  _Left_ and _Top_ can be negative.

In Datasheet View or Print Preview, changes made using the  **Move** method are saved if the user explicitly saves the database, but Access does not prompt the user to save such changes.


## Example

The following example determines whether or not the first form in the current project can be moved; if it can, the example moves the form.


```vb
If Forms(0).Moveable Then 
    Forms(0).Move _ 
       Left:=0, Top:=0, Width:=400, Height:=300 
Else 
    MsgBox "The form cannot be moved." 
End If
```


## See also


#### Concepts


[Form Object](form-object-access.md)

