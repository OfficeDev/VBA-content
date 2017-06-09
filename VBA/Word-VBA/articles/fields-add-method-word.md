---
title: Fields.Add Method (Word)
keywords: vbawd10.chm154140777
f1_keywords:
- vbawd10.chm154140777
ms.prod: word
api_name:
- Word.Fields.Add
ms.assetid: e4633cf9-394c-5af1-1a3f-02e3387ae8a1
ms.date: 06/08/2017
---


# Fields.Add Method (Word)

Adds a  **[Field](field-object-word.md)** object to the **Fields** collection. Returns the **[Field](field-object-word.md)** object at the specified range.


## Syntax

 _expression_ . **Add**( **_Range_** , **_Type_** , **_Text_** , **_PreserveFormatting_** )

 _expression_ Required. A variable that represents a **[Fields](fields-object-word.md)** collection.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Range_|Required| **Range object**|The range where you want to add the field. If the range isn't collapsed, the field replaces the range.|
| _Type_|Optional| **Variant**|Can be any  **WdFieldType** constant. For a list of valid constants, consult the Object Browser. The default value is **wdFieldEmpty** .|
| _Text_|Optional| **Variant**|Additional text needed for the field. For example, if you want to specify a switch for the field, you would add it here.|
| _PreserveFormatting_|Optional| **Variant**| **True** to have the formatting that's applied to the field preserved during updates.|

### Return Value

Field


## Remarks

You cannot insert some fields (such as  **wdFieldOCX** and **wdFieldFormCheckBox** ) by using the **Add** method of the **Fields** collection. Instead, you must use specific methods such as the **AddOLEControl** method and the **Add** method for the **FormFields** collection.


## Example

This example inserts a USERNAME field at the beginning of the selection.


```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myField = ActiveDocument.Fields.Add(Range:=Selection.Range, _ 
 Type:=wdFieldUserName)
```

This example inserts a LISTNUM field at the end of the selection. The starting switch is set to begin at 3.




```
Selection.Collapse Direction:=wdCollapseEnd 
ActiveDocument.Fields.Add Range:=Selection.Range, _ 
 Type:=wdFieldListNum, Text:="\s 3"
```

This example inserts a DATE field at the beginning of the selection and then displays the result.




```vb
Selection.Collapse Direction:=wdCollapseStart 
Set myField = ActiveDocument.Fields.Add(Range:=Selection.Range, _ 
 Type:=wdFieldDate) 
MsgBox myField.Result
```


## See also


#### Concepts


[Fields Collection Object](fields-object-word.md)

