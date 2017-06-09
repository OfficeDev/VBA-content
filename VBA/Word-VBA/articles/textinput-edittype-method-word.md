---
title: TextInput.EditType Method (Word)
keywords: vbawd10.chm153550950
f1_keywords:
- vbawd10.chm153550950
ms.prod: word
api_name:
- Word.TextInput.EditType
ms.assetid: edd9efba-ca77-3f2f-021e-89e86ac9efc8
ms.date: 06/08/2017
---


# TextInput.EditType Method (Word)

Sets options for the specified text form field.


## Syntax

 _expression_ . **EditType**( **_Type_** , **_Default_** , **_Format_** , **_Enabled_** )

 _expression_ Required. A variable that represents a **[TextInput](textinput-object-word.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Type_|Required| **WdTextFormFieldType**|The text box type.|
| _Default_|Optional| **Variant**|The default text that appears in the text box.|
| _Format_|Optional| **Variant**|The formatting string used to format the text, number, or date (for example, "0.00," "Title Case," or "M/d/yy"). For more examples of formats, see the list of formats for the specified text form field type in the  **Text Form Field Options** dialog box.|
| _Enabled_|Optional| **Variant**| **True** to enable the form field for text entry.|

## Example

This example adds a text form field named "Today" at the beginning of the active document. The  **EditType** method is used to set the type to **wdCurrentDateText** and set the date format to "M/d/yy."


```vb
With ActiveDocument.FormFields.Add _ 
 (Range:=ActiveDocument.Range(0, 0), _ 
 Type:=wdFieldFormTextInput) 
 .Name = "Today" 
 .TextInput.EditType Type:=wdCurrentDateText, _ 
 Format:="M/d/yy", Enabled:=False 
End With
```


## See also


#### Concepts


[TextInput Object](textinput-object-word.md)

