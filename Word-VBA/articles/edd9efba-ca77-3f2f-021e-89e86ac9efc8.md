
# TextInput.EditType Method (Word)

 **Last modified:** July 28, 2015

Sets options for the specified text form field.

## Syntax

 _expression_. **EditType**( **_Type_**,  **_Default_**,  **_Format_**,  **_Enabled_**)

 _expression_Required. A variable that represents a  ** [TextInput](d7f6531a-4da2-ccc4-29b3-ad79ca7b18de.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Type|Required| **WdTextFormFieldType**|The text box type.|
|Default|Optional| **Variant**|The default text that appears in the text box.|
|Format|Optional| **Variant**|The formatting string used to format the text, number, or date (for example, "0.00," "Title Case," or "M/d/yy"). For more examples of formats, see the list of formats for the specified text form field type in the  **Text Form Field Options** dialog box.|
|Enabled|Optional| **Variant**| **True** to enable the form field for text entry.|

## Example

This example adds a text form field named "Today" at the beginning of the active document. The  **EditType** method is used to set the type to **wdCurrentDateText** and set the date format to "M/d/yy."


```
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


 [TextInput Object](d7f6531a-4da2-ccc4-29b3-ad79ca7b18de.md)
#### Other resources


 [TextInput Object Members](d21b3150-6a32-3212-d144-9fc72a866187.md)
