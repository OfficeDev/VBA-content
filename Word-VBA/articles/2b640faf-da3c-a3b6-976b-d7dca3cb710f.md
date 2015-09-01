
# Range.ConvertHangulAndHanja Method (Word)

 **Last modified:** July 28, 2015

Converts the specified range from hangul to hanja or vice versa.

## Syntax

 _expression_. **ConvertHangulAndHanja**( **_ConversionsMode_**,  **_FastConversion_**,  **_CheckHangulEnding_**,  **_EnableRecentOrdering_**,  **_CustomDictionary_**)

 _expression_Required. A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|ConversionsMode|Optional| **Variant**|Sets the direction for the conversion between hangul and hanja. Can be either of the following  **WdMultipleWordConversionsMode** constants: **wdHangulToHanja** or **wdHanjaToHangul**. The default value is the current value of the  **MultipleWordConversionsMode** property.|
|FastConversion|Optional| **Variant**| **True** if Microsoft Word automatically converts a word with only one suggestion for conversion. The default value is the current value of the **HangulHanjaFastConversion** property.|
|CheckHangulEnding|Optional| **Variant**| **True** if Word automatically detects hangul endings and ignores them. The default value is the current value of the **CheckHangulEndings** property. This argument is ignored if the ConversionsMode argument is set to **wdHanjaToHangul**.|
|EnableRecentOrdering|Optional| **Variant**| **True** if Word displays the most recently used words at the top of the suggestions list. The default value is the current value of the **EnableHangulHanjaRecentOrdering** property.|
|CustomDictionary|Optional| **Variant**|The name of a custom hangul-hanja conversion dictionary. Use this argument to use a custom dictionary with hangul-hanja conversions not contained in the main dictionary.|

## Example

This example converts the current selection from hangul to hanja.


```
Selection.Range.ConvertHangulAndHanja _ 
 ConversionsMode:=wdHangulToHanja, _ 
 FastConversion:=True, _ 
 EnableRecentOrdering:= True
```


## See also


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
