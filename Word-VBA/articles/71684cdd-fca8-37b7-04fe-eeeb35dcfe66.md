
# Range.TCSCConverter Method (Word)

 **Last modified:** July 28, 2015

Converts the specified range from Traditional Chinese to Simplified Chinese or vice versa.

## Syntax

 _expression_. **TCSCConverter**( **_WdTCSCConverterDirection_**,  **_CommonTerms_**,  **_UseVariants_**)

 _expression_Required. A variable that represents a  ** [Range](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|WdTCSCConverterDirection|Optional| **WdTCSCConverterDirection**|Specifies the direction in which text is converted. If omitted, the default value is  **wdTCSCConverterDirectionAuto**, which converts in the appropriate direction based on the detected language of the specified range.|
|UseVariants|Optional| **Boolean**| **True** if Word uses Taiwan, Hong Kong SAR, and Macao SAR character variants. Can only be used if translating from Simplified Chinese to Traditional Chinese.|

## Example

This example converts the current selection from Simplified Chinese to Traditional Chinese. It converts common expressions intact and uses regional character variants.


```
Selection.Range.TCSCConverter _ 
 wdTCSCConverterDirectionSCTC, True, True
```


## See also


#### Concepts


 [Range Object](15a7a1c4-5f3f-5b6e-60e9-29688de3f274.md)
#### Other resources


 [Range Object Members](3c4a36d9-2a80-5aaf-827b-275a52bfa193.md)
