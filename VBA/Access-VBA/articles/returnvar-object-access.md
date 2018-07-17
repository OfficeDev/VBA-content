---
title: ReturnVar Object (Access)
keywords: vbaac10.chm14689
f1_keywords:
- vbaac10.chm14689
ms.prod: access
api_name:
- Access.ReturnVar
ms.assetid: 8ad5254d-a249-46ba-ac5d-14943179ce05
ms.date: 06/08/2017
---


# ReturnVar Object (Access)

Represents a variable that was initialized by the  **SetReturnVar** function in a Data Macro.


## Remarks

A  **ReturnVar** object provides a convenient way to use values set in a Data Macro.

Although a  **ReturnVar** object can be used to store information for use in VBA procedures, it does not have the same functionality as a VBA variable.


- By default, a  **ReturnVar** object remains in memory until the next time that the **[RunDataMacro](docmd-rundatamacro-method-access.md)** method is used.
    
    Use the Value poperty of the ReturnVar
    
- A  **ReturnVar** object can store only text or numeric data. **ReturnVar** objects cannot store objects.
    
To refer to a TempVar object in a collection by its ordinal number or by its Name property setting, use the following syntax form:




```
ReturnVars![name] 

```


## Properties



|**Name**|
|:-----|
|[Name](returnvar-name-property-access.md)|
|[Value](returnvar-value-property-access.md)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
