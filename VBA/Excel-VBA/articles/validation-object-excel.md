---
title: Validation Object (Excel)
keywords: vbaxl10.chm531072
f1_keywords:
- vbaxl10.chm531072
ms.prod: excel
api_name:
- Excel.Validation
ms.assetid: 59d29d1e-92d3-373e-04d0-0d7fe97e1878
ms.date: 06/08/2017
---


# Validation Object (Excel)

Represents data validation for a worksheet range.


## Example

Use the  **[Validation](range-validation-property-excel.md)** property to return the **Validation** object. The following example changes the data validation for cell E5.


```
Range("e5").Validation _ 
 .Modify xlValidateList, xlValidAlertStop, "=$A$1:$A$10"
```

Use the  **[Add](validation-add-method-excel.md)** method to add data validation to a range and create a new **Validation** object. The following example adds data validation to cell E5.




```
With Range("e5").Validation 
 .Add Type:=xlValidateWholeNumber, _ 
 AlertStyle:=xlValidAlertInformation, _ 
 Minimum:="5", Maximum:="10" 
 .InputTitle = "Integers" 
 .ErrorTitle = "Integers" 
 .InputMessage = "Enter an integer from five to ten" 
 .ErrorMessage = "You must enter a number from five to ten" 
End With 

```


## Methods



|**Name**|
|:-----|
|[Add](validation-add-method-excel.md)|
|[Delete](validation-delete-method-excel.md)|
|[Modify](validation-modify-method-excel.md)|

## Properties



|**Name**|
|:-----|
|[AlertStyle](validation-alertstyle-property-excel.md)|
|[Application](validation-application-property-excel.md)|
|[Creator](validation-creator-property-excel.md)|
|[ErrorMessage](validation-errormessage-property-excel.md)|
|[ErrorTitle](validation-errortitle-property-excel.md)|
|[Formula1](validation-formula1-property-excel.md)|
|[Formula2](validation-formula2-property-excel.md)|
|[IgnoreBlank](validation-ignoreblank-property-excel.md)|
|[IMEMode](validation-imemode-property-excel.md)|
|[InCellDropdown](validation-incelldropdown-property-excel.md)|
|[InputMessage](validation-inputmessage-property-excel.md)|
|[InputTitle](validation-inputtitle-property-excel.md)|
|[Operator](validation-operator-property-excel.md)|
|[Parent](validation-parent-property-excel.md)|
|[ShowError](validation-showerror-property-excel.md)|
|[ShowInput](validation-showinput-property-excel.md)|
|[Type](validation-type-property-excel.md)|
|[Value](validation-value-property-excel.md)|

## See also


#### Other resources


[Excel Object Model Reference](http://msdn.microsoft.com/library/11ea8598-8a20-92d5-f98b-0da04263bf2c%28Office.15%29.aspx)
