---
title: WorksheetFunction.Forecast_Linear Method (Excel)
keywords: vbaxl10.chm137471
f1_keywords:
- vbaxl10.chm137471
ms.assetid: 71b85d12-0c81-f82d-99fe-ad712f2530e5
ms.date: 06/08/2017
ms.prod: excel
---


# WorksheetFunction.Forecast_Linear Method (Excel)

Calculates, or predicts, a future value by using existing values. The predicted value is a y-value for a given x-value. The known values are existing x-values and y-values, and the new value is predicted by using linear regression. You can use this function to predict future sales, inventory requirements, or consumer trends.


## Syntax

 _expression_ . **Forecast_Linear**( **_Arg1_** , **_Arg2_** , **_Arg3_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|x - the data point for which you want to predict a value.|
| _Arg2_|Required| **Variant**|known_y's - the dependent array or range of data.|
| _Arg3_|Required| **Variant**|known_x's - the independent array or range of data.|
|Name|Required/Optional|Data type|Description|

### Return Value

Double


## Remarks


- If x is nonnumeric,  **Forecast_Linear** returns the #VALUE! error value.
    
- If  _known_y_ and _known_x_ parameters are empty or contain a different number of data points, **Forecast_Linear** returns the #N/A error value.
    
- If the variance of  **known_x** parameters equals zero, **Forecast_Linear** returns the #DIV/0! error value.
    
- The equation for  **Forecast_Linear** is a+bx, where:
![Formula](images/awfintc1_ZA06051174.gif)and: 
![Formula](images/awfintc2_ZA06051175.gif)and where x and y are the sample means AVERAGE (all  _known_x_ ) and AVERAGE(all _known_y_ ).
    

## Example


```vb
Dim instance As WorksheetFunction
Dim Arg1 As Double
Dim Arg2 As Object
Dim Arg3 As Object
Dim returnValue As Double

returnValue = instance.Forecast_Linear(Arg1, Arg2, Arg3)

```


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

