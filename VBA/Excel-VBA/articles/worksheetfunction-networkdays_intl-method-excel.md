---
title: WorksheetFunction.NetworkDays_Intl Method (Excel)
keywords: vbaxl10.chm137391
f1_keywords:
- vbaxl10.chm137391
ms.prod: excel
api_name:
- Excel.WorksheetFunction.NetworkDays_Intl
ms.assetid: 04f1b585-396c-f981-9491-70d1b7948e6e
ms.date: 06/08/2017
---


# WorksheetFunction.NetworkDays_Intl Method (Excel)

Returns the number of whole workdays between two dates using parameters to indicate which and how many days are weekend days. Weekend days and any days that are specified as holidays are not considered as workdays.


## Syntax

 _expression_ . **NetworkDays_Intl**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters

|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Start_date - The start date for which the difference is to be computed. The start_date can be earlier than, the same as, or later than the end_date.|
| _Arg2_|Required| **Variant**|End_date - The end date for which the difference is to be computed. The start_date can be earlier than, the same as, or later than the end_date.|
| _Arg3_|Optional| **Variant**|Weekend - Indicates the days of the week that are weekend days and are not included in the number of whole working days between start_date and end_date. Weekend is a weekend number or string that specifies when weekends occur. Weekend number values indicate the weekend days listed in the following table.<table><tr><th>**Weekend number**</th><th>**Weekend days**</th> </tr><tr><td>1 or omitted</td><td>Saturday, Sunday</td></tr><tr><td>2</td><td>Sunday, Monday</td></tr><tr><td>3</td><td>Monday, Tuesday </td> </tr><tr><td>4</td><td>Tuesday, Wednesday</td> </tr><tr><td>5</td><td>Wednesday, Thursday</td> </tr><tr><td>6</td><td>Thursday, Friday</td></tr><tr><td>7</td><td>Friday, Saturday</td></tr><tr><td>11</td><td>Sunday only</td></tr><tr><td>12</td><td>Monday only</td></tr><tr><td>13</td><td>Tuesday only</td> </tr><tr><td>14</td><td>Wednesday only</td></tr><tr><td>15</td><td>Thursday only</td> </tr><tr><td>16</td><td>Friday only</td></tr><tr><td>17</td><td>Saturday only</td> </tr></table>|
| _Arg4_|Optional| **Variant**|Holidays - An optional set of one or more dates that are to be excluded from the working day calendar. Holidays is a range of cells that contain the dates, or an array constant of the serial values that represent those dates. The ordering of dates or serial values in holidays can be arbitrary.|

### Return Value

Double


## Remarks




- If start_date is later than end_date, the return value will be negative, and the magnitude will be the number of whole workdays.
    
-  If start_date is out of range for the current date base value, NETWORKDAYS_INTL returns the #NUM! error value.
    
- If end_date is out of range for the current date base value, NETWORKDAYS_INTL returns the #NUM! error value.
    
- If a weekend string is of invalid length or contains invalid characters, NETWORKDAYS_INTL returns the #VALUE! error value.
    

## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

