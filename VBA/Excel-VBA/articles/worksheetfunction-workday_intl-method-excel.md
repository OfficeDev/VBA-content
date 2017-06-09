---
title: WorksheetFunction.WorkDay_Intl Method (Excel)
keywords: vbaxl10.chm137392
f1_keywords:
- vbaxl10.chm137392
ms.prod: excel
api_name:
- Excel.WorksheetFunction.WorkDay_Intl
ms.assetid: 0a9091a1-c6d4-06c4-a00d-7477474bddf0
ms.date: 06/08/2017
---


# WorksheetFunction.WorkDay_Intl Method (Excel)

Returns the serial number of the date before or after a specified number of workdays with custom weekend parameters. Weekend parameters indicate which and how many days are weekend days. Weekend days and any days that are specified as holidays are not considered as workdays.


## Syntax

 _expression_ . **WorkDay_Intl**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Variant**|Start_date - The start date, truncated to integer.|
| _Arg2_|Required| **Variant**|Days - The number of workdays before or after the start_date. A positive value yields a future date; a negative value yields a past date; a 0 (zero) value yields the start_date. Day-offset is truncated to an integer.|
| _Arg3_|Optional| **Variant**|Weekend - Indicates the days of the week that are weekend days and are not considered working days. Weekend is a weekend number or string that specifies when weekends occur. Weekend number values indicate the following weekend days.<table><tr><th>**Weekend number**</th><th>**Weekend days**</th> </tr><tr><td>1 or omitted</td><td>Saturday, Sunday</td></tr><tr><td>2</td><td>Sunday, Monday</td></tr><tr><td>3</td><td>Monday, Tuesday </td> </tr><tr><td>4</td><td>Tuesday, Wednesday</td> </tr><tr><td>5</td><td>Wednesday, Thursday</td> </tr><tr><td>6</td><td>Thursday, Friday</td></tr><tr><td>7</td><td>Friday, Saturday</td></tr><tr><td>11</td><td>Sunday only</td></tr><tr><td>12</td><td>Monday only</td></tr><tr><td>13</td><td>Tuesday only</td> </tr><tr><td>14</td><td>Wednesday only</td></tr><tr><td>15</td><td>Thursday only</td> </tr><tr><td>16</td><td>Friday only</td></tr><tr><td>17</td><td>Saturday only</td> </tr></table>|
| _Arg4_|Optional| **Variant**|Holidays - An optional set of one or more dates that are to be excluded from the working day calendar. Holidays is a range of cells that contain the dates, or an array constant of the serial values that represent those dates. The ordering of dates or serial values in holidays can be arbitrary.|

### Return Value

Double


## Remarks




- If start_date is out of range for the current date base value, WORKDAY_INTL returns the #NUM! error value.
    
- If any date in holidays is out of range for the current date base value, WORKDAY_INTL returns the #NUM! error value.
    
- If start_date plus day-offset yields an invalid date, WORKDAY_INTL returns the #NUM! error value.
    
- If a weekend string is of invalid length or contains invalid characters, WORKDAY_INTL returns the #VALUE! error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

