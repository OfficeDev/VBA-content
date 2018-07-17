---
title: WorksheetFunction.Aggregate Method (Excel)
keywords: vbaxl10.chm137358
f1_keywords:
- vbaxl10.chm137358
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Aggregate
ms.assetid: 261e51bf-44d4-900c-2a5d-c6612ec9f98c
ms.date: 06/08/2017
---


# WorksheetFunction.Aggregate Method (Excel)

Returns an aggregate in a list or database.


## Syntax

 _expression_ . **Aggregate**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** , **_Arg5_** , **_Arg6_** , **_Arg7_** , **_Arg8_** , **_Arg9_** , **_Arg10_** , **_Arg11_** , **_Arg12_** , **_Arg13_** , **_Arg14_** , **_Arg15_** , **_Arg16_** , **_Arg17_** , **_Arg18_** , **_Arg19_** , **_Arg20_** , **_Arg21_** , **_Arg22_** , **_Arg23_** , **_Arg24_** , **_Arg25_** , **_Arg26_** , **_Arg27_** , **_Arg28_** , **_Arg29_** , **_Arg30_** )

 _expression_ A variable that represents a **[WorksheetFunction](worksheetfunction-object-excel.md)** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **Double**|Function_num - A number from 1 to 19 that specifies which function to use.<table><tr><th>**Function_num**</th><th>**Function**</th></tr><tr><td>1</td><td>AVERAGE</td></tr><tr><td>2</td><td>COUNT</td></tr><tr><td>3</td><td>COUNTA</td></tr><tr><td>4</td><td>MAX</td></tr><tr><td>5</td><td>MIN</td></tr><tr><td>6</td><td>PRODUCT</td></tr><tr><td>7</td><td>STDEV.S</td></tr><tr><td>8</td><td>STDEV.P</td></tr><tr><td>9</td><td>SUM</td></tr><tr><td>10</td><td>VAR.S</td></tr><tr><td>11</td><td>VAR.P</td></tr><tr><td>2</td><td>MEDIAN</td></tr><tr><td>13</td><td>MODE.SNGL</td></tr><tr><td>14</td><td>LARGE</td></tr><tr><td>15</td><td>SMALL</td></tr><tr><td>16</td><td>PERCENTILE.INC </td></tr><tr><td>17</td><td>QUARTILE.INC</td></tr><tr><td>18</td><td>PERCENTILE.EXC</td></tr><tr><td>19</td><td>QUARTILE.EXC</td></tr></table>|
| _Arg2_|Required| **Double**|Options - A numerical value that determines which values to ignore in the evaluation range for the function.<table><tr><th>**Option**</th><th>**Behavior**</th> </tr><tr><td>0 or omitted</td><td>Ignore nested SUBTOTAL and AGGREGATE functions</td> </tr><tr><td>1</td><td>Ignore hidden rows, nested SUBTOTAL and AGGREGATE functions</td> </tr><tr><td>2</td><td>Ignore error values, nested SUBTOTAL and AGGREGATE functions</td> </tr><tr><td>3</td><td>Ignore hidden rows, error values, nested SUBTOTAL and AGGREGATE functions</td> </tr><tr><td>4</td><td>Ignore nothing</td> </tr><tr><td>5</td><td>Ignore hidden rows</td> </tr><tr><td>6</td><td>Ignore error values</td> </tr><tr><td>7</td><td>Ignore hidden rows and error values</td> </tr></table>|
| _Arg3_|Required| **Range**|Ref1 - The first numeric argument for functions that take multiple numeric arguments for which you want the aggregate value.|
| _Arg4 - Arg 30_|Optional| **Variant**|Ref2 - Ref30 - Numeric arguments 2 to 30 for which you want the aggregate value.|

### Return Value

Double


## Remarks

- The following constraints apply to the Ref arguments ( _Arg3 - Arg 30_ ) based on the **Function_num** value.
    

|**Function_num**|**Ref1**|**Ref2**|**Ref3, Ref4, ?**|
|:-----|:-----|:-----|:-----|
|1-13| **Valid types:**<ul><li><p>Any cell reference</p></li><li><p>Unions</p></li><li><p>Intersections</p></li><li><p>Defined names</p></li><li><p>Structured references</p></li></ul> **Invalid types:**<ul><li><p>Any cell reference</p></li><li><p>Unions</p></li><li><p>Intersections</p></li><li><p>Defined names</p></li><li><p>Structured references</p></li></ul><ul><li><p>Actual data</p></li><li><p>Arrays</p></li></ul> **Invalid types:**<ul><li><p>Actual data</p></li><li><p>Arrays</p></li></ul>| **Valid types:**<ul><li><p>Any cell reference</p></li><li><p>Unions</p></li><li><p>Intersections</p></li><li><p>Defined names</p></li><li><p>Structured references</p></li></ul> **Invalid types:**<ul><li><p>Any cell reference</p></li><li><p>Unions</p></li><li><p>Intersections</p></li><li><p>Defined names</p></li><li><p>Structured references</p></li></ul><ul><li><p>Actual data</p></li><li><p>Arrays</p></li></ul> **Invalid types:**<ul><li><p>Actual data</p></li><li><p>Arrays</p></li></ul>| **Valid types:**<ul><li><p>Any cell reference</p></li><li><p>Unions</p></li><li><p>Intersections</p></li><li><p>Defined names</p></li><li><p>Structured references</p></li></ul> **Invalid types:**<ul><li><p>Any cell reference</p></li><li><p>Unions</p></li><li><p>Intersections</p></li><li><p>Defined names</p></li><li><p>Structured references</p></li></ul><ul><li><p>Actual data</p></li><li><p>Arrays</p></li></ul> **Invalid types:**<ul><li><p>Actual data</p></li><li><p>Arrays</p></li></ul>|
|14-17| **Valid types:**<ul><li><p>Any cell reference</p></li><li><p>Unions</p></li><li><p>Intersections</p></li><li><p>Defined names</p></li><li><p>Structured references</p></li><li><p>Actual data</p></li><li><p>Arrays</p></li></ul>| **Valid types:**<ul><li><p>Any cell reference</p></li><li><p>Unions</p></li><li><p>Intersections</p></li><li><p>Defined names</p></li><li><p>Structured references</p></li><li><p>Actual data</p></li><li><p>Arrays</p></li></ul>| **No references are allowed**|
||

- If a second ref argument is required but not provided, AGGREGATE returns a #VALUE! error.
    
- If one or more of the references are 3-D references, AGGREGATE returns the #VALUE! error value.
    



## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

