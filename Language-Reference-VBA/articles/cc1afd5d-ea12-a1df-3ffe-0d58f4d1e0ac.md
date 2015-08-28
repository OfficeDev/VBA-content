
# Mod Operator

 **Last modified:** July 28, 2015


Used to divide two numbers and return only the remainder.
 **Syntax**
 _result_**=**_number1_**Mod**_number2_
The  **Mod** operator syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _result_|Required; any numeric  [variable](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).|
| _number1_|Required; any  [numeric expression](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md).|
| _number2_|Required; any numeric expression.|
 **Remarks**
The modulus, or remainder, operator divides  _number1_ by _number2_ (rounding floating-point numbers to integers) and returns only the remainder as _result_. For example, in the following  [expression](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), A ( _result_) equals 5.
Usually, the  [data type](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) of _result_ is a [Byte](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md),  **Byte** variant, [Integer](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md),  **Integer** variant, [Long](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md), or  [Variant](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) containing a **Long**, regardless of whether or not  _result_ is a whole number. Any fractional portion is truncated. However, if any expression is [Null](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md),  _result_ is **Null**. Any expression that is  [Empty](b8bdf64f-5920-1ae9-16d0-b26d09524a30.md) is treated as 0.

## Example

This example uses the  **Mod** operator to divide two numbers and return only the remainder. If either number is a floating-point number, it is first rounded to an integer.


```
Dim MyResult
MyResult = 10 Mod 5    ' Returns 0.
MyResult = 10 Mod 3    ' Returns 1.
MyResult = 12 Mod 4.3    ' Returns 0.
MyResult = 12.6 Mod 5    ' Returns 3.
```

