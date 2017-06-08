---
title: Rnd Function
keywords: vblr6.chm1009008
f1_keywords:
- vblr6.chm1009008
ms.prod: office
ms.assetid: 57b9e8f9-6e3e-e68b-f5a4-c9c312b74426
ms.date: 06/08/2017
---


# Rnd Function



Returns a  **Single** containing a random number.
 **Syntax**
 **Rnd** [ **(**_number_**)** ]
The optional  _number_[argument](vbe-glossary.md) is a[Single](vbe-glossary.md) or any valid[numeric expression](vbe-glossary.md).
 **Return Values**


|**If  _number_ is**|**Rnd generates**|
|:-----|:-----|
|Less than zero|The same number every time, using  _number_ as the[seed](vbe-glossary.md).|
|Greater than zero|The next random number in the sequence.|
|Equal to zero|The most recently generated number.|
|Not supplied|The next random number in the sequence.|
 **Remarks**
The  **Rnd** function returns a value less than 1 but greater than or equal to zero.
The value of  _number_ determines how **Rnd** generates a random number:
For any given initial seed, the same number sequence is generated because each successive call to the  **Rnd** function uses the previous number as a seed for the next number in the sequence.
Before calling  **Rnd,** use the **Randomize** statement without an argument to initialize the random-number generator with a seed based on the system timer.
To produce random integers in a given range, use this formula:



```
Int((upperbound - lowerbound + 1) * Rnd + lowerbound)

```

Here,  _upperbound_ is the highest number in the range, and _lowerbound_ is the lowest number in the range.

 **Note**  To repeat sequences of random numbers, call  **Rnd** with a negative argument immediately before using **Randomize** with a numeric argument. Using **Randomize** with the same value for _number_ does not repeat the previous sequence.


## Example

This example uses the  **Rnd** function to generate a random integer value from 1 to 6.


```vb
Dim MyValue
MyValue = Int((6 * Rnd) + 1)    ' Generate random value between 1 and 6.


```


