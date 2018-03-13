---
title: For...Next Statement
keywords: vblr6.chm1008924
f1_keywords:
- vblr6.chm1008924
ms.prod: office
ms.assetid: 53e92bd3-1933-5bc7-f7a4-4e6a3d9bef4a
ms.date: 06/08/2017
---


# For...Next Statement

Repeats a group of [statements](vbe-glossary.md) a specified number of times.

 **Syntax**

 **For**_counter_**=**_start_**To**_end_ [ **Step**_step_ ]
[ _statements_ ]
[ **Exit For** ]
[ _statements_ ]

 **Next** [ _counter_ ]
The  **For…Next** statement syntax has these parts:


| <strong>Part</strong> | <strong>Description</strong>                                                                                                                                   |
|:----------------------|:---------------------------------------------------------------------------------------------------------------------------------------------------------------|
| <em>counter</em>      | Required. Numeric [variable](vbe-glossary.md) used as a loop counter. The variable can't be a[Boolean](vbe-glossary.md) or an[array](vbe-glossary.md) element. |
| <em>start</em>        | Required. Initial value of  <em>counter</em>.                                                                                                                  |
| <em>end</em>          | Required. Final value of  <em>counter</em>.                                                                                                                    |
| <em>step</em>         | Optional. Amount  <em>counter</em> is changed each time through the loop. If not specified, <em>step</em> defaults to one.                                     |
| <em>statements</em>   | Optional. One or more statements between  <strong>For</strong> and <strong>Next</strong> that are executed the specified number of times.                      |

 **Remarks**
The  _step_[argument](vbe-glossary.md) can be either positive or negative. The value of the _step_ argument determines loop processing as follows:


| <strong>Value</strong> | <strong>Loop executes if</strong> |
|:-----------------------|:----------------------------------|
| Positive or 0          | <em>counter</em> <= <em>end</em>  |
| Negative               | <em>counter</em> >= <em>end</em>  |

After all statements in the loop have executed,  _step_ is added to _counter_. At this point, either the statements in the loop execute again (based on the same test that caused the loop to execute initially), or the loop is exited and execution continues with the statement following the **Next** statement.

 **Tip**  Changing the value of  _counter_ while inside a loop can make it more difficult to read and debug your code.

Any number of  **Exit For** statements may be placed anywhere in the loop as an alternate way to exit. **Exit For** is often used after evaluating of some condition, for example **If...Then**, and transfers control to the statement immediately following **Next**.
You can nest  **For...Next** loops by placing one **For...Next** loop within another. Give each loop a unique variable name as its _counter_. The following construction is correct:



```
For I = 1 To 10 
 For J = 1 To 10 
 For K = 1 To 10 
 ... 
 Next K 
 Next J 
Next I 
```


 **Note**  If you omit  _counter_ in a **Next** statement, execution continues as if _counter_ is included. If a **Next** statement is encountered before its corresponding **For** statement, an error occurs.


## Example

This example uses the  **For...Next** statement to create a string that contains 10 instances of the numbers 0 through 9, each string separated from the other by a single space. The outer loop uses a loop counter variable that is decremented each time through the loop.


```vb
Dim Words, Chars, MyString 
For Words = 10 To 1 Step -1 ' Set up 10 repetitions. 
 For Chars = 0 To 9 ' Set up 10 repetitions. 
 MyString = MyString &; Chars ' Append number to string. 
 Next Chars ' Increment counter 
 MyString = MyString &; " " ' Append a space. 
Next Words 
```


