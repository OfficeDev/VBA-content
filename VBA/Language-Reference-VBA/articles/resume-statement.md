---
title: Resume Statement
keywords: vblr6.chm1009004
f1_keywords:
- vblr6.chm1009004
ms.prod: office
ms.assetid: 57fa9eb3-7e8d-2f7e-20d7-47e468b7836a
ms.date: 06/08/2017
---


# Resume Statement

Resumes execution after an error-handling routine is finished.

 **Syntax**

 **Resume** [ **0** ]

 **Resume** **Next**
 **Resume**_line_
The  **Resume** statement syntax can have any of the following forms:


|**Statement**|**Description**|
|:-----|:-----|
|**Resume**|If the error occurred in the same [procedure](vbe-glossary.md) as the error handler, execution resumes with the statement that caused the error. If the error occurred in a called procedure, execution resumes at the[statement](vbe-glossary.md) that last called out of the procedure containing the error-handling routine.|
|**Resume** **Next**|If the error occurred in the same procedure as the error handler, execution resumes with the statement immediately following the statement that caused the error. If the error occurred in a called procedure, execution resumes with the statement immediately following the statement that last called out of the procedure containing the error-handling routine (or  **On Error Resume Next** statement).|
|**Resume**_line_|Execution resumes at  _line_ specified in the required _line_[argument](vbe-glossary.md). The  _line_ argument is a[line label](vbe-glossary.md) or[line number](vbe-glossary.md) and must be in the same procedure as the error handler.|
 **Remarks**
If you use a  **Resume** statement anywhere except in an error-handling routine, an error occurs.

## Example

This example uses the  **Resume** statement to end error handling in a procedure, and then resume execution with the statement that caused the error. Error number 55 is generated to illustrate using the **Resume** statement.


```vb
Sub ResumeStatementDemo() 
 On Error GoTo ErrorHandler ' Enable error-handling routine. 
 Open "TESTFILE" For Output As #1 ' Open file for output. 
 Kill "TESTFILE" ' Attempt to delete open file. 
 Exit Sub ' Exit Sub to avoid error handler. 
ErrorHandler: ' Error-handling routine. 
 Select Case Err.Number ' Evaluate error number. 
 Case 55 ' "File already open" error. 
 Close #1 ' Close open file. 
 Case Else 
 ' Handle other situations here.... 
 End Select 
 Resume ' Resume execution at same line 
 ' that caused the error. 
End Sub
```


