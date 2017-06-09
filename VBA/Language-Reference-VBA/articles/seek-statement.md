---
title: Seek Statement
keywords: vblr6.chm1009013
f1_keywords:
- vblr6.chm1009013
ms.prod: office
ms.assetid: 08fff310-85a2-d860-2198-3a0b032c77bc
ms.date: 06/08/2017
---


# Seek Statement

Sets the position for the next read/write operation within a file opened using the  **Open** statement.

 **Syntax**

 **Seek** [ **#** ] _filenumber_, _position_

The  **Seek** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](vbe-glossary.md).|
| _position_|Required. Number in the range 1 - 2,147,483,647, inclusive, that indicates where the next read/write operation should occur.|
 **Remarks**
Record numbers specified in  **Get** and **Put** statements override file positioning performed by **Seek**.
Performing a file-write operation after a  **Seek** operation beyond the end of a file extends the file. If you attempt a **Seek** operation to a negative or zero position, an error occurs.

## Example

This example uses the  **Seek** statement to set the position for the next read or write within a file. This example assumes `TESTFILE` is a file containing records of the user-defined type is a file containing records of the user-defined type `Record`.


```vb
Type Record ' Define user-defined type. 
 ID As Integer 
 Name As String * 20 
End Type 

```

For files opened in Random mode,  **Seek** sets the next record.




```vb
Dim MyRecord As Record, MaxSize, RecordNumber ' Declare variables. 
' Open file in random-file mode. 
Open "TESTFILE" For Random As #1 Len = Len(MyRecord) 
MaxSize = LOF(1) \ Len(MyRecord) ' Get number of records in file. 
' The loop reads all records starting from the last. 
For RecordNumber = MaxSize To 1 Step - 1 
 Seek #1, RecordNumber ' Set position. 
 Get #1, , MyRecord ' Read record. 
Next RecordNumber 
Close #1 ' Close file. 

```

For files opened in modes other than Random mode,  **Seek** sets the byte position at which the next operation takes place. Assume is a file containing a few lines of text.




```vb
Dim MaxSize, NextChar, MyChar 
Open "TESTFILE" For Input As #1 ' Open file for input. 
MaxSize = LOF(1) ' Get size of file in bytes. 
' The loop reads all characters starting from the last. 
For NextChar = MaxSize To 1 Step -1 
 Seek #1, NextChar ' Set position. 
 MyChar = Input(1, #1) ' Read character. 
Next NextChar 
Close #1 ' Close file. 

```


