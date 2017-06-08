---
title: Put Statement
keywords: vblr6.chm1008997
f1_keywords:
- vblr6.chm1008997
ms.prod: office
ms.assetid: 6eb7c5bc-0332-9b4c-7ac0-52ddc9bb9dec
ms.date: 06/08/2017
---


# Put Statement

Writes data from a [variable](vbe-glossary.md) to a disk file.

 **Syntax**

 **Put** [ **#** ] _filenumber_, [ _recnumber_ ], _varname_

The  **Put** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](vbe-glossary.md).|
| _recnumber_|Optional.  **Variant** ( **Long** ). Record number ( **Random** mode files) or byte number ( **Binary** mode files) at which writing begins.|
| _varname_|Required. Name of variable containing data to be written to disk.|
 **Remarks**
Data written with  **Put** is usually read from a file with **Get**.
The first record or byte in a file is at position 1, the second record or byte is at position 2, and so on. If you omit  _recnumber_, the next record or byte after the last **Get** or **Put** statement or pointed to by the last **Seek** function is written. You must include delimiting commas, for example:



```
Put #4,,FileBuffer 

```

For files opened in  **Random** mode, the following rules apply:


- If the length of the data being written is less than the length specified in the  **Len** clause of the **Open** statement, **Put** writes subsequent records on record-length boundaries. The space between the end of one record and the beginning of the next record is padded with the existing contents of the file buffer. Because the amount of padding data can't be determined with any certainty, it is generally a good idea to have the record length match the length of the data being written. If the length of the data being written is greater than the length specified in the **Len** clause of the **Open** statement, an error occurs.
    
- If the variable being written is a variable-length string,  **Put** writes a 2-byte descriptor containing the string length and then the variable. The record length specified by the **Len** clause in the **Open** statement must be at least 2 bytes greater than the actual length of the string.
    
- If the variable being written is a [Variant](vbe-glossary.md) of a[numeric type](vbe-glossary.md),  **Put** writes 2 bytes identifying the **VarType** of the **Variant** and then writes the variable. For example, when writing a **Variant** of **VarType** 3, **Put** writes 6 bytes: 2 bytes identifying the **Variant** as **VarType** 3 ( **Long** ) and 4 bytes containing the **Long** data. The record length specified by the **Len** clause in the **Open** statement must be at least 2 bytes greater than the actual number of bytes required to store the variable.
    
     **Note**  You can use the  **Put** statement to write a **Variant**[array](vbe-glossary.md) to disk, but you can't use **Put** to write a scalar **Variant** containing an array to disk. You also can't use **Put** to write objects to disk.
- If the variable being written is a  **Variant** of **VarType** 8 ( **String** ), **Put** writes 2 bytes identifying the **VarType**, 2 bytes indicating the length of the string, and then writes the string data. The record length specified by the **Len** clause in the **Open** statement must be at least 4 bytes greater than the actual length of the string.
    
- If the variable being written is a dynamic array,  **Put** writes a descriptor whose length equals 2 plus 8 times the number of dimensions, that is, 2 + 8 * _NumberOfDimensions_. The record length specified by the **Len** clause in the **Open** statement must be greater than or equal to the sum of all the bytes required to write the array data and the array descriptor. For example, the following array declaration requires 118 bytes when the array is written to disk.
    
```vb
Dim MyArray(1 To 5,1 To 10) As Integer 

  ```


    
    
- The 118 bytes are distributed as follows: 18 bytes for the descriptor (2 + 8 * 2), and 100 bytes for the data (5 * 10 * 2).
    
- If the variable being written is a fixed-size array,  **Put** writes only the data. No descriptor is written to disk.
    
- If the variable being written is any other type of variable (not a variable-length string or a  **Variant** ), **Put** writes only the variable data. The record length specified by the **Len** clause in the **Open** statement must be greater than or equal to the length of the data being written.
    
-  **Put** writes elements of[user-defined types](vbe-glossary.md) as if each were written individually, except there is no padding between elements. On disk, a dynamic array in a user-defined type written with **Put** is prefixed by a descriptor whose length equals 2 plus 8 times the number of dimensions, that is, 2 + 8 * _NumberOfDimensions_. The record length specified by the **Len** clause in the **Open** statement must be greater than or equal to the sum of all the bytes required to write the individual elements, including any arrays and their descriptors.
    

For files opened in  **Binary** mode, all of the **Random** rules apply, except:


- The  **Len** clause in the **Open** statement has no effect. **Put** writes all variables to disk contiguously; that is, with no padding between records.
    
- For any array other than an array in a user-defined type,  **Put** writes only the data. No descriptor is written.
    
-  **Put** writes variable-length strings that are not elements of user-defined types without the 2-byte length descriptor. The number of bytes written equals the number of characters in the string. For example, the following statements write 10 bytes to file number 1:
    
  ```
  VarString$ = String$(10," ") 
Put #1,,VarString$ 

  ```


    
    


## Example

This example uses the  **Put** statement to write data to a file. Five records of the user-defined type are written to the file.


```vb
Type Record ' Define user-defined type. 
 ID As Integer 
 Name As String * 20 
End Type 
 
Dim MyRecord As Record, RecordNumber ' Declare variables. 
' Open file for random access. 
Open "TESTFILE" For Random As #1 Len = Len(MyRecord) 
For RecordNumber = 1 To 5 ' Loop 5 times. 
 MyRecord.ID = RecordNumber ' Define ID. 
 MyRecord.Name = "My Name" &; RecordNumber ' Create a string. 
 Put #1, RecordNumber, MyRecord ' Write record to file. 
Next RecordNumber 
Close #1 ' Close file. 

```


