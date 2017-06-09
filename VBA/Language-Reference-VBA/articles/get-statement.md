---
title: Get Statement
keywords: vblr6.chm1008928
f1_keywords:
- vblr6.chm1008928
ms.prod: office
ms.assetid: 73b44467-c9e6-3cd4-8d35-b2c19176bf80
ms.date: 06/08/2017
---


# Get Statement

Reads data from an open disk file into a [variable](vbe-glossary.md).

 **Syntax**

 **Get** [ **#** ] _filenumber_**,** [ _recnumber_ ] **,**_varname_

The  **Get** statement syntax has these parts:


|**Part**|**Description**|
|:-----|:-----|
| _filenumber_|Required. Any valid [file number](vbe-glossary.md).|
| _recnumber_|Optional.  **Variant** ( **Long** ). Record number ( **Random** mode files) or byte number ( **Binary** mode files) at which reading begins.|
| _varname_|Required. Valid variable name into which data is read.|
 **Remarks**
Data read with  **Get** is usually written to a file with **Put**.
The first record or byte in a file is at position 1, the second record or byte is at position 2, and so on. If you omit  _recnumber_, the next record or byte following the last **Get** or **Put**[statement](vbe-glossary.md) (or pointed to by the last **Seek** function) is read. You must include delimiting commas, for example:



```
Get #4,,FileBuffer 

```

For files opened in  **Random** mode, the following rules apply:


- If the length of the data being read is less than the length specified in the  **Len** clause of the **Open** statement, **Get** reads subsequent records on record-length boundaries. The space between the end of one record and the beginning of the next record is padded with the existing contents of the file buffer. Because the amount of padding data can't be determined with any certainty, it is generally a good idea to have the record length match the length of the data being read.
    
- If the variable being read into is a variable-length string,  **Get** reads a 2-byte descriptor containing the string length and then reads the data that goes into the variable. Therefore, the record length specified by the **Len** clause in the **Open** statement must be at least 2 bytes greater than the actual length of the string.
    
- If the variable being read into is a [Variant](vbe-glossary.md) of[numeric type](vbe-glossary.md),  **Get** reads 2 bytes identifying the **VarType** of the **Variant** and then the data that goes into the variable. For example, when reading a **Variant** of **VarType** 3, **Get** reads 6 bytes: 2 bytes identifying the **Variant** as **VarType** 3 ( **Long** ) and 4 bytes containing the[Long](vbe-glossary.md) data. The record length specified by the **Len** clause in the **Open** statement must be at least 2 bytes greater than the actual number of bytes required to store the variable.
    
     **Note**  You can use the  **Get** statement to read a **Variant**[array](vbe-glossary.md) from disk, but you can't use **Get** to read a scalar **Variant** containing an array. You also can't use **Get** to read objects from disk.
- If the variable being read into is a  **Variant** of **VarType** 8 ( **String** ), **Get** reads 2 bytes identifying the **VarType**, 2 bytes indicating the length of the string, and then reads the string data. The record length specified by the **Len** clause in the **Open** statement must be at least 4 bytes greater than the actual length of the string.
    
- If the variable being read into is a dynamic array,  **Get** reads a descriptor whose length equals 2 plus 8 times the number of dimensions, that is, 2 + 8 * _NumberOfDimensions_. The record length specified by the **Len** clause in the **Open** statement must be greater than or equal to the sum of all the bytes required to read the array data and the array descriptor. For example, the following array declaration requires 118 bytes when the array is written to disk.
    
```vb
Dim MyArray(1 To 5,1 To 10) As Integer 

  ```


    The 118 bytes are distributed as follows: 18 bytes for the descriptor (2 + 8 * 2), and 100 bytes for the data (5 * 10 * 2).
    
- If the variable being read into is a fixed-size array,  **Get** reads only the data. No descriptor is read.
    
- If the variable being read into is any other type of variable (not a variable-length string or a  **Variant** ), **Get** reads only the variable data. The record length specified by the **Len** clause in the **Open** statement must be greater than or equal to the length of the data being read.
    
-  **Get** reads elements of[user-defined types](vbe-glossary.md) as if each were being read individually, except that there is no padding between elements. On disk, a dynamic array in a user-defined type (written with **Put** ) is prefixed by a descriptor whose length equals 2 plus 8 times the number of dimensions, that is, 2 + 8 * _NumberOfDimensions_. The record length specified by the **Len** clause in the **Open** statement must be greater than or equal to the sum of all the bytes required to read the individual elements, including any arrays and their descriptors.
    

For files opened in  **Binary** mode, all of the **Random** rules apply, except:


- The  **Len** clause in the **Open** statement has no effect. **Get** reads all variables from disk contiguously; that is, with no padding between records.
    
- For any array other than an array in a user-defined type,  **Get** reads only the data. No descriptor is read.
    
-  **Get** reads variable-length strings that aren't elements of user-defined types without expecting the 2-byte length descriptor. The number of bytes read equals the number of characters already in the string. For example, the following statements read 10 bytes from[file number](vbe-glossary.md) 1:
    
  ```
  VarString = String(10," ") 
Get #1,,VarString 

  ```


    
    


## Example

This example uses the  **Get** statement to read data from a file into a variable. This example assumes that `TESTFILE` is a file containing five records of the user-defined type is a file containing five records of the user-defined type `Record`.


```vb
Type Record ' Define user-defined type. 
 ID As Integer 
 Name As String * 20 
End Type 
 
Dim MyRecord As Record, Position ' Declare variables. 
' Open sample file for random access. 
Open "TESTFILE" For Random As #1 Len = Len(MyRecord) 
' Read the sample file using the Get statement. 
Position = 3 ' Define record number. 
Get #1, Position, MyRecord ' Read third record. 
Close #1 ' Close file. 

```


