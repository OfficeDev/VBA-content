---
title: WorksheetFunction.Replace Method (Excel)
keywords: vbaxl10.chm137127
f1_keywords:
- vbaxl10.chm137127
ms.prod: excel
api_name:
- Excel.WorksheetFunction.Replace
ms.assetid: 1cca39db-c4ab-f7d4-dd71-0844d0bb44cd
ms.date: 06/08/2017
---


# WorksheetFunction.Replace Method (Excel)

Replaces part of a text string, based on the number of characters you specify, with a different text string.


## Syntax

 _expression_ . **Replace**( **_Arg1_** , **_Arg2_** , **_Arg3_** , **_Arg4_** )

 _expression_ A variable that represents a **WorksheetFunction** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Arg1_|Required| **String**|Text in which you want to replace some characters.|
| _Arg2_|Required| **Double**|The position of the character in  **Arg1** that you want to replace with **Arg4** .|
| _Arg3_|Required| **Double**|The number of characters in  **Arg1** that you want the **Replace** method to replace with **Arg4** .|
| _Arg4_|Required| **String**|Text that will replace characters in  **Arg1** .|

### Return Value

A String value that represents the new string, after replacement.


## Example

This example replaces abcdef with ac-ef and notifies the user during this process.


```vb
Sub UseReplace() 
 
 Dim strCurrent As String 
 Dim strReplaced As String 
 
 strCurrent = "abcdef" 
 
 ' Notify user and display current string. 
 MsgBox "The current string is: " &; strCurrent 
 
 ' Replace "cd" with "-". 
 strReplaced = Application.WorksheetFunction.Replace _ 
 (Arg1:=strCurrent, Arg2:=3, _ 
 Arg3:=2, Arg4:="-") 
 
 ' Notify user and display replaced string. 
 MsgBox "The replaced string is: " &; strReplaced 
 
End Sub
```


## See also


#### Concepts


[WorksheetFunction Object](worksheetfunction-object-excel.md)

