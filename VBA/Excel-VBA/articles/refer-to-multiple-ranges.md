---
title: Refer to Multiple Ranges
keywords: vbaxl10.chm5204435
f1_keywords:
- vbaxl10.chm5204435
ms.prod: excel
ms.assetid: 11ac8eec-c754-d4e9-373c-84f04355d198
ms.date: 06/08/2017
---


# Refer to Multiple Ranges

By using the appropriate method, you can easily refer to multiple ranges. Use the  **Range** and **Union** methods to refer to any group of ranges. Use the **Areas** property to refer to the group of ranges selected on a worksheet.


## Using the Range Property

You can refer to multiple ranges with the  **Range** property by inserting commas between two or more references. The following example clears the contents of three ranges on Sheet1.


```vb
Sub ClearRanges() 
 Worksheets("Sheet1").Range("C5:D9,G9:H16,B14:D18"). _ 
 ClearContents 
End Sub
```

Named ranges make it easier to use the  **Range** property to work with multiple ranges. The following example works when all three named ranges are on the same sheet.




```vb
Sub ClearNamed() 
 Range("MyRange, YourRange, HisRange").ClearContents 
End Sub
```


## Using the Union Method

You can combine multiple ranges into one  **Range** object by using the **Union** method. The following example creates a **Range** object called `myMultipleRange`, defines it as the ranges A1:B2 and C3:D4, and then formats the combined ranges as bold.


```vb
Sub MultipleRange() 
 Dim r1, r2, myMultipleRange As Range 
 Set r1 = Sheets("Sheet1").Range("A1:B2") 
 Set r2 = Sheets("Sheet1").Range("C3:D4") 
 Set myMultipleRange = Union(r1, r2) 
 myMultipleRange.Font.Bold = True 
End Sub
```


## Using the Areas Property

You can use the  **Areas** property to refer to the selected range or to the collection of ranges in a multiple-area selection. The following procedure counts the areas in the selection. If there is more than one area, a warning message is displayed.


```vb
Sub FindMultiple() 
 If Selection.Areas.Count > 1 Then 
 MsgBox "Cannot do this to a multiple selection." 
 End If 
End Sub
```


