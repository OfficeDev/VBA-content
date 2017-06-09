---
title: Sort Records in Case-Sensitive ( ASCII) Order
ms.prod: access
ms.assetid: 92c74803-2ff3-82b3-ca20-8bef5bfd6004
ms.date: 06/08/2017
---


# Sort Records in Case-Sensitive ( ASCII) Order

Access sorts records in ascending or descending order without regard to case. However, you can use a user-defined function in a query to sort text data by its ASCII character values. This results in a case-sensitive order.

The following table demonstrates how the ascending order in Access differs from a case-sensitive order.


|**Pre-Sort Order**|**Ascending Order**|**Case-Sensitive Order**|
|:-----|:-----|:-----|
|c|A|A|
|D|a|B|
|a|B|C|
|d|b|D|
|B|C|a|
|C|c|b|
|A|D|c|
|b|d|d|
Although the results in the Ascending Order column might at first appear somewhat unpredictable, they are not. In the Ascending Order column, "a" appears before "A," and " B" appears before "b." This occurs because, when evaluated as text values, "A" = "a" and "B" = "b," whether lowercase or uppercase. Access takes into account the original order of the values. In the Pre-Sort Order column, "a" precedes "A" and "B" precedes "b."
When the case-sensitive sort operation is performed, the text values are replaced with their ASCII values. For example, A = 65, a = 97, B = 66, b = 98, and so on.
The following user-defined function can be used to sort data in case-sensitive order. 



```vb
Function StrToHex(S As Variant) As Variant 
 Dim Temp As String, I As Integer 
 
 If VarType(S) <> 8 Then 
 
 StrToHex = S 
 Else 
 Temp = "" 
 For I = 1 To Len(S) 
 Temp = Temp &; Format(Hex(Asc(Mid(S, I, 1))), "00") 
 Next I 
 StrToHex = Temp 
 End If 
End Function
```


## Using the StrToHex Function

The preceding user-defined function, StrToHex, can be called from a query. When you pass the name of the sort field to this function, it will sort the field values in case-sensitive order. The following steps illustrate how to use the function.


1. Create a query from which you will call this function.
    
2. In the  **Show Table** dialog box, click the table that you want to sort, and then click **Add**.
    
3. Drag the fields you want to the grid.
    
4. In the first blank column, in the  **Field** row, type **Expr1: StrToHex([ _SortField_ ])**.StrToHex is the user-defined function you created earlier. SortField is the name of the field that contains the case-sensitive values.
    
5. In the  **Sort** cell, click **Ascending** or **Descending**.If you choose ascending order, values beginning with uppercase letters will appear before those that begin with lowercase letters. Applying a descending-order sort does the opposite.
    
6. Switch to Datasheet view.Access displays the records, sorted in case-sensitive order.
    

