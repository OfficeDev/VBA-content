---
title: Excel performance - Tips for optimizing performance obstructions
description: Learn about ways to optimize Excel functionality such as links, lookups, formulas, functions, and VBA code to avoid common obstructions and improve performance.
ms.date: 10/06/2017 
author: FastExcel
---

# Excel performance: Tips for optimizing performance obstructions

**Applies to:** Excel | Excel 2013 | Office 2016 | VBA

Follow these tips for optimizing many frequently occurring performance obstructions in Excel. 

## Optimize references and links

Learn how to improve performance related to types of references and links.

### Do not use forward referencing and backward referencing

To increase clarity and avoid errors, design your formulas so that they do not refer forward (to the right or below) to other formulas or cells. Forward referencing usually does not affect calculation performance, except in extreme cases for the first calculation of a workbook, where it might take longer to establish a sensible calculation sequence if there are many formulas that need to have their calculation deferred.

### Minimize use of circular references with iteration

Calculating circular references with iterations is slow because multiple calculations are needed, and these calculations are single-threaded. Frequently you can "unroll" the circular references by using algebra so that iterative calculation is no longer needed. For example, in cash flow and interest calculations, try to calculate the cash flow before interest, calculate the interest, and then calculate the cash flow including the interest.
 
Excel calculates circular references sheet-by-sheet without considering dependencies. Therefore, you usually get slow calculation if your circular references span more than one worksheet. Try to move the circular calculations onto a single worksheet or optimize the worksheet calculation sequence to avoid unnecessary calculations.
  
Before the iterative calculations start, Excel must recalculate the workbook to identify all the circular references and their dependents. This process is equal to two or three iterations of the calculation. 
  
After the circular references and their dependents are identified, each iteration requires Excel to calculate not only all the cells in the circular reference, but also any cells that depend on the cells in the circular reference chain, together with volatile cells and their dependents. If you have a complex calculation that depends on cells in the circular reference, it can be faster to isolate this into a separate closed workbook and open it for recalculation after the circular calculation has converged.
  
It is important to reduce the number of cells in the circular calculation and the calculation time that is taken by these cells.

### Avoid links between workbooks

Avoid inter-workbook links when it is possible; they can be slow, easily broken, and not always easy to find and fix.

Using fewer larger workbooks is usually, but not always, better than using many smaller workbooks. Some exceptions to this might be when you have many front-end calculations that are so rarely recalculated that it makes sense to put them in a separate workbook, or when you have insufficient RAM.

Try to use simple direct cell references that work on closed workbooks. By doing this, you can avoid recalculating *all* your linked workbooks when you recalculate *any* workbook. Also, you can see the values Excel has read from the closed workbook, which is frequently important for debugging and auditing the workbook.
 
If you cannot avoid using linked workbooks, try to have them all open instead of closed, and open the workbooks that are linked to before you open the workbooks that are linked from.

### Minimize links between worksheets

Using many worksheets can make your workbook easier to use, but generally it is slower to calculate references to other worksheets than references within worksheets.

## Minimize the used range

To save memory and reduce file size, Excel tries to store information about only the area on a worksheet that was used. This is called the *used range*. Sometimes various editing and formatting operations extend the used range significantly beyond the range that you would currently consider used. This can cause performance obstructions and file-size obstructions.
 
You can check the visible used range on a worksheet by using Ctrl+End. Where this is excessive, you should consider deleting all the rows and columns below and to the right of your real last used cell, and then saving the workbook. Create a backup copy first. If you have formulas with ranges that extend into or refer to the deleted area, these ranges will be reduced in size or changed to **#N/A**.

## Allow for extra data

When you frequently add rows or columns of data to your worksheets, you need to find a way of having your Excel formulas automatically refer to the new data area, instead of trying to find and change your formulas every time.
  
You can do this by using a large range in your formulas that extends well beyond your current data boundaries. However, this can cause inefficient calculation under certain circumstances, and it is difficult to maintain because deleting rows and columns can decrease the range without you noticing.

### Use structured table references (recommended)

Starting in Excel 2007, you can use structured table references, which automatically expand and contract as the size of the referenced table increases or decreases. 

This solution has several advantages:

- Fewer performance disadvantages exist than the alternatives of whole column referencing and dynamic ranges.

- It is easy to have multiple tables of data on a single worksheet.

- Formulas that are embedded in the table also expand and contract with the data.

### Alternatively, use whole column and row references

An alternative approach is to use a whole column reference, for example **$A:$A**. This reference returns all the rows in Column A. Therefore, you can add as much data as you want, and the reference will always include it.

This solution has both advantages and disadvantages:

- Many Excel built-in functions (**SUM**, **SUMIF**) calculate whole column references efficiently because they automatically recognize the last used row in the column. However, array calculation functions like **SUMPRODUCT** either cannot handle whole column references or calculate all the cells in the column.

- User-defined functions do not automatically recognize the last-used row in the column and, therefore, frequently calculate whole column references inefficiently. However, it is easy to program user-defined functions so that they recognize the last-used row.

- It is difficult to use whole column references when you have multiple tables of data on a single worksheet.

- In Excel 2007 and later versions, array formulas can handle whole-column references, but this forces calculation for all the cells in the column, including empty cells. This can be slow to calculate, especially for 1 million rows.

### Alternatively, use dynamic ranges

By using the **OFFSET** or **INDEX** and **COUNTA** functions in the definition of a named range, you can make the area that the named range refers to dynamically expand and contract. For example, create a defined name using one of the following formulas:

```
  =OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)
  =Sheet1!$A$1:INDEX(Sheet1!$A:$A,COUNTA(Sheet1!$A:$A)+ROW(Sheet1!$A$1) - 1,1)
```

When you use the dynamic range name in a formula, it automatically expands to include new entries.
 
Using the **INDEX** formula for a dynamic range is generally preferable to the **OFFSET** formula because **OFFSET** has the disadvantage of being a volatile function that will be calculated at every recalculation.

Performance decreases because the **COUNTA** function inside the dynamic range formula must examine many rows. You can minimize this performance decrease by storing the **COUNTA** part of the formula in a separate cell or defined name, and then referring to the cell or name in the dynamic range:

```
 Counts!z1=COUNTA(Sheet1!$A:$A)
 OffsetDynamicRange=OFFSET(Sheet1!$A$1,0,0,Counts!$Z$1,1)
 IndexDynamicRange=Sheet1!$A$1:INDEX(Sheet1!$A:$A,Counts!$Z$1+ROW(Sheet1!$A$1) - 1,1)
```

You can also use functions such as **INDIRECT** to construct dynamic ranges, but **INDIRECT** is volatile and always calculates single-threaded. 

Dynamic ranges have the following advantages and disadvantages:

- Dynamic ranges work well to limit the number of calculations performed by array formulas.

- Using multiple dynamic ranges within a single column requires special-purpose counting functions.

- Using many dynamic ranges can decrease performance.

## Improve lookup calculation time

Lookups are frequently significant calculation obstructions. Fortunately, there are many ways of improving lookup calculation time. If you use the exact match option, the calculation time for the function is proportional to the number of cells scanned before a match is found. For lookups over large ranges, this time can be significant.

Lookup time using the approximate match options of **VLOOKUP**, **HLOOKUP**, and **MATCH** on sorted data is fast and is not significantly increased by the length of the range you are looking up. Characteristics are the same as binary search.

### Understand lookup options

Ensure that you understand the match-type and range-lookup options in **MATCH**, **VLOOKUP**, and **HLOOKUP**.

The following code example shows the syntax for the **MATCH** function. For more information, see the [Match](http://msdn.microsoft.com/library/901cdd78-e8fc-f149-66ff-5887f7099c96%28Office.14%29.aspx) method of the [WorksheetFunction](http://msdn.microsoft.com/library/7b1d5639-363d-632c-2cf0-2232562646b6%28Office.14%29.aspx) object.

```
  MATCH(lookup value, lookup array, matchtype)
```

- **Matchtype=1** returns the largest match less than or equal to the lookup value when the lookup array is sorted ascending (approximate match). If the lookup array is not sorted ascending, MATCH will return an incorrect answer. The default option is *approximate match sorted ascending*.

- **Matchtype=0** requests an exact match and assumes that the data is not sorted.

- **Matchtype=-1** returns the smallest match greater than or equal to the lookup value if the lookup array is sorted descending (approximate match).
    
The following code example shows the syntax for the **VLOOKUP** and **HLOOKUP** functions. For more information, see the [VLOOKUP](http://msdn.microsoft.com/library/1b84b1f5-b557-3a5c-0787-7c19a9800580%28Office.14%29.aspx) and [HLOOKUP](http://msdn.microsoft.com/library/6e7b5ad0-3f70-d7a8-b161-ce418107d2a1%28Office.14%29.aspx) methods of the [WorksheetFunction](http://msdn.microsoft.com/library/7b1d5639-363d-632c-2cf0-2232562646b6%28Office.14%29.aspx) object.

```
  VLOOKUP(lookup value, table array, col index num, range-lookup)
  HLOOKUP(lookup value, table array, row index num, range-lookup)
```

- **Range-lookup=TRUE** returns the largest match less than or equal to the lookup value (approximate match). This is the default option. Table array must be sorted ascending.
     
- **Range-lookup=FALSE** requests an exact match and assumes the data is not sorted.
    
Avoid performing lookups on unsorted data where possible because it is slow. If your data is sorted, but you want an exact match, see [Use two lookups for sorted data with missing values](#use-two-lookups-for-sorted-data-with-missing-values).

### Use INDEX and MATCH or OFFSET instead of VLOOKUP 

Try using the **INDEX** and **MATCH** functions instead of **VLOOKUP**. Although **VLOOKUP** is slightly faster (approximately 5 percent faster), simpler, and uses less memory than a combination of **MATCH** and **INDEX**, or **OFFSET**, the additional flexibility that **MATCH** and **INDEX** offer often enables you to significantly save time. For example, you can store the result of an exact **MATCH** in a cell and reuse it in several **INDEX** statements.

The **INDEX** function is fast and is a non-volatile function, which speeds up recalculation. The **OFFSET** function is also fast; however, it is a volatile function, and it sometimes significantly increases the time taken to process the calculation chain.

It is easy to convert **VLOOKUP** to **INDEX** and **MATCH**. The following two statements return the same answer:

```
 VLOOKUP(A1, Data!$A$2:$F$1000,3,False)

 INDEX(Data!$A$2:$F$1000,MATCH(A1,$A$1:$A$1000,0),3)
```

### Speed up lookups

Because exact match lookups can be slow, consider the following options for improving performance: 

- Use one worksheet. It is faster to keep lookups and data on the same sheet.    
  
- When you can, **SORT** the data first (**SORT** is fast), and use approximate match.  
  
- When you must use an exact match lookup, restrict the range of cells to be scanned to a minimum. Use tables and structured references or dynamic range names rather than referring to a large number of rows or columns. Sometimes you can pre-calculate a lower-range limit and upper-range limit for the lookup.  

### Use two lookups for sorted data with missing values

*Two approximate matches are significantly faster than one exact match for a lookup over more than a few rows.* (The breakeven point is about 10-20 rows.)

If you can sort your data but still cannot use approximate match because you cannot be sure that the value you are looking up exists in the lookup range, you can use this formula:

```
  IF(VLOOKUP(lookup_val ,lookup_array,1,True)=lookup_val, _
      VLOOKUP(lookup_val, lookup_array, column, True), "notexist")
```

The first part of the formula works by doing an approximate lookup on the lookup column itself.

```
  VLOOKUP(lookup_val ,lookup_array,1,True)
```

You can check if the answer from the lookup column is the same as the lookup value (in which case you have an exact match) by using the following formula:

```
  IF(VLOOKUP(lookup_val ,lookup_array,1,True)=lookup_val,
```

If this formula returns True, you have found an exact match, so you can do the approximate lookup again, but this time, return the answer from the column you want. 

```
  VLOOKUP(lookup_val, lookup_array, column, True)
```

If the answer from the lookup column did not match the lookup value, you have a missing value, and the formula returns "notexist".   
    
Be aware that if you look up a value smaller than the smallest value in the list, you receive an error. You can handle this error by using **IFERROR**, or by adding a small test value to the list.

### Use IFERROR function for unsorted data with missing values

If you must use exact match lookup on unsorted data, and you cannot be sure whether the lookup value exists, you often must handle the #N/A that is returned if no match is found. Beginning with Excel 2007, you can use the **IFERROR** function, which is both simple and fast.  

```
  IF IFERROR(VLOOKUP(lookupval, table, 2 FALSE),0)
```

In earlier versions, a simple but slow way is to use an **IF** function that contains two lookups.  

```
  IF(ISNA(VLOOKUP(lookupval,table,2,FALSE)),0,_
      VLOOKUP(lookupval,table,2,FALSE))
```

You can avoid the double exact lookup if you use exact **MATCH** once, store the result in a cell, and then test the result before doing an **INDEX**.

```
  In A1 =MATCH(lookupvalue,lookuparray,0)
  In B1 =IF(ISNA(A1),0,INDEX(tablearray,A1,column))
```

If you cannot use two cells, use **COUNTIF**. It is generally faster than an exact match lookup. 

```
  IF (COUNTIF(lookuparray,lookupvalue)=0, 0, _
      VLOOKUP(lookupval, table, 2 FALSE))
```

### Use MATCH and INDEX for exact match lookups on multiple columns

You can often reuse a stored exact **MATCH** many times. For example, if you are doing exact lookups on multiple result columns, you can save time by using one **MATCH** and many **INDEX** statements rather than many **VLOOKUP** statements.
 
Add an extra column for the **MATCH** to store the result (`stored_row`), and for each result column use the following: 

```
  INDEX(Lookup_Range,stored_row,column_number)
```

Alternatively, you can use **VLOOKUP** in an array formula. (Array formulas must be entered by using Ctrl+-Shift+Enter. Excel will add the { and } to show you that this is an array formula).

```
  {VLOOKUP(lookupvalue,{4,2},FALSE)}
```

### Use INDEX for a set of contiguous rows or columns

You can also return many cells from one lookup operation. To look up several contiguous columns, you can use the **INDEX** function in an array formula to return multiple columns at once (use 0 as the column number). You can also use the **INDEX** function to return multiple rows at one time.

```
  {INDEX($A$1:$J$1000,stored_row,0)}
```

This returns column A to column J from the stored row created by a previous **MATCH** statement.

### Use MATCH to return a rectangular block of cells

You can use the **MATCH** and **OFFSET** functions to return a rectangular block of cells.

### Use MATCH and INDEX for two-dimensional lookup

You can efficiently do a two-dimensional table lookup by using separate lookups on the rows and columns of a table by using an **INDEX** function with two embedded **MATCH** functions, one for the row and one for the column.

### Use a subset range for multiple-index lookup

In large worksheets, you may frequently need to look up by using multiple indexes, such as looking up product volumes in a country. To do this, you can concatenate the indexes and perform the lookup by using concatenated lookup values. However, this is inefficient for two reasons:

- Concatenating strings is a calculation-intensive operation.

- The lookup will cover a large range.

It is often more efficient to calculate a subset range for the lookup (for example, by finding the first and last row for the country, and then looking up the product within that subset range).

### Consider options for three-dimensional lookup

To look up the table to use in addition to the row and the column, you can use the following techniques, focusing on how to make Excel look up or choose the table.

If each table that you want to look up (the third dimension) is stored as a set of named structured tables, range names, or as a table of text strings that represent ranges, you might be able to use the **CHOOSE** or **INDIRECT** functions.

- Using **CHOOSE** and range names can be an efficient method. **CHOOSE** is not volatile, but it is best-suited to a relatively small number of tables. This example dynamically uses `TableLookup_Value` to choose which range name (`TableName1, TableName2, ...`) to use for the lookup table.

  ```
    INDEX(CHOOSE(TableLookup_Value,TableName1,TableName2,TableName3), _
    MATCH(RowLookup_Value,$A$2:$A$1000),MATCH(colLookup_value,$B$1:$Z$1))
  ```

- The following example uses the **INDIRECT** function and `TableLookup_Value` to dynamically create the sheet name to use for the lookup table. This method has the advantage of being simple and able to handle a large number of tables. Because **INDIRECT** is a volatile single-threaded function, the lookup is single-thread calculated at every calculation even if no data has changed. Using this method is slow.
  
  ```
    INDEX(INDIRECT("Sheet" &amp; TableLookup_Value &amp; "!$B$2:$Z$1000"), _
    MATCH(RowLookup_Value,$A$2:$A$1000),MATCH(colLookup_value,$B$1:$Z$1))
  ```

- You could also use the **VLOOKUP** function to find the name of the sheet or the text string to use for the table, and then use the **INDIRECT** function to convert the resulting text into a range.

  ```
    INDEX(INDIRECT(VLOOKUP(TableLookup_Value,TableOfTAbles,1)),MATCH(RowLookup_Value,$A$2:$A$1000),MATCH(colLookup_value,$B$1:$Z$1))
  ```

Another technique is to aggregate all your tables into one giant table that has an additional column that identifies the individual tables. You can then use the techniques for multiple-index lookup shown in the previous examples.

### Use wildcard lookup

The **MATCH**, **VLOOKUP**, and **HLOOKUP** functions allow you to use the wildcard characters **?** (any single character) and **\*** (no character or any number of characters) on alphabetical exact matches. Sometimes you can use this method to avoid multiple matches.

## Optimize array formulas and SUMPRODUCT

Array formulas and the **SUMPRODUCT** function are powerful, but you must handle them carefully. A single array formula might require many calculations. 
    
The key to optimizing the calculation speed of array formulas is to ensure that the number of cells and expressions that are evaluated in the array formula is as small as possible. Remember that an array formula is a bit like a volatile formula: if any one of the cells that it references has changed, is volatile, or has been recalculated, the array formula calculates all the cells in the formula and evaluates all the virtual cells it needs to do the calculation.
 
To optimize the calculation speed of array formulas:

- Take expressions and range references out of the array formulas into separate helper columns and rows. This makes much better use of the smart recalculation process in Excel.    
  
- Do not reference complete rows, or more rows and columns than you need. Array formulas are forced to calculate all the cell references in the formula even if the cells are empty or unused. With 1 million rows available starting in Excel 2007, an array formula that references a whole column is extremely slow to calculate.
     
- Starting in Excel 2007, use structured references where you can to keep the number of cells that are evaluated by the array formula to a minimum.
     
- In versions earlier than Excel 2007, use dynamic range names where possible. Although they are volatile, it is worthwhile because they minimize the size of the ranges.
    
- Be careful with array formulas that reference both a row and a column: this forces the calculation of a rectangular range.
    
- Use **SUMPRODUCT** if possible; it is slightly faster than the equivalent array formula.
    

### Consider options for using SUM for multiple-condition array formulas

You should always use the **SUMIFS**, **COUNTIFS**, and **AVERAGEIFS** functions instead of array formulas where you can because they are much faster to calculate. Excel 2016 introduces fast **MAXIFS** and **MINIFS** functions.  
      
In versions earlier than Excel 2007, array formulas are often used to calculate a sum with multiple conditions. This is relatively easy to do, especially if you use the **Conditional Sum Wizard** in Excel, but it is often slow. Usually there are much faster ways of getting the same result. If you have only a few multiple-condition SUMs, you may be able to use the **DSUM** function, which is much faster than the equivalent array formula.
 
If you must use array formulas, some good methods of speeding them up are as follows:

- Use dynamic range names or structured table references to minimize the number of cells.

- Split out the multiple conditions into a column of helper formulas that return **True** or **False** for each row, and then reference the helper column in a **SUMIF** or array formula. This might not appear to reduce the number of calculations for a single array formula; however, most of the time it enables the smart recalculation process to recalculate only the formulas in the helper column that need to be recalculated.

- Consider concatenating together all the conditions into a single condition, and then using **SUMIF**.

- If the data can be sorted, count groups of rows and limit the array formulas to looking at the subset groups.
    
### Prioritize multiple-condition SUMIFS, COUNTIFS, and other IFS family functions
  
These functions evaluate each of the conditions from left to right in turn. Therefore, it is more efficient to put the most restrictive condition first, so that subsequent conditions only need to look at the smallest number of rows.

### Consider options for using SUMPRODUCT for multiple-condition array formulas

Starting in Excel 2007, you should always use the **SUMIFS**, **COUNTIFS**, and **AVERAGEIFS** functions, and in Excel 2016 **MAXIFS** and **MINIFS** functions, instead of **SUMPRODUCT** formulas where possible.

In earlier versions, there are a few advantages to using **SUMPRODUCT** instead of **SUM** array formulas:

- **SUMPRODUCT** does not have to be array-entered by using Ctrl+Shift+Enter.

- **SUMPRODUCT** is usually slightly faster (5 to 10 percent).

You can use **SUMPRODUCT** for multiple-condition array formulas as follows:

```
  SUMPRODUCT(--(Condition1),--(Condition2),RangetoSum)
```

In this example, `Condition1` and `Condition2` are conditional expressions such as `$A$1:$A$10000<=$Z4`. Because conditional expressions return **True** or **False** instead of numbers, they must be coerced to numbers inside the **SUMPRODUCT** function. You can do this by using two minus signs (**--**), or by adding 0 (**+0**), or by multiplying by 1 (**x1**). Using **--** is slightly faster than **+0** or **x1**. 
 
Note that the size and shape of the ranges or arrays that are used in the conditional expressions and range to sum must be the same, and they cannot contain entire columns.
  
You can also directly multiply the terms inside **SUMPRODUCT** rather than separate them by commas:

```
  SUMPRODUCT((Condition1)*(Condition2)*RangetoSum)
```

This is usually slightly slower than using the comma syntax, and it gives an error if the range to sum contains a text value. However, it is slightly more flexible in that the range to sum may have, for example, multiple columns when the conditions have only one column. 

### Use SUMPRODUCT to multiply and add ranges and arrays

In cases like weighted average calculations, where you need to multiply a range of numbers by another range of numbers and sum the results, using the comma syntax for **SUMPRODUCT** can be 20 to 25 percent faster than an array-entered **SUM**.

```
  {=SUM($D$2:$D$10301*$E$2:$E$10301)}
  =SUMPRODUCT($D$2:$D$10301*$E$2:$E$10301)
  =SUMPRODUCT($D$2:$D$10301,$E$2:$E$10301)
```

These three formulas all produce the same result, but the third formula, which uses the comma syntax for **SUMPRODUCT**, takes only about 77 percent of the calculation time that the other two formulas need.

### Be aware of potential array and function calculation obstructions

The calculation engine in Excel is optimized to exploit array formulas and functions that reference ranges. However, some unusual arrangements of these formulas and functions can sometimes, but not always, cause significantly increased calculation time.

If you find a calculation obstruction that involves array formulas and range functions, you should look for the following:

- Partially overlapping references.

- Array formulas and range functions that reference part of a block of cells that are calculated in another array formula or range function. This situation can frequently occur in time series analysis.

- One set of formulas referencing by row, and a second set of formulas referencing the first set by column.

- A large set of single-row array formulas covering a block of columns, with **SUM** functions at the foot of each column.

## Use functions efficiently

Functions significantly extend the power of Excel, but the way in which you use them can often affect calculation time.
    
### Avoid single-threaded functions
Most native Excel functions work well with multi-threaded calculation. However, where possible, avoid using the following single-threaded functions:

- VBA and Automation user-defined functions (UDFs), but XLL-based UDFs can be multi-threaded
- PHONETIC
- CELL when either the "format" or "address" argument is used
- INDIRECT
- GETPIVOTDATA
- CUBEMEMBER
- CUBEVALUE
- CUBEMEMBERPROPERTY
- CUBESET
- CUBERANKEDMEMBER
- CUBEKPIMEMBER
- CUBESETCOUNT
- ADDRESS where the fifth parameter (the `sheet_name`) is given
- Any database function (DSUM, DAVERAGE, and so on) that refers to a pivot table
- ERROR.TYPE
- HYPERLINK

<br/>

### Use tables for functions that handle ranges

For functions like **SUM**, **SUMIF**, and **SUMIFS** that handle ranges, the calculation time is proportional to the number of used cells you are summing or counting. Unused cells are not examined, so whole column references are relatively efficient, but it is better to ensure that you do not include more used cells than you need. Use tables, or calculate subset ranges or dynamic ranges.

### Reduce volatile functions

Volatile functions can slow recalculation because they increase the number of formulas that must be recalculated at each calculation.

You can often reduce the number of volatile functions by using **INDEX** instead of **OFFSET**, and **CHOOSE** instead of **INDIRECT**. However, **OFFSET** is a fast function and can often be used in creative ways that give fast calculation.

### Use C or C++ user-defined functions

User-defined functions that are programmed in C or C++ and that use the C API (XLL add-in functions) generally perform faster than user-defined functions that are developed by using VBA or Automation (XLA or Automation add-ins). For more information, see  [Developing Excel 2010 XLLs](http://msdn.microsoft.com/library/dd27ae4d-ef97-47db-885c-ddd955816900%28Office.14%29.aspx).

The performance of VBA user-defined functions is sensitive to how you program and call them.

### Use faster VBA user-defined functions

It is usually faster to use the Excel formula calculations and worksheet functions than to use VBA user-defined functions. This is because there is a small overhead for each user-defined function call and significant overhead transferring information from Excel to the user-defined function. But well-designed and called user-defined functions can be much faster than complex array formulas.
  
Ensure that you have put all the references to worksheet cells in the user-defined function input parameters instead of in the body of the user-defined function, so that you can avoid adding **Application.Volatile** unnecessarily.
  
If you must have many formulas that use user-defined functions, ensure that you are in manual calculation mode, and that the calculation is initiated from VBA. VBA user-defined functions calculate much more slowly if the calculation is *not* called from VBA (for example, in automatic mode or when you press F9 in manual mode). This is particularly true when the Visual Basic Editor (Alt+F11) is open or has been opened in the current Excel session.
 
You can trap F9 and redirect it to a VBA calculation subroutine as follows. Add this subroutine to the *Thisworkbook* module.

```
  Private Sub Workbook_Open()
      Application.OnKey "{F9}", "Recalc"
  End Sub
```

Add this subroutine to a standard module.  
   
```
  Sub Recalc()
      Application.Calculate
      MsgBox "hello"
  End Sub
```

User-defined functions in Automation add-ins (Excel 2002 and later versions) do not incur the Visual Basic Editor overhead because they do not use the integrated editor. Other performance characteristics of Visual Basic 6 user-defined functions in Automation add-ins are similar to VBA functions.

If your user-defined function processes each cell in a range, declare the input as a range, assign it to a variant that contains an array, and loop on that. If you want to handle whole column references efficiently, you must make a subset of the input range, dividing it at its intersection with the used range, as in this example.

```
  Public Function DemoUDF(theInputRange as Range)
      Dim vArr as Variant
      Dim vCell as Variant
      Dim oRange as Range
      Set oRange=Union(theInputRange, theRange.Parent.UsedRange)
      vArr=oRange
      For Each vCell in vArr
          If IsNumeric(vCell) then DemoUDF=DemoUDF+vCell
      Next vCell
  End Function
```

If your user-defined function is using worksheet functions or Excel object model methods to process a range, it is generally more efficient to keep the range as an object variable than to transfer all the data from Excel to the user-defined function.

```
  Function uLOOKUP(lookup_value As Variant, lookup_array As Range, _
                   col_num As Variant, sorted As Variant, _
                   NotFound As Variant)
      Dim vAnsa As Variant
      vAnsa = Application.VLookup(lookup_value, lookup_array, _
                                  col_num, sorted)
      If Not IsError(vAnsa) Then
          uLOOKUP = vAnsa
      Else
          uLOOKUP = NotFound
      End If
  End Function
```

If your user-defined function is called early in the calculation chain, it can be passed as uncalculated arguments. Inside a user-defined function, you can detect uncalculated cells by using the following test for empty cells that contain a formula:

```
  If ISEMPTY(Cell.Value) AND Len(Cell.formula)>0 then
```

A time overhead exists for each call to a user-defined function and for each transfer of data from Excel to VBA. Sometimes one multi-cell array formula user-defined function can help you minimize these overheads by combining multiple function calls into a single function with a multi-cell input range that returns a range of answers. 

### Minimize range of cells that SUM and SUMIF reference

The Excel **SUM** and **SUMIF** functions are frequently used over a large number of cells. Calculation time for these functions is proportionate to the number of cells covered, so try to minimize the range of cells that the functions are referencing.

### Use wildcard SUMIF, COUNTIF, SUMIFS, COUNTIFS, and other IFS functions

You can use the wildcard characters **?** (any single character) and **\*** (no character or any number of characters) in the criteria for alphabetical ranges as part of the **SUMIF**, **COUNTIF**, **SUMIFS**, **COUNTIFS**, and other **IFS** functions.

### Choose method for period-to-date and cumulative SUMs

There are two methods of doing period-to-date or cumulative SUMs. Suppose the numbers that you want to cumulatively **SUM** are in column A, and you want column B to contain the cumulative sum; you can do either of the following:

- You can create a formula in column B such as `=SUM($A$1:$A2)` and drag it down as far as you need. The beginning cell of the SUM is anchored in A1, but because the finishing cell has a relative row reference, it automatically increases for each row.

- You can create a formula such as `=$A1` in cell B1 and `=$B1+$A2` in cell B2 and drag it down as far as you need. This calculates the cumulative cell by adding this row's number to the previous cumulative **SUM**.

For 1,000 rows, the first method makes Excel do about 500,000 calculations, but the second method makes Excel do only about 2,000 calculations.

### Calculate subset sums

When you have multiple sorted indexes to a table (for example, Site within Area) you can often save significant calculation time by dynamically calculating the address of a subset range of rows (or columns) to use in the **SUM** or **SUMIF** function.

**To calculate the address of a subset range of row or columns:**

1. Count the number of rows for each subset block.
  
2. Add the counts cumulatively for each block to determine its start row.
  
3. Use **OFFSET** with the start row and count to return a subset range to the **SUM** or **SUMIF** that covers only the subset block of rows.

### Use SUBTOTAL for filtered lists

Use the **SUBTOTAL** function to **SUM** filtered lists. The **SUBTOTAL** function is useful because, unlike **SUM**, it ignores the following:

- Hidden rows that result from filtering a list. Starting in Excel 2003, you can also make **SUBTOTAL** ignore all hidden rows, not just filtered rows.

- Other **SUBTOTAL** functions.
    
### Use the AGGREGATE function

The AGGREGATE function is a powerful and efficient way of calculating 19 different methods of aggregating data (such as **SUM**, **MEDIAN**, **PERCENTILE** and **LARGE**). **AGGREGATE** has options for ignoring hidden or filtered rows, error values, and nested **SUBTOTAL** and **AGGREGATE** functions.

### Avoid using DFunctions

The DFunctions **DSUM**, **DCOUNT**, **DAVERAGE**, and so on are significantly faster than equivalent array formulas. The disadvantage of the DFunctions is that the criteria must be in a separate range, which makes them impractical to use and maintain in many circumstances. Starting in Excel 2007, you should use **SUMIFS**, **COUNTIFS**, and **AVERAGEIFS** functions instead of the DFunctions.

## Create faster VBA macros

Use the following tips to create faster VBA macros.

### Turn off everything but the essentials while code is running

To improve performance for VBA macros, explicitly turn off the functionality that is not required while your code executes. Often, one recalculation or one redraw after your code runs is all that is necessary and can improve performance. After your code executes, restore the functionality to its original state.

The following functionality can usually be turned off while your VBA macro executes:

- **Application.ScreenUpdating** Turn off screen updating. If **Application.ScreenUpdating** is set to **False**, Excel does not redraw the screen. While your code runs, the screen updates quickly, and it is usually not necessary for the user to see each update. Updating the screen once, after the code executes, improves performance.   
  
- **Application.DisplayStatusBar** Turn off the status bar. If **Application.DisplayStatusBar** is set to **False**, Excel does not display the status bar. The status bar setting is separate from the screen updating setting so that you can still display the status of the current operation even while the screen is not updating. However, if you do not need to display the status of every operation, turning off the status bar while your code runs also improves performance.  
  
- **Application.Calculation** Switch to manual calculation. If **Application.Calculation** is set to **xlCalculationManual**, Excel only calculates the workbook when the user explicitly initiates the calculation. In automatic calculation mode, Excel determines when to calculate. For example, every time a cell value that is related to a formula changes, Excel recalculates the formula. If you switch the calculation mode to manual, you can wait until all the cells associated with the formula are updated before recalculating the workbook. By only recalculating the workbook when necessary while your code runs, you can improve performance. 
  
- **Application.EnableEvents** Turn off events. If **Application.EnableEvents** is set to **False**, Excel does not raise events. If there are add-ins listening for Excel events, those add-ins consume resources on the computer as they record the events. If it is not necessary for the add-in to record the events that occur while your code runs, turning off events improves performance.
  
- **ActiveSheet.DisplayPageBreaks** Turn off page breaks. If **ActiveSheet.DisplayPageBreaks** is set to **False**, Excel does not display page breaks. It is not necessary to recalculate page breaks while your code runs, and calculating the page breaks after the code executes improves performance.

> [!IMPORTANT]
> Remember to restore this functionality to its original state after your code executes. 

The following example shows the functionality that you can turn off while your VBA macro executes.

```
  ' Save the current state of Excel settings.
  screenUpdateState = Application.ScreenUpdating
  statusBarState = Application.DisplayStatusBar
  calcState = Application.Calculation
  eventsState = Application.EnableEvents
  ' Note: this is a sheet-level setting.
  displayPageBreakState = ActiveSheet.DisplayPageBreaks 

  ' Turn off Excel functionality to improve performance.
  Application.ScreenUpdating = False
  Application.DisplayStatusBar = False
  Application.Calculation = xlCalculationManual
  Application.EnableEvents = False
  ' Note: this is a sheet-level setting.
  ActiveSheet.DisplayPageBreaks = False

  ' Insert your code here.

  ' Restore Excel settings to original state.
  Application.ScreenUpdating = screenUpdateState
  Application.DisplayStatusBar = statusBarState
  Application.Calculation = calcState
  Application.EnableEvents = eventsState
  ' Note: this is a sheet-level setting
  ActiveSheet.DisplayPageBreaks = displayPageBreaksState
```

<br/>

### Read and write large blocks of data in a single operation

Optimize your code by explicitly reducing the number of times data is transferred between Excel and your code. Instead of looping through cells one at a time to get or set a value, get or set the values in the entire range of cells in one line, using a variant containing a two-dimensional array to store values as needed. The following code examples compare these two methods.

The following code example shows non-optimized code that loops through cells one at a time to get and set the values of cells A1:C10000. These cells do not contain formulas.

```
  Dim DataRange as Range
  Dim Irow as Long
  Dim Icol as Integer 
  Dim MyVar as Double 
  Set DataRange=Range("A1:C10000") 

  For Irow=1 to 10000 
      For icol=1 to 3
          ' Read the values from the Excel grid 30,000 times.
          MyVar=DataRange(Irow,Icol) 
          If MyVar > 0 then 
              ' Change the value.
              MyVar=MyVar*Myvar 
              ' Write the values back into the Excel grid 30,000 times.
              DataRange(Irow,Icol)=MyVar
          End If 
      Next Icol 
  Next Irow
```

The following code example shows optimized code that uses an array to get and set the values of cells A1:C10000 all at the same time. These cells do not contain formulas.  

```
  Dim DataRange As Variant
  Dim Irow As Long 
  Dim Icol As Integer 
  Dim MyVar As Double 
  ' Read all the values at once from the Excel grid and put them into an array.
  DataRange = Range("A1:C10000").Value2 

  For Irow = 1 To 10000 
      For Icol = 1 To 3 
          MyVar = DataRange(Irow, Icol) 
          If MyVar > 0 Then 
              ' Change the values in the array.
              MyVar=MyVar*Myvar 
              DataRange(Irow, Icol) = MyVar 
          End If 
      Next Icol 
  Next Irow 
  ' Write all the values back into the range at once.
  Range("A1:C10000").Value2 = DataRange 
```

<br/>

### Use .Value2 rather than .Value or .Text when reading data from an Excel range

- **.Text** returns the formatted value of a cell. This is slow, can return ### if the user zooms, and can lose precision.
- **.Value** returns a VBA currency or VBA date variable if the range was formatted as Date or Currency. This is slow, can lose precision, and can cause errors when calling worksheet functions.
- **.Value2** is fast and does not alter the data being retrieved from Excel.

### Avoid selecting and activating objects

Selecting and activating objects is more processing intensive than referencing objects directly. By referencing an object such as a **Range** or a **Shape** directly, you can improve performance. The following code examples compare the two methods.
    
The following code example shows non-optimized code that selects each Shape on the active sheet and changes the text to "Hello".
  
```
  For i = 0 To ActiveSheet.Shapes.Count
      ActiveSheet.Shapes(i).Select
      Selection.Text = "Hello"
  Next i
```

The following code example shows optimized code that references each Shape directly and changes the text to "Hello".

```
  For i = 0 To ActiveSheet.Shapes.Count
      ActiveSheet.Shapes(i).TextEffect.Text = "Hello"
  Next i
```

### Use these additional VBA performance optimizations

The following is a list of additional performance optimizations you can use in your VBA code:

- Return results by assigning an array directly to a **Range**.

- Declare variables with explicit types to avoid the overhead of determining the data type, possibly multiple times in a loop, during code execution.

- For simple functions that you use frequently in your code, implement the functions yourself in VBA instead of using the **WorksheetFunction** object. For more information, see [Use faster VBA user-defined functions](#use-faster-vba-user-defined-functions).

- Use the **Range.SpecialCells** method to scope down the number of cells with which your code interacts.

- Consider the performance gains if you implemented your functionality by using the C API in the XLL SDK. For more information, see the  [Excel 2010 XLL SDK Documentation](http://msdn.microsoft.com/library/abfc9d76-6f22-49b9-ba45-eb7a54b082e0%28Office.14%29.aspx).

## Consider performance and size of Excel file formats

Starting in Excel 2007, Excel contains a wide variety of file formats compared to earlier versions. Ignoring the Macro, Template, Add-in, PDF, and XPS file format variations, the three main formats are XLS, XLSB, and XLSX.

- **XLS format**

  The XLS format is the same format as earlier versions. When you use this format, you are restricted to 256 columns and 65,536 rows. When you save an Excel 2007 or Excel 2010 workbook in XLS format, Excel runs a compatibility check. File size is almost the same as earlier versions (some additional information may be stored), and performance is slightly slower than earlier versions. Any multi-threaded optimization Excel does with respect to cell calculation order is not saved in the XLS format. Therefore, calculation of a workbook can be slower after saving the workbook in the XLS format, closing, and re-opening the workbook.

- **XLSB format**

  XLSB is the binary format starting in Excel 2007. It is structured as a compressed folder that contains many binary files. It is much more compact than the XLS format, but the amount of compression depends on the contents of the workbook. For example, ten workbooks show a size reduction factor ranging from two to eight with an average reduction factor of four. Starting in Excel 2007, opening and saving performance is only slightly slower than the XLS format.

- **XLSX format**

  XLSX is the XML format starting in Excel 2007, and is the default format starting in Excel 2007. The XLSX format is a compressed folder that contains many XML files (if you change the file name extension to .zip, you can open the compressed folder and examine its contents). Typically, the XLSX format creates larger files than the XLSB format (1.5 times larger on average), but they are still significantly smaller than the XLS files. You should expect opening and saving times to be slightly longer than for XLSB files. 

## Open, close, and save workbooks

You may find that opening, closing, and saving workbooks is much slower than calculating them. Sometimes this is just because you have a large workbook, but there can also be other reasons.

If one or more of your workbooks open and close more slowly than is reasonable, it might be caused by one of the following issues.

- **Temporary files**

  Temporary files can accumulate in your \\Windows\\Temp directory (in Windows 95, Windows 98, and Windows ME), or your \\Documents and Settings\\User Name\\Local Settings\\Temp directory (in Windows 2000 and Windows XP). Excel creates these files for the workbook and for controls that are used by open workbooks. Software installation programs also create temporary files. If Excel stops responding for any reason, you might need to delete these files. 

  Too many temporary files can cause problems, so you should occasionally clean them out. However, if you have installed software that requires that you restart your computer, and you have not yet done so, you should restart before deleting the temporary files. 

  An easy way to open your temp directory is from the Windows **Start** menu: Click **Start**, and then click **Run**. In the text box, type **%temp%**, and then click **OK**.

- **Tracking changes in a shared workbook**

  Tracking changes in a shared workbook causes your workbook file-size to increase rapidly.

- **Fragmented swap file**

  Be sure that your Windows swap file is located on a disk that has a lot of space and that you defragment the disk periodically.

- **Workbook with password-protected structure**

  A workbook that has its structure protected with a password (**Tools** menu > **Protection** > **Protect Workbook** > enter the optional password) opens and closes much slower than one that is protected without the optional password.

- **Used range problems**

  Oversized used ranges can cause slow opening and increased file size, especially if they are caused by hidden rows or columns that have non-standard height or width. For more information about used range problems, see [Minimize the used range](#minimize-the-used-range).

- **Large number of controls on worksheets**

  A large number of controls (check boxes, hyperlinks, and so on) on worksheets can slow down opening a workbook because of the number of temporary files that are used. This might also cause problems opening or saving a workbook on a WAN (or even a LAN). If you have this problem, you should consider redesigning your workbook.  

- **Large number of links to other workbooks**

  If possible, open the workbooks that you are linking to before you open the workbook that contains the links. Often it is faster to open a workbook than to read the links from a closed workbook.

- **Virus scanner settings**

  Some virus scanner settings can cause problems or slowness with opening, closing, or saving, especially on a server. If you think that this might be the problem, try temporarily switching the virus scanner off.  

- **Slow calculation causing slow open and save**

  Under some circumstances, Excel recalculates your workbook when it opens or saves it. If the calculation time for your workbook is long and is causing a problem, ensure that you have calculation set to **manual**, and consider turning off the **calculate before save** option (**Tools** > **Options** > **Calculation**). 

- **Toolbar files (.xlb)**

  Check the size of your toolbar file. A typical toolbar file is between 10 KB and 20 KB. You can find your XLB files by searching for `*.xlb` by using Windows search. Each user has a unique XLB file. Adding, changing, or customizing toolbars increases the size of your toolbar.xlb file. Deleting the file removes all your toolbar customizations (renaming it "toolbar.OLD" is safer). A new XLB file is created the next time you open Excel.

## Make additional performance optimizations

You can make performance improvements in the following areas.

- **PivotTables**

  PivotTables provide an efficient way to summarize large amounts of data.

  - **Totals as final results**. If you need to produce totals and subtotals as part of the final results of your workbook, try using PivotTables.

  - **Totals as intermediate results**. PivotTables are a great way to produce summary reports, but try to avoid creating formulas that use PivotTable results as intermediate totals and subtotals in your calculation chain unless you can ensure the following conditions:

   - The PivotTable has been refreshed correctly during the calculation. 

   - The PivotTable has not been changed so that the information is still visible.

  If you still want to use PivotTables as intermediate results, use the **GETPIVOTDATA** function.

- **Conditional formats and data validation**

  Conditional formats and data validation are great, but using a lot of them can significantly slow down calculation. If the cell is displayed, every conditional format formula is evaluated at each calculation and when the display of the cell that contains the conditional format is refreshed. The Excel object model has a **Worksheet.EnableFormatConditionsCalculation** property so that you can enable or disable the calculation of conditional formats.

- **Defined names**

  Defined names are one of the most powerful features in Excel, but they do take additional calculation time. Using names that refer to other worksheets adds an additional level of complexity to the calculation process. Also, you should try to avoid nested names (names that refer to other names).

  Because names are calculated every time a formula that refers to them is calculated, you should avoid putting calculation-intensive formulas or functions in defined names. In these cases, it can be significantly faster to put your calculation-intensive formula or function in a spare cell somewhere and refer to that cell instead, either directly or by using a name.

- **Formulas that are used only occasionally**

  Many workbooks contain a significant number of formulas and lookups that are concerned with getting the input data into the appropriate shape for the calculations, or are being used as defensive measures against changes in the size or shape of the data. When you have blocks of formulas that are used only occasionally, you can copy and paste special values to temporarily eliminate the formulas, or you can put them in a separate, rarely opened workbook. Because worksheet errors are often caused by not noticing that formulas have been converted to values, the separate workbook method may be preferable.

- **Use enough memory**

  The 32-bit version of Excel can use up to 2 GB of RAM or up to 4 GB of RAM for Large Address Aware 32-bit versions of Excel 2013 and 2016. However, the computer that is running Excel also requires memory resources. Therefore, if you only have 2 GB of RAM on your computer, Excel cannot take advantage of the full 2 GB because a portion of the memory is allocated to the operating system and other programs that are running. To optimize the performance of Excel on a 32-bit computer, we recommend that the computer have at least 3 GB of RAM.

  The 64-bit version of Excel does not have a 2 GB or up to 4 GB limit. For more information, see the "Large data sets and the 64-bit version of Excel" section in [Excel performance: Performance and limit improvements](excel-performance-and-limit-improvements.md).

## Conclusion

This article covered ways to optimize Excel functionality such as links, lookups, formulas, functions, and VBA code to avoid common obstructions and improve performance.

## Additional resources

-  [Excel performance: Improving calculation performance](excel-improving-calcuation-performance.md)
    
-  [Excel performance: Performance and limit improvements](excel-performance-and-limit-improvements.md)
    
-  [Excel Developer Portal](http://msdn.microsoft.com/en-us/office/aa905411.aspx)

