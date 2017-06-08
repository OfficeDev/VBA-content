---
title: TempVar Object (Access)
keywords: vbaac10.chm14063
f1_keywords:
- vbaac10.chm14063
ms.prod: access
api_name:
- Access.TempVar
ms.assetid: 4a0429e6-bcfa-7a8b-7030-6e88c2f1a71d
ms.date: 06/08/2017
---


# TempVar Object (Access)

Represents a variable that be used in Visual Basic for Applications (VBA) code or from a macro. 


## Remarks

A  **TempVar** objects provide a convenient way to exchange data between VBA procedures and macros.

Although a  **TempVar** object can be used to store information for use in VBA procedures, it does not have the same funcitonality as a VBA variable.


- By default, a  **TempVar** object remains in memory until Access is closed. You can use the **[Remove](http://msdn.microsoft.com/library/a9ab9ff2-5bfc-d001-f5eb-9929907bc1b2%28Office.15%29.aspx)** method or the[RemoveTempVar](http://msdn.microsoft.com/library/7bcc5010-3e30-ecef-2c5d-a35e73c8e325%28Office.15%29.aspx) macro action to remove a **TempVar** object.
    
- In VBA, a  **TempVar** object is accessible only to the members of the Access **[Application](http://msdn.microsoft.com/library/aefb0713-97e6-e2c7-e530-8fd2e1316a55%28Office.15%29.aspx)** object, referenced databases, or add-ins.
    
- A  **TempVar** object can store only text or numeric data. **TempVar** objects cannot store objects.
    
To refer to a  **TempVar** object in a collection by its ordinal number or by its **Name** property setting, use the following syntax form:


-  **TempVar** ![name]
    

|**Name**|
|:-----|
|[Name](http://msdn.microsoft.com/library/ce0983ec-1f12-d60e-4bfd-3960b5c10316%28Office.15%29.aspx)|
|[Value](http://msdn.microsoft.com/library/3bb66c34-2975-451e-6634-c23977753cb5%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](http://msdn.microsoft.com/library/2de134a4-6c5c-d2a3-8377-f4dd973ba650%28Office.15%29.aspx)
[TempVar Object Members](http://msdn.microsoft.com/library/1d8ac3a8-3116-6ce5-90c0-83265d7b79c4%28Office.15%29.aspx)
