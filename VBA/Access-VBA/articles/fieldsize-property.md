---
title: FieldSize Property
keywords: vbaac10.chm4349
f1_keywords:
- vbaac10.chm4349
ms.prod: access
api_name:
- Access.FieldSize
ms.assetid: 5cf8d67a-d81d-33d8-4afd-17e61abd3b08
ms.date: 06/08/2017
---


# FieldSize Property

  

**Applies to:** Access 2013 | Access 2016

You can use the  **FieldSize** property to set the maximum size for data stored in a field set to the Text, Number, or AutoNumber data type.


## Setting

If the  **[DataType](datatype-property.md)** property is set to Text, enter a number from 0 to 255. The default setting is 50.

If the  **DataType** property is set to AutoNumber, the **FieldSize** property can be set to Long Integer or Replication ID.

If the  **DataType** property is set to Number, the **FieldSize** property settings and their values are related in the following way.



|**Setting**|**Description**|**Decimal precision**|**Storage size**|
|:-----|:-----|:-----|:-----|
|Byte|Stores numbers from 0 to 255 (no fractions).|None|1 byte|
|Decimal|Stores numbers from -10^38-1 through 10^38-1 (.adp)Stores numbers from -10^28-1 through 10^28-1 (.mdb, .accdb)|28|2 bytes|
|Integer|Stores numbers from -32,768 to 32,767 (no fractions).|None|2 bytes|
|Long Integer|(Default) Stores numbers from -2,147,483,648 to 2,147,483,647 (no fractions).|None|4 bytes|
|Single|Stores numbers from -3.402823E38 to -1.401298E-45 for negative values and from 1.401298E-45 to 3.402823E38 for positive values.|7|4 bytes|
|Double|Stores numbers from -1.79769313486231E308 to -4.94065645841247E-324 for negative values and from4.94065645841247E-324 to 1.79769313486231E308 for positive values.|15|8 bytes|
|Replication ID|Globally unique identifier (GUID)|N/A|16 bytes|
You can set this property only from the table's property sheet.

To set the size of a field from Visual Basic , use the DAO  **Size** property to read and set the maximum size of Text fields (for data types other than Text, the DAO **Type** property setting automatically determines the **Size** property setting).


## Remarks

You should use the smallest possible  **FieldSize** property setting because smaller data sizes can be processed faster and require less memory.


 **Caution**  If you convert a large  **FieldSize** setting to a smaller one in a field that already contains data, you might lose data. For example, if you change the **FieldSize** setting for a Text data type field from 255 to 50, data beyond the new 50-character setting will be discarded.

If the data in a Number data type field doesn't fit in a new  **FieldSize** setting, fractional numbers may be rounded or you might get a Null value. For example, if you change from a Single to an Integer field size, fractional values will be rounded to the nearest whole number and values greater than 32,767 or less than -32,768 will result in null fields.

You can't undo changes to data that result from a change to the  **FieldSize** property after saving those changes in table Design view.


 **Note**  You can use the Currency data type if you plan to perform many calculations on a field that contains data with one to four decimal places. Single and Double data type fields require floating-point calculation. Currency data type fields use a faster fixed-point calculation.

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

