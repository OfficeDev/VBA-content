---
title: Comparison of Data Types
keywords: vbaac10.chm5186571
f1_keywords:
- vbaac10.chm5186571
ms.prod: ACCESS
ms.assetid: b1e954f8-e407-ae13-e3d9-a24d3502290a
---


# Comparison of Data Types

The Access database engine recognizes several overlapping sets of data types. In Access, there are four different contexts in which you may need to specify a data type — in table Design view, in the  **Query Parameters** dialog box, in Visual Basic, and in SQL view in a query.

The following table compares the five sets of data types that correspond to each context. The first column lists the  **Type** property settings available in table Design view and the five **FieldSize** property settings for the Number data type. The second column lists the corresponding query parameter data types available for designing parameter queries in the **Query Parameters** dialog box. The third column lists the corresponding Visual Basic data types. The fourth column lists ADO **Field** object data types. The fifth column lists the corresponding Jet database engine SQL data types defined by the Access database engine along with their valid synonyms.


|**Table fields**|**Query parameters**|**Visual Basic**|**ADO Data Type property constants**|**Access database engine SQL and synonyms**|
|:-----|:-----|:-----|:-----|:-----|
| _Not supported_|Binary| _Not supported_|**adBinary**|BINARY (See Notes)
 (Synonym: VARBINARY)|
|Yes/No|Yes/No|**Boolean**|**adBoolean**|BOOLEAN
 (Synonyms: BIT, LOGICAL, LOGICAL1, YESNO)|
|Number
 ( **FieldSize** = Byte)|Byte|**Byte**|**adUnsignedTinyInt**|BYTE
 (Synonym: INTEGER1)|
|AutoNumber
 ( **FieldSize** = Long Integer)|Long Integer|**Long**|**adInteger**|COUNTER
 (Synonym: AUTOINCREMENT)|
|Currency|Currency|**Currency**|**adCurrency**|CURRENCY
 (Synonym: MONEY)|
|Date/Time|Date/Time|**Date**|**adDate**|DATETIME
 (Synonyms: DATE, TIME, TIMESTAMP)|
|Number
 ( **FieldSize** =
 Double)|Double|**Double**|**adDouble**|DOUBLE
 (Synonyms: FLOAT, FLOAT8, IEEEDOUBLE, NUMBER, NUMERIC)|
|AutoNumber /GUID ( **FieldSize** =
 Replication ID)|Replication ID| _Not supported_|**adGUID**|GUID|
|Number
 ( **FieldSize** =
 Long Integer)|Long Integer|**Long**|**adInteger**|LONG (See Notes)
 (Synonyms: INT, INTEGER, INTEGER4)|
|OLE Object|OLE Object|**String**|**adLongVarBinary**|LONGBINARY
 (Synonyms: GENERAL, OLEOBJECT)|
|Memo|Memo|**String**|**adLongVarWChar**|LONGTEXT
 (Synonyms: LONGCHAR, MEMO, NOTE)|
|Number
 ( **FieldSize** =
 Single)|Single|**Single**|**adSingle**|SINGLE
 (Synonyms: FLOAT4, IEEESINGLE, REAL)|
|Number
 ( **FieldSize** =
 Integer)|Integer|**Integer**|**adSmallInt**|SHORT (See Notes)
 (Synonyms: INTEGER2, SMALLINT)|
|Text|Text|**String**|**adVarWChar**|TEXT
 (Synonyms: ALPHANUMERIC, CHAR, CHARACTER, STRING, VARCHAR)|
|Hyperlink|Memo|**String**|**adLongVarWChar**|LONGTEXT
 (Synonyms: LONGCHAR, MEMO, NOTE)|
| _Not supported_|Value|**Variant**|**adVariant**|VALUE (See Notes)|

 **Note**  

* Access itself doesn't use the BINARY data type. It's recognized only for use in queries on linked tables from other database products that support the BINARY data type.

* The INTEGER data type in the Access database engine SQL doesn't correspond to the **Integer** data type for table fields, query parameters, or Visual Basic. Instead, in SQL, the INTEGER data type corresponds to a **Long Integer** data type for table fields and query parameters and to a **Long** data type in Visual Basic.

* The VALUE reserved word doesn't represent a data type defined by the Access database engine. However, in Access or SQL queries, the VALUE reserved word can be considered a valid synonym for the Visual Basic **Variant** data type.

* If you are setting the data type for a DAO object in Visual Basic code, you must set the object's  **Type** property.

