---
title: Comparison of Data Types
keywords: vbaac10.chm5186571
f1_keywords:
- vbaac10.chm5186571
ms.prod: access
ms.assetid: b1e954f8-e407-ae13-e3d9-a24d3502290a
ms.date: 06/08/2017
---


# Comparison of Data Types

The Access database engine recognizes several overlapping sets of data types. In Access, there are four different contexts in which you may need to specify a data type — in table Design view, in the  **Query Parameters** dialog box, in Visual Basic, and in SQL view in a query.

The following table compares the five sets of data types that correspond to each context. The first column lists the  **Type** property settings available in table Design view and the five **FieldSize** property settings for the Number data type. The second column lists the corresponding query parameter data types available for designing parameter queries in the **Query Parameters** dialog box. The third column lists the corresponding Visual Basic data types. The fourth column lists ADO **Field** object data types. The fifth column lists the corresponding Jet database engine SQL data types defined by the Access database engine along with their valid synonyms.


|**Table fields**|**Query parameters**|**Visual Basic**|**ADO Data Type property constants**|**Access database engine SQL and synonyms**|
|:-----|:-----|:-----|:-----|:-----|
| _Not supported_|Binary| _Not supported_|**adBinary**|<p>BINARY (See Notes)</p><p>(Synonym: VARBINARY)</p>|
|Yes/No|Yes/No|**Boolean**|**adBoolean**|<p>BOOLEAN</p><p>(Synonyms: BIT, LOGICAL, LOGICAL1, YESNO)</p>|
|<p>Number</p><p>( **FieldSize** = Byte)</p>|Byte|**Byte**|**adUnsignedTinyInt**|<p>BYTE</p><p>(Synonym: INTEGER1)</p>|
|<p>AutoNumber</p><p>( **FieldSize** = Long Integer)</p>|Long Integer|**Long**|**adInteger**|<p>COUNTER</p><p>(Synonym: AUTOINCREMENT)</p>|
|Currency|Currency|**Currency**|**adCurrency**|<p>CURRENCY</p><p>(Synonym: MONEY)</p>|
|Date/Time|Date/Time|**Date**|**adDate**|<p>DATETIME</p><p>(Synonyms: DATE, TIME, TIMESTAMP)</p>|
|<p>Number</p><p>( **FieldSize** = Double)</p>|Double|**Double**|**adDouble**|<p>DOUBLE</p><p>(Synonyms: FLOAT, FLOAT8, IEEEDOUBLE, NUMBER, NUMERIC)</p>|
|<p>AutoNumber /GUID </p><p>( **FieldSize** = Replication ID)</p>|Replication ID| _Not supported_|**adGUID**|GUID|
|<p>Number</p><p>( **FieldSize** = Long Integer)</p>|Long Integer|**Long**|**adInteger**|<p>LONG (See Notes)</p><p>(Synonyms: INT, INTEGER, INTEGER4)</p>|
|OLE Object|OLE Object|**String**|**adLongVarBinary**|<p>LONGBINARY</p><p>(Synonyms: GENERAL, OLEOBJECT)</p>|
|Memo|Memo|**String**|**adLongVarWChar**|<p>LONGTEXT</p><p>(Synonyms: LONGCHAR, MEMO, NOTE)</p>|
|<p>Number</p><p>( **FieldSize** = Single)</p>|Single|**Single**|**adSingle**|<p>SINGLE</p><p>(Synonyms: FLOAT4, IEEESINGLE, REAL)</p>|
|<p>Number</p><p>( **FieldSize** = Integer)</p>|Integer|**Integer**|**adSmallInt**|<p>SHORT (See Notes)</p><p>(Synonyms: INTEGER2, SMALLINT)</p>|
|Text|Text|**String**|**adVarWChar**|<p>TEXT</p><p>(Synonyms: ALPHANUMERIC, CHAR, CHARACTER, STRING, VARCHAR)</p>|
|Hyperlink|Memo|**String**|**adLongVarWChar**|<p>LONGTEXT</p><p>(Synonyms: LONGCHAR, MEMO, NOTE)</p>|
| _Not supported_|Value|**Variant**|**adVariant**|VALUE (See Notes)|

|**Note**|
|:-----|  
|<ul><li>Access itself doesn't use the BINARY data type. It's recognized only for use in queries on linked tables from other database products that support the BINARY data type.</li><li>The INTEGER data type in the Access database engine SQL doesn't correspond to the **Integer** data type for table fields, query parameters, or Visual Basic. Instead, in SQL, the INTEGER data type corresponds to a **Long Integer** data type for table fields and query parameters and to a **Long** data type in Visual Basic.</li><li>The VALUE reserved word doesn't represent a data type defined by the Access database engine. However, in Access or SQL queries, the VALUE reserved word can be considered a valid synonym for the Visual Basic **Variant** data type.</li><li>If you are setting the data type for a DAO object in Visual Basic code, you must set the object's  **Type** property.</li></ul>|

