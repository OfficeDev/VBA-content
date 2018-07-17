---
title: Application.HyperlinkPart Method (Access)
keywords: vbaac10.chm12569
f1_keywords:
- vbaac10.chm12569
ms.prod: access
api_name:
- Access.Application.HyperlinkPart
ms.assetid: 011665ea-c650-fab3-a736-f26a3de1b65e
ms.date: 06/08/2017
---


# Application.HyperlinkPart Method (Access)

The  **HyperlinkPart** method returns information about data stored as a Hyperlink data type. .


## Syntax

 _expression_. **HyperlinkPart**( ** _Hyperlink_**, ** _Part_** )

 _expression_ A variable that represents an **Application** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
| _Hyperlink_|Required|**Variant**|The data stored in a Hyperlink field.|
| _Part_|Optional|**AcHyperlinkPart**|A  **[AcHyperlinkPart](achyperlinkpart-enumeration-access.md)** constant representing the information you want returned by the **HyperlinkPart** method.|

### Return Value

String


## Remarks

You use the  **HyperlinkPart** method to return one of three values from a Hyperlink field or the displayed value. The value returned depends on the setting of the _part_ argument. The _part_ argument is optional. If it's not used, the function returns the value Microsoft Access displays for the hyperlink (which corresponds to the **acDisplayedValue** setting for the _part_ argument). The returned values can be one of the four parts of the Hyperlink field ( _displaytext_,  _address_,  _subaddress_, or  _screentip_), the full address,  _address_# _subaddress_, or the value Microsoft Access displays for the hyperlink.


 **Note**  If you use the  **HyperlinkPart** method in a query, the _part_ argument is required and you can't use the constants listed above but must use the actual value instead.

When a value is provided in the  _displaytext_ part of a Hyperlink field, the value displayed by Microsoft Access will be the same as the _displaytext_ setting. When there's no value in the _displaytext_ part of a Hyperlink field, the value displayed will be the _address_ or _subaddress_ part of the Hyperlink field, depending on which value is first present in the field.

The following table shows the values returned by the  **HyperlinkPart** method for data stored in a Hyperlink field.



|**Hyperlink field data**|**HyperlinkPart method returned values**|
|:-----|:-----|
|#http://www.microsoft.com#|**acDisplayedValue**: http://www.microsoft.com **acDisplayText**: **acAddress**: http://www.microsoft.com **acSubAddress**: **acScreenTip**: **acFullAddress**: http://www.microsoft.com|
|Microsoft#http://www.microsoft.com#|**acDisplayedValue**: Microsoft **acDisplayText**: Microsoft **acAddress**: http://www.microsoft.com **acSubAddress**: **acScreenTip**: **acFullAddress**: http://www.microsoft.com|
|Customers#http://www.microsoft.com#Form Customers|**acDisplayedValue**: Customers **acDisplayText**: Customers **acAddress**: http://www.microsoft.com **acSubAddress**: Form Customers **acScreenTip**: **acFullAddress**: http://www.microsoft.com#Form Customer|
|##Form Customers#Enter Information|**acDisplayedValue**: Form Customers **acDisplayText**: **acAddress**: **acSubAddress**: Form Customers **acScreenTip**: Enter Information **acFullAddress**: #FormCustomer|
When you add an  _address_ part to a Hyperlink field by using the **Insert Hyperlink** dialog box (available by clicking **Hyperlink** on the **Insert** menu) or by typing an address part directly into a Hyperlink field, Microsoft Access adds the two # symbols that delimit parts of the hyperlink data.

You can add or edit the  _displaytext_ part of a hyperlink field by right-clicking a hyperlink in a table, form, or report, pointing to **Hyperlink** on the shortcut menu, and then typing the display text in the **Text to display** box.

When you add data to a Hyperlink field directly, you must include the two # symbols to delimit the parts of the hyperlink data.


## Example

The following example uses all four of the  _part_ argument constants to display information returned by the **HyperlinkPart** method for each record in a table containing a Hyperlink field. To try this example, paste the DisplayHyperlinkParts procedure into the Declarations section of a module. You can call the DisplayHyperlinkParts procedure from the Debug window, passing to it the name of a table containing hyperlinks and the name of the field containing Hyperlink data. For example:


```vb
DisplayHyperlinkParts "MyHyperlinkTableName", "MyHyperlinkFieldName" 
 
Public Sub DisplayHyperlinkParts(ByVal strTable As String, _ 
 ByVal strField As String) 
 
 Dim rst As New ADODB.Recordset 
 Dim strMsg As String 
 
 
 rst.Open strTable, CurrentProject.Connection, _ 
 adOpenForwardOnly, adLockReadOnly 
 
 ' For each record in table. 
 Do Until rst.EOF 
 strMsg = "DisplayValue = " _ 
 &; HyperlinkPart(rst(strField), acDisplayedValue) _ 
 &; vbCrLf &; "DisplayText = " _ 
 &; HyperlinkPart(rst(strField), acDisplayText) _ 
 &; vbCrLf &; "Address = " _ 
 &; HyperlinkPart(rst(strField), acAddress) _ 
 &; vbCrLf &; "SubAddress = " _ 
 &; HyperlinkPart(rst(strField), acSubAddress) _ 
 &; vbCrLf &; "ScreenTip = " _ 
 &; HyperlinkPart(rst(strField), acScreenTip) _ 
 &; vbCrLf &; "Full Address = " _ 
 &; HyperlinkPart(rst(strField), acFullAddress) 
 
 ' Show parts returned by HyperlinkPart function. 
 MsgBox strMsg 
 rst.MoveNext 
 Loop 
 
End Sub
```

When you use the  **HyperlinkPart** method in a query, the _part_ argument is required. For example, the following SQL statement uses the **HyperlinkPart** method to return information about data stored as a Hyperlink data type in the URL field of the Links table:




```sql
SELECT Links.URL, HyperlinkPart([URL],0) 
 AS Display, HyperlinkPart([URL],1) 
 AS Name, HyperlinkPart([URL],2) 
 AS Addr, HyperlinkPart([URL],3) 
 AS SubAddr, HyperlinkPart([URL],4) 
 AS ScreenTip 
 FROM Links
```


## See also


#### Concepts


[Application Object](application-object-access.md)

