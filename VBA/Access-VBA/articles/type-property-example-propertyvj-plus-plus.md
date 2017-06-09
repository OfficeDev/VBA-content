---
title: Type Property Example (Property) (VJ++)
ms.prod: access
ms.assetid: 405f1769-f661-24e7-22db-0c725ee55576
ms.date: 06/08/2017
---


# Type Property Example (Property) (VJ++)

  

**Applies to:** Access 2013 | Access 2016

This example demonstrates the [Type](http://msdn.microsoft.com/library/14d99172-2145-05ae-620b-459ba097f05c%28Office.15%29.aspx) property. It is a model of a utility for listing the names and types of a collection, like[Properties](http://msdn.microsoft.com/library/4d662790-1252-c930-e6f9-edf6a38636af%28Office.15%29.aspx), [Fields](http://msdn.microsoft.com/library/029aa738-8726-54a6-1813-b152813948bc%28Office.15%29.aspx), etc.

We do not need to open the [Recordset](http://msdn.microsoft.com/library/0f963bf8-f066-dc8a-b754-f427de712df1%28Office.15%29.aspx) to access its **Properties** collection; they come into existence when the **Recordset** object is instantiated. However, setting the[CursorLocation](http://msdn.microsoft.com/library/8a048bd4-ae25-a555-1c07-14364b7e6560%28Office.15%29.aspx) property to **adUseClient** adds several dynamic properties to the **Recordset** object's **Properties** collection, making the example a little more interesting. For sake of illustration, we explicitly use the[Item](http://msdn.microsoft.com/library/793c305f-0e5b-a529-e21f-b7ab0843ed49%28Office.15%29.aspx) property to access each[Property](http://msdn.microsoft.com/library/eec318fd-f5ed-d9ef-9830-848439a8914d%28Office.15%29.aspx) object.



```c#
 
// BegintTypePropertyJ 
import com.ms.wfc.data.*; 
import java.io.* ; 
 
public class TypePropertyX 
{ 
 // The main entry point for the application. 
 
 public static void main (String[] args) 
 { 
 TypePropertyX(); 
 System.exit(0); 
 } 
 
 // TypePropertyX function 
 static void TypePropertyX() 
 { 
 // Define ADO Objects. 
 Recordset rst = null; 
 AdoProperty prop = null; 
 
 // Declarations. 
 BufferedReader in = 
 new BufferedReader (new InputStreamReader(System.in)); 
 String strCnn = "DSN='Pubs';Provider='MSDASQL';Integrated Security='SSPI';"; 
 String strMsg; 
 int intIndex; 
 int intDisplaysize = 15; 
 
 try 
 { 
 rst = new Recordset(); 
 rst.setCursorLocation(AdoEnums.CursorLocation.CLIENT); 
 for(intIndex = 0; 
 intIndex <= rst.getProperties().getCount() - 1;intIndex++) 
 { 
 prop = rst.getProperties().getItem(intIndex); 
 switch(prop.getType()) 
 { 
 case AdoEnums.DataType.BIGINT : 
 strMsg = "adBigInt"; 
 break; 
 case AdoEnums.DataType.BINARY : 
 strMsg = "adBinary"; 
 break; 
 case AdoEnums.DataType.BOOLEAN : 
 strMsg = "adBoolean"; 
 break; 
 case AdoEnums.DataType.BSTR : 
 strMsg = "adBSTR"; 
 break; 
 case AdoEnums.DataType.CHAPTER : 
 strMsg = "adChapter"; 
 break; 
 case AdoEnums.DataType.CHAR : 
 strMsg = "adChar"; 
 break; 
 case AdoEnums.DataType.CURRENCY : 
 strMsg = "adCurrency"; 
 break; 
 case AdoEnums.DataType.DATE : 
 strMsg = "adDate"; 
 break; 
 case AdoEnums.DataType.DBDATE : 
 strMsg = "adDBDate"; 
 break; 
 case AdoEnums.DataType.DBTIME : 
 strMsg = "adDBTime"; 
 break; 
 case AdoEnums.DataType.DBTIMESTAMP : 
 strMsg = "adDBTimeStamp"; 
 break; 
 case AdoEnums.DataType.DECIMAL : 
 strMsg = "adDecimal"; 
 break; 
 case AdoEnums.DataType.DOUBLE : 
 strMsg = "adDouble"; 
 break; 
 case AdoEnums.DataType.EMPTY : 
 strMsg = "adEmpty"; 
 break; 
 case AdoEnums.DataType.ERROR : 
 strMsg = "adError"; 
 break; 
 case AdoEnums.DataType.FILETIME : 
 strMsg = "adFileTime"; 
 break; 
 case AdoEnums.DataType.GUID : 
 strMsg = "adGUID"; 
 break; 
 case AdoEnums.DataType.IDISPATCH : 
 strMsg = "adIDispatch"; 
 break; 
 case AdoEnums.DataType.INTEGER : 
 strMsg = "adInteger"; 
 break; 
 case AdoEnums.DataType.IUNKNOWN : 
 strMsg = "adIUnknown"; 
 break; 
 case AdoEnums.DataType.LONGVARBINARY : 
 strMsg = "adLongVarBinary"; 
 break; 
 case AdoEnums.DataType.LONGVARCHAR : 
 strMsg = "adLongVarChar"; 
 break; 
 case AdoEnums.DataType.LONGVARWCHAR : 
 strMsg = "adLongVarWChar"; 
 break; 
 case AdoEnums.DataType.NUMERIC : 
 strMsg = "adNumeric"; 
 break; 
 case AdoEnums.DataType.PROPVARIANT : 
 strMsg = "adPropVariant"; 
 break; 
 case AdoEnums.DataType.SINGLE : 
 strMsg = "adSingle"; 
 break; 
 case AdoEnums.DataType.SMALLINT : 
 strMsg = "adSmallInt"; 
 break; 
 case AdoEnums.DataType.TINYINT : 
 strMsg = "adTinyInt"; 
 break; 
 case AdoEnums.DataType.UNSIGNEDBIGINT : 
 strMsg = "adUnsignedBigInt"; 
 break; 
 case AdoEnums.DataType.UNSIGNEDINT : 
 strMsg = "adUnsignedInt"; 
 break; 
 case AdoEnums.DataType.UNSIGNEDSMALLINT : 
 strMsg = "adUnsignedSmallInt"; 
 break; 
 case AdoEnums.DataType.UNSIGNEDTINYINT : 
 strMsg = "adUnsignedTinyInt"; 
 break; 
 case AdoEnums.DataType.USERDEFINED : 
 strMsg = "adUserDefined"; 
 break; 
 case AdoEnums.DataType.VARBINARY : 
 strMsg = "adVarBinary"; 
 break; 
 case AdoEnums.DataType.VARCHAR : 
 strMsg = "adVarChar"; 
 break; 
 case AdoEnums.DataType.VARIANT : 
 strMsg = "adVariant"; 
 break; 
 case AdoEnums.DataType.VARNUMERIC : 
 strMsg = "adVarNumeric"; 
 break; 
 case AdoEnums.DataType.VARWCHAR : 
 strMsg = "adVarWChar"; 
 break; 
 case AdoEnums.DataType.WCHAR : 
 strMsg = "adWChar"; 
 break; 
 default: 
 strMsg = "*UNKNOWN*"; 
 break; 
 } 
 System.out.println("Property " + 
 Integer.toString(intIndex) + 
 " : " + 
 prop.getName() + 
 ", Type = " + 
 strMsg); 
 if(intIndex % intDisplaysize == 0 &;&; intIndex != 0) 
 { 
 System.out.println("\nPress <Enter> to continue.."); 
 in.readLine(); 
 } 
 } 
 
 System.out.println("\nPress <Enter> to continue.."); 
 in.readLine(); 
 } 
 catch( AdoException ae ) 
 { 
 // Notify user of any errors that result from ADO. 
 
 // As passing a Recordset, check for null pointer first. 
 if (rst != null) 
 { 
 PrintProviderError(rst.getActiveConnection()); 
 } 
 else 
 { 
 System.out.println("Exception: " + ae.getMessage()); 
 } 
 } 
 
 // System read requires this catch. 
 catch( java.io.IOException je) 
 { 
 PrintIOError(je); 
 } 
 
 finally 
 { 
 // Cleanup objects before exit. 
 if (rst != null) 
 if (rst.getState() == 1) 
 rst.close(); 
 } 
 } 
 
 // PrintProviderError Function 
 static void PrintProviderError(Connection cnn1) 
 { 
 // Print Provider Errors from Connection Object. 
 // ErrItem is an item object in the Connections Errors Collection. 
 com.ms.wfc.data.Error ErrItem = null; 
 long nCount = 0; 
 int i = 0; 
 
 nCount = cnn1.getErrors().getCount(); 
 
 // If there are any errors in the collection, print them. 
 if ( nCount > 0) 
 { 
 // Collection ranges from 0 to nCount-1. 
 for ( i=0;i<nCount; i++) 
 { 
 ErrItem = cnn1.getErrors().getItem(i); 
 System.out.println("\t Error Number: " + ErrItem.getNumber() 
 + "\t" + ErrItem.getDescription()); 
 } 
 } 
 } 
 // PrintIOError Function 
 static void PrintIOError(java.io.IOException je) 
 { 
 System.out.println("Error: \n"); 
 System.out.println("\t Source: " + je.getClass() + "\n"); 
 System.out.println("\t Description: "+ je.getMessage() + "\n"); 
 } 
} 
// EndTypePropertyJ 

```

 **ACCESS SUPPORT RESOURCES**<br>
[Access for developers forum on MSDN](https://social.msdn.microsoft.com/Forums/office/en-US/home?forum=accessdev)<br>
[Access help on support.office.com](https://support.office.com/search/results?query=Access)<br>
[Access help on answers.microsoft.com](http://answers.microsoft.com/en-us/office/forum/access?page=1&;tab=question&;status=all&;auth=1)<br>
[Search for specific Access error codes on Bing](http://www.bing.com/)<br>
[Access forums on UtterAccess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access wiki on UtterAcess](http://www.utteraccess.com/forum/index.php?act=idx)<br>
[Access developer and VBA programming help center (FMS)](http://www.fmsinc.com/MicrosoftAccess/developer/)<br>
[Access posts on StackOverflow](http://stackoverflow.com/questions/tagged/ms-access)

