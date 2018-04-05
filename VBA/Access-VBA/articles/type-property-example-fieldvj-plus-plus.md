---
title: Type Property Example (Field) (VJ++)
ms.prod: access
ms.assetid: ee010723-d429-e190-e8e2-b1d8c2cfcb3a
ms.date: 06/08/2017
---


# Type Property Example (Field) (VJ++)

  

**Applies to:** Access 2013 | Access 2016

This example demonstrates the [Type](http://msdn.microsoft.com/library/14d99172-2145-05ae-620b-459ba097f05c%28Office.15%29.aspx) property by displaying the name of the constant that corresponds to the value of the **Type** property of all the[Field](http://msdn.microsoft.com/library/1dbd535e-48ad-a5c8-a1b2-6776c1e3e19d%28Office.15%29.aspx) objects in the ** _Employees_** table. The FieldType function is required for this procedure to run.




```c#
 
// BeginFieldTypeJ 
import java.io.*; 
import com.ms.wfc.data.*; 
 
public class TypeFieldX 
{ 
 
 // The main entry point of the application. 
 
 public static void main (String[] args) 
 { 
 TypeFieldX(); 
 System.exit(0); 
 } 
 
 // TypeFieldX Function 
 static void TypeFieldX() 
 { 
 // Define ADO Objects. 
 Recordset rstEmployees = null; 
 Field fldLoop = null; 
 
 // Declarations. 
 String strCnn = "Provider='sqloledb';Data Source='MySqlServer';"+ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';"; 
 int intLoop; 
 int intRecordCount = 0; 
 BufferedReader in = 
 new BufferedReader(new InputStreamReader(System.in)); 
 
 try 
 { 
 // Open the Recordset with data from Employees table. 
 rstEmployees = new Recordset(); 
 rstEmployees.open("employee", strCnn, 
 AdoEnums.CursorType.FORWARDONLY, AdoEnums.LockType.READONLY, 
 AdoEnums.CommandType.TABLE); 
 
 System.out.println("Fields in the Employees table:\n"); 
 
 // Enumerate fields collection of Employees table. 
 for(intLoop = 0;intLoop < 
 rstEmployees.getFields().getCount();intLoop++) 
 { 
 intRecordCount++; 
 // Loop needed for display of records 
 if((intRecordCount % 6)==0) 
 { 
 System.out.println("Press <Enter> to continue.."); 
 in.readLine(); 
 } 
 
 fldLoop = rstEmployees.getFields().getItem(intLoop); 
 System.out.println(" Name:" + fldLoop.getName() + "\n"+ 
 " Type:" + FieldType(fldLoop.getType()) + "\n"); 
 
 } 
 System.out.println("Press <Enter> to continue"); 
 in.readLine(); 
 } 
 catch(AdoException ae) 
 { 
 // Notify the user of any errors that result from ADO. 
 
 // As passing a Recordset, check for the null pointer first. 
 if(rstEmployees != null) 
 { 
 PrintProviderError(rstEmployees.getActiveConnection()); 
 } 
 else 
 { 
 System.out.println("Exception: " + ae.getMessage()); 
 } 
 } 
 // System read requires this catch. 
 catch(java.io.IOException je) 
 { 
 PrintIOError(je); 
 } 
 
 finally 
 { 
 // Cleanup objects before exit. 
 if (rstEmployees != null) 
 if (rstEmployees.getState() == 1) 
 rstEmployees.close(); 
 } 
 } 
 // FieldType Function 
 static String FieldType( int intType ) 
 { 
 String strLoop = null; 
 
 switch(intType) 
 { 
 case AdoEnums.DataType.CHAR: 
 strLoop = "adChar"; 
 break; 
 case AdoEnums.DataType.VARCHAR: 
 strLoop ="adVarChar"; 
 break; 
 case AdoEnums.DataType.SMALLINT: 
 strLoop = "adSmallInt"; 
 break; 
 case AdoEnums.DataType.UNSIGNEDTINYINT: 
 strLoop = "adUnsignedTinyInt" ; 
 break; 
 case AdoEnums.DataType.DBTIMESTAMP: 
 strLoop = "adDBTimeStamp"; 
 break; 
 default: 
 break; 
 } 
 
 return strLoop; 
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
// EndFieldTypeJ 

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

