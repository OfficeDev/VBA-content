
# Status Property Example (VJ++)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

This example uses the [Status](bf3ccb36-c985-5fae-4f76-c48a0e20e6f7.md) property to display which records have been modified in a batch operation before a batch update has occurred.




```java
// BeginStatusJ 
import java.io.*; 
import com.ms.wfc.data.*; 
 
public class StatusX 
{ 
 // The main entry point of the application. 
 
 public static void main (String[] args) 
 { 
 StatusX(); 
 System.exit(0); 
 } 
 // StatusX Function 
 
 static void StatusX() 
 { 
 // Define ADO Objects. 
 Recordset rstTitles = null; 
 
 // Declarations. 
 String strCnn = "Provider='sqloledb';Data Source='MySqlServer';"+ 
 "Initial Catalog='Pubs';Integrated Security='SSPI';"; 
 BufferedReader in = 
 new BufferedReader(new InputStreamReader(System.in)); 
 
 try 
 { 
 // Open Recordset for batch update. 
 rstTitles = new Recordset(); 
 rstTitles.setCursorType(AdoEnums.CursorType.KEYSET); 
 rstTitles.setLockType(AdoEnums.LockType.BATCHOPTIMISTIC); 
 rstTitles.open("Titles", strCnn, AdoEnums.CursorType.KEYSET, 
 AdoEnums.LockType.BATCHOPTIMISTIC, 
 AdoEnums.CommandType.TABLE); 
 
 // Change the type of psychology titles. 
 while(!rstTitles.getEOF()) 
 { 
 if(rstTitles.getField("Type").getString().trim(). 
 equals(new String("psychology"))) 
 rstTitles.getField("Type").setString("self_help"); 
 
 rstTitles.moveNext(); 
 } 
 
 // Display Title ID and status. 
 rstTitles.moveFirst(); 
 
 while(!rstTitles.getEOF()) 
 { 
 if(rstTitles.getStatus()==AdoEnums.RecordStatus.MODIFIED) 
 System.out.println(rstTitles.getField("title_id"). 
 getString() + "- Modified"); 
 else 
 System.out.println(rstTitles.getField("title_id"). 
 getString()); 
 rstTitles.moveNext(); 
 } 
 
 // Cancel the update because this is a demonstration. 
 rstTitles.cancelBatch(); 
 
 System.out.println("Press <Enter> to continue.."); 
 in.readLine(); 
 } 
 catch(AdoException ae) 
 { 
 // Notify the user of any errors that result from ADO. 
 
 // As passing a Recordset, check for the null pointer first. 
 if(rstTitles != null) 
 { 
 PrintProviderError(rstTitles.getActiveConnection()); 
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
 if (rstTitles != null) 
 if (rstTitles.getState() == 1) 
 rstTitles.close(); 
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
// EndStatusJ 

```

