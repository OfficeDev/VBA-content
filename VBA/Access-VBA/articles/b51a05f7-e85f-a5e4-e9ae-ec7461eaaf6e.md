
# OriginalValue and UnderlyingValue Properties Example (VC++)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

This example demonstrates the [OriginalValue](02ffc728-4692-d439-e2a6-2f02cca53a71.md) and[UnderlyingValue](f84f4c1c-2bd4-a725-3575-ed063ead13c8.md) properties by displaying a message if a record's underlying data has changed during a[Recordset](0f963bf8-f066-dc8a-b754-f427de712df1.md) batch update.




```cpp
// BeginOriginalValueCpp 
#import "c:\Program Files\Common Files\System\ADO\msado15.dll" no_namespace rename("EOF", "EndOfFile") 
 
#include <ole2.h> 
#include <stdio.h> 
#include <conio.h> 
#include "OriginalValueX.h" 
 
// Function declarations 
inline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);}; 
void OriginalValueX(void); 
void PrintProviderError(_ConnectionPtr pConnection); 
void PrintComError(_com_error &;e); 
 
/////////////////////////////////////////////////////////// 
// // 
// Main Function // 
// // 
/////////////////////////////////////////////////////////// 
 
void main() 
{ 
 if(FAILED(::CoInitialize(NULL))) 
 return; 
 
 OriginalValueX(); 
 
 ::CoUninitialize(); 
} 
 
/////////////////////////////////////////////////////////// 
// // 
// OriginalValueX Function // 
// // 
/////////////////////////////////////////////////////////// 
 
void OriginalValueX(void) 
{ 
 // Define ADO object pointers. 
 // Initialize pointers on define. 
 // These are in the ADODB:: namespace. 
 _ConnectionPtr pConnection = NULL; 
 FieldPtr pFldType = NULL; 
 _RecordsetPtr pRstTitles = NULL; 
 
 // Define string variables. 
 _bstr_t strSQLChange("UPDATE Titles SET Type = " 
 "'sociology' WHERE Type = 'psychology'"); 
 _bstr_t strSQLRestore("UPDATE Titles SET Type = " 
 "'psychology' WHERE Type = 'sociology'"); 
 
 // Define Other Variables 
 HRESULT hr = S_OK; 
 IADORecordBinding *picRs = NULL; //Interface Pointer declared 
 CTitleRs titlers; // C++ Class object 
 
 try 
 { 
 _bstr_t strCnn("Provider='sqloledb';Data Source='MySqlServer';" 
 "Initial Catalog='pubs';Integrated Security='SSPI';"); 
 
 // Open connection. 
 TESTHR(pConnection.CreateInstance(__uuidof(Connection))); 
 pConnection->Open (strCnn, "", "", adConnectUnspecified); 
 
 // Open Recordset for batch update. 
 TESTHR(pRstTitles.CreateInstance(__uuidof(Recordset))); 
 pRstTitles->PutActiveConnection( 
 _variant_t((IDispatch *)pConnection,true)); 
 pRstTitles->CursorType = adOpenKeyset; 
 pRstTitles->LockType = adLockBatchOptimistic; 
 
 // Cast Connection pointer to an IDispatch type so converted 
 // to correct type of variant. 
 pRstTitles->Open("Titles", 
 _variant_t((IDispatch *)pConnection,true), 
 adOpenKeyset, adLockBatchOptimistic, adCmdTable); 
 
 //Open an IADORecordBinding interface pointer which 
 //we'll use for Binding Recordset to a class. 
 TESTHR(pRstTitles->QueryInterface( 
 __uuidof(IADORecordBinding),(LPVOID*)&;picRs)); 
 
 //Bind the Recordset to a C++ Class here 
 TESTHR(picRs->BindToRecordset(&;titlers)); 
 
 // Set field object variable for Type field. 
 pFldType = pRstTitles->Fields->GetItem("type"); 
 
 // Change the type of psychology titles. 
 while(!(pRstTitles->EndOfFile)) 
 { 
 if (!strcmp(strtok((char *)titlers.m_szau_Type," "), 
 "psychology")) 
 { 
 pFldType->Value = "self_help"; 
 } 
 pRstTitles->MoveNext(); 
 } 
 
 // Simulate a change by another user by updating data 
 // using a command string. 
 pConnection->Execute(strSQLChange,NULL,0); 
 
 // Check for changes. 
 pRstTitles->MoveFirst(); 
 while(!(pRstTitles->EndOfFile)) 
 { 
 if (strcmp(pFldType->OriginalValue.pcVal, 
 pFldType->UnderlyingValue.pcVal)) 
 { 
 printf("\n\nData has changed!"); 
 
 printf("\n\nTitle ID: %s",titlers.lau_Title_idStatus == 
 adFldOK ? titlers.m_szau_Title_id : "<NULL>"); 
 
 printf("\n\nCurrent Value: %s", 
 (LPCSTR) (_bstr_t) pFldType->Value); 
 
 printf("\n\nOriginal Value: %s", 
 (LPCSTR) (_bstr_t) pFldType->OriginalValue); 
 
 printf("\n\nUnderlying Value: %s\n\n", 
 (LPCSTR) (_bstr_t) pFldType->UnderlyingValue); 
 
 printf("Press any key to continue..."); 
 getch(); 
 
 system("cls"); 
 } 
 pRstTitles->MoveNext(); 
 } 
 } 
 catch (_com_error &;e) 
 { 
 // Notify the user of errors if any. 
 // Pass a connection pointer accessed from the Connection. 
 PrintProviderError(pConnection); 
 PrintComError(e); 
 } 
 
 // Clean up objects before exit. 
 //Release the IADORecordset Interface here 
 if (picRs) 
 picRs->Release(); 
 
 if (pRstTitles) 
 if (pRstTitles->State == adStateOpen) 
 { 
 // Cancel the update because this is a demonstration. 
 pRstTitles->CancelBatch(adAffectAll); 
 pRstTitles->Close(); 
 } 
 if (pConnection) 
 if (pConnection->State == adStateOpen) 
 { 
 // Restore Original Values. 
 pConnection->Execute(strSQLRestore,NULL,0); 
 pConnection->Close(); 
 } 
}; 
 
/////////////////////////////////////////////////////////// 
// // 
// PrintProviderError Function // 
// // 
/////////////////////////////////////////////////////////// 
void PrintProviderError(_ConnectionPtr pConnection) 
{ 
 // Print Provider Errors from Connection object. 
 // pErr is a record object in the Connection's Error collection. 
 ErrorPtr pErr = NULL; 
 
 if( (pConnection->Errors->Count) > 0) 
 { 
 long nCount = pConnection->Errors->Count; 
 
 // Collection ranges from 0 to nCount -1. 
 for(long i = 0;i < nCount;i++) 
 { 
 pErr = pConnection->Errors->GetItem(i); 
 printf("\t Error number: %x\t%s", pErr->Number, 
 pErr->Description); 
 } 
 } 
} 
 
/////////////////////////////////////////////////////////// 
// // 
// PrintComError Function // 
// // 
/////////////////////////////////////////////////////////// 
void PrintComError(_com_error &;e) 
{ 
 _bstr_t bstrSource(e.Source()); 
 _bstr_t bstrDescription(e.Description()); 
 
 // Print COM errors. 
 printf("Error\n"); 
 printf("\tCode = %08lx\n", e.Error()); 
 printf("\tCode meaning = %s\n", e.ErrorMessage()); 
 printf("\tSource = %s\n", (LPCSTR) bstrSource); 
 printf("\tDescription = %s\n", (LPCSTR) bstrDescription); 
} 
// EndOriginalValueCpp 

```

 **OriginalValueX.h**



```c#
cpp
// BeginOriginalValueH 
#include "icrsint.h" 
 
//This class extracts title_id and type from titles table 
class CTitleRs : public CADORecordBinding 
{ 
BEGIN_ADO_BINDING(CTitleRs) 
 // Column title_id is the 1st field in the Recordset 
 ADO_VARIABLE_LENGTH_ENTRY2(1, adVarChar, m_szau_Title_id, 
 sizeof(m_szau_Title_id), lau_Title_idStatus, FALSE) 
 
 // Column type is the 3rd field in the Recordset 
 ADO_VARIABLE_LENGTH_ENTRY2(3, adVarChar, m_szau_Type, 
 sizeof(m_szau_Type), lau_TypeStatus, TRUE) 
END_ADO_BINDING() 
 
public: 
 CHAR m_szau_Title_id[7]; 
 ULONG lau_Title_idStatus; 
 CHAR m_szau_Type[13]; 
 ULONG lau_TypeStatus; 
}; 
// EndOriginalValueH 

```

