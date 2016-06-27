
# Connection Close Method, Table Type Property Example (VC++)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

Setting the [ActiveConnection](c1d90eca-9d62-4d7e-c275-5094e914ecb4.md) property to **Nothing** should "close" the catalog. Associated collections will be empty. Any objects that were created from schema objects in the catalog will be orphaned. Any properties on those objects that have been cached will still be available, but attempting to read properties that require a call to the provider will fail.




```cpp
// BeginCloseConnectionCpp 
#import "c:\Program Files\Common Files\system\ado\msadox.dll" no_namespace 
#import "c:\Program Files\Common Files\system\ado\msado15.dll" 
 
#include "iostream.h" 
#include "stdio.h" 
#include "conio.h" 
 
//Function declarations 
inline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);}; 
void CloseConnectionByNothingX(void); 
void CloseConnectionX(void); 
 
////////////////////////////////////////////////////////// 
// // 
// Main Function // 
// // 
////////////////////////////////////////////////////////// 
void main() 
{ 
 if(FAILED(::CoInitialize(NULL))) 
 return; 
 
 CloseConnectionByNothingX(); 
 CloseConnectionX(); 
 
 ::CoUninitialize(); 
} 
 
////////////////////////////////////////////////////////// 
// // 
// CloseConnectionByNothingX Function // 
// // 
////////////////////////////////////////////////////////// 
void CloseConnectionByNothingX(void) 
{ 
 HRESULT hr = S_OK; 
 
 // Define ADOX object pointers. 
 // Initialize pointers on define. 
 // These are in the ADOX:: namespace. 
 
 _CatalogPtr m_pCatalog = NULL; 
 _TablePtr m_pTable = NULL; 
 
 //Define ADODB object pointers 
 ADODB::_ConnectionPtr m_pCnn = NULL; 
 
 //Define other variables 
 _variant_t vIndex = (short) 0; 
 
 try 
 { 
 TESTHR(hr = m_pCnn.CreateInstance(__uuidof(ADODB::Connection))); 
 
 TESTHR(hr = m_pCatalog.CreateInstance(__uuidof(Catalog))); 
 
 m_pCnn->Open("Provider='Microsoft.JET.OLEDB.4.0';" 
 "Data Source= 'c:\\Program Files\\Microsoft Office\\" 
 "Office\\Samples\\Northwind.mdb';","","",NULL); 
 
 m_pCatalog->PutActiveConnection(_variant_t((IDispatch *)m_pCnn)); 
 
 m_pTable = m_pCatalog->Tables->GetItem(vIndex); 
 
 // Cache m_pTable.Type info 
 cout << m_pTable->Type << endl; 
 
 _variant_t vCnn; 
 vCnn.vt = VT_DISPATCH; 
 vCnn.pdispVal = NULL; 
 m_pCatalog->PutActiveConnection(vCnn); 
 
 // m_pTable is orphaned 
 cout << m_pTable->Type << endl; 
 // Previous line will succeed if this was cached 
 
 cout << m_pTable->Columns->GetItem(vIndex)->DefinedSize << endl; 
 // Previous line will fail if this info has not been cached 
 } 
 catch(_com_error &;e) 
 { 
 // Notify the user of errors if any. 
 _bstr_t bstrSource(e.Source()); 
 _bstr_t bstrDescription(e.Description()); 
 
 printf("\nError\n\tSource : %s \n\tdescription : %s \n", 
 (LPCSTR)bstrSource,(LPCSTR)bstrDescription); 
 } 
 catch(...) 
 { 
 cout << "Error occured in CloseConnectionByNothingX...."<< endl; 
 } 
} 
 
////////////////////////////////////////////////////////// 
// // 
// CloseConnectionX Function // 
// // 
////////////////////////////////////////////////////////// 
void CloseConnectionX() 
{ 
 HRESULT hr = S_OK; 
 
 // Define ADOX object pointers. 
 // Initialize pointers on define. 
 // These are in the ADOX:: namespace. 
 
 _CatalogPtr m_pCatalog = NULL; 
 _TablePtr m_pTable = NULL; 
 
 //Define ADODB object pointers 
 ADODB::_ConnectionPtr m_pCnn = NULL; 
 
 //Define other variables 
 _variant_t vIndex = (short) 0; 
 try 
 { 
 TESTHR(hr = m_pCnn.CreateInstance(__uuidof(ADODB::Connection))); 
 m_pCnn->Open("Provider='Microsoft.JET.OLEDB.4.0';" 
 "Data Source= 'c:\\Program Files\\Microsoft Office\\" 
 "Office\\Samples\\Northwind.mdb';","","",NULL); 
 
 TESTHR(hr = m_pCatalog.CreateInstance(__uuidof(Catalog))); 
 m_pCatalog->PutActiveConnection(_variant_t((IDispatch *)m_pCnn)); 
 
 m_pTable = m_pCatalog->Tables->GetItem(vIndex); 
 
 // Cache m_pTable.Type info 
 cout << m_pTable->Type << endl; 
 
 m_pCnn->Close(); 
 
 // m_pTable is orphaned 
 cout << m_pTable->Type << endl; 
 // Previous line will succeed if this was cached 
 
 cout << m_pTable->Columns->GetItem(vIndex)->DefinedSize << endl; 
 // Previous line will fail if this info has not been cached 
 } 
 catch(_com_error &;e) 
 { 
 // Notify the user of errors if any. 
 _bstr_t bstrSource(e.Source()); 
 _bstr_t bstrDescription(e.Description()); 
 
 printf("\nError\n\tSource : %s \n\tdescription : %s \n", 
 (LPCSTR)bstrSource,(LPCSTR)bstrDescription); 
 } 
 catch(...) 
 { 
 cout << "Error occured in CloseConnectionX...."<< endl; 
 } 
} 
// EndCloseConnectionCpp 

```

