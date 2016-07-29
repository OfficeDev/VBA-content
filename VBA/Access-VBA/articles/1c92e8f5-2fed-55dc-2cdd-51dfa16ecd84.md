
# DateCreated and DateModified Properties Example (VC++)

 **Last modified:** June 29, 2011

 _ **Applies to:** Access 2013 | Access 2016_

This example demonstrates the [DateCreated](ee975bf5-7d44-a993-d1c0-077993515698.md) and[DateModified](aebe8818-82e7-84a4-24d7-d97afa32e193.md) properties by adding a new[Column](ad38c2df-f704-0599-4b7a-8556e430ba46.md) to an existing[Table](53a3e2f9-4ec0-8fed-d482-4f995921587b.md) and by creating a new **Table**. The DateOutput procedure is required for this example to run.




```cpp
// BeginDateCreatedCpp 
#import "c:\Program Files\Common Files\system\ado\msadox.dll" no_namespace 
#import "c:\Program Files\Common Files\system\ado\msado15.dll" 
 
#include "iostream.h" 
#include "stdio.h" 
#include "conio.h" 
 
//Function declarations 
inline void TESTHR(HRESULT x) {if FAILED(x) _com_issue_error(x);}; 
void DateCreatedX(void); 
void DateOutPut(_bstr_t strTemp , _TablePtr tblTemp); 
 
////////////////////////////////////////////////////////// 
// // 
// Main Function // 
// // 
////////////////////////////////////////////////////////// 
void main() 
{ 
 if(FAILED(::CoInitialize(NULL))) 
 return; 
 
 DateCreatedX(); 
 
 ::CoUninitialize(); 
} 
 
////////////////////////////////////////////////////////// 
// // 
// DateCreatedX Function // 
// // 
////////////////////////////////////////////////////////// 
void DateCreatedX() 
{ 
 HRESULT hr = S_OK; 
 
 // Define ADOX object pointers. 
 // Initialize pointers on define. 
 // These are in the ADOX:: namespace. 
 _CatalogPtr m_pCatalog = NULL; 
 _TablePtr m_pTblEmployees = NULL; 
 _TablePtr m_pTblNew = NULL; 
 
 //Set ActiveConnection of Catalog to this string 
 _bstr_t strCnn("Provider='Microsoft.JET.OLEDB.4.0';" 
 "Data Source= 'c:\\Program Files\\Microsoft Office\\" 
 "Office\\Samples\\Northwind.mdb';"); 
 
 try 
 { 
 TESTHR(hr = m_pCatalog.CreateInstance(__uuidof (Catalog))); 
 
 // Connect the catalog. 
 m_pCatalog->PutActiveConnection(strCnn); 
 
 m_pTblEmployees = m_pCatalog->Tables->GetItem("Employees"); 
 
 // Print current information about the Employees table. 
 DateOutPut((_bstr_t)"Current properties", m_pTblEmployees); 
 
 // Create and append column to the Employees table. 
 m_pTblEmployees->Columns->Append("NewColumn", adInteger,0); 
 
 m_pCatalog->Tables->Refresh(); 
 
 // Print new information about the Employees table. 
 DateOutPut((_bstr_t)"After creating a new column", 
 m_pTblEmployees); 
 
 // Delete new column because this is a demonstration. 
 m_pTblEmployees->Columns->Delete("NewColumn"); 
 
 // Create and append new Table object to the Northwind database. 
 TESTHR(hr = m_pTblNew.CreateInstance(__uuidof(Table))); 
 
 m_pTblNew->Name = "NewTable"; 
 m_pTblNew->Columns->Append("NewColumn", adInteger,0); 
 m_pCatalog->Tables->Append(_variant_t((IDispatch*)m_pTblNew)); 
 m_pCatalog->Tables->Refresh(); 
 
 //Print information about the new Table object. 
 DateOutPut((_bstr_t)"After creating a new table", m_pCatalog-> 
 Tables->GetItem("NewTable")); 
 
 // Delete new Table object because this is a demonstration. 
 m_pCatalog->Tables->Delete(m_pTblNew->Name); 
 } 
 
 catch(_com_error &;e) 
 { 
 // Notify the user of errors if any. 
 _bstr_t bstrSource(e.Source()); 
 _bstr_t bstrDescription(e.Description()); 
 
 printf("\n\tSource : %s \n\tdescription : %s \n ", 
 (LPCSTR)bstrSource,(LPCSTR)bstrDescription); 
 } 
 
 catch(...) 
 { 
 cout << "Error occured in include files...."<< endl; 
 } 
} 
 
////////////////////////////////////////////////////////// 
// // 
// DateOutPut Function // 
// // 
////////////////////////////////////////////////////////// 
void DateOutPut(_bstr_t strTemp , _TablePtr tblTemp) 
{ 
 // Print DateCreated and DateModified information about 
 // specified Table object. 
 cout << strTemp << endl; 
 cout << " Table: " << tblTemp->GetName() << endl; 
 cout << " DateCreated = " << (_bstr_t)tblTemp-> 
 GetDateCreated() << endl; 
 cout << " DateModified = " << (_bstr_t)tblTemp-> 
 GetDateModified() << endl; 
} 
// EndDateCreatedCpp 

```

