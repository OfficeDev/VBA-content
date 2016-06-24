
# ADORecordConstruction Interface (ADO)

 **Last modified:** March 09, 2015

 _ **Applies to:** Access 2013 | Access 2016_

 **In this article**
[Properties](#sectionSection1)
[Methods](#sectionSection2)
[Events](#sectionSection3)
[Remarks](#sectionSection4)
[Requirements](#sectionSection5)



The  **ADORecordConstruction** interface is used to construct an ADO **Record** object from an OLE DB **Row** object in a C/C++ application.
This interface supports the following properties:

## Properties
<a name="sectionSection1"> </a>


|||
|:-----|:-----|
|[ParentRow](c7520353-9428-9c8f-9d21-ff42e30e1193.md)|Write-only. Sets the container of an OLE DB **Row** object on this ADO **Record** object.|
|[Row](1c2b0e27-7232-4b1c-826c-9dc15d758851.md)|Read/Write. Gets/sets an OLE DB **Row** object from/on this ADO **Record** object.|

## Methods
<a name="sectionSection2"> </a>

None.


## Events
<a name="sectionSection3"> </a>

None.


## Remarks
<a name="sectionSection4"> </a>

Given an OLE DB  **Row** object ( `pRow`), the construction of an ADO  **Record** object (), the construction of an ADO **Record** object ( `adoR`), amounts to the following three basic operations:


1. Create an ADO  **Record** object:
    
```
  _RecordPtr adoR;
adoRs.CreateInstance(__uuidof(_Record));

```

2. Query the  **IADORecordConstruction** interface on the **Record** object:
    
```
  adoRecordConstructionPtr adoRConstruct=NULL;
adoR->QueryInterface(__uuidof(ADORecordConstruction),
                    (void**)&;adoRConstruct);

```

3. Call the  **IADORecordConstruction::put_Row** property method to set the OLE DB **Row** object on the ADO **Record** object:
    
```cpp
  IUnknown *pUnk=NULL;
pRow->QueryInterface(IID_IUnknown, (void**)&;pUnk);
adoRConstruct->put_Row(pUnk);

```

The resultant  **adoR** object now represents the ADO **Record** object constructed from the OLE DB **Row** object.

An ADO  **Record** object can also be constructed from the container of an OLE DB **Row** object.


## Requirements
<a name="sectionSection5"> </a>

 **Version:** ADO 2.0 and later

 **Library:** msado15.dll

 **UUID:** 00000567-0000-0010-8000-00AA006D2EA4

