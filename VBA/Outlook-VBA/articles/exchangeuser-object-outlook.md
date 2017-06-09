---
title: ExchangeUser Object (Outlook)
keywords: vbaol11.chm3158
f1_keywords:
- vbaol11.chm3158
ms.prod: outlook
api_name:
- Outlook.ExchangeUser
ms.assetid: 6ec117d1-7fdb-aa36-b567-1242f8238df0
ms.date: 06/08/2017
---


# ExchangeUser Object (Outlook)

Provides detailed information about an  **[AddressEntry](addressentry-object-outlook.md)** that represents a Microsoft Exchange mailbox user.


## Remarks

 **ExchangeUser** is derived from the **AddressEntry** object, and is returned instead of an **AddressEntry** when the caller performs a query interface on the **AddressEntry** object.

This object provides first-class access to properties applicable to Exchange users such as  **[FirstName](http://msdn.microsoft.com/library/6a72812a-31fd-aa6a-be08-f765018208ab%28Office.15%29.aspx)**, **[JobTitle](http://msdn.microsoft.com/library/2cfa5301-3164-c472-3f8e-831c1eebc810%28Office.15%29.aspx)**, **[LastName](http://msdn.microsoft.com/library/1f9f9675-3e72-da56-d654-a1473f4f71a7%28Office.15%29.aspx)**, and **[OfficeLocation](http://msdn.microsoft.com/library/b37d5622-27ba-b2c4-cfd3-6aa1e9e9296b%28Office.15%29.aspx)**. You can also access other properties specific to the Exchange user that are not exposed in the object model through the **[PropertyAccessor](propertyaccessor-object-outlook.md)** object. Note that some of the explicit built-in properties are read-write properties. Setting these properties requires the code to be running under an appropriate Exchange administrator account; without sufficient permissions, calling the **[ExchangeUser.Update](http://msdn.microsoft.com/library/a2672fbf-f32a-f120-227c-24ee5c361f35%28Office.15%29.aspx)** method will result in a "permission denied" error.


## Example

The following code sample shows how to obtain the business phone number, office location, and job title for all entries in the Exchange Global Address List.


```
Sub DemoAE() 
 
 Dim colAL As Outlook.AddressLists 
 
 Dim oAL As Outlook.AddressList 
 
 Dim colAE As Outlook.AddressEntries 
 
 Dim oAE As Outlook.AddressEntry 
 
 Dim oExUser As Outlook.ExchangeUser 
 
 Set colAL = Application.Session.AddressLists 
 
 For Each oAL In colAL 
 
 'Address list is an Exchange Global Address List 
 
 If oAL.AddressListType = olExchangeGlobalAddressList Then 
 
 Set colAE = oAL.AddressEntries 
 
 For Each oAE In colAE 
 
 If oAE.AddressEntryUserType = _ 
 
 olExchangeUserAddressEntry Then 
 
 Set oExUser = oAE.GetExchangeUser 
 
 Debug.Print(oExUser.JobTitle) 
 
 Debug.Print(oExUser.OfficeLocation) 
 
 Debug.Print(oExUser.BusinessTelephoneNumber) 
 
 End If 
 
 Next 
 
 End If 
 
 Next 
 
End Sub
```


## Methods



|**Name**|
|:-----|
|[Delete](http://msdn.microsoft.com/library/d11a82c4-28de-efef-5170-20f999f2bf08%28Office.15%29.aspx)|
|[Details](http://msdn.microsoft.com/library/6c93a583-cc61-e527-7832-88dba525854a%28Office.15%29.aspx)|
|[GetContact](http://msdn.microsoft.com/library/443fb23a-cd26-e385-bd9d-e978aec56458%28Office.15%29.aspx)|
|[GetDirectReports](http://msdn.microsoft.com/library/753201ad-8001-3185-7d68-fda15907099d%28Office.15%29.aspx)|
|[GetExchangeDistributionList](http://msdn.microsoft.com/library/4ebc0448-97a9-ca5c-35f0-ef852de27324%28Office.15%29.aspx)|
|[GetExchangeUser](http://msdn.microsoft.com/library/ff0babba-895f-8205-fefb-c587ee53ea91%28Office.15%29.aspx)|
|[GetExchangeUserManager](http://msdn.microsoft.com/library/ead5e950-7f7a-b213-0daf-c2bff890a30c%28Office.15%29.aspx)|
|[GetFreeBusy](http://msdn.microsoft.com/library/0dcd36af-e9d7-ca1e-334f-c540c46254f7%28Office.15%29.aspx)|
|[GetMemberOfList](http://msdn.microsoft.com/library/1f4e8910-8998-85ab-05dc-d06f6fd323c3%28Office.15%29.aspx)|
|[GetPicture](http://msdn.microsoft.com/library/4298db85-0576-4982-9592-6eae666d966a%28Office.15%29.aspx)|
|[Update](http://msdn.microsoft.com/library/a2672fbf-f32a-f120-227c-24ee5c361f35%28Office.15%29.aspx)|
|[GetUnifiedGroup](http://msdn.microsoft.com/library/ec0f58fa-969d-ed38-705b-2c99ccbf3c86%28Office.15%29.aspx)|
|[GetUnifiedGroupFromStore](http://msdn.microsoft.com/library/38a901d3-670f-afd2-a385-3b2bb859cb81%28Office.15%29.aspx)|
|[IsUnifiedGroup](http://msdn.microsoft.com/library/46f9564a-1c0a-fe6c-3f06-989fb5f36adf%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[Address](http://msdn.microsoft.com/library/b3a36b16-e652-9e3f-86fd-7cea0c72d78c%28Office.15%29.aspx)|
|[AddressEntryUserType](http://msdn.microsoft.com/library/fb5b16be-8846-7c9f-22bf-847d2cfc0a54%28Office.15%29.aspx)|
|[Alias](http://msdn.microsoft.com/library/ea87a061-4f09-e0ed-fd3d-bfb44cccaf15%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/17331aa1-d926-ada9-a782-02291cd6f720%28Office.15%29.aspx)|
|[AssistantName](http://msdn.microsoft.com/library/cca35d99-7031-ba46-4171-8c89b9ea2e2b%28Office.15%29.aspx)|
|[BusinessTelephoneNumber](http://msdn.microsoft.com/library/c01f85bb-24a2-c08f-df4c-9e6506ca2077%28Office.15%29.aspx)|
|[City](http://msdn.microsoft.com/library/fcec3330-39fb-61e9-e447-4adca0146171%28Office.15%29.aspx)|
|[Class](http://msdn.microsoft.com/library/eea4ce34-a957-3771-ae7b-d8fdd959a37d%28Office.15%29.aspx)|
|[Comments](http://msdn.microsoft.com/library/b55f865c-c564-f209-7648-9977512dd5a5%28Office.15%29.aspx)|
|[CompanyName](http://msdn.microsoft.com/library/d7a630ec-0fbf-78ea-5f2a-51be6d001c23%28Office.15%29.aspx)|
|[Department](http://msdn.microsoft.com/library/3b2512ff-d741-53b2-6f1d-a0f74ffbbce1%28Office.15%29.aspx)|
|[DisplayType](http://msdn.microsoft.com/library/3060a00b-9a99-7833-1a8a-5c18123d7d98%28Office.15%29.aspx)|
|[FirstName](http://msdn.microsoft.com/library/6a72812a-31fd-aa6a-be08-f765018208ab%28Office.15%29.aspx)|
|[ID](http://msdn.microsoft.com/library/b26ae0d3-ba96-f3ad-cd74-92ce5305e702%28Office.15%29.aspx)|
|[JobTitle](http://msdn.microsoft.com/library/2cfa5301-3164-c472-3f8e-831c1eebc810%28Office.15%29.aspx)|
|[LastName](http://msdn.microsoft.com/library/1f9f9675-3e72-da56-d654-a1473f4f71a7%28Office.15%29.aspx)|
|[MobileTelephoneNumber](http://msdn.microsoft.com/library/9c76ef68-1f85-d072-11d9-015fbbe1658e%28Office.15%29.aspx)|
|[Name](http://msdn.microsoft.com/library/8b93c5a3-7c6a-4193-7fc3-621e1d0dda18%28Office.15%29.aspx)|
|[OfficeLocation](http://msdn.microsoft.com/library/b37d5622-27ba-b2c4-cfd3-6aa1e9e9296b%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/18a2505c-14aa-7924-ec59-74c8e85ac92e%28Office.15%29.aspx)|
|[PostalCode](http://msdn.microsoft.com/library/b135d7a6-daa1-4154-d6e7-506c59860a81%28Office.15%29.aspx)|
|[PrimarySmtpAddress](http://msdn.microsoft.com/library/2dda21da-44a2-fbfe-babc-58646c76689d%28Office.15%29.aspx)|
|[PropertyAccessor](http://msdn.microsoft.com/library/d1427525-8f6a-04a2-9cfa-b91ee0a89ec2%28Office.15%29.aspx)|
|[Session](http://msdn.microsoft.com/library/7d2d23f0-c441-281a-1784-fe63dfa47b9f%28Office.15%29.aspx)|
|[StateOrProvince](http://msdn.microsoft.com/library/abac4889-800a-5573-5851-095f5b5176c5%28Office.15%29.aspx)|
|[StreetAddress](http://msdn.microsoft.com/library/155399c8-7d99-6537-ba30-84145b26ef21%28Office.15%29.aspx)|
|[Type](http://msdn.microsoft.com/library/de3652a8-023c-5d2c-9ced-88f768c22a87%28Office.15%29.aspx)|
|[YomiCompanyName](http://msdn.microsoft.com/library/481fec99-c2ab-c4c7-8e05-ede9e6846d1e%28Office.15%29.aspx)|
|[YomiDepartment](http://msdn.microsoft.com/library/6bc06cf2-7dee-fa50-7380-73df8022ff18%28Office.15%29.aspx)|
|[YomiDisplayName](http://msdn.microsoft.com/library/71e97add-9cf1-86c7-3e94-985d2333ebbd%28Office.15%29.aspx)|
|[YomiFirstName](http://msdn.microsoft.com/library/b44094df-af5a-21fd-0c09-ada48e51cfd8%28Office.15%29.aspx)|
|[YomiLastName](http://msdn.microsoft.com/library/079ba8e7-4a3a-2f8c-fa17-0db5ab8f47c2%28Office.15%29.aspx)|

## See also


#### Other resources


[ExchangeUser Object Members](http://msdn.microsoft.com/library/b9489e9d-0b8e-1c8d-d5df-8def4b1ee5e8%28Office.15%29.aspx)
[Outlook Object Model Reference](http://msdn.microsoft.com/library/73221b13-d8d8-99b8-3394-b95dbbfd5ddc%28Office.15%29.aspx)
