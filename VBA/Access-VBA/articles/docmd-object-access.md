---
title: DoCmd Object (Access)
keywords: vbaac10.chm4241
f1_keywords:
- vbaac10.chm4241
ms.prod: access
api_name:
- Access.DoCmd
ms.assetid: 3ce44cca-9979-0a1e-9787-079a52ce528f
ms.date: 06/08/2017
---


# DoCmd Object (Access)

You can use the methods of the  **DoCmd** object to run Microsoft Office Access actions from Visual Basic. An action performs tasks such as closing windows, opening forms, and setting the value of controls.


## Remarks

For example, you can use the  **OpenForm** method of the **DoCmd** object to open a form, or use the **Hourglass** method to change the mouse pointer to an hourglass icon.

Most of the methods of the  **DoCmd** object have arguments — some are required, while others are optional. If you omit optional arguments, the arguments assume the default values for the particular method. For example, the **OpenForm** method uses seven arguments, but only the first argument, _FormName_, is required. The following example shows how you can open the Employees form in the current database. Only employees with the title Sales Representative are included.




```
DoCmd.OpenForm "Employees", , ,"[Title] = 'Sales Representative'"
```

The  **DoCmd** object doesn't support methods corresponding to the following actions:
    
- MsgBox. Use the  **MsgBox** function.
    
- RunApp. Use the  **Shell** function to run another application.
    
- RunCode. Run the function directly in Visual Basic.
    
- SendKeys. Use the  **SendKeys** statement.
    
- SetValue. Set the value directly in Visual Basic.
    
- StopAllMacros.
    
- StopMacro.
    

## Example

The following example opens a form in Form view and moves to a new record.


```
Sub ShowNewRecord() 
 DoCmd.OpenForm "Employees", acNormal 
 DoCmd.GoToRecord , , acNewRec 
End Sub
```


## Methods



|**Name**|
|:-----|
|[AddMenu](http://msdn.microsoft.com/library/d2db2143-fd15-56b3-ee99-b895bc6b21f8%28Office.15%29.aspx)|
|[ApplyFilter](http://msdn.microsoft.com/library/926c7135-131b-1a7c-465b-a9b2ed71cd7b%28Office.15%29.aspx)|
|[Beep](http://msdn.microsoft.com/library/822a565d-89d9-fdc1-eb01-b8535e363714%28Office.15%29.aspx)|
|[BrowseTo](http://msdn.microsoft.com/library/7cfd2cc5-ad2d-4bf8-ed90-1fb6adf1890a%28Office.15%29.aspx)|
|[CancelEvent](http://msdn.microsoft.com/library/f8c0d2ff-9bf3-09d5-d15b-d3134bb6df80%28Office.15%29.aspx)|
|[ClearMacroError](http://msdn.microsoft.com/library/2784bfc8-f61a-a461-e067-640a4244436d%28Office.15%29.aspx)|
|[Close](http://msdn.microsoft.com/library/3fdb2fa2-31d8-baf7-89f3-f9ef330280b3%28Office.15%29.aspx)|
|[CloseDatabase](http://msdn.microsoft.com/library/0150a029-176c-7385-71ee-0d76d6fb9ca3%28Office.15%29.aspx)|
|[CopyDatabaseFile](http://msdn.microsoft.com/library/15a820d9-fbcb-d803-d58a-5718924e6c73%28Office.15%29.aspx)|
|[CopyObject](http://msdn.microsoft.com/library/003e5b47-f8a2-2b6a-5e0c-7fb3e87b3258%28Office.15%29.aspx)|
|[DeleteObject](http://msdn.microsoft.com/library/8e59c5a8-89bd-0d90-9fd1-a1178c73c1c1%28Office.15%29.aspx)|
|[DoMenuItem](http://msdn.microsoft.com/library/b897bfdb-7f03-2b42-2bfd-219a2f4aa21b%28Office.15%29.aspx)|
|[Echo](http://msdn.microsoft.com/library/519b4fe7-ff48-7ab3-3117-43da2278aa66%28Office.15%29.aspx)|
|[FindNext](http://msdn.microsoft.com/library/7edd2936-85d2-27f1-e72e-2408338fa740%28Office.15%29.aspx)|
|[FindRecord](http://msdn.microsoft.com/library/dc48bc3d-5408-40a8-509b-e52b48b26187%28Office.15%29.aspx)|
|[GoToControl](http://msdn.microsoft.com/library/2b51231d-f6a4-4891-d49d-bedb68f85b04%28Office.15%29.aspx)|
|[GoToPage](http://msdn.microsoft.com/library/37fe25b3-85b2-f681-acfd-96dab039e58f%28Office.15%29.aspx)|
|[GoToRecord](http://msdn.microsoft.com/library/5494b6fc-112f-e944-9072-873b00271ab1%28Office.15%29.aspx)|
|[Hourglass](http://msdn.microsoft.com/library/e032e879-6ce4-982d-08cb-f9622c000b11%28Office.15%29.aspx)|
|[LockNavigationPane](http://msdn.microsoft.com/library/64b44d9b-4cbd-182c-9bfb-89b4ca04dbf9%28Office.15%29.aspx)|
|[Maximize](http://msdn.microsoft.com/library/6b1103f5-07b8-fbcf-ff7e-ccbfd6945768%28Office.15%29.aspx)|
|[Minimize](http://msdn.microsoft.com/library/fa29ccaa-9d61-c5c3-fc32-f53a5d96ff05%28Office.15%29.aspx)|
|[MoveSize](http://msdn.microsoft.com/library/8fe8fc60-023e-26ce-c11a-2c29ffc21fbb%28Office.15%29.aspx)|
|[NavigateTo](http://msdn.microsoft.com/library/27a6e4ee-1c03-2652-3c5a-73c45f3109df%28Office.15%29.aspx)|
|[OpenDataAccessPage](http://msdn.microsoft.com/library/130dcb88-e3e6-25a6-186c-bf541d114169%28Office.15%29.aspx)|
|[OpenDiagram](http://msdn.microsoft.com/library/a9736e57-eb82-77d7-c57a-8c793333392a%28Office.15%29.aspx)|
|[OpenForm](http://msdn.microsoft.com/library/a1c9d3a9-2af8-c30a-acb0-6428c70dcdb0%28Office.15%29.aspx)|
|[OpenFunction](http://msdn.microsoft.com/library/56168394-9e83-f620-8b5e-680e824ec941%28Office.15%29.aspx)|
|[OpenModule](http://msdn.microsoft.com/library/3d0b1599-6f52-e369-55e4-7fdc1c370953%28Office.15%29.aspx)|
|[OpenQuery](http://msdn.microsoft.com/library/3ea20a28-8dd4-e54c-831b-e7e5444aa793%28Office.15%29.aspx)|
|[OpenReport](http://msdn.microsoft.com/library/3c08755a-5116-f085-d498-725dc12e62f1%28Office.15%29.aspx)|
|[OpenStoredProcedure](http://msdn.microsoft.com/library/90e229f9-072a-8d41-4c9b-363501770c8c%28Office.15%29.aspx)|
|[OpenTable](http://msdn.microsoft.com/library/6461c8c1-7452-f812-8914-e46406c58eae%28Office.15%29.aspx)|
|[OpenView](http://msdn.microsoft.com/library/8d2970dd-9a06-f917-04da-850b085126dd%28Office.15%29.aspx)|
|[OutputTo](http://msdn.microsoft.com/library/2a21a7c3-0846-cbec-d5dd-a1648f705557%28Office.15%29.aspx)|
|[PrintOut](http://msdn.microsoft.com/library/3b7c1ab7-1a60-cab3-2d4e-c95d6b5bd4aa%28Office.15%29.aspx)|
|[Quit](http://msdn.microsoft.com/library/2644084a-fd24-6271-7679-46c5f1b206d5%28Office.15%29.aspx)|
|[RefreshRecord](http://msdn.microsoft.com/library/2707cdf2-7458-7ef2-8c20-26fed3eda3ce%28Office.15%29.aspx)|
|[Rename](http://msdn.microsoft.com/library/c9286727-a172-b7c5-c8b4-6e63012db98a%28Office.15%29.aspx)|
|[RepaintObject](http://msdn.microsoft.com/library/6def040f-ae34-ce49-d3a0-786ad09bdc20%28Office.15%29.aspx)|
|[Requery](http://msdn.microsoft.com/library/6869c39f-b43f-ad83-4140-67b763342605%28Office.15%29.aspx)|
|[Restore](http://msdn.microsoft.com/library/455c2589-6d1a-aa87-d338-37bcb0abe608%28Office.15%29.aspx)|
|[RunCommand](http://msdn.microsoft.com/library/5d4a4a3c-cea0-7f2c-8af7-51b65f7bdcf8%28Office.15%29.aspx)|
|[RunDataMacro](http://msdn.microsoft.com/library/e95b7a8e-a502-67c6-1941-dd5a06c08ef7%28Office.15%29.aspx)|
|[RunMacro](http://msdn.microsoft.com/library/2abb0056-3f8a-337b-307f-6d653aa2b963%28Office.15%29.aspx)|
|[RunSavedImportExport](http://msdn.microsoft.com/library/cb0ade9a-5cd4-1225-5231-8266fdfb3690%28Office.15%29.aspx)|
|[RunSQL](http://msdn.microsoft.com/library/5d61f75a-b220-cc2c-edea-51a6d4f9f106%28Office.15%29.aspx)|
|[Save](http://msdn.microsoft.com/library/7e01f370-36c9-9f4d-b506-61bc8886ee18%28Office.15%29.aspx)|
|[SearchForRecord](http://msdn.microsoft.com/library/eb7a82b0-1ecb-cbfe-94b0-e2d6742de8b4%28Office.15%29.aspx)|
|[SelectObject](http://msdn.microsoft.com/library/def1bac5-57b1-0b2c-d39a-f0c10962880c%28Office.15%29.aspx)|
|[SendObject](http://msdn.microsoft.com/library/881004c6-2dd7-55f1-2a16-2d28034125a8%28Office.15%29.aspx)|
|[SetDisplayedCategories](http://msdn.microsoft.com/library/ae2290c3-43ff-c19d-63f8-41427aacd9ce%28Office.15%29.aspx)|
|[SetFilter](http://msdn.microsoft.com/library/98c3e202-8581-2215-7fb2-4a006a97d38f%28Office.15%29.aspx)|
|[SetMenuItem](http://msdn.microsoft.com/library/690263c1-5e0f-54cd-1032-b2f718d82075%28Office.15%29.aspx)|
|[SetOrderBy](http://msdn.microsoft.com/library/020fde6d-4809-79f6-3da5-fc5f6a315a83%28Office.15%29.aspx)|
|[SetParameter](http://msdn.microsoft.com/library/55e64bab-1c5e-9da0-5425-c8ed7b0bb1c2%28Office.15%29.aspx)|
|[SetProperty](http://msdn.microsoft.com/library/32347eb6-115d-36c5-4c18-eab7e7422b78%28Office.15%29.aspx)|
|[SetWarnings](http://msdn.microsoft.com/library/fe8cbd54-fa63-4057-8ea2-da9ba79ed1a6%28Office.15%29.aspx)|
|[ShowAllRecords](http://msdn.microsoft.com/library/765ead1a-d626-3a54-1831-1490fc8daacc%28Office.15%29.aspx)|
|[ShowToolbar](http://msdn.microsoft.com/library/63663cc5-a591-c847-25c8-25777cf7806a%28Office.15%29.aspx)|
|[SingleStep](http://msdn.microsoft.com/library/fa355661-9605-9477-15f6-10f0a163ba67%28Office.15%29.aspx)|
|[TransferDatabase](http://msdn.microsoft.com/library/7eff4d0c-f660-72db-ee99-b6a3158f01de%28Office.15%29.aspx)|
|[TransferSharePointList](http://msdn.microsoft.com/library/9cbd8de6-dc1a-47b0-c1f4-62959a66faf4%28Office.15%29.aspx)|
|[TransferSpreadsheet](http://msdn.microsoft.com/library/0349d8e0-9363-0eda-4efb-a73c9e643823%28Office.15%29.aspx)|
|[TransferSQLDatabase](http://msdn.microsoft.com/library/d6a88496-9137-b190-8357-316fd580a036%28Office.15%29.aspx)|
|[TransferText](http://msdn.microsoft.com/library/e59f26dc-2df8-8d87-b73d-f3004eed0719%28Office.15%29.aspx)|

## See also


#### Other resources


[Access Object Model Reference](object-model-access-vba-reference.md)

