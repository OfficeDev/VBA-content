---
title: Assignments Object (Project)
ms.prod: project-server
ms.assetid: 83661095-030c-0488-5763-320b6de6f381
ms.date: 06/08/2017
---


# Assignments Object (Project)

Contains a collection of  **[Assignment](assignment-object-project.md)** objects for a task or resource.


## Example

 **Using the Assignment Object**

Use  **Assignments(** _Index_ **)**, where _Index_ is the assignment index number, to return a single **Assignment** object. The following example displays the name of the first resource assigned to the specified task.




```
MsgBox ActiveProject.Tasks(1).Assignments(1).ResourceName
```

 **Using the Assignments Collection**

Use the  **[Assignments](http://msdn.microsoft.com/library/a481e813-8f02-c58b-2910-6995aaaafa09%28Office.15%29.aspx)** property to return an **Assignments** collection. The following example displays all the resources assigned to the specified task.




```
Dim A As Assignment 

 

For Each A In ActiveProject.Tasks(1).Assignments 

 MsgBox A.ResourceName 

Next A
```

Use the  **[Add](http://msdn.microsoft.com/library/c135a80e-1fb9-32e3-864e-f701c1947ca4%28Office.15%29.aspx)** method to add an **Assignment** object to the **Assignments** collection. The following example adds a resource identified by the number 212 as a new assignment for the specified task.




```
ActiveProject.Tasks(1).Assignments.Add ResourceID:=212
```


## Methods



|**Name**|
|:-----|
|[AppendNotes](http://msdn.microsoft.com/library/78ccad76-ac3f-c11e-9d88-2ed133358671%28Office.15%29.aspx)|
|[Delete](http://msdn.microsoft.com/library/3147c0e0-239c-75d2-cae9-c299412190e2%28Office.15%29.aspx)|
|[EnterpriseTeamMember](http://msdn.microsoft.com/library/706a7f8b-b545-7398-7c09-f29f6b8d225d%28Office.15%29.aspx)|
|[Replan](http://msdn.microsoft.com/library/29ec0102-b4e4-c9dc-d930-4f8ff4069bd6%28Office.15%29.aspx)|
|[TimeScaleData](http://msdn.microsoft.com/library/ff948754-cc0e-8bf0-31e8-30b19dbcb08d%28Office.15%29.aspx)|

## Properties



|**Name**|
|:-----|
|[ActualCost](http://msdn.microsoft.com/library/45bf4d44-bce7-474a-7093-ff0c97d3b7f6%28Office.15%29.aspx)|
|[ActualFinish](http://msdn.microsoft.com/library/b1ef2626-4fa2-a036-28f0-fbbff5c06407%28Office.15%29.aspx)|
|[ActualOvertimeCost](http://msdn.microsoft.com/library/ee89c244-f153-e42c-3e56-a1d363b62f9c%28Office.15%29.aspx)|
|[ActualOvertimeWork](http://msdn.microsoft.com/library/cc427c88-18f4-5235-f787-d8366c3e3a23%28Office.15%29.aspx)|
|[ActualStart](http://msdn.microsoft.com/library/0a20d560-ce64-4696-e9d4-61bf2a7dda04%28Office.15%29.aspx)|
|[ActualWork](http://msdn.microsoft.com/library/10a4102c-0549-a9b3-94bd-5aa1c5d8b813%28Office.15%29.aspx)|
|[ACWP](http://msdn.microsoft.com/library/a28a370c-f7ee-56e4-e11b-a40553dcaec0%28Office.15%29.aspx)|
|[Application](http://msdn.microsoft.com/library/c6dbff13-0f33-7f78-b603-fa7084889bed%28Office.15%29.aspx)|
|[Baseline10BudgetCost](http://msdn.microsoft.com/library/75705ad0-4da0-2fd3-1dda-33042313d9c1%28Office.15%29.aspx)|
|[Baseline10BudgetWork](http://msdn.microsoft.com/library/6392d966-1ce4-fa4d-28ac-5bced525ba10%28Office.15%29.aspx)|
|[Baseline10Cost](http://msdn.microsoft.com/library/590ec3c4-417f-e407-c0da-786f7512f2c1%28Office.15%29.aspx)|
|[Baseline10Finish](http://msdn.microsoft.com/library/0d67a0c2-035e-80be-a588-4ea95b2da4c0%28Office.15%29.aspx)|
|[Baseline10Start](http://msdn.microsoft.com/library/7ecc2bc8-607a-5d9f-8bdd-a2b7b34c985d%28Office.15%29.aspx)|
|[Baseline10Work](http://msdn.microsoft.com/library/e6b020f7-c2cd-cb15-d77f-bc384ed1d934%28Office.15%29.aspx)|
|[Baseline1BudgetCost](http://msdn.microsoft.com/library/b58491e6-11f2-3f85-4e9a-ba686c353304%28Office.15%29.aspx)|
|[Baseline1BudgetWork](http://msdn.microsoft.com/library/7df3330c-0397-0075-0c3c-d4bfffc6ed20%28Office.15%29.aspx)|
|[Baseline1Cost](http://msdn.microsoft.com/library/9c20db71-484d-810f-24e5-a972e86f29a9%28Office.15%29.aspx)|
|[Baseline1Finish](http://msdn.microsoft.com/library/92141961-5d2c-4fb8-8924-065e1b3bddb6%28Office.15%29.aspx)|
|[Baseline1Start](http://msdn.microsoft.com/library/16afebc0-3856-46e3-cdbb-875bd0904ceb%28Office.15%29.aspx)|
|[Baseline1Work](http://msdn.microsoft.com/library/6584b8d7-96f0-905b-9b22-19917c1452ae%28Office.15%29.aspx)|
|[Baseline2BudgetCost](http://msdn.microsoft.com/library/44a3bd58-a6dc-6fe6-5ecb-61b35077a660%28Office.15%29.aspx)|
|[Baseline2BudgetWork](http://msdn.microsoft.com/library/aeda3d79-e129-78db-c6b9-38a5fdd7a1fc%28Office.15%29.aspx)|
|[Baseline2Cost](http://msdn.microsoft.com/library/827ab8e6-0e4f-84a7-e77a-2966747c8d59%28Office.15%29.aspx)|
|[Baseline2Finish](http://msdn.microsoft.com/library/95760bcd-8072-143a-478a-12bdfa1a9f16%28Office.15%29.aspx)|
|[Baseline2Start](http://msdn.microsoft.com/library/e62326eb-590b-6df4-362e-3cd00220557f%28Office.15%29.aspx)|
|[Baseline2Work](http://msdn.microsoft.com/library/40be106a-90ea-8240-d6ee-a485663bcbec%28Office.15%29.aspx)|
|[Baseline3BudgetCost](http://msdn.microsoft.com/library/e55e4f8e-5e14-8e7a-67f9-d6e721d7b671%28Office.15%29.aspx)|
|[Baseline3BudgetWork](http://msdn.microsoft.com/library/2bc8234e-bb10-0f46-ad88-797755318319%28Office.15%29.aspx)|
|[Baseline3Cost](http://msdn.microsoft.com/library/e752f055-1e29-b7a3-5e72-020daa867388%28Office.15%29.aspx)|
|[Baseline3Finish](http://msdn.microsoft.com/library/a52d9f03-e7f0-b1a0-69bd-cc563162bb69%28Office.15%29.aspx)|
|[Baseline3Start](http://msdn.microsoft.com/library/106ce677-8c42-6974-490c-f72f8095621b%28Office.15%29.aspx)|
|[Baseline3Work](http://msdn.microsoft.com/library/f834160a-40e3-d6e9-66ed-0f9b9f6a1698%28Office.15%29.aspx)|
|[Baseline4BudgetCost](http://msdn.microsoft.com/library/7ebc26fa-dbd3-2372-4566-68c854990038%28Office.15%29.aspx)|
|[Baseline4BudgetWork](http://msdn.microsoft.com/library/5efff144-fb05-2108-8260-f4195c4ea54d%28Office.15%29.aspx)|
|[Baseline4Cost](http://msdn.microsoft.com/library/2bab26ff-0d68-6258-3978-45fc6faf3e9d%28Office.15%29.aspx)|
|[Baseline4Finish](http://msdn.microsoft.com/library/3339d680-94b3-48d6-86a1-cab385465bd9%28Office.15%29.aspx)|
|[Baseline4Start](http://msdn.microsoft.com/library/656122d8-4228-667e-7dec-bdfd7774cc80%28Office.15%29.aspx)|
|[Baseline4Work](http://msdn.microsoft.com/library/d1d075e6-c248-1b7c-470c-95ae2241def7%28Office.15%29.aspx)|
|[Baseline5BudgetCost](http://msdn.microsoft.com/library/af5f4183-4db9-9f83-2a13-9ff8cb66df3e%28Office.15%29.aspx)|
|[Baseline5BudgetWork](http://msdn.microsoft.com/library/aebaa0d4-4484-6718-b0b5-ba58972d8f0e%28Office.15%29.aspx)|
|[Baseline5Cost](http://msdn.microsoft.com/library/1cad6c8b-2e0a-2a76-0888-11f487e481a1%28Office.15%29.aspx)|
|[Baseline5Finish](http://msdn.microsoft.com/library/210c4b18-119d-5bdd-20ff-8a27e6c03fc1%28Office.15%29.aspx)|
|[Baseline5Start](http://msdn.microsoft.com/library/4d2a1a50-5e71-78b2-f2d6-55dc0bca7494%28Office.15%29.aspx)|
|[Baseline5Work](http://msdn.microsoft.com/library/16893da5-816f-4cdc-c256-09c3860532a6%28Office.15%29.aspx)|
|[Baseline6BudgetCost](http://msdn.microsoft.com/library/df07aa02-bd67-8be3-f3de-1f6988e7f806%28Office.15%29.aspx)|
|[Baseline6BudgetWork](http://msdn.microsoft.com/library/1a7ca85e-5f9c-ee43-a34c-43aa645cf66f%28Office.15%29.aspx)|
|[Baseline6Cost](http://msdn.microsoft.com/library/4daa1d9c-48b1-044a-745e-409e4a6247b3%28Office.15%29.aspx)|
|[Baseline6Finish](http://msdn.microsoft.com/library/00de68e1-0d22-821b-3e4b-7bd863d70d25%28Office.15%29.aspx)|
|[Baseline6Start](http://msdn.microsoft.com/library/f132de0f-a3d2-dea4-444b-ec25d7eac234%28Office.15%29.aspx)|
|[Baseline6Work](http://msdn.microsoft.com/library/57952e9c-9cb9-e507-3788-266240974b93%28Office.15%29.aspx)|
|[Baseline7BudgetCost](http://msdn.microsoft.com/library/b3710f3b-8502-5af3-76df-4b87d22ce5ea%28Office.15%29.aspx)|
|[Baseline7BudgetWork](http://msdn.microsoft.com/library/0e21c0e9-8dca-91b4-6a63-d373eea6c7e9%28Office.15%29.aspx)|
|[Baseline7Cost](http://msdn.microsoft.com/library/ca6f21e7-7430-24c3-cef5-e94565acb98e%28Office.15%29.aspx)|
|[Baseline7Finish](http://msdn.microsoft.com/library/c982594c-0086-8468-ce6e-51e8c2a46f4f%28Office.15%29.aspx)|
|[Baseline7Start](http://msdn.microsoft.com/library/82062a92-b922-0f71-f145-bac9161cdcd4%28Office.15%29.aspx)|
|[Baseline7Work](http://msdn.microsoft.com/library/fce7b332-6890-f951-28cc-c766a4baba20%28Office.15%29.aspx)|
|[Baseline8BudgetCost](http://msdn.microsoft.com/library/bd8febca-06f7-29f7-6b94-e7ca72f3c1c6%28Office.15%29.aspx)|
|[Baseline8BudgetWork](http://msdn.microsoft.com/library/b4f81a07-1442-bcec-867e-86ae9af8c207%28Office.15%29.aspx)|
|[Baseline8Cost](http://msdn.microsoft.com/library/25ad0e71-a2e8-959c-ac6b-a77425121a28%28Office.15%29.aspx)|
|[Baseline8Finish](http://msdn.microsoft.com/library/19f921df-4785-1963-2dcc-297c11518494%28Office.15%29.aspx)|
|[Baseline8Start](http://msdn.microsoft.com/library/888fcd06-cd02-0743-8f85-1038abddf9a8%28Office.15%29.aspx)|
|[Baseline8Work](http://msdn.microsoft.com/library/1b1572de-4d01-be5a-3093-626783004033%28Office.15%29.aspx)|
|[Baseline9BudgetCost](http://msdn.microsoft.com/library/1e89b6be-9a75-28b4-6d1f-79e31825fa8d%28Office.15%29.aspx)|
|[Baseline9BudgetWork](http://msdn.microsoft.com/library/8c76d3e1-0ff1-6ada-0bfc-20a22cdc1ca3%28Office.15%29.aspx)|
|[Baseline9Cost](http://msdn.microsoft.com/library/fbcd0b8e-e153-6e1e-efa4-877dca6d70c0%28Office.15%29.aspx)|
|[Baseline9Finish](http://msdn.microsoft.com/library/57889822-a28e-4ed5-d972-0c63bef29fc2%28Office.15%29.aspx)|
|[Baseline9Start](http://msdn.microsoft.com/library/78fee6d3-2645-62be-0173-9f35b58b4b0c%28Office.15%29.aspx)|
|[Baseline9Work](http://msdn.microsoft.com/library/777a8d7a-d9d4-e0fb-5b5b-2c78302e5fa4%28Office.15%29.aspx)|
|[BaselineBudgetCost](http://msdn.microsoft.com/library/65053c03-5b36-41a8-7857-c987c10d63ea%28Office.15%29.aspx)|
|[BaselineBudgetWork](http://msdn.microsoft.com/library/d10ddcdc-0879-1567-2697-e55ebcd4675b%28Office.15%29.aspx)|
|[BaselineCost](http://msdn.microsoft.com/library/80077930-4bc7-f5f3-9c59-c6477db779fd%28Office.15%29.aspx)|
|[BaselineFinish](http://msdn.microsoft.com/library/9e062dc8-fed3-446f-776c-2d10179a6c3b%28Office.15%29.aspx)|
|[BaselineStart](http://msdn.microsoft.com/library/95586824-b281-cefd-c360-f8a951c86088%28Office.15%29.aspx)|
|[BaselineWork](http://msdn.microsoft.com/library/9399ca50-e952-0ac0-3677-f0bee2a71ec7%28Office.15%29.aspx)|
|[BCWP](http://msdn.microsoft.com/library/4e8f5b89-8e71-bd05-3681-63e56d6969b2%28Office.15%29.aspx)|
|[BCWS](http://msdn.microsoft.com/library/22ffb05e-6e36-061b-771b-f8fc3bf8217e%28Office.15%29.aspx)|
|[BookingType](http://msdn.microsoft.com/library/9effb3b1-42eb-8adb-9c26-7103df375c88%28Office.15%29.aspx)|
|[BudgetCost](http://msdn.microsoft.com/library/1f7ec7dd-8733-7050-e038-29a917f155ff%28Office.15%29.aspx)|
|[BudgetWork](http://msdn.microsoft.com/library/21c73cbb-4bca-1eea-4900-6e575cd298a7%28Office.15%29.aspx)|
|[Confirmed](http://msdn.microsoft.com/library/67d562c2-139a-3bf1-8a50-8e44adad657e%28Office.15%29.aspx)|
|[Cost](http://msdn.microsoft.com/library/286f8677-2dc9-a3e0-5b24-8b48d1099819%28Office.15%29.aspx)|
|[Cost1](http://msdn.microsoft.com/library/71757dbd-e42b-cfe1-459c-663e1475e643%28Office.15%29.aspx)|
|[Cost10](http://msdn.microsoft.com/library/1c68b400-cc7c-3e54-94b4-6c791ab52579%28Office.15%29.aspx)|
|[Cost2](http://msdn.microsoft.com/library/ce7dd57d-7a43-1753-5470-2fade9aa68f2%28Office.15%29.aspx)|
|[Cost3](http://msdn.microsoft.com/library/6da4eddf-fc32-5b03-79a9-951fa0aab941%28Office.15%29.aspx)|
|[Cost4](http://msdn.microsoft.com/library/f8876853-af81-c359-c230-8ea1c9a6f184%28Office.15%29.aspx)|
|[Cost5](http://msdn.microsoft.com/library/54217131-6d53-7568-6f98-4f1266bbbf9d%28Office.15%29.aspx)|
|[Cost6](http://msdn.microsoft.com/library/d0ad1074-caf9-c160-042b-2bca5ea220e4%28Office.15%29.aspx)|
|[Cost7](http://msdn.microsoft.com/library/14d2f7b3-b90b-67ae-7418-44e1d7836f90%28Office.15%29.aspx)|
|[Cost8](http://msdn.microsoft.com/library/08c1c081-81af-37f7-00b8-cfc4d29df4e0%28Office.15%29.aspx)|
|[Cost9](http://msdn.microsoft.com/library/f81c1aea-625a-ac7d-c837-7cde27d3f3bc%28Office.15%29.aspx)|
|[CostRateTable](http://msdn.microsoft.com/library/03d615e2-6dea-849f-a9a5-c20e1c35bee8%28Office.15%29.aspx)|
|[CostVariance](http://msdn.microsoft.com/library/140fe7d6-cfd6-7521-e11b-24d5dbe09d1a%28Office.15%29.aspx)|
|[Created](http://msdn.microsoft.com/library/6ad7a628-8841-716f-0de9-a6f13aa61e85%28Office.15%29.aspx)|
|[CV](http://msdn.microsoft.com/library/15028dc8-1226-333f-e4f4-9e31f9970481%28Office.15%29.aspx)|
|[Date1](http://msdn.microsoft.com/library/d06bbeb2-2b3d-eded-195e-dcab6ccd50a7%28Office.15%29.aspx)|
|[Date10](http://msdn.microsoft.com/library/795c71e1-5dfb-4044-3679-6db2bf2b30b5%28Office.15%29.aspx)|
|[Date2](http://msdn.microsoft.com/library/be8665ce-ffd6-fc0e-6b0d-17dc0bcdac65%28Office.15%29.aspx)|
|[Date3](http://msdn.microsoft.com/library/7ddf378a-2ea4-0c66-4266-4ca77d86e18f%28Office.15%29.aspx)|
|[Date4](http://msdn.microsoft.com/library/02e92640-d5c1-15c5-fda9-01f5df33d6f2%28Office.15%29.aspx)|
|[Date5](http://msdn.microsoft.com/library/3d144835-0bc0-6021-9ed5-13846c568ca2%28Office.15%29.aspx)|
|[Date6](http://msdn.microsoft.com/library/0651e923-132a-933e-9191-5dd8e4c9c222%28Office.15%29.aspx)|
|[Date7](http://msdn.microsoft.com/library/1d50befd-3087-2584-b41a-f96a2cfa8fa7%28Office.15%29.aspx)|
|[Date8](http://msdn.microsoft.com/library/cc1af84d-7b97-de6a-72c4-334fd6183303%28Office.15%29.aspx)|
|[Date9](http://msdn.microsoft.com/library/a53e08a9-cd7e-2652-60d8-b1adc90e926c%28Office.15%29.aspx)|
|[Delay](http://msdn.microsoft.com/library/55b07677-2937-90f8-aa71-314732f27354%28Office.15%29.aspx)|
|[Duration1](http://msdn.microsoft.com/library/a6d57e54-cad2-0edf-994b-65405d47c0d9%28Office.15%29.aspx)|
|[Duration10](http://msdn.microsoft.com/library/f6ad9b7e-41e0-9929-879a-51c12e89d56f%28Office.15%29.aspx)|
|[Duration2](http://msdn.microsoft.com/library/d51247c6-1270-ba93-13ac-7b5dabb38ccd%28Office.15%29.aspx)|
|[Duration3](http://msdn.microsoft.com/library/aafc2f78-fa61-2c44-d7ca-0c6499e97632%28Office.15%29.aspx)|
|[Duration4](http://msdn.microsoft.com/library/e33d3fd0-a9bb-9766-76c4-4b0cb148ec8a%28Office.15%29.aspx)|
|[Duration5](http://msdn.microsoft.com/library/4aabfaec-f98a-709f-733f-4fec28e37b2d%28Office.15%29.aspx)|
|[Duration6](http://msdn.microsoft.com/library/6d04b8ab-d5f7-6a93-36e5-4b9c9f57cb23%28Office.15%29.aspx)|
|[Duration7](http://msdn.microsoft.com/library/7fc5c07a-a832-444a-3865-402401e10a94%28Office.15%29.aspx)|
|[Duration8](http://msdn.microsoft.com/library/0be92dfc-bfa2-629f-b7a0-65643ad5902e%28Office.15%29.aspx)|
|[Duration9](http://msdn.microsoft.com/library/5b7d66df-21e6-cbf0-788d-260ec048f062%28Office.15%29.aspx)|
|[Finish](http://msdn.microsoft.com/library/c67224ed-0bfc-2119-b68c-5d7bd290b357%28Office.15%29.aspx)|
|[Finish1](http://msdn.microsoft.com/library/ed5c64e4-60d9-c6aa-33cf-570d76170cb7%28Office.15%29.aspx)|
|[Finish10](http://msdn.microsoft.com/library/8d4bb42d-a83f-9fc3-2318-1f6df8f8ee1f%28Office.15%29.aspx)|
|[Finish2](http://msdn.microsoft.com/library/7b620a85-cf0e-8394-bf0f-5b9d27750c46%28Office.15%29.aspx)|
|[Finish3](http://msdn.microsoft.com/library/d76d6820-68b7-1742-1b7c-c8ab69d928cf%28Office.15%29.aspx)|
|[Finish4](http://msdn.microsoft.com/library/ae4a0294-5ab2-4308-2243-39d6524178a7%28Office.15%29.aspx)|
|[Finish5](http://msdn.microsoft.com/library/14e669f5-3918-d4f0-33b2-1284c75a129a%28Office.15%29.aspx)|
|[Finish6](http://msdn.microsoft.com/library/4fa7d458-ea66-632d-957f-67a136e49284%28Office.15%29.aspx)|
|[Finish7](http://msdn.microsoft.com/library/80bba55c-67f7-442b-215c-ecdef96b219b%28Office.15%29.aspx)|
|[Finish8](http://msdn.microsoft.com/library/3609260a-515a-734f-4eaf-d7b55d20963e%28Office.15%29.aspx)|
|[Finish9](http://msdn.microsoft.com/library/fb169e42-d24d-6818-b73b-40f7a513b6f6%28Office.15%29.aspx)|
|[FinishVariance](http://msdn.microsoft.com/library/3ec68258-b79b-9c19-63e9-e018bb506dc4%28Office.15%29.aspx)|
|[FixedMaterialAssignment](http://msdn.microsoft.com/library/16593466-1d5e-27b3-110d-e5cfeb165355%28Office.15%29.aspx)|
|[Flag1](http://msdn.microsoft.com/library/167a2a3b-7118-1f36-0fa8-9323f530c965%28Office.15%29.aspx)|
|[Flag10](http://msdn.microsoft.com/library/204a3d12-fb71-2277-c613-f9427402dff1%28Office.15%29.aspx)|
|[Flag11](http://msdn.microsoft.com/library/225eeb44-621d-0468-5cfc-e5ce80b3a861%28Office.15%29.aspx)|
|[Flag12](http://msdn.microsoft.com/library/b4f07f88-1e02-70d4-79cf-bc0d5f8ba0d4%28Office.15%29.aspx)|
|[Flag13](http://msdn.microsoft.com/library/c79abd66-88b4-8592-6cad-1d567770e95c%28Office.15%29.aspx)|
|[Flag14](http://msdn.microsoft.com/library/8067c60f-bd67-6625-e127-badb32e7453d%28Office.15%29.aspx)|
|[Flag15](http://msdn.microsoft.com/library/d9c0e683-007c-99c7-fb5a-b8085e51c491%28Office.15%29.aspx)|
|[Flag16](http://msdn.microsoft.com/library/fc4034ce-15b2-42fa-a292-453f5b2abacd%28Office.15%29.aspx)|
|[Flag17](http://msdn.microsoft.com/library/cda8dbba-c35c-86a8-348b-ed0ac4a15db5%28Office.15%29.aspx)|
|[Flag18](http://msdn.microsoft.com/library/46e6a314-ef73-8db8-1422-340e7dd05d1d%28Office.15%29.aspx)|
|[Flag19](http://msdn.microsoft.com/library/aaa6e052-743c-ca3d-78c9-2a1ae6881e01%28Office.15%29.aspx)|
|[Flag2](http://msdn.microsoft.com/library/a1659a3c-e5a9-0409-217c-3cb0be5c0818%28Office.15%29.aspx)|
|[Flag20](http://msdn.microsoft.com/library/dd7420f0-f949-805c-5d06-928c62fc2c75%28Office.15%29.aspx)|
|[Flag3](http://msdn.microsoft.com/library/00dbf405-bed1-60fa-8b36-e7111f0519b4%28Office.15%29.aspx)|
|[Flag4](http://msdn.microsoft.com/library/16af5669-ced4-3f4b-063a-0755fcefbeb7%28Office.15%29.aspx)|
|[Flag5](http://msdn.microsoft.com/library/d05594c1-f117-e623-7145-788d60ba6eb5%28Office.15%29.aspx)|
|[Flag6](http://msdn.microsoft.com/library/7acf802a-94e5-f0ec-cfc7-5cc861987872%28Office.15%29.aspx)|
|[Flag7](http://msdn.microsoft.com/library/8613ebea-1029-e66f-cbf9-6ff29d4063a5%28Office.15%29.aspx)|
|[Flag8](http://msdn.microsoft.com/library/053c6f11-3881-8872-39b8-40c61ab621f1%28Office.15%29.aspx)|
|[Flag9](http://msdn.microsoft.com/library/516292ee-c93a-61ff-be24-c1e620d9088f%28Office.15%29.aspx)|
|[Guid](http://msdn.microsoft.com/library/c6db05fe-e2f1-edb7-e622-5b2d5e791237%28Office.15%29.aspx)|
|[Hyperlink](http://msdn.microsoft.com/library/00c0d49f-7888-8f1f-42cf-380caf6dd672%28Office.15%29.aspx)|
|[HyperlinkAddress](http://msdn.microsoft.com/library/ead317d6-aa1a-57a1-4d58-189ccf551b40%28Office.15%29.aspx)|
|[HyperlinkHREF](http://msdn.microsoft.com/library/7e8f761d-3167-2e43-fb73-40528f567153%28Office.15%29.aspx)|
|[HyperlinkScreenTip](http://msdn.microsoft.com/library/48b8b03c-4662-3ea8-646e-22a1ce268f81%28Office.15%29.aspx)|
|[HyperlinkSubAddress](http://msdn.microsoft.com/library/c26ca17d-f038-0c54-2868-4aacb381fd49%28Office.15%29.aspx)|
|[Index](http://msdn.microsoft.com/library/eea6d62f-e896-7a5e-dd33-dadc15d5ce03%28Office.15%29.aspx)|
|[LevelingDelay](http://msdn.microsoft.com/library/b01087ec-9440-9288-3afe-6c0ed87e4a50%28Office.15%29.aspx)|
|[LinkedFields](http://msdn.microsoft.com/library/72db7318-589e-bb65-a7ee-0e5031fb1122%28Office.15%29.aspx)|
|[Notes](http://msdn.microsoft.com/library/91915e62-bd93-3671-a232-05cb99836428%28Office.15%29.aspx)|
|[Number1](http://msdn.microsoft.com/library/5cfe0434-a7ef-2f5d-ed61-6262e475288c%28Office.15%29.aspx)|
|[Number10](http://msdn.microsoft.com/library/ed85359b-394e-c0c3-c8e5-926f25243fcc%28Office.15%29.aspx)|
|[Number11](http://msdn.microsoft.com/library/fcb31200-1139-3c55-0413-40a6619a2b07%28Office.15%29.aspx)|
|[Number12](http://msdn.microsoft.com/library/aa305f50-5145-69c2-5038-8884ac2cb2c6%28Office.15%29.aspx)|
|[Number13](http://msdn.microsoft.com/library/853d3dea-6085-3088-04d1-18a28c3bae7e%28Office.15%29.aspx)|
|[Number14](http://msdn.microsoft.com/library/4e91d926-0bb5-034f-da83-9770517f0762%28Office.15%29.aspx)|
|[Number15](http://msdn.microsoft.com/library/05037ca0-7343-f793-8c86-abfaeba5c0b7%28Office.15%29.aspx)|
|[Number16](http://msdn.microsoft.com/library/9af9d070-bb06-9ba4-da6e-34e9f7e04dfe%28Office.15%29.aspx)|
|[Number17](http://msdn.microsoft.com/library/e1e789d4-3dbb-ca47-ca46-786ded7c8b46%28Office.15%29.aspx)|
|[Number18](http://msdn.microsoft.com/library/7d38aa2a-1075-63ec-0377-7f06917918e2%28Office.15%29.aspx)|
|[Number19](http://msdn.microsoft.com/library/8cac7db2-2b9e-3ee2-628d-9981f6799518%28Office.15%29.aspx)|
|[Number2](http://msdn.microsoft.com/library/a588c314-3950-f0e5-3fa9-5bd24cbb6ff4%28Office.15%29.aspx)|
|[Number20](http://msdn.microsoft.com/library/b5d944bb-b69b-d0d8-ffe8-7c95205a3b6f%28Office.15%29.aspx)|
|[Number3](http://msdn.microsoft.com/library/51d0e7be-aea8-4fda-df9c-e3f855584ccd%28Office.15%29.aspx)|
|[Number4](http://msdn.microsoft.com/library/0e954fb2-bea7-e6ef-5070-87cab4f714c8%28Office.15%29.aspx)|
|[Number5](http://msdn.microsoft.com/library/7c3595ad-caa9-2bce-6d31-8f7e114d4445%28Office.15%29.aspx)|
|[Number6](http://msdn.microsoft.com/library/5e124fd9-cbc7-dd94-d744-55d15d1406b1%28Office.15%29.aspx)|
|[Number7](http://msdn.microsoft.com/library/37d38dc3-cab1-a92c-c56f-f0c6a8065de3%28Office.15%29.aspx)|
|[Number8](http://msdn.microsoft.com/library/1e009c3c-b37e-1ceb-5472-ec1145b82e9e%28Office.15%29.aspx)|
|[Number9](http://msdn.microsoft.com/library/656b64f7-a08c-2d4a-9b3c-01cbd7f02885%28Office.15%29.aspx)|
|[Overallocated](http://msdn.microsoft.com/library/739fcdcd-5ef0-754b-8868-ef3e0662a2e2%28Office.15%29.aspx)|
|[OvertimeCost](http://msdn.microsoft.com/library/5c5ab221-104d-147b-320c-9514acc98447%28Office.15%29.aspx)|
|[OvertimeWork](http://msdn.microsoft.com/library/df885955-c919-82c7-e3c1-5ee6b66440e4%28Office.15%29.aspx)|
|[Owner](http://msdn.microsoft.com/library/d5051b82-a56a-93bb-cf85-81f3f99d3a11%28Office.15%29.aspx)|
|[Parent](http://msdn.microsoft.com/library/0bc76866-8710-6c8b-a7eb-e8650a3baed7%28Office.15%29.aspx)|
|[Peak](http://msdn.microsoft.com/library/52b5d301-6034-b207-c5ae-dfadb56ecd73%28Office.15%29.aspx)|
|[PercentWorkComplete](http://msdn.microsoft.com/library/9535e887-e15c-ebd7-c65f-a3e8d80b8f99%28Office.15%29.aspx)|
|[Project](http://msdn.microsoft.com/library/a51ccbec-7fd9-f296-6f42-f538992d8973%28Office.15%29.aspx)|
|[RegularWork](http://msdn.microsoft.com/library/af65d263-f5e2-158d-bfe0-99062ea1b53c%28Office.15%29.aspx)|
|[RemainingCost](http://msdn.microsoft.com/library/ae7310f7-ac16-fe2f-2efd-4020c114ddab%28Office.15%29.aspx)|
|[RemainingOvertimeCost](http://msdn.microsoft.com/library/6f13f7f0-bc3f-9f58-8047-0fabfa2eccb7%28Office.15%29.aspx)|
|[RemainingOvertimeWork](http://msdn.microsoft.com/library/6db49689-8fb9-e42c-d279-aadca2154bc6%28Office.15%29.aspx)|
|[RemainingWork](http://msdn.microsoft.com/library/94ff4bd9-502c-69f0-a2c2-ac457e677558%28Office.15%29.aspx)|
|[Resource](http://msdn.microsoft.com/library/c24adc5c-9481-5b94-951b-a43fdafaf153%28Office.15%29.aspx)|
|[ResourceGuid](http://msdn.microsoft.com/library/d3def8ce-3eed-700a-2021-71c2b4669697%28Office.15%29.aspx)|
|[ResourceID](http://msdn.microsoft.com/library/8f2a5c6f-a674-5c63-4795-a72b14685d2d%28Office.15%29.aspx)|
|[ResourceName](http://msdn.microsoft.com/library/f0d4e7ff-99b0-70d2-d302-a995a793afbc%28Office.15%29.aspx)|
|[ResourceRequestType](http://msdn.microsoft.com/library/1662d049-5e7e-4a33-528e-784df78a8f5f%28Office.15%29.aspx)|
|[ResourceType](http://msdn.microsoft.com/library/c4a99c35-4241-0739-2b42-05a57cf64ced%28Office.15%29.aspx)|
|[ResourceUniqueID](http://msdn.microsoft.com/library/b6c8b37a-e851-d419-2a28-59d61a640226%28Office.15%29.aspx)|
|[ResponsePending](http://msdn.microsoft.com/library/19fde907-327b-7ecf-3132-9192a2c223aa%28Office.15%29.aspx)|
|[Start](http://msdn.microsoft.com/library/44b132f6-a76a-f5dc-3ac9-28f83a52c4c0%28Office.15%29.aspx)|
|[Start1](http://msdn.microsoft.com/library/06c9ff33-867e-872b-9421-8a8058de3524%28Office.15%29.aspx)|
|[Start10](http://msdn.microsoft.com/library/ef9bc83e-30b4-f46e-d6b4-e908a7e773c9%28Office.15%29.aspx)|
|[Start2](http://msdn.microsoft.com/library/7ce47332-963f-125e-8759-d881b056c0b7%28Office.15%29.aspx)|
|[Start3](http://msdn.microsoft.com/library/2e9998ab-3579-12b6-d3e1-98df62a39a14%28Office.15%29.aspx)|
|[Start4](http://msdn.microsoft.com/library/22750cd1-fa23-1925-1d8e-234c4acf2804%28Office.15%29.aspx)|
|[Start5](http://msdn.microsoft.com/library/6eda3fa3-873c-6920-5cf0-dd15e16c0cb9%28Office.15%29.aspx)|
|[Start6](http://msdn.microsoft.com/library/677a30a3-1f69-0488-ee40-ee336ef7f501%28Office.15%29.aspx)|
|[Start7](http://msdn.microsoft.com/library/0860961d-93d9-a738-7ee7-d0f049b5eb02%28Office.15%29.aspx)|
|[Start8](http://msdn.microsoft.com/library/f6f2dc3d-bc59-cbf5-8cb7-e0604e974e83%28Office.15%29.aspx)|
|[Start9](http://msdn.microsoft.com/library/c533d79f-e78d-94da-f481-043fb91624dc%28Office.15%29.aspx)|
|[StartVariance](http://msdn.microsoft.com/library/080f4dea-76aa-5438-e44a-ab71732b30b1%28Office.15%29.aspx)|
|[Summary](http://msdn.microsoft.com/library/7f8f38f3-c712-0f4e-6b46-0d8eb02119f4%28Office.15%29.aspx)|
|[SV](http://msdn.microsoft.com/library/c63cd139-5a5e-2111-ed52-f239d401f227%28Office.15%29.aspx)|
|[Task](http://msdn.microsoft.com/library/e86d5f79-1e8f-5416-8795-db31cb50eede%28Office.15%29.aspx)|
|[TaskGuid](http://msdn.microsoft.com/library/e08a97f7-6504-b15d-157f-e641112b61c2%28Office.15%29.aspx)|
|[TaskID](http://msdn.microsoft.com/library/71044e84-1388-1b9a-a374-d34f8cdef73b%28Office.15%29.aspx)|
|[TaskName](http://msdn.microsoft.com/library/9fb4480c-520d-1a8b-a07f-b83497e07467%28Office.15%29.aspx)|
|[TaskOutlineNumber](http://msdn.microsoft.com/library/0e356f68-76a8-11df-a723-718c93e61a2c%28Office.15%29.aspx)|
|[TaskSummaryName](http://msdn.microsoft.com/library/a206d327-1ae2-4a09-7029-ac52a517a0a9%28Office.15%29.aspx)|
|[TaskUniqueID](http://msdn.microsoft.com/library/76fef662-2199-7c70-7d69-e97ea8cebb8b%28Office.15%29.aspx)|
|[TeamStatusPending](http://msdn.microsoft.com/library/8e403925-225e-a1e9-121c-6f9353578150%28Office.15%29.aspx)|
|[Text1](http://msdn.microsoft.com/library/67f01a8c-facb-cbfc-64df-e32a053dcab3%28Office.15%29.aspx)|
|[Text10](http://msdn.microsoft.com/library/5d6cc09f-4ef8-7aa9-7840-6a4ba341f55f%28Office.15%29.aspx)|
|[Text11](http://msdn.microsoft.com/library/d4c37d9a-610b-10cd-8811-5ad649fbcaaa%28Office.15%29.aspx)|
|[Text12](http://msdn.microsoft.com/library/93ef9135-d0c5-6961-899d-606c7ec73bc3%28Office.15%29.aspx)|
|[Text13](http://msdn.microsoft.com/library/f00d17b1-a749-8d19-98c5-7cb301005721%28Office.15%29.aspx)|
|[Text14](http://msdn.microsoft.com/library/44456fa9-47c5-d8a7-0bcc-f01d9cd08344%28Office.15%29.aspx)|
|[Text15](http://msdn.microsoft.com/library/98f6ac6f-c443-e7b7-cdaa-e6ddb1046623%28Office.15%29.aspx)|
|[Text16](http://msdn.microsoft.com/library/cd01c1a8-73f9-4fd1-aea4-434256492dbf%28Office.15%29.aspx)|
|[Text17](http://msdn.microsoft.com/library/e5ada6ee-f41f-b7f2-661a-08b84f0a4e71%28Office.15%29.aspx)|
|[Text18](http://msdn.microsoft.com/library/a346d796-70cf-213f-4b0e-6083803215b5%28Office.15%29.aspx)|
|[Text19](http://msdn.microsoft.com/library/288bf010-c3af-047b-459b-75461ec928f5%28Office.15%29.aspx)|
|[Text2](http://msdn.microsoft.com/library/f9111a21-6a9c-d5c9-bff8-235fd2c05b11%28Office.15%29.aspx)|
|[Text20](http://msdn.microsoft.com/library/12bf936c-c4cb-9224-fcc8-ab8b952f6364%28Office.15%29.aspx)|
|[Text21](http://msdn.microsoft.com/library/f74a6191-36e3-fa12-326c-5bd65d1741e1%28Office.15%29.aspx)|
|[Text22](http://msdn.microsoft.com/library/bf9aaf5c-7544-1449-e374-72a368bf6605%28Office.15%29.aspx)|
|[Text23](http://msdn.microsoft.com/library/73a481bb-4a05-6bdc-2a9f-553295c742e6%28Office.15%29.aspx)|
|[Text24](http://msdn.microsoft.com/library/0cb73f81-293b-4281-19fa-022d0af71609%28Office.15%29.aspx)|
|[Text25](http://msdn.microsoft.com/library/67cd48cc-5517-37e4-64a9-2ce4fc609963%28Office.15%29.aspx)|
|[Text26](http://msdn.microsoft.com/library/e01ed7b0-88f1-818f-8548-150945b3bc1f%28Office.15%29.aspx)|
|[Text27](http://msdn.microsoft.com/library/f8c5d733-7a20-979e-7494-e35f52ae6ece%28Office.15%29.aspx)|
|[Text28](http://msdn.microsoft.com/library/70dd5ef5-d25b-4b9e-97d7-b894b1649242%28Office.15%29.aspx)|
|[Text29](http://msdn.microsoft.com/library/11cc5c17-92f0-67f4-1f2d-9e3fb96561b1%28Office.15%29.aspx)|
|[Text3](http://msdn.microsoft.com/library/a2121c88-a787-4118-9451-89024ebe3048%28Office.15%29.aspx)|
|[Text30](http://msdn.microsoft.com/library/62fca21f-d9f2-dbf0-1260-2b5b5cb7f3f5%28Office.15%29.aspx)|
|[Text4](http://msdn.microsoft.com/library/1690718d-d1f2-f4fb-eff1-50719a6cc05c%28Office.15%29.aspx)|
|[Text5](http://msdn.microsoft.com/library/70e4e5d0-c780-1151-688a-59a10df4262f%28Office.15%29.aspx)|
|[Text6](http://msdn.microsoft.com/library/6bb2ea40-e75b-290c-79c7-91702de041e9%28Office.15%29.aspx)|
|[Text7](http://msdn.microsoft.com/library/ad7878f8-8d09-8c4b-d620-ab47c5a40ad0%28Office.15%29.aspx)|
|[Text8](http://msdn.microsoft.com/library/83c2ec8a-a3ad-4f0d-ab72-f9f7c3c1d444%28Office.15%29.aspx)|
|[Text9](http://msdn.microsoft.com/library/f1eb39f5-8403-fa1a-763e-aa3c429414a5%28Office.15%29.aspx)|
|[UniqueID](http://msdn.microsoft.com/library/694aa1b6-eb88-e921-bc4a-b2dfe47df817%28Office.15%29.aspx)|
|[Units](http://msdn.microsoft.com/library/feab9879-5566-a7b6-061d-47e231ac64a1%28Office.15%29.aspx)|
|[UpdateNeeded](http://msdn.microsoft.com/library/5a98cd9e-b467-6bdf-e17f-cf96ee7cf15e%28Office.15%29.aspx)|
|[VAC](http://msdn.microsoft.com/library/27188491-ee6a-f9cf-60d9-ec2876b0c528%28Office.15%29.aspx)|
|[WBS](http://msdn.microsoft.com/library/c3974263-87e9-3102-3c16-712946c926ad%28Office.15%29.aspx)|
|[Work](http://msdn.microsoft.com/library/fe7b1700-2dc4-fcbb-a288-ef3e540319d4%28Office.15%29.aspx)|
|[WorkContour](http://msdn.microsoft.com/library/a47a3012-7e5e-febb-d023-368c7c01e065%28Office.15%29.aspx)|
|[WorkVariance](http://msdn.microsoft.com/library/e92fce82-213f-b412-cc4a-f3c93d11ad8f%28Office.15%29.aspx)|
|[Compliant](http://msdn.microsoft.com/library/bceddf30-8cb4-4098-c354-46c044a97b0a%28Office.15%29.aspx)|

## See also


#### Other resources


[Project Object Model](http://msdn.microsoft.com/library/900b167b-88ec-ea88-15b7-27bb90c22ac6%28Office.15%29.aspx)
