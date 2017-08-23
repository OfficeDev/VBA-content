---
title: "Объект вывода (издатель)"
keywords: vbapb10.chm6356991
f1_keywords: vbapb10.chm6356991
ms.prod: publisher
api_name: Publisher.MailMergeDataSource
ms.assetid: a02eb4fb-7db7-e533-c3ca-95bc4ca68e82
ms.date: 06/08/2017
ms.openlocfilehash: 9b7dd098d16c95ebc9523d2e2d2929b73106d265
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasource-object-publisher"></a>Объект вывода (издатель)

Представляет источник данных в операции объединения слияния почты и каталогов.
 


## <a name="example"></a>Пример

Свойство **[DataSource](mailmerge-datasource-property-publisher.md)** используется для возврата объекта **вывода** . Следующий пример отображает имя источника данных, связанного с активной публикации.
 

 

```
Sub ShowDataSourceName() 
 If ActiveDocument.MailMerge.DataSource.Name <> "" Then _ 
 MsgBox ActiveDocument.MailMerge.DataSource.Name 
End Sub
```

Следующий пример проверяет источник open данных, связанных с active публикации, чтобы определить, содержит ли фамилия имя Белова.
 

 



```
Sub FindSelectedRecord() 
 With ActiveDocument.MailMerge 
 If .DataSource.FindRecord(FindText:="Fuller", _ 
 Field:="LastName") = True Then 
 MsgBox "Data was found" 
 End If 
 End With 
End Sub
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[ApplyFilter](mailmergedatasource-applyfilter-method-publisher.md)|
|[Закрыть](mailmergedatasource-close-method-publisher.md)|
|[ИзменитьЗапись](mailmergedatasource-editrecord-method-publisher.md)|
|[НайтиЗапись](mailmergedatasource-findrecord-method-publisher.md)|
|[OpenRecipientsDialog](mailmergedatasource-openrecipientsdialog-method-publisher.md)|
|[SetAllErrorFlags](mailmergedatasource-setallerrorflags-method-publisher.md)|
|[SetAllIncludedFlags](mailmergedatasource-setallincludedflags-method-publisher.md)|
|[SetSortOrder](mailmergedatasource-setsortorder-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[ActiveRecord](mailmergedatasource-activerecord-property-publisher.md)|
|[Приложения](mailmergedatasource-application-property-publisher.md)|
|[Стрсоедин](mailmergedatasource-connectstring-property-publisher.md)|
|[DataFields](mailmergedatasource-datafields-property-publisher.md)|
|[Источники данных](mailmergedatasource-datasources-property-publisher.md)|
|[EverValidated](mailmergedatasource-evervalidated-property-publisher.md)|
|[Фильтры](mailmergedatasource-filters-property-publisher.md)|
|[FirstRecord](mailmergedatasource-firstrecord-property-publisher.md)|
|[Включенные](mailmergedatasource-included-property-publisher.md)|
|[InvalidAddress](mailmergedatasource-invalidaddress-property-publisher.md)|
|[InvalidComments](mailmergedatasource-invalidcomments-property-publisher.md)|
|[IsMaster](mailmergedatasource-ismaster-property-publisher.md)|
|[LastRecord](mailmergedatasource-lastrecord-property-publisher.md)|
|[MappedDataFields](mailmergedatasource-mappeddatafields-property-publisher.md)|
|[Name](mailmergedatasource-name-property-publisher.md)|
|[Родительский раздел](mailmergedatasource-parent-property-publisher.md)|
|[RecordCount](mailmergedatasource-recordcount-property-publisher.md)|
|[TableName](mailmergedatasource-tablename-property-publisher.md)|
|[Type](mailmergedatasource-type-property-publisher.md)|
|[ValidatedClean](mailmergedatasource-validatedclean-property-publisher.md)|

