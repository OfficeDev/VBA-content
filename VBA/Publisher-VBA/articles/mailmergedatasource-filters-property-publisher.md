---
title: "Свойство MailMergeDataSource.Filters (издатель)"
keywords: vbapb10.chm6291463
f1_keywords: vbapb10.chm6291463
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.Filters
ms.assetid: 7b8fa974-08e5-9691-c69d-314eb6a5c651
ms.date: 06/08/2017
ms.openlocfilehash: 367f0f96d18b12de7a372eefdb6b19ed97c01366
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourcefilters-property-publisher"></a>Свойство MailMergeDataSource.Filters (издатель)

Возвращает объект **[MailMergeFilters](mailmergefilters-object-publisher.md)** , представляющий фильтры для слияния почты и каталогов источник данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Фильтры**

 переменная _expression_A, представляющий объект **вывода** .


### <a name="return-value"></a>Возвращаемое значение

MailMergeFilters


## <a name="example"></a>Пример

В этом примере добавляется новый фильтр, который удаляет все записи с пустым полем региона и затем применяет фильтр для активной публикации. В этом примере предполагается, что источник данных слияния почты подключенный к активной публикации.


```vb
Sub FilterDataSource() 
 With ActiveDocument.MailMerge.DataSource 
 .Filters.Add Column:="Region", _ 
 Comparison:=msoFilterComparisonIsBlank, _ 
 Conjunction:=msoFilterConjunctionAnd 
 .ApplyFilter 
 End With 
End Sub
```


