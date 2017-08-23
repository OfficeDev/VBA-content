---
title: "Метод MailMergeDataSource.ApplyFilter (издатель)"
keywords: vbapb10.chm6291492
f1_keywords: vbapb10.chm6291492
ms.prod: publisher
api_name: Publisher.MailMergeDataSource.ApplyFilter
ms.assetid: a94af75c-e558-7160-76c9-c0f8c3fb317d
ms.date: 06/08/2017
ms.openlocfilehash: 31059d58942dde02e2b1e342b902929b11224e97
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergedatasourceapplyfilter-method-publisher"></a>Метод MailMergeDataSource.ApplyFilter (издатель)

Применяет фильтр для слияния почты источник данных, чтобы удалить (или отфильтровать) указан записей содержащего (или не содержащие) определенного данных.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ApplyFilter**

 переменная _expression_A, представляющий объект **вывода** .


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


