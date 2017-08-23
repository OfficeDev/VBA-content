---
title: "Свойство MailMergeFilterCriterion.CompareTo (издатель)"
keywords: vbapb10.chm6815749
f1_keywords: vbapb10.chm6815749
ms.prod: publisher
api_name: Publisher.MailMergeFilterCriterion.CompareTo
ms.assetid: 6e81fa38-a5d7-8421-6722-a18c5e9a8229
ms.date: 06/08/2017
ms.openlocfilehash: 984f53efccef550a1ba08547c1a2ec87c8a18197
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefiltercriterioncompareto-property-publisher"></a>Свойство MailMergeFilterCriterion.CompareTo (издатель)

Возвращает или задает **строку** , представляющую текст для сравнения в критерий фильтра запроса. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **CompareTo**

 переменная _expression_A, представляет собой объект- **MailMergeFilterCriterion** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В следующем примере изменяется существующий фильтр для удаления из слияния почты все записи, у которых нет поля региона, равное «WA». В этом примере предполагается, что источник данных слияния почты подключенный к активной публикации.


```vb
Sub SetQueryCriterion() 
 Dim intItem As Integer 
 With ActiveDocument.MailMerge.DataSource.Filters 
 For intItem = 1 To .Count 
 With .Item(intItem) 
 If .Column = "Region" Then 
 .Comparison = msoFilterComparisonNotEqual 
 .CompareTo = "WA" 
 If .Conjunction = "Or" Then .Conjunction = "And" 
 End If 
 End With 
 Next intItem 
 End With 
End Sub
```


