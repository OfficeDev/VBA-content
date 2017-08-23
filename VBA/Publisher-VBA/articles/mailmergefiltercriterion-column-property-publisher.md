---
title: "Свойство MailMergeFilterCriterion.Column (издатель)"
keywords: vbapb10.chm6815747
f1_keywords: vbapb10.chm6815747
ms.prod: publisher
api_name: Publisher.MailMergeFilterCriterion.Column
ms.assetid: 000b4b4c-73a1-ea9f-6f44-bc6eac15cb4b
ms.date: 06/08/2017
ms.openlocfilehash: b012327e7da28938bcfcab36e23966e6d3515b13
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefiltercriterioncolumn-property-publisher"></a>Свойство MailMergeFilterCriterion.Column (издатель)

Возвращает **строку** , представляющую имя поля в источнике данных слияния почты для использования в фильтре. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Столбец**

 переменная _expression_A, представляет собой объект- **MailMergeFilterCriterion** .


## <a name="example"></a>Пример

В следующем примере изменяется существующий фильтр для удаления из слияния почты все записи, у которых нет поля региона, равное «WA».


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


