---
title: "Свойство MailMergeFilterCriterion.Comparison (издатель)"
keywords: vbapb10.chm6815748
f1_keywords: vbapb10.chm6815748
ms.prod: publisher
api_name: Publisher.MailMergeFilterCriterion.Comparison
ms.assetid: ba815a39-35d6-803e-39c4-deba30646e66
ms.date: 06/08/2017
ms.openlocfilehash: 4c7a89afb316eccef5d2c23f215cb5755f844c3f
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefiltercriterioncomparison-property-publisher"></a>Свойство MailMergeFilterCriterion.Comparison (издатель)

Возвращает или задает константой **MsoFilterComparison** , представляющий сравнение свойств [столбца](cell-column-property-publisher.md) и **[CompareTo](mailmergefiltercriterion-compareto-property-publisher.md)** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сравнение**

 переменная _expression_A, представляет собой объект- **MailMergeFilterCriterion** .


### <a name="return-value"></a>Возвращаемое значение

MsoFilterComparison


## <a name="remarks"></a>Заметки

Значение свойства **сравнения** может иметь одно из ** [MsoFilterComparison](http://msdn.microsoft.com/library/12650101-777b-2142-e985-cc34d5e2fb16%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


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


