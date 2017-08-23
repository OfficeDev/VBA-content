---
title: "Свойство MailMergeFilterCriterion.Conjunction (издатель)"
keywords: vbapb10.chm6815750
f1_keywords: vbapb10.chm6815750
ms.prod: publisher
api_name: Publisher.MailMergeFilterCriterion.Conjunction
ms.assetid: 79365a25-97fd-a18f-7815-eaccf4c5bdca
ms.date: 06/08/2017
ms.openlocfilehash: 624a1b0ce2b211a3462842bf504de8e3e171ef2b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="mailmergefiltercriterionconjunction-property-publisher"></a>Свойство MailMergeFilterCriterion.Conjunction (издатель)

Возвращает или задает константой **MsoFilterConjunction** , представляющий как условиям фильтра относится к других критериев фильтрации в объекте **[MailMergeFilters](mailmergefilters-object-publisher.md)** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Совместно**

 переменная _expression_A, представляет собой объект- **MailMergeFilterCriterion** .


### <a name="return-value"></a>Возвращаемое значение

MsoFilterConjunction


## <a name="remarks"></a>Заметки

Значение свойства **совместно** может иметь одно из следующих **MsoFilterConjunction** константы, описанные в библиотеке типов, Microsoft Office.



| **msoFilterConjunctionAnd**|| **msoFilterConjunctionOr**|

## <a name="example"></a>Пример

В следующем примере изменяет существующий фильтр для удаления из слияния почты, все записи, не связанные с полем региона равно «WA» и затем добавляет фильтр для следующих фильтра, чтобы условия фильтра должно соответствовать фильтры объединенный и не только один или другое.


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
 Next 
 End With 
End Sub
```


