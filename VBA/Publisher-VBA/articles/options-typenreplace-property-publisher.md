---
title: "Свойство Options.TypeNReplace (издатель)"
keywords: vbapb10.chm1048626
f1_keywords: vbapb10.chm1048626
ms.prod: publisher
api_name: Publisher.Options.TypeNReplace
ms.assetid: 0eb378d2-3554-6a46-8b6b-4a990b4638db
ms.date: 06/08/2017
ms.openlocfilehash: d4a59b41a0b03f30dddd036fa9a5f76f95148324
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="optionstypenreplace-property-publisher"></a>Свойство Options.TypeNReplace (издатель)

 **Значение true** для Microsoft Publisher для замены кластеры не будут считываться азиатских символов из последовательности недопустимый клавиатуры. Чтение и запись **типа Boolean**.


## <a name="syntax"></a>Синтаксис

 _выражение_. **TypeNReplace**

 переменная _expression_A, представляет собой объект- **Параметры** .


### <a name="return-value"></a>Возвращаемое значение

Логический


## <a name="example"></a>Пример

В этом примере указывает, что Publisher для замены кластеры не будут считываться азиатских символов из последовательности недопустимый клавиатуры.


```vb
Sub TypeReplace() 
 Options.TypeNReplace = True 
End Sub
```


