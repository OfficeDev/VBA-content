---
title: "Свойство Document.Stories (издатель)"
keywords: vbapb10.chm196659
f1_keywords: vbapb10.chm196659
ms.prod: publisher
api_name: Publisher.Document.Stories
ms.assetid: 4ffc7d20-eb11-942e-e28a-81c2caa19a50
ms.date: 06/08/2017
ms.openlocfilehash: fc546dec01d16c29cf35a5bbe7b5d76515cafd98
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentstories-property-publisher"></a>Свойство Document.Stories (издатель)

Возвращает коллекцию **[истории (en)](stories-object-publisher.md)** , содержащую всех статьях публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Истории (en)**

 переменная _expression_A, представляющий объект **Document** .


### <a name="return-value"></a>Возвращаемое значение

Истории (en)


## <a name="example"></a>Пример

В этом примере первая статья в коллекции **функциональности** присваивается переменной.


```vb
Sub FirstStory() 
 
 Dim stFirst As Story 
 
 stFirst = Application.ActiveDocument.Stories(1) 
 
End Sub
```


