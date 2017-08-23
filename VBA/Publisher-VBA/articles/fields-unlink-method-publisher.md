---
title: "Метод Fields.Unlink (издатель)"
keywords: vbapb10.chm6029316
f1_keywords: vbapb10.chm6029316
ms.prod: publisher
api_name: Publisher.Fields.Unlink
ms.assetid: 7a40909f-5fc1-84ef-6679-969a98a8a668
ms.date: 06/08/2017
ms.openlocfilehash: 333ea283bad6ee2f04437d570e71b5ee59a1bdb0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="fieldsunlink-method-publisher"></a>Метод Fields.Unlink (издатель)

Заменяет указанное поле или коллекции **[полей](fields-object-publisher.md)** с их последние результаты.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Разорвать связь**

 переменная _expression_A, представляющий объект **поля** .


### <a name="return-value"></a>Возвращаемое значение

Значение Nothing


## <a name="remarks"></a>Заметки

Когда вы разорвать связь с полем, его текущего результат преобразуется в текст или графику и уже не могут быть обновлены автоматически.


## <a name="example"></a>Пример

В этом примере удаляется связь первого поля в форме одну на первой странице active публикации.


```vb
ActiveDocument.Pages(1).Shapes(1) _ 
 .TextFrame.TextRange.Fields(1).Unlink
```


