---
title: "Метод Document.Save (издатель)"
keywords: vbapb10.chm196695
f1_keywords: vbapb10.chm196695
ms.prod: publisher
api_name: Publisher.Document.Save
ms.assetid: 89eae461-d1c2-b3ca-58b7-9528df8801d8
ms.date: 06/08/2017
ms.openlocfilehash: 69527684f9a46999d027a4d492920279774c2989
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentsave-method-publisher"></a>Метод Document.Save (издатель)

Сохранение указанной публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Сохранение**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

При публикации ранее не была сохранена, вызвав метод **Save** эквивалентен вызову метода **[SaveAs](document-saveas-method-publisher.md)** в аргументе **_FileName_** , значение свойства **[Name](application-name-property-publisher.md)** публикации. Если ранее сохранить публикацию метод **Save** сохраняет текущую версию публикации в формате, в котором было открыто и в расположение, к которому последнего сохранения.

При вызове метода **Save** всегда выполняется сохранение на переднем плане независимо от того, включен ли фоновое сохранение.


## <a name="example"></a>Пример

В этом примере сохраняет active публикации, если он был изменен с момента последнего сохранения.


```vb
If ActiveDocument.Saved = False Then ActiveDocument.Save
```


