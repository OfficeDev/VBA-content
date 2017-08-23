---
title: "Событие Document.ShapesAdded (издатель)"
keywords: vbapb10.chm285212675
f1_keywords: vbapb10.chm285212675
ms.prod: publisher
api_name: Publisher.Document.ShapesAdded
ms.assetid: f6573f7c-56fa-1efa-9dba-39cde3859cc0
ms.date: 06/08/2017
ms.openlocfilehash: ce90caaf8d86dbfd0dedec3d25c6e68ff030206e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentshapesadded-event-publisher"></a>Событие Document.ShapesAdded (издатель)

Происходит, когда один или несколько новых фигур добавляются к публикации. Это событие происходит, добавляются ли фигуры автоматически или вручную.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShapesAdded**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

В этом примере выводится сообщение, новые фигуры при добавлении активных публикации. Для работы этого примера необходимо поместить этот код в модуле **ThisDocument** .


```vb
Private Sub PubDoc_ShapesAdded() 
 MsgBox "You just added a new shape." 
End Sub
```


