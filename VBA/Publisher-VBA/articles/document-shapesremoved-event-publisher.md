---
title: "Событие Document.ShapesRemoved (издатель)"
keywords: vbapb10.chm285212677
f1_keywords: vbapb10.chm285212677
ms.prod: publisher
api_name: Publisher.Document.ShapesRemoved
ms.assetid: e2a67359-5673-2c72-e1fc-e3e3a3b564f9
ms.date: 06/08/2017
ms.openlocfilehash: c067ce7d418980f1645232d4528e50eb2339674d
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentshapesremoved-event-publisher"></a>Событие Document.ShapesRemoved (издатель)

Происходит, когда фигура удаляется из публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShapesRemoved**

 переменная _expression_A, представляющий объект **Document** .


## <a name="example"></a>Пример

В этом примере отображается сообщение при фигуры удаляется из активной публикации. Для работы этого примера необходимо поместить этот код в модуле **ThisDocument** .


```vb
Private Sub Document_ShapesRemoved() 
 MsgBox "You just deleted one or more shapes." 
End Sub
```


