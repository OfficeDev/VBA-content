---
title: "Метод Document.UndoClear (издатель)"
keywords: vbapb10.chm196705
f1_keywords: vbapb10.chm196705
ms.prod: publisher
api_name: Publisher.Document.UndoClear
ms.assetid: 63e9bb00-950f-3e30-3897-434362b9efbf
ms.date: 06/08/2017
ms.openlocfilehash: 6066bb50a217319178b43383540bdbba9212ee22
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentundoclear-method-publisher"></a>Метод Document.UndoClear (издатель)

Удаляет список действий, которые можно отменить для указанной публикации. Соответствует список элементов, который отображается, если щелкнуть стрелку рядом с кнопкой " **Отмена** " на панели инструментов **Стандартная** .


## <a name="syntax"></a>Синтаксис

 _выражение_. **UndoClear**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

Включайте данный метод в конце макроса для хранения отображение в поле **Отменить** (например, «VBA-Selection.InsertAfter») действий Microsoft Visual Basic.


## <a name="example"></a>Пример

Этот пример удаляет список действий, которые можно отменить для активной публикации.


```vb
ActiveDocument.UndoClear
```


