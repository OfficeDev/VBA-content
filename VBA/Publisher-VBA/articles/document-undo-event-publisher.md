---
title: "Событие Document.Undo (издатель)"
keywords: vbapb10.chm285212678
f1_keywords: vbapb10.chm285212678
ms.prod: publisher
api_name: Publisher.Document.Undo
ms.assetid: 9789e469-dc84-a0b7-ffe0-405d4e7ad861
ms.date: 06/08/2017
ms.openlocfilehash: a956876a39505069cb0fc38cd0d8dbf68c3a68af
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentundo-event-publisher"></a>Событие Document.Undo (издатель)

Происходит, когда пользователь отменяет последнее действие выполнять.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Отменить**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

**Отменить** событие происходит сразу же после отменить действие.

Если несколько действий будут отменены, **Отменить** событие происходит только один раз после отменяются все действия.

Дополнительные сведения об использовании событий с помощью объекта **Document** содержатся в разделе [С помощью событий с помощью объекта Document](using-events-with-the-document-object-publisher.md).


## <a name="example"></a>Пример

В этом примере выводится сообщение, когда пользователь нажимает кнопку **Отмена** на панели инструментов **Стандартная** или выбирает **Отменить** в меню **Правка** . Для этой процедуры для работы с текущей публикации необходимо поместить его в модуле ThisDocument.


```vb
Private Sub DocPub_Undo() 
 MsgBox "Your last action has been reversed." 
End Sub
```

Для перехвата это событие из Microsoft Publisher проекта, необходимо поместить следующий код в раздел общих объявлений модуля и запуска процедуры InitiatePubApp.




```vb
Private WithEvents DocPub As Publisher.Document 
 
Sub InitiatePubApp() 
 Set DocPub = Publisher.ActiveDocument 
End Sub
```


