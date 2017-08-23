---
title: "Событие Document.Redo (издатель)"
keywords: vbapb10.chm285212679
f1_keywords: vbapb10.chm285212679
ms.prod: publisher
api_name: Publisher.Document.Redo
ms.assetid: c00db13d-1c03-2536-8923-bd7d9393fee2
ms.date: 06/08/2017
ms.openlocfilehash: 21b8b4df75155aafa450c2410708533470c3347e
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="documentredo-event-publisher"></a>Событие Document.Redo (издатель)

Происходит, когда реверсировании последнее действие, которое было отменено.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Повторное применение**

 переменная _expression_A, представляющий объект **Document** .


## <a name="remarks"></a>Заметки

Событие **Повторить** происходит сразу же после применить действие.

Если повторно выполняется несколько действий, событие **Повторить** происходит только один раз, после завершения всех действий.

Дополнительные сведения об использовании событий с помощью объекта **Document** содержатся в разделе [С помощью событий с помощью объекта Document](using-events-with-the-document-object-publisher.md).


## <a name="example"></a>Пример

В этом примере выводится сообщение, когда пользователь нажимает кнопку **Отмена** на панели инструментов **Стандартная** или выбирает **Вернуть** в меню **Правка** . Для этой процедуры для работы с текущей публикации необходимо поместить его в модуле ThisDocument.


```vb
Private Sub DocPub_Redo() 
 MsgBox "Your last undo has been reversed." 
End Sub
```

Для перехвата это событие из Microsoft Publisher проекта, необходимо поместить следующий код в раздел общих объявлений модуля и запуска процедуры InitiatePubApp.




```vb
Private WithEvents DocPub As Publisher.Document 
 
Sub InitiatePubApp() 
 Set DocPub = Publisher.ActiveDocument 
End Sub
```


