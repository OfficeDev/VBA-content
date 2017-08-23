---
title: "Событие Application.HideCatalogUI (издатель)"
keywords: vbapb10.chm268435494
f1_keywords: vbapb10.chm268435494
ms.prod: publisher
api_name: Publisher.Application.HideCatalogUI
ms.assetid: a7ac7594-18fe-355e-d270-d205c405862a
ms.date: 06/08/2017
ms.openlocfilehash: ab711cb964973a43df86aa869e89f49d7158fd61
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationhidecatalogui-event-publisher"></a>Событие Application.HideCatalogUI (издатель)

Происходит, когда каталога мастера публикации скрыта в пользовательском интерфейсе Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **HideCatalogUI**

 _expression_An выражение, возвращающее объект **приложения** .


## <a name="remarks"></a>Заметки

Дополнительные сведения об использовании событий с помощью объекта **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как обрабатывать события **HideCatalogUI** . Будет выведено сообщение о том, что скрыто каталога пользовательского интерфейса.


```vb
Private Sub pubApplication_HideCatalogUI() 
 MsgBox "The Wizard Catalog is hidden." 
End Sub
```

Для чтобы произошло это событие необходимо включить следующую строку кода в разделе **Общие описаний** модуля.




```vb
Private WithEvents pubApplication As Application
```

Затем выполните следующую процедуру инициализации.




```vb
Public Sub Initialize_pubApplication() 
 Set pubApplication = Publisher.Application 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

