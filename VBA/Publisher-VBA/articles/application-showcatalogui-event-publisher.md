---
title: "Событие Application.ShowCatalogUI (издатель)"
keywords: vbapb10.chm268435493
f1_keywords: vbapb10.chm268435493
ms.prod: publisher
api_name: Publisher.Application.ShowCatalogUI
ms.assetid: 8a5a3798-4b95-d77f-70f6-d69dd9dc8f99
ms.date: 06/08/2017
ms.openlocfilehash: 2d043e6a43509ddd2f63085f13a3bf4e37db4f97
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationshowcatalogui-event-publisher"></a>Событие Application.ShowCatalogUI (издатель)

Активируется, когда каталога мастера публикации отображается в интерфейсе пользователя Microsoft Publisher.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ShowCatalogUI**

 _expression_An выражение, возвращающее объект **приложения** .


## <a name="remarks"></a>Заметки

Можно использовать ** [Application.ShowWizardCatalog](application-showwizardcatalog-method-publisher.md)** метод для отображения мастера каталога в пользовательском интерфейсе.

**ShowCatalogUI** события не возникают при публикации каталога отображается при первом запуске Publisher. Чтобы определить, если каталог отображается в это время, можно использовать свойство **[WizardCatalogVisible](application-wizardcatalogvisible-property-publisher.md)** .

Дополнительные сведения об использовании событий с помощью объекта **приложения** [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как обрабатывать события **ShowCatalogUI** . Будет выведено сообщение уведомления об этом пользователя отображение пользовательского интерфейса каталога.


```vb
Private Sub pubApplication_ShowCatalogUI() 
 MsgBox "The Wizard Catalog is displayed." 
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

