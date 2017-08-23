---
title: "Событие Application.WindowPageChange (издатель)"
keywords: vbapb10.chm268435460
f1_keywords: vbapb10.chm268435460
ms.prod: publisher
api_name: Publisher.Application.WindowPageChange
ms.assetid: bb636f6e-da4b-7271-9f59-2b7000270c16
ms.date: 06/08/2017
ms.openlocfilehash: 3e4b08b47c43a179b9b2ff7b060c6b3f259a5924
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationwindowpagechange-event-publisher"></a>Событие Application.WindowPageChange (издатель)

Происходит при переключении в представление с одной страницы на другую страницу в публикации.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WindowPageChange** ( **_Фольцваген_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Фольцваген|Обязательное свойство.| **Просмотр**|Новое представление, которое включает в себя страницы, для которого было выполнено переключение.|

## <a name="example"></a>Пример

В этом примере изменяется представление для отображения в пределах всей страницы при переходе на новую страницу в публикации. Для работы этого примера необходимо поместить объявления **ключевое слово WithEvents** в раздел общих объявлений модуль класса и выполнить процедуру InitializeEvents.


```vb
Private WithEvents PubApp As Publisher.Application 
 
Sub InitializeEvents() 
 Set PubApp = Publisher.Application 
End Sub 
 
Private Sub PubApp_WindowPageChange(ByVal Vw As View) 
 Vw.Zoom = pbZoomWholePage 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

