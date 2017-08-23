---
title: "Событие Application.WindowDeactivate (издатель)"
keywords: vbapb10.chm268435458
f1_keywords: vbapb10.chm268435458
ms.prod: publisher
api_name: Publisher.Application.WindowDeactivate
ms.assetid: 84473784-7c03-4c9e-3e1b-9bf6ec7e1fbc
ms.date: 06/08/2017
ms.openlocfilehash: e4c571705ca8da0349f5f64e21208efd15077078
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="applicationwindowdeactivate-event-publisher"></a>Событие Application.WindowDeactivate (издатель)

Происходит при отключении окна приложения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WindowDeactivate** ( **_Низ_**)

 переменная _expression_A, представляющий объект **приложения** .


### <a name="parameters"></a>Параметры



|**Имя**|**Обязательный или необязательный**|**Тип данных**|**Описание**|
|:-----|:-----|:-----|:-----|
|Низ|Обязательное свойство.| **Окно**|Окно, — происходит деактивация функции.|

## <a name="remarks"></a>Заметки

Сведения об использовании событий с помощью объекта приложения [С помощью событий объекта](using-events-with-the-application-object-publisher.md)см.


## <a name="example"></a>Пример

В этом примере Свертывание окна, если он отключен. Этот код должен находиться в модуле класса и экземпляр класса необходимо правильно инициализировать данный пример работы; просмотра [Событий с помощью объекта](using-events-with-the-application-object-publisher.md)для указания о том, как это сделать.


```vb
Public WithEvents appPublisher as Publisher.Application 
 
Private Sub appPublisher_WindowDeactivate _ 
 (ByVal Wn As Window) 
 Wn.WindowState = pbWindowStateMinimize 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект приложения](application-object-publisher.md)

