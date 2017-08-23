---
title: "Свойство Window.WindowState (издатель)"
keywords: vbapb10.chm262160
f1_keywords: vbapb10.chm262160
ms.prod: publisher
api_name: Publisher.Window.WindowState
ms.assetid: 063ede5e-f279-09e3-5672-b634c752b927
ms.date: 06/08/2017
ms.openlocfilehash: 1554212f6cdc43bf0ecccbff61948aaef4b27cc7
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowwindowstate-property-publisher"></a>Свойство Window.WindowState (издатель)

Возвращает или задает значение, указывающее состояние окна, Microsoft Publisher константы **PbWindowState** . Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **WindowState**

 переменная _expression_A, представляющий объект **Window** .


### <a name="return-value"></a>Возвращаемое значение

PbWindowState


## <a name="remarks"></a>Заметки

Значение свойства **WindowState** может иметь одно из следующих констант **PbWindowState** .



| **pbWindowStateMaximize**|| **pbWindowStateMinimize**|| **pbWindowStateNormal**| Когда состояние окна **pbWindowStateNormal**, окно не развернуто и не свернуто.


## <a name="example"></a>Пример

В этом примере разворачивает окно Publisher.


```vb
ActiveWindow.WindowState = pbWindowStateMaximized
```


