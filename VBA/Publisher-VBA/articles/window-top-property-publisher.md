---
title: "Свойство Window.Top (издатель)"
keywords: vbapb10.chm262148
f1_keywords: vbapb10.chm262148
ms.prod: publisher
api_name: Publisher.Window.Top
ms.assetid: 22fe0170-7433-a917-87ca-f418c2aefc70
ms.date: 06/08/2017
ms.openlocfilehash: a07f15a38560b4daf3a8b931ed6e076df5384b07
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowtop-property-publisher"></a>Свойство Window.Top (издатель)

Возвращает или задает типа **Long** , представляющее расстояние между верхнего края экрана и окна приложения. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **В начало**

 переменная _expression_A, представляющий объект **Window** .


## <a name="example"></a>Пример

В этом примере проверяется состояние окна приложения — ни развернуто, ни свернуто и затем изменяет размер окна и переводит его 150 точек из верхней части экрана.


```vb
Sub MoveWindow() 
 With ActiveWindow 
 If .WindowState = pbWindowStateNormal Then 
 .Top = 150 
 .Resize Width:=500, Height:=500 
 End If 
 End With 
End Sub
```


