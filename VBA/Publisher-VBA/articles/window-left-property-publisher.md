---
title: "Свойство Window.Left (издатель)"
keywords: vbapb10.chm262149
f1_keywords: vbapb10.chm262149
ms.prod: publisher
api_name: Publisher.Window.Left
ms.assetid: 8d61331a-a70f-4a8a-8dc7-12d93ec51bfc
ms.date: 06/08/2017
ms.openlocfilehash: 643c142e153f1df59f5ab9b948f73c8e677969d0
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="windowleft-property-publisher"></a>Свойство Window.Left (издатель)

Возвращает или задает типа **Long** , указывающее положение (в точках) левого края окна приложения относительно левого края экрана. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Слева**

 переменная _expression_A, представляющий объект **Window** .


## <a name="example"></a>Пример

В этом примере задается горизонтальную позицию окна 100 точек.


```vb
With ActiveDocument.ActiveWindow 
 .WindowState = pbWindowStateNormal 
 .Left = 100 
 .Top = 0 
End With
```


