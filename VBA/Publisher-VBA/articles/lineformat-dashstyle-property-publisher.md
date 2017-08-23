---
title: "Свойство LineFormat.DashStyle (издатель)"
keywords: vbapb10.chm3408132
f1_keywords: vbapb10.chm3408132
ms.prod: publisher
api_name: Publisher.LineFormat.DashStyle
ms.assetid: c2904350-89c1-2fc0-5bae-86f5193c8732
ms.date: 06/08/2017
ms.openlocfilehash: 542c8049efd62f442aa1f3b01dceb51f04f4eb19
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="lineformatdashstyle-property-publisher"></a>Свойство LineFormat.DashStyle (издатель)

Возвращает или задает константой **MsoLineDashStyle** , указывающее, помощник для указанной строки. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **DashStyle**

 переменная _expression_A, представляет собой объект- **LineFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoLineDashStyle


## <a name="remarks"></a>Заметки

Значение свойства **DashStyle** может иметь одно из ** [MsoLineDashStyle](http://msdn.microsoft.com/library/aba7f9d7-1689-c4a8-3b1e-e8dfb4a81d44%28Office.15%29.aspx)** объявленные константы в библиотеке типов, Microsoft Office.


## <a name="example"></a>Пример

В этом примере добавляет синий пунктирная линия active публикации.


```vb
With ActiveDocument.Pages(1).Shapes _ 
 .AddLine(BeginX:=10, BeginY:=10, _ 
 EndX:=250, EndY:=250).Line 
 .DashStyle = msoLineDashDotDot 
 .ForeColor.RGB = RGB(50, 0, 128) 
End With 

```


