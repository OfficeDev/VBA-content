---
title: "Свойство PictureFormat.TransparencyColor (издатель)"
keywords: vbapb10.chm3604743
f1_keywords: vbapb10.chm3604743
ms.prod: publisher
api_name: Publisher.PictureFormat.TransparencyColor
ms.assetid: 908d2e21-3e2a-b75b-a82d-454686b7ecb8
ms.date: 06/08/2017
ms.openlocfilehash: 09f7b44c84709ca3cdc40ff8b23bce1ee77723ee
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformattransparencycolor-property-publisher"></a>Свойство PictureFormat.TransparencyColor (издатель)

Возвращает или задает константой **MsoRGBType** , представляющий прозрачность цвета. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **Прозрачного цвета**

 переменная _expression_A, представляет собой объект- **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

MsoRGBType


## <a name="example"></a>Пример

В этом примере создается изображение на первой странице и задает цвет прозрачность в черный цвет.


```vb
Sub SetTransparentColor() 
 With ActiveDocument.Pages(1).Shapes.AddPicture( _ 
 FileName:="C:\My Pictures\Sample.gif", LinkToFile:=msoFalse, _ 
 SaveWithDocument:=msoTrue, Left:=36, Top:=36) 
 .PictureFormat.TransparencyColor = RGB(Red:=255, Green:=255, Blue:=255) 
 End With 
End Sub
```


