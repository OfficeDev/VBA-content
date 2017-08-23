---
title: "Свойство AdvancedPrintOptions.GraphicsResolution (издатель)"
keywords: vbapb10.chm7077909
f1_keywords: vbapb10.chm7077909
ms.prod: publisher
api_name: Publisher.AdvancedPrintOptions.GraphicsResolution
ms.assetid: 1e4e06aa-327b-5689-ff97-eea9f866260a
ms.date: 06/08/2017
ms.openlocfilehash: 78e6c0000d7c55bdff9f6a35a88e9a4fd97d40d2
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="advancedprintoptionsgraphicsresolution-property-publisher"></a>Свойство AdvancedPrintOptions.GraphicsResolution (издатель)

Возвращает или задает значение константы **PbPrintGraphics** , представляющее разрешения, на которой можно вставленные изображения на печать в указанной публикации. Чтение и запись.


## <a name="syntax"></a>Синтаксис

 _выражение_. **GraphicsResolution**

 переменная _expression_A, представляет собой объект- **AdvancedPrintOptions** .


### <a name="return-value"></a>Возвращаемое значение

PbPrintGraphics


## <a name="remarks"></a>Заметки

Установка для этого свойства влияет только на вставленных изображений (ли связанных или внедренных) и Коллекция картинок. Картинка автофигуры и рамка всегда печатается.

Печать полей вместо графики полезен при печати быстрого обоснования разметки, отображающего исключительно размещение рисунков.

Это свойство соответствует **графики** элементов управления на вкладке **графики и шрифты** диалоговое окно **Дополнительные параметры печати** .

Значение свойства **GraphicsResolution** может иметь одно из **[PbPrintGraphics](pbprintgraphics-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.


## <a name="example"></a>Пример

В следующем примере задается графики для печати в виде полях active публикации.


```vb
Sub PrintGraphicAsBoxes 
 With ActiveDocument.AdvancedPrintOptions 
 If .GraphicsResolution <> pbPrintNoGraphics Then 
 .GraphicsResolution = pbPrintNoGraphics 
 End If 
 End With 
End Sub
```


## <a name="see-also"></a>См. также


#### <a name="concepts"></a>Основные понятия


 [Объект AdvancedPrintOptions](advancedprintoptions-object-publisher.md)

