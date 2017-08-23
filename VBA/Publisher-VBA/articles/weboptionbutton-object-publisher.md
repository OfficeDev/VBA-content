---
title: "Объект WebOptionButton (издатель)"
keywords: vbapb10.chm4325375
f1_keywords: vbapb10.chm4325375
ms.prod: publisher
api_name: Publisher.WebOptionButton
ms.assetid: acdbaebd-b333-02b1-bf4d-d7e92148a275
ms.date: 06/08/2017
ms.openlocfilehash: 0904d9154044208f3916a7bda2fb6ea1b65816e5
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="weboptionbutton-object-publisher"></a>Объект WebOptionButton (издатель)

Представляет элемент управления button параметр Web. Объект **WebOptionButton** является членом объекта **Shape** .
 


## <a name="example"></a>Пример

Используйте метод **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** для создания новой кнопки параметр Web. Используйте свойство **[WebOptionButton](shape-weboptionbutton-property-publisher.md)** для доступа к кнопки управления Web параметр фигуры. В этом примере создается новая кнопка параметр Web и указывает, что выбран состояние по умолчанию; затем добавляется текстовое поле рядом с ним описать его.
 

 

```
Sub CreateNewWebOptionButton() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlOptionButton, Left:=100, _ 
 Top:=123, Width:=16, Height:=10).WebOptionButton 
 .Selected = msoTrue 
 End With 
 With .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=120, Top:=120, Width:=70, Height:=15) 
 .TextFrame.TextRange.Text = "Advanced User" 
 End With 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](weboptionbutton-application-property-publisher.md)|
|[Родительский раздел](weboptionbutton-parent-property-publisher.md)|
|[ReturnDataLabel](weboptionbutton-returndatalabel-property-publisher.md)|
|[Выбранные](weboptionbutton-selected-property-publisher.md)|
|[Значение](weboptionbutton-value-property-publisher.md)|

