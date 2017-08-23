---
title: "Объект ColorFormat (издатель)"
keywords: vbapb10.chm2621439
f1_keywords: vbapb10.chm2621439
ms.prod: publisher
api_name: Publisher.ColorFormat
ms.assetid: 659069e1-e359-94d7-de06-a1d98378193b
ms.date: 06/08/2017
ms.openlocfilehash: 6bec281e254dc06c3ec9c31f3f1488b52346d973
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorformat-object-publisher"></a>Объект ColorFormat (издатель)

Представляет цвет переднего плана или объект один цвет или цвет фона объекта с градиента и узорные заливки. Можно задать цвета явного красный зеленый синий значения с помощью свойства **[RGB](colorformat-rgb-property-publisher.md)** .
 


## <a name="remarks"></a>Заметки

Возвращает объект **ColorFormat** , используйте один из свойств, перечисленных в следующей таблице.
 

 


|**Использование этого свойства**|**С помощью этого объекта**|**Возвращает объект ColorFormat, который представляет это**|
|:-----|:-----|:-----|
|**[Цвет фона](fillformat-backcolor-property-publisher.md)**|**[FillFormat](fillformat-object-publisher.md)**|Цвет заливки фона (используется в затемненные или узорную заливку)|
|**[Цвет текста](fillformat-forecolor-property-publisher.md)**|**FillFormat**|Цвет заливки (или цвет заливки для сплошной заливке)|
|**[Цвет фона](lineformat-backcolor-property-publisher.md)**|**[LineFormat](lineformat-object-publisher.md)**|Цвет фона строки (используется в узорная линия)|
|**[Цвет текста](lineformat-forecolor-property-publisher.md)**|**LineFormat**|Цвет строки (или цвет линии для сплошной линии)|
|**[Цвет текста](shadowformat-forecolor-property-publisher.md)**|**[ShadowFormat](shadowformat-object-publisher.md)**|Цвет затенения|
|**[ExtrusionColor](threedformat-extrusioncolor-property-publisher.md)**|**[ThreeDFormat](threedformat-object-publisher.md)**|Цвет границы Вытянутый объект|

## <a name="example"></a>Пример

Используйте свойство **RGB** задать цвет явного значения красный зеленый синий. В следующем примере добавляет прямоугольник active публикации и затем задает цвет переднего плана, цвет фона и градиент для заливки прямоугольника.
 

 

```
Sub GradientFill() 
 With ActiveDocument.Pages(1).Shapes _ 
 .AddShape(Type:=msoShapeRectangle, _ 
 Left:=90, Top:=90, Width:=90, Height:=50).Fill 
 .ForeColor.RGB = RGB(128, 0, 0) 
 .BackColor.RGB = RGB(170, 170, 170) 
 .TwoColorGradient msoGradientHorizontal, 1 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](colorformat-application-property-publisher.md)|
|[BaseCMYK](colorformat-basecmyk-property-publisher.md)|
|[BaseRGB](colorformat-basergb-property-publisher.md)|
|[ФОРМАТ CMYK](colorformat-cmyk-property-publisher.md)|
|[Рукописного ввода](colorformat-ink-property-publisher.md)|
|[Родительский раздел](colorformat-parent-property-publisher.md)|
|[RGB](colorformat-rgb-property-publisher.md)|
|[SchemeColor](colorformat-schemecolor-property-publisher.md)|
|[TintAndShade](colorformat-tintandshade-property-publisher.md)|
|[Прозрачность](colorformat-transparency-property-publisher.md)|
|[Type](colorformat-type-property-publisher.md)|

