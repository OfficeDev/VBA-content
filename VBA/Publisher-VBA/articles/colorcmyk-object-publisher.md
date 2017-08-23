---
title: "Объект ColorCMYK (издатель)"
keywords: vbapb10.chm2686975
f1_keywords: vbapb10.chm2686975
ms.prod: publisher
api_name: Publisher.ColorCMYK
ms.assetid: e1a39f6f-f440-e375-4f8c-e81093e5a451
ms.date: 06/08/2017
ms.openlocfilehash: 2f33c5bef8e24d4f5f987dcbb346590d2849de2b
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="colorcmyk-object-publisher"></a>Объект ColorCMYK (издатель)

Представляет значение голубой пурпурный желтый черный цвет (CMYK).
 


## <a name="example"></a>Пример

Используйте свойство **CMYK** объекта **ColorFormat** возвращает объект **ColorCMYK** . Используйте **голубой**, **пурпурный**, **желтый**и **черный** свойства объекта **ColorCMYK** по отдельности установка каждого из четырех цветов в формат CMYK значение цвета. Используйте метод **SetCMYK** для объекта **ColorCMYK** для установки всех четырех цветов за один раз.
 

 

 

 
В следующем примере показано получение CMYK значение цвета заливки фигуры из них и изменяется на другую CMYK значение цвета.
 

 



```
Dim cmykColor As ColorCMYK Set cmykColor = ActiveDocument.Pages(1).Shapes(1).Fill.ForeColor.CMYK cmykColor.SetCMYK Cyan:=0, Magenta:=255, Yellow:=255, Black:=50
```


## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[SetCMYK](colorcmyk-setcmyk-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](colorcmyk-application-property-publisher.md)|
|[Черный](colorcmyk-black-property-publisher.md)|
|[Голубой](colorcmyk-cyan-property-publisher.md)|
|[Пурпурный](colorcmyk-magenta-property-publisher.md)|
|[Родительский раздел](colorcmyk-parent-property-publisher.md)|
|[Желтый](colorcmyk-yellow-property-publisher.md)|

