---
title: "Объект формы (издатель)"
keywords: vbapb10.chm2949119
f1_keywords: vbapb10.chm2949119
ms.prod: publisher
api_name: Publisher.Plate
ms.assetid: f7d7dbb1-a6a4-780f-814e-8e95aaaeeeea
ms.date: 06/08/2017
ms.openlocfilehash: dcd562397190161e7a2790978e56b242ff4a2087
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="plate-object-publisher"></a>Объект формы (издатель)

Представляет один принтер формы. Объект **формы** является элементом коллекции **[формы](plates-object-publisher.md)** .
 


## <a name="example"></a>Пример

Используйте метод **[Add](plates-add-method-publisher.md)** коллекции **[формы](plates-object-publisher.md)** для создания новой формы. В этом примере создается коллекция цвет область формы и добавляет форму.
 

 

```
Sub AddNewPlates() 
 Dim plts As Plates 
 Set plts = ActiveDocument.CreatePlateCollection(Mode:=pbColorModeSpot) 
 plts.Add 
 With plts(1) 
 .Color.RGB = RGB(Red:=255, Green:=0, Blue:=0) 
 .Luminance = 4 
 End With 
End Sub
```

Используйте метод **[FindPlateByInkName](plates-findplatebyinkname-method-publisher.md)** для возврата определенного формы с учетом его рукописного ввода имени. Процесс цвета назначены разные номера индекса в коллекции **формы** , чем в коллекции **[PrintablePlates](printableplates-object-publisher.md)** . Используйте метод **FindPlateByInkName** для гарантии на желаемую **формы** или получить доступ к объекту **[PrintablePlate](printableplate-object-publisher.md)** .
 

 

## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[ConvertToProcess](plate-converttoprocess-method-publisher.md)|
|[Delete](plate-delete-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](plate-application-property-publisher.md)|
|[Цвет](plate-color-property-publisher.md)|
|[Index](plate-index-property-publisher.md)|
|[InkName](plate-inkname-property-publisher.md)|
|[Может быть каталогом](plate-inuse-property-publisher.md)|
|[Яркости](plate-luminance-property-publisher.md)|
|[Name](plate-name-property-publisher.md)|
|[Родительский раздел](plate-parent-property-publisher.md)|

