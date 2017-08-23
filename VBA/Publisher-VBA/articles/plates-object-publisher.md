---
title: "Объект формы (издатель)"
keywords: vbapb10.chm2883583
f1_keywords: vbapb10.chm2883583
ms.prod: publisher
api_name: Publisher.Plates
ms.assetid: 7da44b06-c94f-dadc-da91-09b757d5a076
ms.date: 06/08/2017
ms.openlocfilehash: c7fd59eb06abd46fb17f95717713cfd1ab2b3e76
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="plates-object-publisher"></a>Объект формы (издатель)

Коллекция объектов **формы** в публикации.
 


## <a name="example"></a>Пример

Набор **печатных форм** состоит из объекты **формы** для различных режимах цвет публикации. Для каждой публикации можно использовать только один режим цвета. Например нельзя задать режим цвета место в процедуре и затем укажите режим цвета. Используйте метод **[CreatePlateCollection](http://msdn.microsoft.com/library/339c2c90-d1b7-808e-2b3c-c52c000e4908%28Office.15%29.aspx)** объекта **[Document](document-object-publisher.md)** для указания способ цвет для использования в семействе сайтов публикации формы. Используйте метод **[Add](plates-add-method-publisher.md)** коллекции **формы** для добавления новой формы в коллекцию **форм** . В этом примере создается коллекция цвет область формы и добавляет форму.
 

 

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

Используйте метод **[EnterColorMode](http://msdn.microsoft.com/library/3c04275d-d274-f681-7391-139a54232a3b%28Office.15%29.aspx)** объекта **[Document](document-object-publisher.md)** , чтобы указать режим цвета и коллекции **формы** для использования с цветовой режим. Используйте свойство **[ColorMode](http://msdn.microsoft.com/library/58befa97-9d9b-9294-18b2-ae10dc87f51c%28Office.15%29.aspx)** , чтобы определить, какой режим цвет используется в публикации. В этом примере создается коллекция цвет область формы, добавляет две формы и затем вводит эти формы в режиме кусочков цвет.
 

 



```
Sub CreateSpotColorMode() 
 Dim plArray As Plates 
 
 With ThisDocument 
 'Creates a color plate collection, 
 'which contains one black plate by default 
 Set plArray = .CreatePlateCollection(Mode:=pbColorModeSpot) 
 
 'Sets the plate color to red 
 plArray(1).Color.RGB = RGB(255, 0, 0) 
 
 'Adds another plate, black by default and 
 'sets the plate color to green 
 plArray.Add 
 plArray(2).Color.RGB = RGB(0, 255, 0) 
 
 'Enters spot-color mode with above 
 'two plates in the plates array 
 .EnterColorMode Mode:=pbColorModeSpot, Plates:=plArray 
 End With 
End Sub
```

Используйте метод **[FindPlateByInkName](plates-findplatebyinkname-method-publisher.md)** для возврата определенного формы с учетом его рукописного ввода имени. Процесс цвета назначены разные номера индекса в коллекции **формы** , чем в коллекции **[PrintablePlates](printableplates-object-publisher.md)** . Используйте метод **FindPlateByInkName** для гарантии на желаемую **[формы](plate-object-publisher.md)** или получить доступ к объекту **[PrintablePlate](printableplate-object-publisher.md)** .
 

 

## <a name="methods"></a>Методы



|**Name**|
|:-----|
|[Добавление](plates-add-method-publisher.md)|
|[FindPlateByInkName](plates-findplatebyinkname-method-publisher.md)|

## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](plates-application-property-publisher.md)|
|[Count](plates-count-property-publisher.md)|
|[Элемент](plates-item-property-publisher.md)|
|[Родительский раздел](plates-parent-property-publisher.md)|

