---
title: "Объект WebCheckBox (издатель)"
keywords: vbapb10.chm4390911
f1_keywords: vbapb10.chm4390911
ms.prod: publisher
api_name: Publisher.WebCheckBox
ms.assetid: adcdf233-50b8-acbe-e52f-1e86e175b31d
ms.date: 06/08/2017
ms.openlocfilehash: 8b85ebbf5520c06823d16c7b623e73b52b170202
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="webcheckbox-object-publisher"></a>Объект WebCheckBox (издатель)

Представляет веб-элемент управления "флажок". Объект **WebCheckBox** является членом объекта **Shape** .
 


## <a name="example"></a>Пример

Используйте метод **[AddWebControl](shapes-addwebcontrol-method-publisher.md)** для создания флажок Web. Используйте свойство **[WebCheckBox](shape-webcheckbox-property-publisher.md)** для доступа к фигуры Web флажок элемента управления. В этом примере создается новый Web флажок и указывает, что проверяется состояние по умолчанию; затем добавляется текстовое поле рядом с ним описать его.
 

 

```
Sub CreateNewWebCheckBox() 
 With ActiveDocument.Pages(1).Shapes 
 With .AddWebControl(Type:=pbWebControlCheckBox, Left:=100, _ 
 Top:=123, Width:=17, Height:=12).WebCheckBox 
 .Selected = msoTrue 
 End With 
 With .AddTextbox(Orientation:=pbTextOrientationHorizontal, _ 
 Left:=118, Top:=120, Width:=70, Height:=15) 
 .TextFrame.TextRange.Text = "Description text for Web check box" 
 End With 
 End With 
End Sub
```


## <a name="properties"></a>Properties



|**Name**|
|:-----|
|[Приложения](webcheckbox-application-property-publisher.md)|
|[Родительский раздел](webcheckbox-parent-property-publisher.md)|
|[ReturnDataLabel](webcheckbox-returndatalabel-property-publisher.md)|
|[Выбранные](webcheckbox-selected-property-publisher.md)|
|[Значение](webcheckbox-value-property-publisher.md)|

