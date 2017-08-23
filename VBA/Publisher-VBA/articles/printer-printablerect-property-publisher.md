---
title: "Свойство Printer.PrintableRect (издатель)"
keywords: vbapb10.chm8978450
f1_keywords: vbapb10.chm8978450
ms.prod: publisher
api_name: Publisher.Printer.PrintableRect
ms.assetid: 9d5b8264-9213-3d89-0613-421a4872c158
ms.date: 06/08/2017
ms.openlocfilehash: 9cb8e29d238f1f6996acc981f3677a91a52a8896
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="printerprintablerect-property-publisher"></a>Свойство Printer.PrintableRect (издатель)

Возвращает объект **[PrintableRect](printablerect-object-publisher.md)** , представляющий принтера области листа, в течение которого указанного печать. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **PrintableRect**

 переменная _expression_A, представляющий объект **Printer** .


### <a name="return-value"></a>Возвращаемое значение

PrintableRect


## <a name="remarks"></a>Заметки

Область печати определяется принтера на основе указанного размера листа. Не следует путать с область внутри поля страницы публикации подготовленных к печати прямоугольника листа принтера. Область печати может быть больше или меньше, чем страницы публикации.


 **Примечание**  При идентичны sheet принтера и размер страницы публикации, страница публикации располагается на листе принтера и ни один из метки Печать "," даже в том случае, если они выбраны.


## <a name="example"></a>Пример

Следующие Microsoft Visual Basic для приложений (VBA) макроса показано, как использовать свойство **PrintableRect** для получения границах подготовленных к печати прямоугольника для листа принтера активного принтера.


```vb
Public Sub PrintableRect_Example() 
 
 Dim pubInstalledPrinters As Publisher.InstalledPrinters 
 Dim pubApplication As Publisher.Application 
 Dim pubPrinter As Publisher.Printer 
 
 Set pubApplication = ThisDocument.Application 
 Set pubInstalledPrinters = pubApplication.InstalledPrinters 
 
 For Each pubPrinter In pubInstalledPrinters 
 If pubPrinter.IsActivePrinter Then 
 With pubPrinter.PrintableRect 
 Debug.Print "Printable area is " &; PointsToInches(.Width) &; " by " &; PointsToInches(.Height) &; " inches." 
 Debug.Print "Left Boundary: " &; PointsToInches(.Left) &; " inches (from left)." 
 Debug.Print "Right Boundary: " &; PointsToInches(.Left + .Width) &; " inches (from left)." 
 Debug.Print "Top Boundary: " &; PointsToInches(.Top) &; " inches(from top)." 
 Debug.Print "Bottom Boundary: " &; PointsToInches(.Top + .Height) &; " inches (from top)." 
 End With 
 End If 
 Next 
 
End Sub
```


