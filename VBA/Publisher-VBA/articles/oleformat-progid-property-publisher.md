---
title: "Свойство OLEFormat.ProgId (издатель)"
keywords: vbapb10.chm4456452
f1_keywords: vbapb10.chm4456452
ms.prod: publisher
api_name: Publisher.OLEFormat.ProgId
ms.assetid: dae7e591-65d2-b956-e598-8746955c4182
ms.date: 06/08/2017
ms.openlocfilehash: 89dd9528d021203812bc6634ad3b80f5e4fb9176
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="oleformatprogid-property-publisher"></a>Свойство OLEFormat.ProgId (издатель)

Возвращает **строку** , представляющую программный идентификатор (ProgID) для указанного объекта OLE. Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ProgId**

 переменная _expression_A, представляющий объект **OLEFormat** .


### <a name="return-value"></a>Возвращаемое значение

String


## <a name="example"></a>Пример

В этом примере циклически просматривает все связанные OLE объект фигуры на первой странице активного документа и обновляет все связанные листы Excel. В этом примере предполагается, что имеется по крайней мере один фигуры на первой странице active публикации.


```vb
Sub UpdateLinkedOLEObject() 
 Dim shp As Shape 
 For Each shp In ActiveDocument.Pages(1).Shapes 
 If shp.Type = msoLinkedOLEObject Then 
 If shp.OLEFormat.ProgId = "Excel.Sheet" Then 
 shp.LinkFormat.Update 
 End If 
 End If 
 Next 
End Sub
```


