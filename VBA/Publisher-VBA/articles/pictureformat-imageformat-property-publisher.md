---
title: "Свойство PictureFormat.ImageFormat (издатель)"
keywords: vbapb10.chm3604761
f1_keywords: vbapb10.chm3604761
ms.prod: publisher
api_name: Publisher.PictureFormat.ImageFormat
ms.assetid: a5523a1e-4dbf-5cd7-ba73-2a5570865ee6
ms.date: 06/08/2017
ms.openlocfilehash: 6218778fb6f7338583be3c4378a70230cc5a04dc
ms.sourcegitcommit: 1102fd44df64f18dc0561d0b3a7103ca81e74318
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/26/2017
---
# <a name="pictureformatimageformat-property-publisher"></a>Свойство PictureFormat.ImageFormat (издатель)

Возвращает константу **PbImageFormat** , представляющий формат изображения, определяемой на интерфейс для графических устройств Microsoft Windows (GDI +). Только для чтения.


## <a name="syntax"></a>Синтаксис

 _выражение_. **ImageFormat**

 переменная _expression_A, представляющий объект **PictureFormat** .


### <a name="return-value"></a>Возвращаемое значение

PbImageFormat


## <a name="remarks"></a>Заметки

Свойство **ImageFormat** применяется к исходного изображения, а не рисунок, если он существует.

Значение свойства **ImageFormat** может иметь одно из **[PbImageFormat](pbimageformat-enumeration-publisher.md)** константы в библиотеке типов, Microsoft Publisher.

Свойство **ImageFormat** указывает формат рисунков после импорта в среде Windows, а не его исходный формат файла. Если формат файла изображения изначально не поддерживается операционной системой Windows, изображение преобразуется в аналогичные формат, который поддерживается изначально. В результате константы **pbImageFormatCMYKJPEG**, **pbImageFormatDIB**, **pbImageFormatEMF**, **pbImageFormatGIF**и **pbImageFormatPICT** редко, если вообще возвращаются. Обратитесь в приведенной ниже таблице преобразования формата файла.



|**Формат файла**|**Константы возвращаются**|
|:-----|:-----|
|.bmp, .dib, .gif, .pict|pbImageFormatPNG|
|.EMF, получение, .epfs|pbImageFormatWMF|
|Формат CMYK .jfif, .jpeg, .jpg|pbImageFormatJPEG|
Windows GDI + — это часть операционной системы Microsoft Windows XP и операционной системы Microsoft Windows Server 2003, который предоставляет двухмерных векторной графики, рисунков и оформление.


## <a name="example"></a>Пример

В следующем примере выводится список JPG и JPEG изображения, используемые в активной публикации.


```vb
Dim pgLoop As Page 
Dim shpLoop As Shape 
 
For Each pgLoop In ActiveDocument.Pages 
 For Each shpLoop In pgLoop.Shapes 
 
 If shpLoop.Type = pbPicture Or shpLoop.Type = pbLinkedPicture Then 
 
 With shpLoop.PictureFormat 
 If .IsEmpty = msoFalse Then 
 
 If .ImageFormat = pbImageFormatJPEG Then 
 Debug.Print .Filename 
 End If 
 
 End If 
 End With 
 
 End If 
 
 Next shpLoop 
Next pgLoop 

```


