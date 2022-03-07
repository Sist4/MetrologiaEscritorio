Imports iTextSharp.text.pdf
Imports iTextSharp.text
Imports System.IO

Public Class CodigoQr

    Public Sub IMG_Qr()
        Dim ruta As String = "C:\archivos_metrologia\Informes\2020\10 - Octubre\ICC-201029 IDEAL ALAMBREC S.A\ICC-201029.pdf" 'ruta del pdf (se creara en la carpeta Debug del proyecto
        Dim textoArial As String = "Ejemplo de personalizacion de fuente Arial"
        Dim textoCalibri As String = "Ejemplo de personalizacion de fuente Calibri"
        Dim textoVerdana As String = "Ejemplo de personalizacion de fuente Verdana"
        Dim documento As New iTextSharp.text.Document(PageSize.LETTER, 72, 72, 72, 72)
        Dim pdfw As iTextSharp.text.pdf.PdfWriter

        'craeación de fuentes arial, calibri y verdana
        Dim fuenteArial As New Font(BaseFont.CreateFont("c:/windows/fonts/arial.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED), 12)
        Dim fuenteCalibri As New Font(BaseFont.CreateFont("c:/windows/fonts/calibri.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED), 12)
        Dim fuenteVerdana As New Font(BaseFont.CreateFont("c:/windows/fonts/verdana.ttf", BaseFont.WINANSI, BaseFont.EMBEDDED), 12)

        pdfw = PdfWriter.GetInstance(documento, New FileStream(ruta, FileMode.Create, FileAccess.Write, FileShare.None))

        'Apertura del documento.
        documento.Open()

        'Agregamos una pagina.
        documento.NewPage()

        'posicionar y redimensionarfranja azul
        Dim imagen As iTextSharp.text.Image 'declaración de imagen
        imagen = iTextSharp.text.Image.GetInstance("C:\archivos_metrologia\Imagenes_QR\201029A.jpg") 'nombre y ruta de la imagen a insertar
        imagen.ScalePercent(16.7) 'escala al tamaño de la imagen
        imagen.SetAbsolutePosition(40, 500) 'posición en la que se inserta. 40 (de izquierda a derecha). 500 (de abajo hacia arriba)

        documento.Add(imagen) 'se agrega la imagen al documento

        imagen.SetAbsolutePosition(140, 500) 'posición en la que se inserta. 140 (de izquierda a derecha). 500 (de abajo hacia arriba)
        imagen.RotationDegrees = 30 ' rotación de la imagen

        documento.Add(imagen) 'se agrega la imagen al documento

        'Forzamos vaciamiento del buffer.
        pdfw.Flush()
        'Cerramos el documento.
        documento.Close()

        pdfw = Nothing
        documento = Nothing

    End Sub

End Class
