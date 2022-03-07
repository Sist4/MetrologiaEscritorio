Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms
Imports System.Security.Cryptography.X509Certificates
Imports Gma.QrCodeNet.Encoding.Windows.Render
Imports Gma.QrCodeNet.Encoding
Imports System.Security.Cryptography
Imports System.Security.Cryptography.Pkcs
Imports Org.BouncyCastle.X509
Imports X509Certificate = System.Security.Cryptography.X509Certificates.X509Certificate
Imports System.IO
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Office
Imports Microsoft.Office.Interop.Word
Imports Application = Microsoft.Office.Interop.Word.Application
Public Class Generar_QR
    'Public Sub reemplazarImagen(ByVal INS_IMAGEN As String, ByVal textoAbuscar As String, ByVal imagenReemplazo As String)

    '    Dim ap As New Microsoft.Office.Interop.Word.Application()
    '    Dim documento As DocumentFormat.OpenXml.Wordprocessing.Document = ap.Documents.Open(INS_IMAGEN)
    '    '  reemplazarImagen(documento, "Codigo_qr", "C:\archivos_metrologia\Imagenes_QR\" & IdeComBpr & ".jpg")

    '    Dim documentoActivo As Microsoft.Office.Interop.Word.Document = (documento)
    '    Dim nullobj As Object = Type.Missing
    '    Dim start As Object = 0
    '    Dim [end] As Object = documentoActivo.Characters.Count
    '    Dim rng As Microsoft.Office.Interop.Word.Range = documentoActivo.Range(start, [end])
    '    Dim fnd As Microsoft.Office.Interop.Word.Find = rng.Find
    '    fnd.ClearFormatting()
    '    fnd.Text = textoAbuscar
    '    fnd.Forward = True
    '    Dim linktoFile As Object = False
    '    Dim SaveWithDoc As Object = True
    '    Dim replaceOption As Object = Microsoft.Office.Interop.Word.WdReplace.wdReplaceOne
    '    Dim range As Object = Type.Missing

    '    ' ubicamos la posicion donde queremos insertar la imagen
    '    fnd.Execute(nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, replaceOption, nullobj, nullobj, nullobj, nullobj)

    '    'Insertamos la imagen en la posicion adecuada
    '    rng.InlineShapes.AddPicture(imagenReemplazo, linktoFile, SaveWithDoc, range)
    'End Sub


    Public Sub reemplazarImagen(ByVal documentoActivo As Microsoft.Office.Interop.Word.Document, ByVal textoAbuscar As String, ByVal imagenReemplazo As String)
        Dim nullobj As Object = Type.Missing
        Dim start As Object = 0
        Dim [end] As Object = documentoActivo.Characters.Count
        Dim rng As Microsoft.Office.Interop.Word.Range = documentoActivo.Range(start, [end])
        Dim fnd As Microsoft.Office.Interop.Word.Find = rng.Find
        fnd.ClearFormatting()
        fnd.Text = textoAbuscar
        fnd.Forward = True
        Dim linktoFile As Object = False
        Dim SaveWithDoc As Object = True
        Dim replaceOption As Object = WdReplace.wdReplaceOne
        Dim range As Object = Type.Missing

        ' ubicamos la posicion donde queremos insertar la imagen
        fnd.Execute(nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, nullobj, replaceOption, nullobj, nullobj, nullobj, nullobj)

        'Insertamos la imagen en la posicion adecuada
        rng.InlineShapes.AddPicture(imagenReemplazo, linktoFile, SaveWithDoc, range)
    End Sub
End Class
