Imports System
Imports System.Collections
Imports System.IO
Imports System.Security.Cryptography
Imports System.Security.Cryptography.Pkcs
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Org.BouncyCastle.X509
Imports SysX509 = System.Security.Cryptography.X509Certificates



Public Class PDF


    Public Shared Sub SignHashed(ByVal Source As String, ByVal Target As String, ByVal Certificate As SysX509.X509Certificate2, ByVal Reason As String, ByVal Location As String, ByVal AddVisibleSign As Boolean)
        '*******************************************
        'creamos la imagen QR
        'Dim DF() As String = objCert.SubjectName.Name.Split(",")

        'QrCodeImgControl1.Text = "FIRMADO POR: " & DF(0).Replace("CN=", "").ToString() & "" & vbCrLf & "RAZON:" & vbCrLf & "LOCALIZACION:" & vbCrLf & "FECHA: 2020-07-08" & vbCrLf & "VALIDAR CON:" & vbCrLf & "www.firmadigital.gob.ec 2.7.0"
        'Dim Img As Image = DirectCast(QrCodeImgControl1.Image.Clone, Image)
        'Img.Save("C:\temp\1.jpg")



        '********************************************

        Dim objCP As X509CertificateParser = New X509CertificateParser
        Dim objChain() As X509Certificate = New X509Certificate() {objCP.ReadCertificate(Certificate.RawData)}
        Dim objReader As PdfReader = New PdfReader(Source)
        Dim objStamper As PdfStamper = PdfStamper.CreateSignature(objReader, New FileStream(Target, FileMode.Create), Microsoft.VisualBasic.ChrW(92))
        Dim objSA As PdfSignatureAppearance = objStamper.SignatureAppearance
        If AddVisibleSign Then
            objSA.SignDate = DateTime.Now
        End If

        objSA.SetCrypto(Nothing, objChain, Nothing, Nothing)
        objSA.Reason = Reason
        objSA.Location = Location
        objSA.Acro6Layers = True
        objSA.Render = PdfSignatureAppearance.SignatureRender.NameAndDescription
        Dim objSignature As PdfSignature = New PdfSignature(PdfName.ADOBE_PPKMS, PdfName.ADBE_PKCS7_SHA1)
        objSignature.Date = New PdfDate(objSA.SignDate)
        objSignature.Name = PdfPKCS7.GetSubjectFields(objChain(0)).GetField("CN")
        If (Not (objSA.Reason) Is Nothing) Then
            objSignature.Reason = objSA.Reason
        End If

        If (Not (objSA.Location) Is Nothing) Then
            objSignature.Location = objSA.Location
        End If

        objSA.CryptoDictionary = objSignature
        Dim intCSize As Integer = 4000
        Dim objTable As Hashtable = New Hashtable
        objTable(PdfName.CONTENTS) = ((intCSize * 2) _
                    + 2)
        objSA.PreClose(objTable)
        Dim objSHA1 As HashAlgorithm = New SHA1CryptoServiceProvider
        Dim objStream As Stream = objSA.RangeStream
        Dim intRead As Integer = 0
        Dim bytBuffer() As Byte = New Byte((8192) - 1) {}


        While __Assign(intRead, objStream.Read(bytBuffer, 0, 8192)) > 0
            objSHA1.TransformBlock(bytBuffer, 0, intRead, bytBuffer, 0)
        End While

        objSHA1.TransformFinalBlock(bytBuffer, 0, 0)
        Dim bytPK() As Byte = SignMsg(objSHA1.Hash, Certificate, False)
        Dim bytOut() As Byte = New Byte((intCSize) - 1) {}
        Dim objDict As PdfDictionary = New PdfDictionary
        Array.Copy(bytPK, 0, bytOut, 0, bytPK.Length)
        objDict.Put(PdfName.CONTENTS, (New PdfString(bytOut).SetHexWriting(True)))
        ' objDict.Put(PdfName.CONTENTS, New PdfString(bytOut).SetHexWriting(True))
        objSA.Close(objDict)
    End Sub

    Private Shared Function SignMsg(ByVal Message As Byte(), ByVal SignerCertificate As SysX509.X509Certificate2, ByVal Detached As Boolean) As Byte()
        'Creamos el contenedor
        Dim contentInfo As New ContentInfo(Message)
        'Instanciamos el objeto SignedCms con el contenedor
        Dim objSignedCms As New SignedCms(contentInfo, Detached)
        'Creamos el "firmante"
        Dim objCmsSigner As New CmsSigner(SignerCertificate)
        ' Include the following line if the top certificate in the
        ' smartcard is not in the trusted list.
        objCmsSigner.IncludeOption = SysX509.X509IncludeOption.EndCertOnly
        '  Sign the CMS/PKCS #7 message. The second argument is
        '  needed to ask for the pin.
        objSignedCms.ComputeSignature(objCmsSigner, False)
        'Encodeamos el mensaje CMS/PKCS #7
        Return objSignedCms.Encode()
    End Function

    Shared Function __Assign(Of T)(ByRef target As T, value As T) As T
        target = value
        Return value
    End Function



End Class
