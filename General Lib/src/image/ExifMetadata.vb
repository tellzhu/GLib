Imports System.Drawing.Imaging
Imports System.Text
Imports dotNet.image.ExifMetadata.ExifTag

Namespace image
    Public Class ExifMetadata
        Friend Structure MetadataDetail
            Friend RawValueAsString As String
            Public DisplayValue As String
        End Structure

        Friend Structure Metadata
            Public EquipmentMake As MetadataDetail
            Public CameraModel As MetadataDetail
            'Public ExposureTime As MetadataDetail
            'Public Fstop As MetadataDetail
            'Public DatePictureTaken As MetadataDetail
            'Public ShutterSpeed As MetadataDetail
            'Public MeteringMode As MetadataDetail
            'Public Flash As MetadataDetail
            'Public XResolution As MetadataDetail
            'Public YResolution As MetadataDetail
            'Public ImageWidth As MetadataDetail
            'Public ImageHeight As MetadataDetail
            'Public FNumber As MetadataDetail
            'Public ExposureProg As MetadataDetail
            'Public SpectralSense As MetadataDetail
            'Public ISOSpeed As MetadataDetail
            'Public OECF As MetadataDetail
            'Public Ver As MetadataDetail
            'Public CompConfig As MetadataDetail
            'Public CompBPP As MetadataDetail
            'Public Aperture As MetadataDetail
            'Public Brightness As MetadataDetail
            'Public ExposureBias As MetadataDetail
            'Public MaxAperture As MetadataDetail
            'Public SubjectDist As MetadataDetail
            'Public LightSource As MetadataDetail
            'Public FocalLength As MetadataDetail
            'Public FPXVer As MetadataDetail
            'Public ColorSpace As MetadataDetail
            'Public Interop As MetadataDetail
            'Public FlashEnergy As MetadataDetail
            'Public SpatialFR As MetadataDetail
            'Public FocalXRes As MetadataDetail
            'Public FocalYRes As MetadataDetail
            'Public FocalResUnit As MetadataDetail
            'Public ExposureIndex As MetadataDetail
            'Public SensingMethod As MetadataDetail
            'Public SceneType As MetadataDetail
            'Public CfaPattern As MetadataDetail
            Public GPSLatitudeRef As MetadataDetail
            Public GPSLatitude As MetadataDetail
            Public GPSLongitudeRef As MetadataDetail
            Public GPSLongitude As MetadataDetail
            Public GPSAltitudeRef As MetadataDetail
            Public GPSAltitude As MetadataDetail
        End Structure

        Private Function LookupEXIFValue(ByVal Description As String, ByVal Value As String) As String
            Dim DescriptionValue As String = Nothing
            Select Case Description
                Case "MeteringMode"
                    Select Case Value
                        Case "0"
                            DescriptionValue = "Unknown"
                        Case "1"
                            DescriptionValue = "Average"
                        Case "2"
                            DescriptionValue = "Center Weighted Average"
                        Case "3"
                            DescriptionValue = "Spot"
                        Case "4"
                            DescriptionValue = "Multi-spot"
                        Case "5"
                            DescriptionValue = "Multi-segment"
                        Case "6"
                            DescriptionValue = "Partial"
                        Case "255"
                            DescriptionValue = "Other"
                    End Select
                Case "ResolutionUnit"
                    Select Case Value
                        Case "1"
                            DescriptionValue = "No Units"
                        Case "2"
                            DescriptionValue = "Inch"
                        Case "3"
                            DescriptionValue = "Centimeter"
                    End Select
            End Select
            Return DescriptionValue
        End Function

        Public Sub UpdateGPSLocation(ByRef gpsLoc As GPSLocation, ByVal NewFileName As String)
            Dim img As Drawing.Image = Drawing.Image.FromFile(m_PhotoName)
            gpsLoc.UpdateImage(img)
            img.Save(NewFileName)
            img.Dispose()
        End Sub

        Public ReadOnly Property GPSMetadata As GPSLocation
            Get
                If ContainsGPS() Then
                    Dim gpsLoc As GPSLocation = New GPSLocation
                    gpsLoc.Add(m_dict.Item(GPSLatitudeRef))
                    gpsLoc.Add(m_dict.Item(GPSLatitude))
                    gpsLoc.Add(m_dict.Item(GPSLongitudeRef))
                    gpsLoc.Add(m_dict.Item(GPSLongitude))
                    gpsLoc.Add(m_dict.Item(GPSAltitudeRef))
                    gpsLoc.Add(m_dict.Item(GPSAltitude))
                    Return gpsLoc
                Else
                    Return Nothing
                End If
            End Get
        End Property

        ''' <summary>
        ''' 定义EXIF规范中常用的标识代码。
        ''' </summary>
        ''' <remarks></remarks>
        Public Enum ExifTag As Integer
            ''' <summary>
            ''' 纬度位置。
            ''' </summary>
            ''' <remarks>EXIF规范标识为1，用于区分北纬或南纬。</remarks>
            GPSLatitudeRef = 1
            ''' <summary>
            ''' 纬度。
            ''' </summary>
            ''' <remarks>EXIF规范标识为2，用于记录纬度的数值。</remarks>
            GPSLatitude = 2
            ''' <summary>
            ''' 经度位置。
            ''' </summary>
            ''' <remarks>EXIF规范标识为3，用于区分东经或西经。</remarks>
            GPSLongitudeRef = 3
            ''' <summary>
            ''' 经度。
            ''' </summary>
            ''' <remarks>EXIF规范标识为4，用于记录经度的数值。</remarks>
            GPSLongitude = 4
            ''' <summary>
            ''' 海平面位置。
            ''' </summary>
            ''' <remarks>EXIF规范标识为5，用于区分海平面以上或海平面以下。</remarks>
            GPSAltitudeRef = 5
            ''' <summary>
            ''' 海平面高度。
            ''' </summary>
            ''' <remarks>EXIF规范标识为6，用于记录海平面高度的数值。</remarks>
            GPSAltitude = 6
            ''' <summary>
            ''' 设备制造商。
            ''' </summary>
            ''' <remarks>EXIF规范标识为10F，用于记录原始设备的制造厂商。</remarks>
            EquipmentMake = 271
            ''' <summary>
            ''' 设备型号。
            ''' </summary>
            ''' <remarks>EXIF规范标识为110，用于记录原始设备的具体型号。</remarks>
            CameraModel = 272
        End Enum

        Private Function GetEXIFMetaData(ByVal PhotoName As String) As Metadata
            Dim srcImg As Drawing.Image = Drawing.Image.FromFile(PhotoName)
            Dim propIdsList As Integer() = srcImg.PropertyIdList
            If propIdsList.Length = 0 Then
                If propIdsList IsNot Nothing Then
                    Array.Clear(propIdsList, 0, propIdsList.Length)
                    propIdsList = Nothing
                End If
                srcImg.Dispose()
                srcImg = Nothing
                Return Nothing
            End If
            Dim md As Metadata = New Metadata()
            'md.DatePictureTaken.Hex = "9003"
            'md.ExposureTime.Hex = "829a"
            'md.Fstop.Hex = "829d"
            'md.ShutterSpeed.Hex = "9201"
            'md.MeteringMode.Hex = "9207"
            'md.Flash.Hex = "9209"
            'md.FNumber.Hex = "829d"
            'md.ExposureProg.Hex = ""
            'md.SpectralSense.Hex = "8824"
            'md.ISOSpeed.Hex = "8827"
            'md.OECF.Hex = "8828"
            'md.Ver.Hex = "9000"
            'md.CompConfig.Hex = "9101"
            'md.CompBPP.Hex = "9102"
            'md.Aperture.Hex = "9202"
            'md.Brightness.Hex = "9203"
            'md.ExposureBias.Hex = "9204"
            'md.MaxAperture.Hex = "9205"
            'md.SubjectDist.Hex = "9206"
            'md.LightSource.Hex = "9208"
            'md.FocalLength.Hex = "920a"
            'md.FPXVer.Hex = "a000"
            'md.ColorSpace.Hex = "a001"
            'md.FocalXRes.Hex = "a20e"
            'md.FocalYRes.Hex = "a20f"
            'md.FocalResUnit.Hex = "a210"
            'md.ExposureIndex.Hex = "a215"
            'md.SensingMethod.Hex = "a217"
            'md.SceneType.Hex = "a301"
            'md.CfaPattern.Hex = "a302"
            Dim encoding As ASCIIEncoding = New ASCIIEncoding()
            Dim prop As PropertyItem = Nothing
            Dim b() As Byte = Nothing
            For Each propId As Integer In propIdsList
                prop = srcImg.GetPropertyItem(propId)
                b = prop.Value
                If Not m_dict.ContainsKey(propId) Then
                    m_dict.Add(propId, prop)
                End If
                Select Case propId
                    Case EquipmentMake
                        With md.EquipmentMake
                            .RawValueAsString = BitConverter.ToString(b)
                            .DisplayValue = encoding.GetString(b)
                        End With
                    Case CameraModel
                        With md.CameraModel
                            .RawValueAsString = BitConverter.ToString(b)
                            .DisplayValue = encoding.GetString(b)
                        End With
                    Case GPSLatitudeRef
                        With md.GPSLatitudeRef
                            .RawValueAsString = BitConverter.ToString(b)
                            .DisplayValue = encoding.GetString(b)
                        End With
                    Case GPSLatitude
                        With md.GPSLatitude
                            .RawValueAsString = BitConverter.ToString(b)
                            .DisplayValue = GetGPSDegree(b)
                        End With
                    Case GPSLongitudeRef
                        With md.GPSLongitudeRef
                            .RawValueAsString = BitConverter.ToString(b)
                            .DisplayValue = encoding.GetString(b)
                        End With
                    Case GPSLongitude
                        With md.GPSLongitude
                            .RawValueAsString = BitConverter.ToString(b)
                            .DisplayValue = GetGPSDegree(b)
                        End With
                    Case GPSLatitudeRef
                        With md.GPSAltitudeRef
                            .RawValueAsString = BitConverter.ToString(b)
                            .DisplayValue = GetGPSAltitudeRef(b)
                        End With
                    Case GPSLatitude
                        With md.GPSAltitude
                            .RawValueAsString = BitConverter.ToString(b)
                            .DisplayValue = GetGPSAltitude(b)
                        End With
                End Select
            Next
            prop = Nothing
            If b IsNot Nothing Then
                Array.Clear(b, 0, b.Length)
                b = Nothing
            End If
            encoding = Nothing
            If propIdsList IsNot Nothing Then
                Array.Clear(propIdsList, 0, propIdsList.Length)
                propIdsList = Nothing
            End If
            srcImg.Dispose()
            srcImg = Nothing
            Return md
        End Function

        Private Function GetGPSAltitude(ByRef b() As Byte) As String
            Dim alt As Double = CDbl(BitConverter.ToInt32(b, 0) / BitConverter.ToInt32(b, 4))
            Return CStr(alt)
        End Function

        Private Function GetGPSDegree(ByRef b() As Byte) As String
            Dim deg As Double = CInt(BitConverter.ToInt32(b, 0) / BitConverter.ToInt32(b, 4))
            Dim min As Double = CInt(BitConverter.ToInt32(b, 8) / BitConverter.ToInt32(b, 12))
            Dim sec As Double = CDbl(BitConverter.ToInt32(b, 16) / BitConverter.ToInt32(b, 20))
            Return CStr(deg & ";" & min & ";" & sec)
        End Function

        Private Function GetGPSAltitudeRef(ByRef b() As Byte) As String
            If b(0) = 0 Then
                Return "Above Sea Level"
            Else
                Return "Below Sea Level"
            End If
        End Function

        Private m_dict As Dictionary(Of Integer, PropertyItem) = Nothing
        Private m_PhotoName As String = Nothing

        Public Function ContainsGPS() As Boolean
            Return m_dict.ContainsKey(GPSLatitudeRef) And m_dict.ContainsKey(GPSLatitude) _
                And m_dict.ContainsKey(GPSLongitudeRef) And m_dict.ContainsKey(GPSLongitude) _
                And m_dict.ContainsKey(GPSAltitudeRef) And m_dict.ContainsKey(GPSAltitude)
        End Function

        Public Function NotContainsGPS() As Boolean
            Return Not m_dict.ContainsKey(GPSLatitudeRef) And Not m_dict.ContainsKey(GPSLatitude) _
                And Not m_dict.ContainsKey(GPSLongitudeRef) And Not m_dict.ContainsKey(GPSLongitude) _
                And Not m_dict.ContainsKey(GPSAltitudeRef) And Not m_dict.ContainsKey(GPSAltitude)
        End Function

        Public Sub New(ByVal PhotoName As String)
            m_PhotoName = PhotoName
            m_dict = New Dictionary(Of Integer, PropertyItem)
            GetEXIFMetaData(m_PhotoName)
        End Sub

        Protected Overrides Sub Finalize()
            m_dict.Clear()
            m_dict = Nothing
            m_PhotoName = Nothing
            MyBase.Finalize()
        End Sub
    End Class

End Namespace

