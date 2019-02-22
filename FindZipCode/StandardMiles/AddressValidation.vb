Imports System.Data.SqlClient
Imports MapPoint
Imports MapPoint.GeoPaneState

Namespace StandardMiles
    ''' <summary>
    ''' Standard Miles is Proprietary code of Better Idea Group
    ''' The inclusion in this project does, in no way grant any rights to any reuse of this code
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    ''' 

    Public Structure StandardAddress
        Dim TransactionId As Double
        Dim Address As String
        Dim City As String
        Dim State As String
        Dim ZIP As String
        Dim latitude As Double
        Dim longtitude As Double
    End Structure

    Public Structure AddressError
        Public Address As StandardAddress
        Public FindResultsQuality As MapPoint.GeoFindResultsQuality
        Public ErrorDescription As String
    End Structure

    Public Class AddressValidation

        Private m_bIsDisposing As Boolean
        Private m_oApp As Application
        Private m_oMap As Map

#Region "Contructor/Destructor"
        Public Sub New()
            m_oApp = DirectCast(CreateObject("MapPoint.Application"), MapPoint.Application)
            'm_oApp = DirectCast(GetObject("MapPoint.Application"), MapPoint.Application)

        End Sub

        Public Sub Destroy()
            ' Release all objects cleanly
            Try
                If Not m_oApp Is Nothing Then
                    m_oApp.ActiveMap.Saved() = True
                    m_oApp.Quit()

                    m_oMap = Nothing
                    m_oApp = Nothing
                End If
            Catch ex As Exception
                MsgBox("Unable to close MapPoint. If you see this message consistantly, report it to applicaiton support.")
            End Try

        End Sub
        Public Property IsDisposing() As Boolean
            Get
                Return m_bIsDisposing
            End Get
            Set(ByVal Value As Boolean)
                m_bIsDisposing = Value
            End Set
        End Property
#End Region

        Public Overloads Function ValidateAddress(ByVal strAddress As String, ByVal strCity As String) As AddressError
            ' Dim ReturnResults As MapPointEngine.MapPointStructs.ResultStatus
            Dim myError As New StandardMiles.AddressError

            m_oApp.PaneState = geoPaneRoutePlanner

            m_oMap = DirectCast(GetObject(, "MapPoint.Application"), MapPoint.Application).ActiveMap

            'strAddress = "9700 village center dr"
            'strCity = "granite bay"

            Dim oFindResults As FindResults
            Dim objLocation As MapPoint.Location

            ' Get the possible address results from MapPoint
            Try
                oFindResults = m_oMap.FindAddressResults(strAddress, strCity, , "CA")

            Catch ex As Exception
                ' Address is not valid
                myError.FindResultsQuality = GeoFindResultsQuality.geoNoResults
                myError.ErrorDescription = "Program Exception: Address is not valid! Exception message is: " & ex.Message
                oFindResults = Nothing
                Return myError
                Exit Function
            End Try

            ' Did MapPoint find a valid address for the particular row? - Check first result is good.
            If oFindResults.ResultsQuality = GeoFindResultsQuality.geoFirstResultGood Or _
                oFindResults.ResultsQuality = GeoFindResultsQuality.geoAllResultsValid Then
                ' Address found! 
                objLocation = oFindResults(1)

                myError.Address.Address = objLocation.StreetAddress.Street
                myError.Address.City = objLocation.StreetAddress.City
                myError.Address.ZIP = objLocation.StreetAddress.PostalCode
                myError.Address.State = objLocation.StreetAddress.Region

                myError.FindResultsQuality = oFindResults.ResultsQuality
                myError.ErrorDescription = ""

                oFindResults = Nothing
                Return myError
                Exit Function
            ElseIf oFindResults.ResultsQuality = GeoFindResultsQuality.geoAmbiguousResults Then

                Dim cnt As Integer = 1

                myError.ErrorDescription = "Multiple Addresses found! First result returned"
                objLocation = oFindResults(1)

                'While cnt < oFindResults.Count
                '    Debug.Print(oFindResults.Item(cnt).name)
                '     return JUST the first result found
                '    objLocation = oFindResults(1)

                '    myError.Address.ZIP = objLocation.StreetAddress.PostalCode

                '    cnt = cnt + 1
                'End While
            Else ' no good results found, total failure

            End If

            myError.FindResultsQuality = oFindResults.ResultsQuality
            myError.ErrorDescription = "Multiple Addresses found! First result returned"
            oFindResults = Nothing
            Return myError
            Exit Function

        End Function

        Public Overloads Function ValidateAddress(ByVal latitude As Double, ByVal longtitude As Double) As AddressError

            Dim myError As New AddressError

            m_oApp.PaneState = geoPaneRoutePlanner

            m_oMap = DirectCast(GetObject(, "MapPoint.Application"), MapPoint.Application).ActiveMap

            Dim oLocationLookupResults As MapPoint.Location

            ' Get the possible address results from MapPoint
            Try
                oLocationLookupResults = m_oMap.GetLocation(latitude, longtitude)
            Catch ex As Exception
                ' Address is not valid
                myError.FindResultsQuality = GeoFindResultsQuality.geoNoGoodResult
                myError.ErrorDescription = "Lat/Long of : " & latitude.ToString & "/" & longtitude.ToString & " is not valid!"
                oLocationLookupResults = Nothing
                Return myError
            End Try

            ' Did MapPoint find a valid address for the particular row? - Check first result is good.
            ' Address found!
            myError.FindResultsQuality = GeoFindResultsQuality.geoFirstResultGood
            myError.ErrorDescription = ""
            oLocationLookupResults = Nothing
            Return myError
        End Function
    End Class

End Namespace
