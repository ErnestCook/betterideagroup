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
    Public Class MileageCalculation

        Private m_bIsDisposing As Boolean
        Private m_oApp As Application
        Private m_oMap As Map
        Dim lErrorDialog As New ErrorDialog

        Public Structure RouteData
            Dim DistanceCalculated As Double
            Dim DrivingTime As Double
        End Structure

#Region "Contructor/Destructor"
        Public Sub New()

            Try
                m_oApp = DirectCast(GetObject("MapPoint.Application"), MapPoint.Application)
            Catch ex As Exception
                m_oApp = DirectCast(CreateObject("MapPoint.Application"), MapPoint.Application)
            End Try
            'm_oApp = DirectCast(CreateObject("MapPoint.Application"), MapPoint.Application)
            'm_oApp = DirectCast(GetObject("MapPoint.Application"), MapPoint.Application)

        End Sub

        Public Sub Destroy()
            ' Release all objects cleanly
            Try
                If Not m_oApp Is Nothing Then
                    m_oApp.ActiveMap.Saved() = True
                    m_oApp.Quit()
                    'm_oRoute = Nothing
                    m_oMap = Nothing
                    m_oApp = Nothing
                End If
            Catch ex As Exception
                MsgBox("Unable to close MapPoint. If you see this message consistantly, report it to application support." & _
                       vbCrLf & ex.Message & vbCrLf & ex.StackTrace)
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

        ''' <summary>
        '''  This function looks for routeprocessed = 0 on dailyactivity rows and then passes those daily activity's
        '''   over to the procedure to calculate (or recalculate) the mileage
        ''' </summary>
        ''' <remarks> 
        ''' Because we allow for reprocessing, mearly setting the flag on the dailyactivity will reprocess the entire route
        ''' In the brave new world of addresses being validated as they are input, address validation is no longer a requirement
        '''   </remarks>
        Public Sub Process_New_Routes()
            Dim tbl As DataTable
            Dim row As DataRow
            Dim sql As String
            Dim lrowcount As Long = 1
            ' Do we have to worry about address issues?
            sql = "select * from dailyactivity where isnull(routeprocessed,0) = 0 and isnull(AddrProcessed,0) <> 0 and  SystemName <> ''"
            Dim myStatus As New frmStatus
            myStatus.MyProgressBar.Value = 1
            myStatus.lblDataType.Text = "Calculating routes for Daily Activity... "

            Try
                tbl = getDataItems(sql)

                If tbl.Rows.Count <> 0 Then         ' ////// Row count for routes to process   //////
                    myStatus.lblMaximumRecord.Text = tbl.Rows.Count.ToString '+ 1
                    myStatus.Show()
                    ' for each daily activity we find that has NOT been processed, 
                    ' go ahead and pass over the dailyactivityid to the process procedure to handle
                    frmStatus.CalculatedOKCounter.Text = 0
                    'Me.StatusLabel.Text = tbl.Rows.Count & " routes to process."
                    For Each row In tbl.Rows
                        Try
                            myStatus.MyProgressBar.Value = lrowcount
                            myStatus.CurrentRecordNumber.Text = lrowcount
                            myStatus.MyProgressBar.Refresh()
                            lrowcount += 1

                            If myStatus.CancelRequested Then
                                myStatus.Close()
                                Exit Sub
                            End If

                            'My.Application.DoEvents()
                            System.Windows.Forms.Application.DoEvents()
                            ' looping thru every dailyactivity that needs attention

                            If CalculateRouteforDailyActivity(row("DailyActivityId")) Then
                                ' shout for joy. UI component TBD
                                myStatus.CalculatedOKCounter.Text += 1
                            Else
                                ' oh no, not happy about this
                                myStatus.lstIssues.Items.Add("Daily# " & row("DailyActivityId") & " failed")
                            End If
                        Catch appex As ApplicationException
                            'log the issue and clean-up by ensuring all rows that were created are delete
                            lErrorDialog.errMessage.Text = "Unable to process route for dailyactivityid= " & row("DailyActivityId") & vbCrLf & "appex:" & appex.Message.ToString
                            lErrorDialog.ShowDialog(Me)
                            'MsgBox("Unable to process route for dailyactivityid= " & row("DailyActivityId") & vbCrLf & "appex:" & appex.Message.ToString)
                        Catch ex As Exception
                            lErrorDialog.errMessage.Text = "Critical exception occured processing route for dailyactivityid= " & row("DailyActivityId") & vbCrLf & "exception:" & ex.Message.ToString
                            lErrorDialog.ShowDialog(Me)
                        End Try
                    Next
                    'myStatus.Close()
                    myStatus.Hide()
                    myStatus.btnCancel.Visible = False
                    myStatus.btnClose.Visible = True
                    myStatus.ShowDialog()
                Else                             ' ////// Row count for routes to process   //////
                    myStatus.lblDataType.Text = "No Routes to process"
                    myStatus.btnCancel.Visible = False
                    myStatus.btnClose.Visible = True
                    '////////////
                    'myStatus.MyProgressBar.Visible = False
                    myStatus.MyProgressBar.Refresh()
                    '////////////////
                    myStatus.Show()
                    myStatus.Visible = False
                    myStatus.MyProgressBar.Maximum = 1
                    myStatus.MyProgressBar.Value = 1
                    myStatus.MyProgressBar.Refresh()
                    myStatus.ShowDialog()
                End If                            ' ////// Row count for routes to process   //////
                ' pulling in the try/catch block here since it is the top tier UI interaction
            Catch appex As ApplicationException
                ' not sure at this point if we need a seperate layer for appex vs ex
                ' so am stubing out in case...
                MsgBox("Critical exception occured processing new routes " & vbCrLf & appex.Message.ToString)
            Catch ex As Exception
                'stub
                MsgBox("Critical exception occured processing new routes " & vbCrLf & ex.Message.ToString)
            Finally
                ' cleanup tasks tbd
                ' myStatus.Close()
            End Try
        End Sub

        ''' <summary>
        '''  Cycles thru all workorders for the dailyactivity and then calculates the distance between segments
        ''' </summary>
        ''' <param name="intDailyActivityId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CalculateRouteforDailyActivity(ByVal intDailyActivityId As Integer) As Boolean
            Try ' Allowing for a public call to this routine
                ' Delete any previous route calculation from the db
                ' Call Delete_RouteData_for_DailyActivity(intDailyActivityId)
                Call Reset_MileageData_for_DailyActivity(intDailyActivityId)

                Dim BeginAddress As StandardAddress
                Dim EndingAddress As StandardAddress
                Dim WarehouseAddress As StandardAddress

                ' First, go get the starting warehouse location as this begins all routes
                ' saving it off in a seperate structure since we will also END our route at the warehouse
                WarehouseAddress = GetStartingWareHouseAddress(intDailyActivityId)
                BeginAddress = WarehouseAddress

                ' Get all the workorders associated so we can loop thru them
                Dim tbl As DataTable
                Dim row As DataRow
                Dim sql As String
                Dim bRouteError As Boolean = True
                'Dim dblSegmentDistance As Double
                Dim lrouteResults As RouteData
                Dim myStatus As New frmStatus
                Dim lErrorDialog As New ErrorDialog

                myStatus.MyProgressBar.Value = 1
                myStatus.lblDataType.Text = "Calculating routes for Daily Activity... "

                sql = "select * from  vw_WorkordersForMileage where DailyActivityId = " & intDailyActivityId & " order by StartTime"
                Dim intRowcounter As Integer = 1
                tbl = getDataItems(sql)
                If tbl.Rows.Count <> 0 Then
                    myStatus.lblMaximumRecord.Text = tbl.Rows.Count.ToString '+ 1
                    myStatus.Show()
                    ' for each daily activity we find that has NOT been processed, 
                    ' go ahead and pass over the dailyactivityid to the process procedure to handle
                    frmStatus.CalculatedOKCounter.Text = 0

                    For Each row In tbl.Rows
                        ' Looping thru every workorder  in this daily activity
                        Try
                            myStatus.MyProgressBar.Value = intRowcounter
                            myStatus.CurrentRecordNumber.Text = intRowcounter
                            myStatus.MyProgressBar.Refresh()

                            If myStatus.CancelRequested Then
                                myStatus.Close()
                                Exit Function
                            End If

                            'My.Application.DoEvents()
                            System.Windows.Forms.Application.DoEvents()
                            ' looping thru every dailyactivity that needs attention
                            '                                 ' set address fields used by MP into standard address format
                            EndingAddress = SetWOAddress(row)
                            '                                 ' Calculate the distance for this segment
                            lrouteResults = MP_CalculateDistanceAddress(BeginAddress, EndingAddress)
                            '                                 ' On success, update the database
                            UpdateDbWithRouteCalculationResult(intDailyActivityId, EndingAddress.TransactionId, intRowcounter, BeginAddress.TransactionId, lrouteResults, False)
                            '                                 ' For the next segment, the startpoint will be the end address of this segment
                            BeginAddress = EndingAddress
                        Catch appex As ApplicationException
                            ' For some Reason we failed to find an address so the update will be skipped and as a result the dailyactivityflag will never be updated
                            'Call Delete_RouteData_for_DailyActivity(intDailyActivityId)
                            Call Reset_MileageData_for_DailyActivity(intDailyActivityId)
                            myStatus.Close()
                            Throw New ApplicationException(appex.Message)
                            Exit Function
                        End Try
                        intRowcounter += 1
                    Next

                    bRouteError = False ' no errors so update flag
                Else ' no detail records to process
                    bRouteError = False
                    ' Update_RouteProcessed_Flag_on_DailyActivity(intDailyActivityId, False)
                End If
                ' update the parent row to show status
                Update_RouteProcessed_Flag_on_DailyActivity(intDailyActivityId, bRouteError)
                myStatus.Close()
                bRouteError = True ' reset so next row will not be updated on an exception
                Return True
            Catch appex As ApplicationException
                Update_RouteProcessed_Flag_on_DailyActivity(intDailyActivityId, True, appex.Message)
                Return False
            Finally
                ' do our cleanup
            End Try
        End Function

        ''' <summary>
        '''   Heart of the program. Interfaces with Map Point to calculate the distance between two points
        ''' </summary>
        ''' <param name="beginAddress"></param>
        ''' <param name="endAddress"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function MP_CalculateDistanceAddress(ByVal beginAddress As StandardAddress, ByVal endAddress As StandardAddress) As RouteData

            Dim lrouteResults As RouteData
            'Dim DistanceCalculated As Double
            Dim BeginObjLoc As MapPoint.Location
            Dim EndObjLoc As MapPoint.Location
            Dim m_oRoute As MapPoint.Route

            ' Ensure we're working in miles
            m_oApp.Units = GeoUnits.geoMiles

            ' Dim fo As MapPoint.FindResults
            Dim oFindResults As FindResults

            m_oMap = DirectCast(GetObject(, "MapPoint.Application"), MapPoint.Application).ActiveMap

            '============ Begining address  ===========================
            Dim oLocationLookupResults As MapPoint.Location
            Dim myError As New StandardMiles.AddressError
            ' Get the possible address results from MapPoint

            If beginAddress.latitude <> 0 Then ' lat long was used, find based on that
                Try
                    oLocationLookupResults = m_oMap.GetLocation(beginAddress.latitude, beginAddress.longtitude)
                    BeginObjLoc = oLocationLookupResults
                Catch ex As Exception
                    ' Address is not valid
                    Throw New ApplicationException("Lat/Long of : " & beginAddress.latitude.ToString & "/" & beginAddress.longtitude.ToString & " is not valid!")
                End Try
            Else
                oFindResults = m_oMap.FindAddressResults(beginAddress.Address, beginAddress.City, , , beginAddress.ZIP, "USA")
                If oFindResults.ResultsQuality = GeoFindResultsQuality.geoFirstResultGood Then
                    BeginObjLoc = DirectCast(oFindResults.Item(1), MapPoint.Location)
                Else
                    Throw New ApplicationException("Invalid Address during route calculation: " & beginAddress.Address & "," & beginAddress.City & " , " & beginAddress.ZIP)
                End If
            End If
            '============ Begining address  ===========================

            '============ Ending address  ===========================
            If endAddress.latitude <> 0 Then ' lat long was used, find based on that
                Try
                    oLocationLookupResults = m_oMap.GetLocation(endAddress.latitude, endAddress.longtitude)
                    EndObjLoc = oLocationLookupResults
                Catch ex As Exception       ' Address is not valid
                    Throw New ApplicationException("Lat/Long of : " & endAddress.latitude.ToString & "/" & endAddress.longtitude.ToString & " is not valid!")
                End Try
            Else
                oFindResults = m_oMap.FindAddressResults(endAddress.Address, endAddress.City, , , endAddress.ZIP, "USA")
                If oFindResults.ResultsQuality = GeoFindResultsQuality.geoFirstResultGood Then
                    EndObjLoc = DirectCast(oFindResults.Item(1), MapPoint.Location)
                Else
                    Throw New ApplicationException("Invalid Address during route calculation: " & endAddress.Address & "," & endAddress.City & " , " & endAddress.ZIP)
                End If

            End If

            ' ///////////  This is a new setting so this section deals with the fact the setting may not exist ////////
            Dim bUseRoute As Boolean = True
            Try
                If My.Settings.UseRoute Then
                    bUseRoute = True
                Else
                    bUseRoute = False
                End If
            Catch ex As Exception
                ' going to ignore this 
            End Try

                    ' ///////////  This is a new setting so this section deals with the fact the setting may not exist ////////
                    Try
                        If bUseRoute Then
                            m_oRoute = m_oMap.ActiveRoute
                            m_oRoute.Clear()
                            m_oRoute.Waypoints.Add(BeginObjLoc)
                            m_oRoute.Waypoints.Add(EndObjLoc)
                            ' Calculate the route distance
                            m_oRoute.Calculate()
                            lrouteResults.DistanceCalculated = m_oRoute.Distance
                            lrouteResults.DrivingTime = m_oRoute.DrivingTime / MapPoint.GeoTimeConstants.geoOneHour
                            'm_oRoute.Waypoints.Count
                            ' Debug.Print(MapPoint.GeoTimeConstants.geoOneDay())
                            m_oRoute.Clear()
                            m_oRoute = Nothing
                        Else
                            ' Calculate the  straight line distance
                            lrouteResults.DistanceCalculated = m_oMap.Distance(BeginObjLoc, EndObjLoc)
                            lrouteResults.DrivingTime = 0

                        End If
                    Catch ex As Exception
                        Throw New ApplicationException("Map Point failure during distance calculation! " & vbCrLf & ex.Message)
                        Exit Function
                    End Try

                    Return lrouteResults
        End Function

        Private Sub Update_RouteProcessed_Flag_on_DailyActivity(ByVal intDailyActivityID As Integer, ByVal bProcessError As Boolean, Optional ByVal strErrReturned As String = "")
            Dim myCommand As SqlCommand
            Try

                Dim myConn = New SqlConnection(My.Settings.ACsquaredConnectionString)

                'if this is being updated then we need to pass in LastUpdatedBy and Timestamp from
                'previous record - otherwise we will let the database do it
                Dim strsql As String = ""
                If bProcessError Then
                    strsql = "update DailyActivity set RouteProcessed = 0, ProcessError = '" & Left(strErrReturned, 255) & "' where dailyActivityID = " & intDailyActivityID
                Else
                    strsql = "update DailyActivity set RouteProcessed = -1, ProcessError = '" & Left(strErrReturned, 255) & "' where dailyActivityID = " & intDailyActivityID
                End If

                myCommand = New SqlCommand(strsql, myConn)
                myCommand.Connection.Open()

                'get the last id of Work order entered in case we need to insert additional
                'codes - ExecuteScalar returns that value
                Dim intresult As Integer = myCommand.ExecuteNonQuery
                If intresult <> 1 Then ' we didn't get the record updated. No reason this should or could ever happen
                    Throw New ApplicationException("For some unknown reason we failed to update the route processed flag on DailyActivityID " & intDailyActivityID.ToString)
                End If

            Catch sqlError As SqlException
                MsgBox("Sorry, we received an error from the database while trying to update it. Save this message as it might give us a clue what happened: " & sqlError.Message.ToString, MsgBoxStyle.Exclamation, My.Application.Info.ProductName.ToString)
                'Throw New ApplicationException(sqlError.Message.ToString)
            Finally
                If Not IsNothing(myCommand) Then
                    myCommand.Connection.Close()
                End If
            End Try
        End Sub

        Private Sub Reset_MileageData_for_DailyActivity(ByVal intDailyActivityID As Integer)
            Dim myCommand As SqlCommand
            Dim strsql As String = ""
            Try
                Dim myConn = New SqlConnection(My.Settings.ACsquaredConnectionString)

                strsql = "update workorders set distance = 0, processed = 0, ProcessError= '' where  DailyActivityID = " & intDailyActivityID

                myCommand = New SqlCommand(strsql, myConn)
                myCommand.Connection.Open()

                Dim intRowsUpdated As Integer = myCommand.ExecuteNonQuery
                If intRowsUpdated = 0 Then ' we didn't get the record updated. No reason this should or could ever happen
                    Throw New ApplicationException("For some unknown reason we failed to update the route processed flag on DailyActivityID " & intDailyActivityID.ToString)
                End If

            Catch sqlError As SqlException
                MsgBox("Sorry, we received an error from the database while trying to delete previous route calculations. Save this message as it might give us a clue what happened: " & sqlError.Message.ToString & "Sql executing was: " & strsql, MsgBoxStyle.Exclamation, My.Application.Info.ProductName.ToString)
                Throw New ApplicationException(sqlError.Message.ToString)
            Finally
                If Not IsNothing(myCommand) Then
                    myCommand.Connection.Close()
                End If
            End Try
        End Sub

        Private Sub Delete_RouteData_for_DailyActivity(ByVal intDailyActivityID As Integer)
            Dim myCommand As SqlCommand
            Dim strsql As String = ""
            Try
                Dim myConn = New SqlConnection(My.Settings.ACsquaredConnectionString)

                strsql = "delete tblDailyRoute where  DailyActivityID = " & intDailyActivityID

                myCommand = New SqlCommand(strsql, myConn)
                myCommand.Connection.Open()

                Dim intRowsDeleted As Integer = myCommand.ExecuteNonQuery
                'If intRowsDeleted <> 1 Then ' we didn't get the record updated. No reason this should or could ever happen
                '    Throw New ApplicationException("For some unknown reason we failed to update the route processed flag on DailyActivityID " & intDailyActivityID.ToString)
                'End If

            Catch sqlError As SqlException
                MsgBox("Sorry, we received an error from the database while trying to delete previous route calculations. Save this message as it might give us a clue what happened: " & sqlError.Message.ToString & "Sql executing was: " & strsql, MsgBoxStyle.Exclamation, My.Application.Info.ProductName.ToString)
                Throw New ApplicationException(sqlError.Message.ToString)
            Finally
                If Not IsNothing(myCommand) Then
                    myCommand.Connection.Close()
                End If
            End Try
        End Sub

        ''' <summary>
        '''  Assuming that the starting warehouse is ALWAYS the warehouse location associated with the 
        '''   systen the work is being done in, fetch the street address and set into the stand address,
        '''   lat /long is not supported for the warehouse location
        ''' </summary>
        ''' <param name="intDailyActivityId"></param>
        ''' <returns>Warehouse location in the StandardAddress structure</returns>
        ''' <remarks></remarks>
        Private Function GetStartingWareHouseAddress(ByVal intDailyActivityId) As StandardAddress

            Dim tbl As DataTable
            Dim row As DataRow

            Dim strsql As String = "SELECT DailyActivity.DailyActivityId, WarehouseLocations.Address, WarehouseLocations.City," & _
                    " WarehouseLocations.State, WarehouseLocations.ZIPCode FROM DailyActivity " & _
                    " INNER JOIN  System ON DailyActivity.SystemName = System.SystemName " & _
                    " INNER JOIN WarehouseLocations ON System.WarehouseName = WarehouseLocations.WarehouseName " & _
                    " where DailyActivityId = " & intDailyActivityId

            tbl = getDataItems(strsql)
            Dim laddress As New StandardAddress

            If tbl.Rows.Count <> 1 Then
                ' hum.. we have a problem. Warehouse location was not found
                Throw New ApplicationException("Exception Occured. Unable to find a warehouse location for DailyActivityId = " & intDailyActivityId)
            Else
                For Each row In tbl.Rows
                    laddress.Address = row("Address")
                    laddress.City = row("City")
                    laddress.State = row("State")
                    laddress.ZIP = row("ZIPCode")
                    laddress.TransactionId = 0
                Next
            End If


            Return laddress

        End Function

        ''' <summary>
        '''  Simple proc that sets the data based on datarow passed into the standard address structure
        ''' </summary>
        ''' <param name="WORow"> Row from a Work Order table containing address fields</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function SetWOAddress(ByVal WORow As System.Data.DataRow) As StandardAddress
            Try

                SetWOAddress.TransactionId = WORow("Transactionid")
                SetWOAddress.Address = WORow("HouseStreetname")
                SetWOAddress.City = WORow("CityName")
                SetWOAddress.ZIP = WORow("HouseZip")
                SetWOAddress.State = "CA"
                If WORow("latitude") Is DBNull.Value Then
                    SetWOAddress.latitude = 0
                Else
                    SetWOAddress.latitude = WORow("latitude")
                End If

                If WORow("Longtitude") Is DBNull.Value Then
                    SetWOAddress.longtitude = 0
                Else
                    SetWOAddress.longtitude = WORow("Longtitude")

                End If

            Catch sqlex As SqlException
                Throw New ApplicationException("Invalid data: Error returned while fetching data. Message is: " & sqlex.Message)
            Finally

            End Try
        End Function
        ''' <summary>
        '''  For any given Work Order row, insert row into the dailyroute table
        ''' (or perhaps later update the actual workorder table if we can sort out what to do with the return trip segment
        ''' </summary>
        ''' <param name="intDailyActivityId"></param>
        ''' <param name="SegmentTransactionId"></param>
        ''' <param name="dblAddressOrder"></param>
        ''' <param name="StartingSegmentTransactionId"></param>
        ''' <param name="MyRouteResults"></param>
        ''' <param name="baddressError"></param>
        ''' <param name="strErrReturned"></param>
        ''' <remarks>
        '''  This insert statement assume the following table structure
        '''CREATE TABLE [dbo].[tblDailyRoute]
        '''(
        '''[DailyActivityId] [int] NOT NULL,
        '''[SegmentTransactionId] [int] NOT NULL,
        '''   //* below here are additional data attributes */
        '''[AddrOrder] [int] NOT null,
        '''[StartingSegmentTransactionId] [int] NOT NULL,
        '''[Distance] [numeric] (18, 4) NULL,
        '''[ZoneID] [int] NULL, --- set to 0 in all instance right now
        '''[Processed] [bit] NULL, --- set to -1 in all instance right now
        '''[ProcessError] [varchar] (255) NULL,
        '''</remarks>
        ''' 
        Private Sub UpdateDbWithRouteCalculationResult _
                        (ByVal intDailyActivityId As Integer, ByVal SegmentTransactionId As Integer, _
                         ByVal dblAddressOrder As Double, ByVal StartingSegmentTransactionId As Integer, ByVal MyRouteResults As RouteData, _
                         ByVal baddressError As Boolean, Optional ByVal strErrReturned As String = "none")
            Dim myCommand As SqlCommand
            Dim myConn = New SqlConnection(My.Settings.ACsquaredConnectionString)
            Dim strsql As String = ""
            Dim intresult As Integer

            Try        '========= insert new row into the database for the new route segment ===================
                'strsql = "insert tblDailyRoute values (" & intDailyActivityId & "," & SegmentTransactionId & "," & _
                '   dblAddressOrder & "," & StartingSegmentTransactionId & "," & dblDistance & "," & _
                '   0 & "," & -1 & ",'" & Left(strErrReturned, 255) & "')"
                Dim errstring As String
                If (strErrReturned = "none") Then
                    errstring = String.Empty
                Else
                    errstring = Left(strErrReturned, 255)
                End If

                strsql = "update workorders set distance =  " & MyRouteResults.DistanceCalculated & ", DriveTime = " & MyRouteResults.DrivingTime & " , ProcessError = '" & errstring & _
                         "', Processed = -1 where TransactionId = " & SegmentTransactionId

                myCommand = New SqlCommand(strsql, myConn)
                myCommand.Connection.Open()
                intresult = myCommand.ExecuteNonQuery
                If intresult <> 1 Then ' we didn't get the record updated. No reason this should or could ever happen
                    '                    Throw New ApplicationException("Failed to insert a database row with SQL: " & strsql)
                    Throw New ApplicationException("Failed to update workorder row with SQL: " & strsql)
                End If

            Catch SQLex As Exception
                If Not myCommand.Connection Is Nothing Then
                    myCommand.Connection.Close()
                End If
                Throw New ApplicationException("SQL Error: " & SQLex.Message & " SQL issued was: " & strsql)
            Finally
                myCommand.Connection.Close()
            End Try         '========= insert new row into the database for the new route segment ===================

        End Sub

        Protected Overrides Sub Finalize()
            MyBase.Finalize()

        End Sub
    End Class
End Namespace