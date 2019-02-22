Imports System.Data.SqlClient
Imports System.Data
Imports System.Data.SqlTypes

Module modACFunctions

    Public Sub main()
        'If Now > CDate("9/30/2009") Then
        '    End  ' no pomp and circumstance, just end the program
        'End If

        Call killTaskJobs("MapPoint")

        'frmMain.ShowDialog()
    End Sub

    Public Sub killTaskJobs(ByVal jobToKill As String)
        For Each p As Process In Process.GetProcessesByName(jobToKill)
            Try
                p.Kill()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Next
    End Sub

    Public okToOverride As Boolean

    ''' <summary>
    ''' Applies standard logic for the system to a name value pair
    ''' </summary>
    ''' <param name="itemName"></param>
    ''' <param name="itemValue"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function VerifyWorkOrderDataItems(ByVal itemName As String, ByVal itemValue As String)
        Dim sqlstr As String = ""
        VerifyWorkOrderDataItems = True

        Select Case UCase(itemName)
            Case "SYSNAME"
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("System Name cannot be blank")
                    ' e.Cancel = True
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                sqlstr = "Select systemName as dataitem from system where systemName = '" & itemValue & "'"
                If Not isDataValid(sqlstr) Then
                    Console.Beep(2000, 100)
                    MsgBox("You entered and invalid System Name")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

            Case "INSTALLER"
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("Installer Name cannot be blank")
                    ' e.Cancel = True
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If
                sqlstr = "Select installerName as dataitem from tblInstallers where installerName = '" & itemValue & "'"
                If Not isDataValid(sqlstr) Then
                    Console.Beep(2000, 100)
                    MsgBox("You entered and invalid Installer Name")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

            Case "INSTALLERNBR"
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("Installer Name cannot be blank")
                    ' e.Cancel = True
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If


            Case "TECHNUMBER"
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("Tech Number cannot be blank")
                    ' e.Cancel = True
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

            Case "LUNCH"
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("Lunch Duration cannot be blank")
                    ' e.Cancel = True
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If
                If Not IsNumeric(itemValue) Then
                    Console.Beep(2000, 100)
                    MsgBox("Invalid Lunch Duration!")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If


            Case "ZIP" ' zip code
                '         Dim strZip As String
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("Zip Code cannot be blank")
                    ' e.Cancel = True
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                If Len(Trim(itemValue)) < 5 Then
                    Console.Beep(2000, 100)
                    MsgBox("Invalid zip code!")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                If Not IsNumeric(itemValue) Then
                    Console.Beep(2000, 100)
                    MsgBox("Invalid zip code!")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

            Case "CITY" ' city
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("City cannot be blank")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                If Len(Trim(itemValue)) > 100 Then
                    Console.Beep(2000, 100)
                    MsgBox("City cannot be larger then 100 characters")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

            Case "ADDR" ' address
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("Address cannot be blank")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                If Len(Trim(itemValue)) > 50 Then
                    Console.Beep(2000, 100)
                    MsgBox("Address cannot be larger then 50 characters")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

            Case "JOBNR" ' job number
                If Len(Trim(itemValue)) > 0 Then
                    If Len(Trim(itemValue)) <> 6 Then
                        Console.Beep(2000, 100)
                        MsgBox("Job Number must be 6 characters or blank")
                        VerifyWorkOrderDataItems = False
                        Exit Function
                    End If
                End If

            Case "TIME"  ' must be valid times
                'lets make sure the is something in the cell
                If Len(Trim(itemValue)) = 0 Then
                    'nothin in the cell exit
                    Console.Beep(2000, 100)
                    MsgBox("Time field cannot be blank")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                'since we are using 24 hour clock time - len cannot be longer 
                'then 4 chars
                If Len(Trim(itemValue)) <> 4 Then
                    'nothin in the cell exit
                    Console.Beep(2000, 100)
                    MsgBox("Valid Time format is hhmm")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                'verify its numeric
                If Not IsNumeric(itemValue) Then
                    Console.Beep(2000, 100)
                    MsgBox("Time format must be numeric")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                If CInt(itemValue) > 2400 Then
                    Console.Beep(2000, 100)
                    MsgBox("Time entered cannot be greater then 2400.")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

                Dim lhour As String = Mid((itemValue), 1, 2)
                If Not IsNumeric(lhour) Then
                    'invalid hour
                    Console.Beep(2000, 100)
                    MsgBox("Valid Time format is hh:mm")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                Else
                    ' check to see if it is in the right range
                    If CInt(lhour) > 0 And CInt(lhour) <= 24 Then 'it is good
                    Else
                        'invalid hour
                        Console.Beep(2000, 100)
                        MsgBox("Valid Time format is hh:mm")
                        VerifyWorkOrderDataItems = False
                        Exit Function
                    End If
                End If

                Dim lMinutes As String = Mid((itemValue), 3, 2)
                If Not IsNumeric(lMinutes) Then
                    'invalid hour
                    Console.Beep(2000, 100)
                    MsgBox("Valid Time format is hh:mm")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                Else
                    ' check to see if it is in the right range
                    If CInt(lMinutes) >= 0 And CInt(lMinutes) <= 60 Then 'it is good
                    Else
                        'invalid hour
                        Console.Beep(2000, 100)
                        MsgBox("Valid Time format is hh:mm")
                        VerifyWorkOrderDataItems = False
                        Exit Function
                    End If
                End If

            Case "SERCD" ' service code
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("Primary Service Code cannot be blank")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If
                sqlstr = "Select LineItemId as dataitem from tblItems where LineItemId = '" & itemValue & "'"
                If Not isDataValid(sqlstr) Then
                    Console.Beep(2000, 100)
                    MsgBox("You entered and invalid Sevice Code")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If

            Case "QTY" 'qty
                If Len(Trim(itemValue)) = 0 Then
                    Console.Beep(2000, 100)
                    MsgBox("Quantity cannot be blank")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If
                If Not IsNumeric(itemValue) Then
                    Console.Beep(2000, 100)
                    MsgBox("Invalid Quantity Entered")
                    VerifyWorkOrderDataItems = False
                    Exit Function
                End If
        End Select
    End Function
    
    ''' <summary>
    '''  This procedure is needed to verify that data entered by the user exists in the
    '''  Tables -this is usually checked from a drop down - since we allow an item to be 
    '''  inactive - we cannot simply check to see if its in the drop down list since
    '''  items that are inactive - do not appear
    ''' </summary>
    ''' <param name="sqlStr"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function isDataValid(ByVal sqlStr As String) As Boolean

        Dim tbl As DataTable
        Dim row As DataRow
        Dim iName As String = ""
        isDataValid = False

        tbl = getDataItems(sqlStr)
        If tbl.Rows.Count <> 0 Then
            'there should only be one
            For Each row In tbl.Rows
                iName = row("DataItem")
            Next
        End If

        If iName <> "" Then
            isDataValid = True
        End If

        Return isDataValid

    End Function

    ''' <summary>
    '''  Better Idea Group generic handler for getting a datatable for a given SQL statement. Uses db connection
    '''  defined in the project 
    ''' </summary>
    ''' <param name="sqlStr"></param>
    ''' <returns>DataTable</returns>
    ''' <remarks></remarks>
    Public Function getDataItems(ByVal sqlStr As String) As DataTable
        Dim dt As DataTable = Nothing
        Dim reader As SqlDataReader = Nothing

        Using mConn As New SqlConnection(My.Settings.ACsquaredConnectionString())
            Dim mCommand As New SqlCommand(sqlStr.ToString, mConn)
            mConn.Open()
            Try
                reader = mCommand.ExecuteReader()
                dt = New DataTable
                dt.Load(reader)
            Catch sqlex As SqlException
                ' *****************************************************
                ' ********** get ALL the data we can from SQL error handler 
                Dim i As Integer
                Dim errorMessages As New System.Text.StringBuilder()
                For i = 0 To sqlex.Errors.Count - 1
                    errorMessages.Append("Index #" & i.ToString() & ControlChars.NewLine _
                        & "Message: " & sqlex.Errors(i).Message & ControlChars.NewLine _
                        & "LineNumber: " & sqlex.Errors(i).LineNumber & ControlChars.NewLine _
                        & "Source: " & sqlex.Errors(i).Source & ControlChars.NewLine _
                        & "Procedure: " & sqlex.Errors(i).Procedure & ControlChars.NewLine)
                Next i
                ' *****************************************************
                Select Case sqlex.Number
                    ' insert any "common" issues here for special handling, i.e. no server connection
                    Case Else ' unhandled
                        '  Throw New ApplicationException("Unable to get data items: " & errorMessages.ToString)
                        Dim myerror As New ErrorDialog
                        myerror.errMessage.Text = "[GetDataItem] Unable to get data items: " & errorMessages.ToString
                        myerror.ShowDialog()
                        myerror.Close()
                End Select

            Catch ex As Exception
                Throw ex
            Finally
                mConn.Close()
            End Try

        End Using

        Return dt
    End Function

    Public Function getInstallerName(ByVal insId As String) As String
        Dim sqlstr As String
        Dim tbl As DataTable
        Dim row As DataRow
        Dim iName As String = ""

        sqlstr = "Select installerName from tblinstallers where installerNo = " & insId
        tbl = getDataItems(sqlstr)
        If tbl.Rows.Count <> 0 Then
            'there should only be one
            For Each row In tbl.Rows
                iName = row("InstallerName")
            Next
        End If

        If iName = "" Then
            getInstallerName = "No Longer Valid"
        Else
            getInstallerName = iName
        End If


    End Function

    Public Function getWareHouseDescription(ByVal Id As String) As String
        Dim sqlstr As String
        Dim tbl As DataTable
        Dim row As DataRow
        Dim iName As String = ""

        sqlstr = "Select warehousedescription from warehouselocations where warehousename = '" & Id & "'"
        tbl = getDataItems(sqlstr)
        If tbl.Rows.Count <> 0 Then
            'there should only be one
            For Each row In tbl.Rows
                iName = row("warehousedescription")
            Next
        End If

        If iName = "" Then
            getWareHouseDescription = "No Longer Valid"
        Else
            getWareHouseDescription = iName
        End If


    End Function

    Public Function deleteDataItems(ByVal sqlStr As String) As Int32

        Dim result As Int32 = 0

        Using mConn As New SqlConnection(My.Settings.ACsquaredConnectionString())
            Dim mCommand As New SqlCommand(sqlStr.ToString, mConn)
            mConn.Open()
            Try
                result = mCommand.ExecuteNonQuery
            Catch ex As Exception
                result = -99
            Finally
                mConn.Close()
            End Try

        End Using

        Return result
    End Function

    Public Sub getZipCodeListBySystemName(ByVal systemNm As String, ByVal theObj As System.Windows.Forms.ComboBox)
        'clear zip codes
        theObj.Items.Clear()

        Dim sql As String
        Dim tbl As New DataTable

        sql = "Select distinct zipcode from vw_systemCityZip where systemName  = '" & systemNm & "'"

        tbl = getDataItems(sql)
        If tbl.Rows.Count <> 0 Then
            For Each row In tbl.Rows
                theObj.Items.Add(row("zipCode"))
            Next

        End If
    End Sub

    Public Function getActivityType(ByVal daId As String)
        Dim atype As Integer
        Dim tbl As New DataTable
        Dim sql As String

        'default is WO
        atype = 1

        sql = "Select activityType from dailyActivity where dailyactivityId  = " & daId
        tbl = getDataItems(sql)
        If tbl.Rows.Count <> 0 Then
            For Each row In tbl.Rows
                'should only be one row returned
                atype = row("activityType")

            Next
        End If

        Return atype
    End Function

    Public Sub getCityListByZipCode(ByVal zc As String, ByVal theObj As System.Windows.Forms.ComboBox)
        'clear combo box
        'theOBJ is the 
        theObj.Items.Clear()

        Dim tbl As New DataTable
        Dim sql As String

        sql = "Select cityname from vw_SystemCityZip where zipcode = '" & zc & "' group by cityname order by cityname"

        tbl = getDataItems(sql)
        If tbl.Rows.Count <> 0 Then
            For Each row In tbl.Rows
                theObj.Items.Add(row("Cityname"))

                'make this the default if only one row for zip selected
                If tbl.Rows.Count = 1 Then
                    theObj.Text = row("Cityname")
                End If
            Next
        Else
            theObj.Text = ""
        End If

    End Sub

    Public Function RouteErrorMessage(ByVal daid As String) As String
        Dim tbl As DataTable
        Dim row As DataRow
        Dim strResponse As String = ""
        Dim sql As String
        Try

            sql = "Select processerror from dailyactivity where dailyactivityid = " & daid

            tbl = getDataItems(sql)
            If tbl.Rows.Count <> 0 Then
                For Each row In tbl.Rows
                    strResponse += (row("processerror").ToString)
                Next
            Else
                strResponse = "No data found for " & daid
            End If

        Catch sqlex As SqlException
            strResponse = "Lookup of error message failed: " & sqlex.Message
        End Try

        Return strResponse
    End Function

    Function EndGreaterThenStart(ByVal stime As String, ByVal etime As String) As Boolean

        'validate whether end time is greater then start time
        EndGreaterThenStart = True
        If stime >= etime Then
            EndGreaterThenStart = False
            Console.Beep(2000, 100)
            MsgBox("End Time must be Greater then Start Time!")
        End If

        Return EndGreaterThenStart
    End Function

    Public Sub getSystemStartEndTimes(ByVal systemNm As String, ByVal systemStartTime As String, ByVal systemEndTime As String)

        Dim sql As String
        Dim tbl As New DataTable

        sql = "Select * from system where systemName = '" & systemNm & "'"

        tbl = getDataItems(sql)
        If tbl.Rows.Count <> 0 Then
            For Each row In tbl.Rows

                systemStartTime = FormatDateTime(row("startTime"), DateFormat.ShortTime).ToString
                systemEndTime = FormatDateTime(row("endTime"), DateFormat.ShortTime).ToString
            Next
        End If
    End Sub

    Public Sub UpdateDailyActivityAudit(ByVal vDAID, ByVal vAction, ByVal vdtOfWork, ByVal vSysName, ByVal vInstallerName, ByVal vInstallerNbr, ByVal vSysTechNbr, ByVal vDtStart, ByVal vDtEnd, ByVal vLunchDur)
        Dim sConn As SqlConnection = New SqlConnection()
        Dim auditIntresult As Integer = 0

        Try
            sConn.ConnectionString = My.Settings.ACsquaredConnectionString()
            Dim sDA As New SqlCommand("ac_sp_UpdateDailyActivity", sConn)
            sDA.CommandType = CommandType.StoredProcedure
            sDA = New SqlCommand("ac_sp_UpdateDailyActivityAudit", sConn)
            sDA.CommandType = CommandType.StoredProcedure

            sDA.Parameters.Add("@pAction", SqlDbType.NChar).Value = vAction
            sDA.Parameters.Add("@pDAId", SqlDbType.BigInt).Value = vDAID
            sDA.Parameters.Add("@pDateOfWork", SqlDbType.NChar).Value = vdtOfWork
            sDA.Parameters.Add("@pSystemName", SqlDbType.NChar).Value = vSysName
            sDA.Parameters.Add("@pInstallerName", SqlDbType.NChar).Value = vInstallerName
            sDA.Parameters.Add("@pInstallerNo", SqlDbType.Int).Value = vInstallerNbr
            sDA.Parameters.Add("@pTechN0", SqlDbType.NChar).Value = vSysTechNbr
            sDA.Parameters.Add("@pShiftSTime", SqlDbType.NChar).Value = vDtStart
            sDA.Parameters.Add("@pShiftETime", SqlDbType.NChar).Value = vDtEnd
            sDA.Parameters.Add("@pLunchDuaration", SqlDbType.NChar).Value = vLunchDur
            sDA.Parameters.Add("@RowCount", SqlDbType.Int)
            sDA.Parameters("@RowCount").Direction = ParameterDirection.ReturnValue

            sConn.Open()
            auditIntresult = sDA.ExecuteNonQuery()
           
        Catch sqlex As SqlException
            MsgBox("BtnSave - SQL Error saving data: " & sqlex.Message.ToString)
        Catch ex As Exception
            MsgBox("BtnSave - Error saving data: " & ex.Message.ToString)
        Finally

            sConn.Close()
        End Try

    End Sub

    Public Sub UpdateWOAudit(ByVal vWOID, ByVal vAction, ByVal vDAId, ByVal vZip, ByVal vCity, ByVal vAddr, ByVal vApt, ByVal vSeq, ByVal vStime, ByVal vEtime, ByVal vPSerCd, ByVal vPQty, ByVal vNotes, ByVal vPDate, ByVal vSerCd1, ByVal vSerQty1, ByVal vSerCd2, ByVal vSerQty2, ByVal vSerCd3, ByVal vSerQty3, ByVal vSerCd4, ByVal vSerQty4, ByVal vSerCd5, ByVal vSerQty5)

        Dim iDAId As Int64
        iDAId = vDAId
        Dim sConn As SqlConnection = New SqlConnection()

        sConn.ConnectionString = My.Settings.ACsquaredConnectionString()
        Dim sWO As New SqlCommand("ac_sp_UpdateWorkOrdersAudit", sConn)
        sWO.CommandType = CommandType.StoredProcedure

        'these parms define the work order
        sWO.Parameters.Add("@pAction", SqlDbType.NChar).Value = vAction
        sWO.Parameters.Add("@pwoToUpdate", SqlDbType.NChar).Value = vWOID
        sWO.Parameters.Add("@pDAId", SqlDbType.BigInt).Value = iDAId
        sWO.Parameters.Add("@pHouseZip", SqlDbType.NChar).Value = vZip
        sWO.Parameters.Add("@pCityName", SqlDbType.NVarChar).Value = vCity
        sWO.Parameters.Add("@pHouseStreetName", SqlDbType.NVarChar).Value = vAddr
        sWO.Parameters.Add("@pHouseApt", SqlDbType.NChar).Value = vApt
        sWO.Parameters.Add("@pJobNo", SqlDbType.NChar).Value = vSeq
        sWO.Parameters.Add("@pStartTime", SqlDbType.NChar).Value = vStime
        sWO.Parameters.Add("@pEndTime", SqlDbType.NChar).Value = vEtime
        sWO.Parameters.Add("@pPrimaryLineItemId", SqlDbType.NVarChar).Value = vPSerCd
        sWO.Parameters.Add("@pPrimaryLineQty", SqlDbType.Real).Value = CDbl(vPQty)
        sWO.Parameters.Add("@pWoNotes", SqlDbType.VarChar).Value = vNotes

        sWO.Parameters.Add("@pPrintDate", SqlDbType.DateTime).Value = CDate(vPDate)

        'these parms define the wo detail
        '     sWO.Parameters.Add("@pLineItem1", SqlDbType.NVarChar).Value = vSerCd1
        '     sWO.Parameters.Add("@pItemQty1", SqlDbType.NChar).Value = vSerQty1
        '     sWO.Parameters.Add("@pLineItem2", SqlDbType.NVarChar).Value = vSerCd2
        '     sWO.Parameters.Add("@pItemQty2", SqlDbType.NChar).Value = vSerQty2
        '     sWO.Parameters.Add("@pLineItem3", SqlDbType.NVarChar).Value = vSerCd3
        '     sWO.Parameters.Add("@pItemQty3", SqlDbType.NChar).Value = vSerQty3
        '     sWO.Parameters.Add("@pLineItem4", SqlDbType.NVarChar).Value = vSerCd4
        '     sWO.Parameters.Add("@pItemQty4", SqlDbType.NChar).Value = vSerQty4
        '     sWO.Parameters.Add("@pLineItem5", SqlDbType.NVarChar).Value = vSerCd5
        '     sWO.Parameters.Add("@pItemQty5", SqlDbType.NChar).Value = vSerQty5

        '      Dim rParm As SqlParameter = sWO.Parameters.Add("@tranCnt", SqlDbType.Int)
        '      rParm.Direction = ParameterDirection.ReturnValue

        sConn.Open()
        Dim intresult As Integer = sWO.ExecuteNonQuery()
        '   Dim rval As Integer = sWO.Parameters("@tranCnt").Value
        '   If (intresult <> -1 Or rval <> 0) Then
        'failed
        ' saveWORecord = 9999
        ' Throw New ApplicationException("Failed to update")
        ' Else
        ' success, do nothing?
        ' End If
        sConn.Close()


    End Sub

End Module