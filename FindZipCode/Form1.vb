Imports System.Data.SqlClient
Imports MapPoint
Imports MapPoint.GeoPaneState

Public Class Form1
    Private MyAddressValidator As StandardMiles.AddressValidation
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

        Call Main()
        ''********************************************************
        'Dim woAddressError As StandardMiles.AddressError

        'MyAddressValidator = New StandardMiles.AddressValidation

        '' This is where I need to iterate thru all of the records that need to get a zip code
        'Try
        '    woAddressError = MyAddressValidator.ValidateAddress(txtAddress.Text, txtCity.Text)

        '    ' this is the output
        '    Select Case woAddressError.FindResultsQuality
        '        Case GeoFindResultsQuality.geoFirstResultGood, GeoFindResultsQuality.geoAllResultsValid
        '            Debug.Print(woAddressError.Address.Address)
        '            Debug.Print(woAddressError.Address.City)
        '            Debug.Print(woAddressError.Address.State)
        '            Debug.Print(woAddressError.Address.ZIP)

        '            txtZip.Text = woAddressError.Address.ZIP

        '        Case GeoFindResultsQuality.geoAmbiguousResults
        '            Debug.Print(woAddressError.Address.Address)
        '            Debug.Print(woAddressError.Address.ZIP)
        '            txtZip.Text = woAddressError.Address.ZIP

        '        Case GeoFindResultsQuality.geoNoGoodResult, GeoFindResultsQuality.geoNoResults
        '            Debug.Print(woAddressError.ErrorDescription)
        '        Case Else
        '            MsgBox("Unexpected issue, Call support!")

        '    End Select


        'Catch ex As Exception
        '    MsgBox(ex.Message)

        'End Try

    End Sub

    Sub Main()
        Dim woAddressError As StandardMiles.AddressError

        MyAddressValidator = New StandardMiles.AddressValidation

        Try
            ' This is where I need to iterate thru all of the records that need to get a zip code
            Dim tbl As New DataTable
            Dim sql As String
            Dim row As DataRow

            sql = "Select * from workorders"

            tbl = getDataItems(sql)
            If tbl.Rows.Count <> 0 Then
                For Each row In tbl.Rows
                    Try
                        Debug.Print("Working with rowid:" & row("ID").ToString())

                        woAddressError = MyAddressValidator.ValidateAddress(row("Address").ToString, row("City").ToString)

                        updateAddressRecord(row("Id"), woAddressError)

                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                Next
            Else
            End If

        Catch sqlex As SqlException
            ' if I am unable to get a good connection to the database, going to throw up a simple
            ' Messagebox just to let the user know they are in trouble
            Dim myError As New ErrorDialog
            myError.errMessage.Text = "Unable to determine current user or permissions. Check database connection" & vbCrLf & "Perhaps this info from SQL will be usefull " & sqlex.Message & vbCrLf & "FullText: " & sqlex.ToString
            myError.lblCaption.Text = "Critical Issue"
            myError.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try



    End Sub

    Sub updateAddressRecord(id As Integer, results As StandardMiles.AddressError)

        Dim myCommand As SqlCommand
        Try

            Dim myConn = New SqlConnection(My.Settings.ACsquaredConnectionString)

            Dim strsql As String = ""

            strsql = "update Workorders set zip = '" & results.Address.ZIP & "' , [FindResultsQuality] = '" & results.FindResultsQuality.ToString & "' where ID = " & id

            myCommand = New SqlCommand(strsql, myConn)
            myCommand.Connection.Open()

            'get the last id of Work order entered in case we need to insert additional
            'codes - ExecuteScalar returns that value
            Dim intresult As Integer = myCommand.ExecuteNonQuery
            If intresult <> 1 Then ' we didn't get the record updated. No reason this should or could ever happen
                Throw New ApplicationException("For some unknown reason we failed to update the database")
            End If

        Catch sqlError As SqlException
            Debug.Print("Sorry, we received an error from the database while trying to update it. Save this message as it might give us a clue what happened: " & sqlError.Message.ToString, MsgBoxStyle.Exclamation, My.Application.Info.ProductName.ToString)
            'Throw New ApplicationException(sqlError.Message.ToString)
        Finally
            If Not IsNothing(myCommand) Then
                myCommand.Connection.Close()
            End If
        End Try
    End Sub

End Class
