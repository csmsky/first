Imports MySql.Data.MySqlClient
Imports WindowsApplication1.ConfigClass

Public Class InputCustomerDetails

    Dim strConn As String = FDEandD()
    Dim conn As New MySqlConnection(strConn)

    ' Variables to remember context
    Private _transType As String = ""
    Private _currentOR As String = "" '<-- Variable to hold the OR number

    ' Call this from your Main POS Form to pass the Type AND the OR Number
    Public Sub SetTransactionDetails(type As String, orNo As String)
        _transType = type
        _currentOR = orNo
        Me.Text = "Enter Details for: " & type
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        auto()

        ' Validation: Ensure Name, ID, and OR Number (Barcode) are filled out
        If String.IsNullOrWhiteSpace(TextBoxGuestName.Text) OrElse String.IsNullOrWhiteSpace(TextBoxGuestID.Text) Then
            MessageBox.Show("Please enter Name and ID.", "Missing Info")
            Exit Sub
        End If

        Using conn As New MySqlConnection(strConn)
            Try
                conn.Open()

                Dim query As String = "INSERT INTO customer_tbl (or_no, name, id_no, tin_no, address, transType) VALUES (@or_no, @name, @id_no, @tin_no, @address, @transType)"

                Using cmd As New MySqlCommand(query, conn)
                    cmd.Parameters.AddWithValue("@or_no", mainform.TextBoxBarcode.Text)
                    cmd.Parameters.AddWithValue("@name", TextBoxGuestName.Text)
                    cmd.Parameters.AddWithValue("@id_no", TextBoxGuestID.Text)
                    cmd.Parameters.AddWithValue("@tin_no", TextBoxGuestTIN.Text)
                    cmd.Parameters.AddWithValue("@address", TextBoxGuestAddress.Text)
                    cmd.Parameters.AddWithValue("@transType", _transType)

                    cmd.ExecuteNonQuery()
                End Using

                MessageBox.Show("Customer saved to database!", "Success")

                ' Clean up the inputs
                Dim cleanTIN As String = TextBoxGuestTIN.Text.Replace("-", "").Trim()

                ' 1. Save to Memory ONLY (Database upload happens during printing)
                Dim newCustomer As New CustomerInfo With {
                    .Name = TextBoxGuestName.Text,
                    .ID = TextBoxGuestID.Text,
                    .TIN = cleanTIN,
                    .Address = TextBoxGuestAddress.Text,
                    .TransType = _transType
                }

                ' Add to your global/shared list
                eJournalCustomerData.CustomerList.Add(newCustomer)


                ' Notify the user it was added to the current transaction queue
                MessageBox.Show($"{_transType} details added to the current transaction!", "Success")

                ' Clear text boxes for the next entry
                TextBoxGuestName.Clear()
                TextBoxGuestID.Clear()
                TextBoxGuestTIN.Clear()
                TextBoxGuestAddress.Clear()

            Catch ex As Exception
                MessageBox.Show("Error: " & ex.Message)
            End Try
        End Using

        ' Close the form and return to POS
        Me.Close()
    End Sub

    Private Sub InputCustomerDetails_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    End Sub

    Private Sub auto()
        Dim seq As Single

        conn.Open()
        Dim qd As String = "SELECT MAX(or_no) AS receipt FROM or_tbl WHERE pos_id = '" & mainform.LabelPOSno.Text & "'"
        Dim cmd As New MySqlCommand(qd) With {.Connection = conn}
        Dim rdr As MySqlDataReader = cmd.ExecuteReader()
        While rdr.Read
            seq = Val(rdr.Item("receipt").ToString) + 1
        End While
        Select Case Len(Trim(seq))
            Case 1 : mainform.TextBoxBarcode.Text = "000000000000000" + Trim(Str(seq))
            Case 2 : mainform.TextBoxBarcode.Text = "00000000000000" + Trim(Str(seq))
            Case 3 : mainform.TextBoxBarcode.Text = "0000000000000" + Trim(Str(seq))
            Case 4 : mainform.TextBoxBarcode.Text = "000000000000" + Trim(Str(seq))
            Case 5 : mainform.TextBoxBarcode.Text = "00000000000" + Trim(Str(seq))
            Case 6 : mainform.TextBoxBarcode.Text = "0000000000" + Trim(Str(seq))
            Case 7 : mainform.TextBoxBarcode.Text = "000000000" + Trim(Str(seq))
            Case 8 : mainform.TextBoxBarcode.Text = "00000000" + Trim(Str(seq))
            Case 9 : mainform.TextBoxBarcode.Text = "0000000" + Trim(Str(seq))
            Case 10 : mainform.TextBoxBarcode.Text = "000000" + Trim(Str(seq))
            Case 11 : mainform.TextBoxBarcode.Text = "00000" + Trim(Str(seq))
            Case 12 : mainform.TextBoxBarcode.Text = "0000" + Trim(Str(seq))
            Case 13 : mainform.TextBoxBarcode.Text = "000" + Trim(Str(seq))
            Case 14 : mainform.TextBoxBarcode.Text = "00" + Trim(Str(seq))
            Case 15 : mainform.TextBoxBarcode.Text = "0" + Trim(Str(seq))
        End Select
        conn.Close()

    End Sub
End Class


Public Class CustomerInfo
    Public Property ORNumber As String
    Public Property Name As String
    Public Property ID As String
    Public Property TIN As String
    Public Property Address As String
    Public Property TransType As String
End Class
