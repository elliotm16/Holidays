Imports System.IO

Public Class Customers

    ' Structure for the customer's data
    Private Structure Customer

        Public CustomerID As String ' Used to uniquely identify a customer
        Public FirstName As String
        Public Surname As String
        Public EmailAddress As String
        Public PhoneNumber As String

    End Structure

    Private Sub Holidays_Load() Handles MyBase.Load

        ' Disables input into the customer ID textbox
        txtCustomerID.Enabled = False

        If Dir$("Customers.txt") = "" Then

            Dim sw As New StreamWriter("Customers.txt", True)

            sw.WriteLine("0")

            ' Needs to be close after use
            sw.Close()

            ' Warning that a new file has been created
            MsgBox("A new file has been created", vbExclamation, "Warning!")

        End If

    End Sub

    Private Sub cmdSave_Click(sender As System.Object, e As System.EventArgs) Handles cmdSave.Click

        Dim CustomerData As New Customer

        Dim CustomersData() As String = File.ReadAllLines(Dir$("Customers.txt"))

        ' Lowest an ID can be
        Dim CurrentCustomerID As Integer = 1

        ' Clear the customer ID textbox
        txtCustomerID.Text = ""

        ' If the user is in read mode
        If chkSaveMode.Checked = False Then

            MsgBox("You must tick save mode if you would like to save data.")

            Exit Sub

        End If

        ' i starts off at zero and incremements until it reaches the upper bound of customer data
        For i = 0 To UBound(CustomersData)

            ' If the highest ID is equal to the current customer ID, add one to it
            If Val(Trim(Mid(CustomersData(i), 1, 4))) = CurrentCustomerID Then

                ' Increment the value of the current customer ID
                CurrentCustomerID = CurrentCustomerID + 1

            End If

        Next

        ' If the data passed the test for validation
        If Validation(CustomersData) = True Then

            Dim sw As New System.IO.StreamWriter("Customers.txt", True)

            ' The Customer ID is padded and saved to the structure
            CustomerData.CustomerID = LSet(CurrentCustomerID, 4)

            ' THE ID is then saved to the textbox for the user to view
            txtCustomerID.Text = CurrentCustomerID

            ' The data can now be saved as it's in the correct format

            CustomerData.FirstName = LSet(txtFirstName.Text, 50)
            CustomerData.Surname = LSet(txtEmailAddress.Text, 50)
            CustomerData.EmailAddress = LSet(txtSurname.Text, 50)
            CustomerData.PhoneNumber = LSet(txtPhoneNumber.Text, 50)

            ' Write the data to a textfile
            sw.WriteLine(CustomerData.CustomerID & CustomerData.FirstName & CustomerData.Surname & CustomerData.EmailAddress & CustomerData.PhoneNumber)

            sw.Close()

            ' Ouput that the file is saved with the customer ID
            MsgBox("File Saved! Customer ID: " & CustomerData.CustomerID)

        End If

    End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click

        ' If the user is in save mode
        If chkSaveMode.Checked = True Then

            MsgBox("You must untick save mode if you would like to search data.")

            Exit Sub

        End If

        Dim CustomerData() As String = File.ReadAllLines("Customers.txt")

        Dim CustomerFound As Boolean
        Dim CustomerCount As Integer
        Dim FoundCustomer As Integer

        CustomerFound = False
        CustomerCount = 0

        ' i starts off at zero and incremements until it reaches the upper bound of customer data
        For i = 0 To UBound(CustomerData)

            ' If the customer is searching a customer ID
            If txtCustomerID.Text <> "" Then

                If Trim(Mid(CustomerData(i), 1, 4)) = txtCustomerID.Text Then

                    MsgBox("A customer with this Customer ID has been found.")

                    ' A customer has been found
                    CustomerFound = True

                    ' Output the customer's data to the textboxes
                    txtCustomerID.Text = Trim(Mid(CustomerData(i), 1, 4))
                    txtFirstName.Text = Trim(Mid(CustomerData(i), 5, 50))
                    txtSurname.Text = Trim(Mid(CustomerData(i), 55, 50))
                    txtEmailAddress.Text = Trim(Mid(CustomerData(i), 105, 50))
                    txtPhoneNumber.Text = Trim(Mid(CustomerData(i), 155, 50))

                End If

            Else

                ' If the data in the textfile match the textboxes
                If (Trim(Mid(CustomerData(i), 5, 50)) = txtFirstName.Text Or txtFirstName.Text = "") And (Trim(Mid(CustomerData(i), 55, 50)) = txtSurname.Text Or txtSurname.Text = "") And (Trim(Mid(CustomerData(i), 105, 50)) = txtEmailAddress.Text Or txtEmailAddress.Text = "") And (Trim(Mid(CustomerData(i), 155, 50)) = txtPhoneNumber.Text Or txtPhoneNumber.Text = "") Then

                    ' A customer has been found, and the data is stored in the variable
                    FoundCustomer = i

                    ' Incrememnt number of customers
                    CustomerCount = CustomerCount + 1

                End If

            End If

        Next i

        ' If the user is not searching for a customer ID
        If txtCustomerID.Text = "" Then

            If CustomerCount = 1 Then

                MsgBox("One customer was found.")

                ' Output the customer's data to the textboxes
                txtCustomerID.Text = Trim(Mid(CustomerData(FoundCustomer), 1, 4))
                txtFirstName.Text = Trim(Mid(CustomerData(FoundCustomer), 5, 50))
                txtSurname.Text = Trim(Mid(CustomerData(FoundCustomer), 55, 50))
                txtEmailAddress.Text = Trim(Mid(CustomerData(FoundCustomer), 105, 50))
                txtPhoneNumber.Text = Trim(Mid(CustomerData(FoundCustomer), 155, 50))

                Exit Sub

            Else

                MsgBox("There were " & CustomerCount & " customers found.")

            End If

        End If

        ' If the user is searching for customer ID and a customer has not been found
        If txtCustomerID.Text <> "" And CustomerFound = False Then

            MsgBox("A customer with this Customer ID has not been found.")

            ClearTextboxes()

        End If

    End Sub

    Private Sub chkSaveMode_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSaveMode.Click

        ' If the user is in save mode, otherwise in read mode
        If chkSaveMode.Checked = True Then

            ' User can now save
            MsgBox("Save Mode enabled.")

            ClearTextboxes()

            ' Customer ID textbox can not be changed as they are in save mode
            txtCustomerID.Enabled = False

        Else

            ' User can now search
            MsgBox("Search Mode enabled.")

            ClearTextboxes()

            ' Customer ID textbox can be changed to allow the user to search
            txtCustomerID.Enabled = True

        End If

    End Sub

    Private Function Validation(CustomersData)

        Dim Validated As Boolean

        Validated = True

        ' If any of the textboxes are blank, other than customer ID as this is generated
        ' If phone number isn't numeric or it's not 11 characters
        If txtFirstName.Text = "" Or txtSurname.Text = "" Or txtEmailAddress.Text = "" Or txtPhoneNumber.Text = "" Then

            ' Data failed validation
            Validated = False

            MsgBox("You must save a customer's full name, email address and phone number.")

        ElseIf (IsNumeric(txtPhoneNumber.Text) = False) Or (Val(txtPhoneNumber.Text.Length) <> 11) Then

            ' Data failed validation
            Validated = False

            MsgBox("Phone number must consist of 11 numbers.")

        End If

        ' i starts off at zero and incremements until it reaches the upper bound of customer data
        For i = 0 To UBound(CustomersData)

            ' If all the data entered matches a set of data previously entered
            If Trim(Mid(CustomersData(i), 5, 50)) = txtFirstName.Text And Trim(Mid(CustomersData(i), 55, 50)) = txtSurname.Text And Trim(Mid(CustomersData(i), 105, 50)) = txtEmailAddress.Text And Trim(Mid(CustomersData(i), 155, 50)) = txtPhoneNumber.Text Then

                ' Data failed validation
                Validated = False

                MsgBox("This data has already been saved previously.")

                Exit For

            End If

        Next

        ' Return true or false on whether the data is validated
        Return Validated

    End Function

    Private Sub ClearTextboxes()

        ' Clears all the textboxes

        txtCustomerID.Text = ""
        txtFirstName.Text = ""
        txtSurname.Text = ""
        txtEmailAddress.Text = ""
        txtPhoneNumber.Text = ""

    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click

        ' Called when the user clicks the 'Clear' button
        ClearTextboxes()

    End Sub

End Class