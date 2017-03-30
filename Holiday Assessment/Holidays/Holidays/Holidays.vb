Imports System.IO

Public Class Holidays

    Private Structure Holiday

        Public HolidayID As String
        Public HolidayName As String
        Public Location As String
        Public HolidayType As String
        Public Rating As String

    End Structure

    Private Sub Holidays_Load() Handles MyBase.Load

        Customers.Show()

        txtHolidayID.Enabled = False

        If Dir$("Holidays.txt") = "" Then

            Dim sw As New StreamWriter("Holidays.txt", True)

            sw.WriteLine("0")

            sw.Close()

            MsgBox("A new file has been created", vbExclamation, "Warning!")

        End If

    End Sub

    Private Sub cmdSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSave.Click

        Dim HolidayData As New Holiday

        Dim HolidaysData() As String = File.ReadAllLines(Dir$("Holidays.txt"))

        Dim CurrentHolidayID As Integer = 1

        txtHolidayID.Text = ""

        If chkSaveMode.Checked = False Then

            MsgBox("You must tick save mode if you would like to save data.")

            Exit Sub

        End If

        For i = 0 To UBound(HolidaysData)

            If Val(Trim(Mid(HolidaysData(i), 1, 4))) = CurrentHolidayID Then

                CurrentHolidayID = CurrentHolidayID + 1

            End If

        Next

        If Validation(HolidaysData) = True Then

            Dim sw As New System.IO.StreamWriter("Holidays.txt", True)

            HolidayData.HolidayID = LSet(CurrentHolidayID, 4)
            txtHolidayID.Text = CurrentHolidayID

            HolidayData.HolidayName = LSet(txtHolidayName.Text, 50)
            HolidayData.HolidayType = LSet(txtHolidayType.Text, 50)
            HolidayData.Location = LSet(txtLocation.Text, 50)
            HolidayData.Rating = LSet(txtRating.Text, 50)

            sw.WriteLine(HolidayData.HolidayID & HolidayData.HolidayName & HolidayData.Location & HolidayData.HolidayType & HolidayData.Rating)

            sw.Close()

            MsgBox("File Saved! Holiday ID: " & HolidayData.HolidayID)

        End If

    End Sub

    Private Sub cmdSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSearch.Click

        If chkSaveMode.Checked = True Then

            MsgBox("You must untick save mode if you would like to search data.")

            Exit Sub

        End If

        Dim HolidayData() As String = File.ReadAllLines("Holidays.txt")

        Dim HolidayFound As Boolean
        Dim HolidayCount As Integer
        Dim FoundHoliday As Integer

        HolidayFound = False
        HolidayCount = 0

        For i = 0 To UBound(HolidayData)

            If txtHolidayID.Text <> "" Then

                If Trim(Mid(HolidayData(i), 1, 4)) = txtHolidayID.Text Then

                    MsgBox("A holiday with this Holiday ID has been found.")

                    HolidayFound = True

                    txtHolidayID.Text = Trim(Mid(HolidayData(i), 1, 4))
                    txtHolidayName.Text = Trim(Mid(HolidayData(i), 5, 50))
                    txtLocation.Text = Trim(Mid(HolidayData(i), 55, 50))
                    txtHolidayType.Text = Trim(Mid(HolidayData(i), 105, 50))
                    txtRating.Text = Trim(Mid(HolidayData(i), 155, 50))

                End If

            Else

                If (Trim(Mid(HolidayData(i), 5, 50)) = txtHolidayName.Text Or txtHolidayName.Text = "") And (Trim(Mid(HolidayData(i), 55, 50)) = txtLocation.Text Or txtLocation.Text = "") And (Trim(Mid(HolidayData(i), 105, 50)) = txtHolidayType.Text Or txtHolidayType.Text = "") And (Trim(Mid(HolidayData(i), 155, 50)) = txtRating.Text Or txtRating.Text = "") Then

                    FoundHoliday = i

                    HolidayCount = HolidayCount + 1

                End If

            End If

        Next i

        If txtHolidayID.Text = "" Then

            If HolidayCount = 1 Then

                MsgBox("One holiday was found.")

                txtHolidayID.Text = Trim(Mid(HolidayData(FoundHoliday), 1, 4))
                txtHolidayName.Text = Trim(Mid(HolidayData(FoundHoliday), 5, 50))
                txtLocation.Text = Trim(Mid(HolidayData(FoundHoliday), 55, 50))
                txtHolidayType.Text = Trim(Mid(HolidayData(FoundHoliday), 105, 50))
                txtRating.Text = Trim(Mid(HolidayData(FoundHoliday), 155, 50))

                Exit Sub

            Else

                MsgBox("There were " & HolidayCount & " holidays found.")

            End If

        End If

        If txtHolidayID.Text <> "" And HolidayFound = False Then

            MsgBox("A holiday with this Holiday ID has not been found.")

            ClearTextboxes()

        End If

    End Sub

    Private Sub chkSaveMode_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkSaveMode.Click

        If chkSaveMode.Checked = True Then

            MsgBox("Save Mode enabled.")

            ClearTextboxes()

            txtHolidayID.Enabled = False

        Else

            MsgBox("Search Mode enabled.")

            ClearTextboxes()

            txtHolidayID.Enabled = True

        End If

    End Sub

    Private Function Validation(HolidaysData)

        Dim Validated As Boolean

        Validated = True

        If txtHolidayName.Text = "" Or txtLocation.Text = "" Or txtHolidayType.Text = "" Or txtRating.Text = "" Then

            Validated = False

            MsgBox("You must save a holiday name, location, type and rating.")

        ElseIf (IsNumeric(txtRating.Text) = False) Or (Val(txtRating.Text) < 0 Or Val(txtRating.Text) > 5) Then

            Validated = False

            MsgBox("Rating must be a number beteen 1 and 5.")

        End If

        For i = 0 To UBound(HolidaysData)

            If Trim(Mid(HolidaysData(i), 5, 50)) = txtHolidayName.Text And Trim(Mid(HolidaysData(i), 55, 50)) = txtLocation.Text And Trim(Mid(HolidaysData(i), 105, 50)) = txtHolidayType.Text And Trim(Mid(HolidaysData(i), 155, 50)) = txtRating.Text Then

                Validated = False

                MsgBox("This data has already been saved previously.")

                Exit For

            End If

        Next

        Return Validated

    End Function

    Private Sub ClearTextboxes()

        txtHolidayID.Text = ""
        txtHolidayName.Text = ""
        txtLocation.Text = ""
        txtHolidayType.Text = ""
        txtRating.Text = ""

    End Sub

    Private Sub btnClear_Click(sender As System.Object, e As System.EventArgs) Handles btnClear.Click

        ClearTextboxes()

    End Sub

End Class