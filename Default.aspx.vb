Imports System.Data
Imports System.Data.Sql
Imports System.Data.SqlClient
Imports System.IO
Imports ClosedXML.Excel

Partial Class _Default
    Inherits Page

    Dim Conn As New SqlConnection(System.Configuration.ConfigurationManager.ConnectionStrings("DB").ConnectionString)
    Private Property cmd As SqlCommand
    Dim rdr As SqlDataReader

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles btnUpload.Click
        Dim val As String

        val = DropDownList1.SelectedValue

        If val = "No Value Selected" Then
            Exit Sub
        Else
            Select Case val
                Case "Advanced Core Study"
                    insertintoACS()
                Case "Atterberg"
                    insertAtt()
                Case "Geochemistry"
                    insertGeo()
                Case "Gravel Sand Silt and Clay"
                    insertGSSC()
                Case "Particle Size Characteristics"
                    insertPSC()
                Case "RBRC"
                    insertRBRC()
                Case "Sieves Passing"
                    insertSP()
            End Select
        End If
    End Sub

    Private Sub insertintoACS()
        Dim wb As XLWorkbook
        Dim insert As String
        Dim rowCount As Integer
        Dim errLog As String

        'save file
        'FileUpload1.SaveAs(Server.MapPath("TempFile"))

        insert = "insert into TBL_ASSAY_ADVCORESTUDY(Lab, DateReceived, TestDate, SampleNumber, TotalPorosityPercent, RBRCPercent, TotalPorosityFraction, DrainablePorosityFraction, FileName, DateLoaded) values(@Lab, @DateReceived, @TestDate, @SampleNumber, @TotalPorosityPercent, @RBRCPercent, @TotalPorosityFraction, @DrainablePorosityFraction, @FileName, @DateLoaded)"
        wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        Conn.Open()
        Dim lastrow = ws.RangeUsed
        errLog = ""

        For Each row In ws.Rows
            rowCount = row.RowNumber
            Using cmd As New SqlCommand
                With cmd
                    .Connection = Conn
                    .CommandType = Data.CommandType.Text
                    .CommandText = insert
                    .Parameters.AddWithValue("@Lab", ws.Cell(rowCount, 1).GetString)
                    .Parameters.AddWithValue("@DateReceived", ws.Cell(rowCount, 2).GetString)
                    .Parameters.AddWithValue("@TestDate", ws.Cell(rowCount, 3).GetString)
                    .Parameters.AddWithValue("@SampleNumber", ws.Cell(rowCount, 4).GetString)
                    .Parameters.AddWithValue("@TotalPorosityPercent", ws.Cell(rowCount, 12).GetString)
                    .Parameters.AddWithValue("@RBRCPercent", ws.Cell(rowCount, 13).GetString)
                    .Parameters.AddWithValue("@TotalPorosityFraction", ws.Cell(rowCount, 14).GetString)
                    .Parameters.AddWithValue("@DrainablePorosityFraction", ws.Cell(rowCount, 15).GetString)
                    .Parameters.AddWithValue("@FileName", ws.Cell(rowCount, 9).GetString)
                    .Parameters.AddWithValue("@DateLoaded", ws.Cell(rowCount, 11).GetString)
                End With
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    errLog = errLog & "Row: " & rowCount & " encountered an error!" & vbNewLine
                End Try
            End Using
        Next
        Conn.Close()

        'Throw New NotImplementedException()
        Dim val As String

        If errLog <> "" Then
            lblError.Text = errLog
        End If
        val = DropDownList1.SelectedValue
        lblSuccess.Visible = True
        lblSuccess.Text = val & " Uploaded Successfully!"

    End Sub

    Private Sub insertSP()
        Dim wb As XLWorkbook
        Dim insert As String
        Dim rowCount As Integer
        Dim errLog As String

        'save file
        'FileUpload1.SaveAs(Server.MapPath("TempFile"))

        insert = "insert into TBL_ASSAY_SIEVES_PASS(Lab, DateReceived, TestDate, SampleNumber, Inch3, Inch2, Inch1_5, Inch1, Inch3_4, Inch3_8, Inch4, Inch10, Inch20, Inch40, Inch60, Inch140, Inch200, FileName, DateLoaded) values(@Lab, @DateReceived, @TestDate, @SampleNumber, @Inch3, @Inch2, @Inch1_5, @Inch1, @Inch3_4, @Inch3_8, @Inch4, @Inch10, @Inch20, @Inch40, @Inch60, @Inch140, @Inch200 @FileName, @DateLoaded)"
        wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        Conn.Open()
        Dim lastrow = ws.RangeUsed
        errLog = ""

        For Each row In ws.Rows
            rowCount = row.RowNumber
            Using cmd As New SqlCommand
                With cmd
                    .Connection = Conn
                    .CommandType = Data.CommandType.Text
                    .CommandText = insert
                    .Parameters.AddWithValue("@Lab", ws.Cell(rowCount, 1).GetString)
                    .Parameters.AddWithValue("@DateReceived", ws.Cell(rowCount, 2).GetString)
                    .Parameters.AddWithValue("@TestDate", ws.Cell(rowCount, 3).GetString)
                    .Parameters.AddWithValue("@SampleNumber", ws.Cell(rowCount, 4).GetString)
                    .Parameters.AddWithValue("@Inch3", ws.Cell(rowCount, 21).GetString)
                    .Parameters.AddWithValue("@Inch2", ws.Cell(rowCount, 22).GetString)
                    .Parameters.AddWithValue("@Inch1_5", ws.Cell(rowCount, 23).GetString)
                    .Parameters.AddWithValue("@Inch1", ws.Cell(rowCount, 24).GetString)
                    .Parameters.AddWithValue("@Inch3_4", ws.Cell(rowCount, 25).GetString)
                    .Parameters.AddWithValue("@Inch3_8", ws.Cell(rowCount, 26).GetString)
                    .Parameters.AddWithValue("@Inch4", ws.Cell(rowCount, 27).GetString)
                    .Parameters.AddWithValue("@Inch10", ws.Cell(rowCount, 28).GetString)
                    .Parameters.AddWithValue("@Inch20", ws.Cell(rowCount, 29).GetString)
                    .Parameters.AddWithValue("@Inch40", ws.Cell(rowCount, 30).GetString)
                    .Parameters.AddWithValue("@Inch60", ws.Cell(rowCount, 31).GetString)
                    .Parameters.AddWithValue("@Inch140", ws.Cell(rowCount, 32).GetString)
                    .Parameters.AddWithValue("@Inch200", ws.Cell(rowCount, 33).GetString)
                    .Parameters.AddWithValue("@FileName", ws.Cell(rowCount, 18).GetString)
                    .Parameters.AddWithValue("@DateLoaded", ws.Cell(rowCount, 20).GetString)
                End With
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    errLog = errLog & "Row: " & rowCount & " encountered an error!" & vbNewLine
                End Try
            End Using
        Next
        Conn.Close()
        'Throw New NotImplementedException()
        Dim val As String

        If errLog <> "" Then
            lblError.Text = errLog
        End If
        val = DropDownList1.SelectedValue
        lblSuccess.Visible = True
        lblSuccess.Text = val & " Uploaded Successfully!"
    End Sub

    Private Sub insertRBRC()
        Dim wb As XLWorkbook
        Dim insert As String
        Dim rowCount As Integer
        Dim errLog As String

        'save file
        'FileUpload1.SaveAs(Server.MapPath("TempFile"))

        insert = "insert into TBL_ASSAY_RBRC(Lab, DateReceived, TestDate, SampleNumber, SaturatedVolumetricBrineContent, VacuumDryVolumetricBrineContent, RelativeBrineReleaseCapacity, SampleIntegrityQuestionable, AppliedParticleDensity, TechnicianVisualDescription, FileName, DateLoaded) values(@Lab, @DateReceived, @TestDate, @SampleNumber, @SaturatedVolumetricBrineContent, @VacuumDryVolumetricBrineContent, @RelativeBrineReleaseCapacity, @SampleIntegrityQuestionable, @AppliedParticleDensity, @TechnicianVisualDescription, @FileName, @DateLoaded)"
        wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        Conn.Open()
        Dim lastrow = ws.RangeUsed
        errLog = ""

        For Each row In ws.Rows
            rowCount = row.RowNumber
            Using cmd As New SqlCommand
                With cmd
                    .Connection = Conn
                    .CommandType = Data.CommandType.Text
                    .CommandText = insert
                    .Parameters.AddWithValue("@Lab", ws.Cell(rowCount, 1).GetString)
                    .Parameters.AddWithValue("@DateReceived", ws.Cell(rowCount, 2).GetString)
                    .Parameters.AddWithValue("@TestDate", ws.Cell(rowCount, 3).GetString)
                    .Parameters.AddWithValue("@SampleNumber", ws.Cell(rowCount, 4).GetString)
                    .Parameters.AddWithValue("@SaturatedVolumetricBrineContent", ws.Cell(rowCount, 14).GetString)
                    .Parameters.AddWithValue("@VacuumDryVolumetricBrineContent", ws.Cell(rowCount, 15).GetString)
                    .Parameters.AddWithValue("@RelativeBrineReleaseCapacity", ws.Cell(rowCount, 16).GetString)
                    .Parameters.AddWithValue("@SampleIntegrityQuestionable", ws.Cell(rowCount, 8).GetString)
                    .Parameters.AddWithValue("@AppliedParticleDensity", ws.Cell(rowCount, 17).GetString)
                    .Parameters.AddWithValue("@TechnicianVisualDescription", ws.Cell(rowCount, 10).GetString)
                    .Parameters.AddWithValue("@FileName", ws.Cell(rowCount, 11).GetString)
                    .Parameters.AddWithValue("@DateLoaded", ws.Cell(rowCount, 13).GetString)
                End With
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    errLog = errLog & "Row: " & rowCount & " encountered an error!" & vbNewLine
                End Try
            End Using
        Next
        Conn.Close()
        'Throw New NotImplementedException()
        Dim val As String

        If errLog <> "" Then
            lblError.Text = errLog
        End If
        val = DropDownList1.SelectedValue
        lblSuccess.Visible = True
        lblSuccess.Text = val & " Uploaded Successfully!"
    End Sub

    Private Sub insertPSC()
        Dim wb As XLWorkbook
        Dim insert As String
        Dim rowCount As Integer
        Dim errLog As String

        'save file
        'FileUpload1.SaveAs(Server.MapPath("TempFile"))

        insert = "insert into TBL_ASSAY_PARTICLE_SIZE_CHAR(Lab, DateReceived, TestDate, SampleNumber, d10, d50, d60, Cu, Cc, Method, ASTMClassification, USDAClassification, Estimated, FileName, DateLoaded) values(@Lab, @DateReceived, @TestDate, @SampleNumber, @d10, @d50, @d60, @Cu, @Cc, @Method, @ASTMClassification, @USDAClassification, @Estimated, @FileName, @DateLoaded)"
        wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        Conn.Open()
        Dim lastrow = ws.RangeUsed
        errLog = ""

        For Each row In ws.Rows
            rowCount = row.RowNumber
            Using cmd As New SqlCommand
                With cmd
                    .Connection = Conn
                    .CommandType = Data.CommandType.Text
                    .CommandText = insert
                    .Parameters.AddWithValue("@Lab", ws.Cell(rowCount, 1).GetString)
                    .Parameters.AddWithValue("@DateReceived", ws.Cell(rowCount, 2).GetString)
                    .Parameters.AddWithValue("@TestDate", ws.Cell(rowCount, 3).GetString)
                    .Parameters.AddWithValue("@SampleNumber", ws.Cell(rowCount, 4).GetString)
                    .Parameters.AddWithValue("@d10", ws.Cell(rowCount, 17).GetString)
                    .Parameters.AddWithValue("@d50", ws.Cell(rowCount, 18).GetString)
                    .Parameters.AddWithValue("@d60", ws.Cell(rowCount, 19).GetString)
                    .Parameters.AddWithValue("@Cu", ws.Cell(rowCount, 20).GetString)
                    .Parameters.AddWithValue("@Cc", ws.Cell(rowCount, 21).GetString)
                    .Parameters.AddWithValue("@Method", ws.Cell(rowCount, 10).GetString)
                    .Parameters.AddWithValue("@ASTMClassification", ws.Cell(rowCount, 11).GetString)
                    .Parameters.AddWithValue("@USDAClassification", ws.Cell(rowCount, 12).GetString)
                    .Parameters.AddWithValue("@Estimated", ws.Cell(rowCount, 13).GetString)
                    .Parameters.AddWithValue("@FileName", ws.Cell(rowCount, 14).GetString)
                    .Parameters.AddWithValue("@DateLoaded", ws.Cell(rowCount, 16).GetString)
                End With
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    errLog = errLog & "Row: " & rowCount & " encountered an error!" & vbNewLine
                End Try
            End Using
        Next
        Conn.Close()
        'Throw New NotImplementedException()
        Dim val As String

        If errLog <> "" Then
            lblError.Text = errLog
        End If
        val = DropDownList1.SelectedValue
        lblSuccess.Visible = True
        lblSuccess.Text = val & " Uploaded Successfully!"
    End Sub

    Private Sub insertGSSC()
        Dim wb As XLWorkbook
        Dim insert As String
        Dim rowCount As Integer
        Dim errLog As String

        'save file
        'FileUpload1.SaveAs(Server.MapPath("TempFile"))

        insert = "insert into TBL_ASSAY_GRAVELSSC(Lab, DateReceived, TestDate, SampleNumber, Gravel, Sand, SiltAndClay, FileName, DateLoaded) values(@Lab, @DateReceived, @TestDate, @SampleNumber, @Gravel, @Sand, @SiltAndClay, @FileName, @DateLoaded)"
        wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        Conn.Open()
        Dim lastrow = ws.RangeUsed
        errLog = ""
        For Each row In ws.Rows
            rowCount = row.RowNumber
            Using cmd As New SqlCommand
                With cmd
                    .Connection = Conn
                    .CommandType = Data.CommandType.Text
                    .CommandText = insert
                    .Parameters.AddWithValue("@Lab", ws.Cell(rowCount, 1).GetString)
                    .Parameters.AddWithValue("@DateReceived", ws.Cell(rowCount, 2).GetString)
                    .Parameters.AddWithValue("@TestDate", ws.Cell(rowCount, 3).GetString)
                    .Parameters.AddWithValue("@SampleNumber", ws.Cell(rowCount, 4).GetString)
                    .Parameters.AddWithValue("@Gravel", ws.Cell(rowCount, 11).GetString)
                    .Parameters.AddWithValue("@Sand", ws.Cell(rowCount, 12).GetString)
                    .Parameters.AddWithValue("@SiltAndClay", ws.Cell(rowCount, 13).GetString)
                    .Parameters.AddWithValue("@FileName", ws.Cell(rowCount, 8).GetString)
                    .Parameters.AddWithValue("@DateLoaded", ws.Cell(rowCount, 10).GetString)
                End With
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    errLog = errLog & "Row: " & rowCount & " encountered an error!" & vbNewLine
                End Try
            End Using
        Next
        Conn.Close()
        'Throw New NotImplementedException()
        Dim val As String

        If errLog <> "" Then
            lblError.Text = errLog
        End If
        val = DropDownList1.SelectedValue
        lblSuccess.Visible = True
        lblSuccess.Text = val & " Uploaded Successfully!"
    End Sub

    Private Sub insertGeo()
        Dim wb As XLWorkbook
        Dim insert As String
        Dim rowCount As Integer
        Dim errLog As String

        'save file
        'FileUpload1.SaveAs(Server.MapPath("TempFile"))

        insert = "insert into TBL_ASSAY_GEOCHEMISTRY(Lab, DateReceived, TestDate, SampleNumber, Li, Ca, Mg, B, Na, K, Ba, Sr, Fe, Mn, Cl, SO4, Alkalinity, CO3, HCO3, Hardness, Density, pH, Conductivity, STD180, FileName, DateLoaded) values(@Lab, @DateReceived, @TestDate, @SampleNumber, @Li, @Ca, @Mg, @B, @Na, @K, @Ba, @Sr, @Fe, @Mn, @Cl, @SO4, @Alkalinity, @CO3, @HCO3, @Hardness, @Density, @pH, @Conductivity, @STD180, @FileName, @DateLoaded)"
        wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        Conn.Open()
        Dim lastrow = ws.RangeUsed
        errLog = ""

        For Each row In ws.Rows
            rowCount = row.RowNumber
            Using cmd As New SqlCommand
                With cmd
                    .Connection = Conn
                    .CommandType = Data.CommandType.Text
                    .CommandText = insert
                    .Parameters.AddWithValue("@Lab", ws.Cell(rowCount, 1).GetString)
                    .Parameters.AddWithValue("@DateReceived", ws.Cell(rowCount, 2).GetString)
                    .Parameters.AddWithValue("@TestDate", ws.Cell(rowCount, 3).GetString)
                    .Parameters.AddWithValue("@SampleNumber", ws.Cell(rowCount, 4).GetString)
                    .Parameters.AddWithValue("@Li", ws.Cell(rowCount, 46).GetString)
                    .Parameters.AddWithValue("@Ca", ws.Cell(rowCount, 47).GetString)
                    .Parameters.AddWithValue("@Mg", ws.Cell(rowCount, 48).GetString)
                    .Parameters.AddWithValue("@B", ws.Cell(rowCount, 49).GetString)
                    .Parameters.AddWithValue("@Na", ws.Cell(rowCount, 50).GetString)
                    .Parameters.AddWithValue("@K", ws.Cell(rowCount, 51).GetString)
                    .Parameters.AddWithValue("@Ba", ws.Cell(rowCount, 52).GetString)
                    .Parameters.AddWithValue("@Sr", ws.Cell(rowCount, 53).GetString)
                    .Parameters.AddWithValue("@Fe", ws.Cell(rowCount, 54).GetString)
                    .Parameters.AddWithValue("@Mn", ws.Cell(rowCount, 55).GetString)
                    .Parameters.AddWithValue("@Cl", ws.Cell(rowCount, 56).GetString)
                    .Parameters.AddWithValue("@SO4", ws.Cell(rowCount, 57).GetString)
                    .Parameters.AddWithValue("@Alkalinity", ws.Cell(rowCount, 58).GetString)
                    .Parameters.AddWithValue("@CO3", ws.Cell(rowCount, 59).GetString)
                    .Parameters.AddWithValue("@HCO3", ws.Cell(rowCount, 60).GetString)
                    .Parameters.AddWithValue("@Hardness", ws.Cell(rowCount, 61).GetString)
                    .Parameters.AddWithValue("@Density", ws.Cell(rowCount, 21).GetString)
                    .Parameters.AddWithValue("@pH", ws.Cell(rowCount, 22).GetString)
                    .Parameters.AddWithValue("@Conductivity", ws.Cell(rowCount, 23).GetString)
                    .Parameters.AddWithValue("@STD180", ws.Cell(rowCount, 62).GetString)
                    .Parameters.AddWithValue("@FileName", ws.Cell(rowCount, 25).GetString)
                    .Parameters.AddWithValue("@DateLoaded", ws.Cell(rowCount, 45).GetString)
                End With
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    errLog = errLog & "Row: " & rowCount & " encountered an error!" & vbNewLine
                End Try
            End Using
        Next
        Conn.Close()
        'Throw New NotImplementedException()
        Dim val As String

        If errLog <> "" Then
            lblError.Text = errLog
        End If
        val = DropDownList1.SelectedValue
        lblSuccess.Visible = True
        lblSuccess.Text = val & " Uploaded Successfully!"
    End Sub

    Private Sub insertAtt()
        Dim wb As XLWorkbook
        Dim insert As String
        Dim rowCount As Integer
        Dim errLog As String

        'save file
        'FileUpload1.SaveAs(Server.MapPath("TempFile"))

        insert = "insert into TBL_ASSAY_ATTERBERG(Lab, DateReceived, TestDate, SampleNumber, LiquidLimit, PlasticLimit, PlasticityIndex, Classification, FileName, DateLoaded) values(@Lab, @DateReceived, @TestDate, @SampleNumber, @LiquidLimit, @PlasticLimit, @PlasticityIndex, @Classification, @FileName, @DateLoaded)"
        wb = New XLWorkbook(FileUpload1.PostedFile.InputStream)
        Dim ws = wb.Worksheet(1)
        Conn.Open()
        Dim lastrow = ws.RangeUsed
        errLog = ""

        For Each row In ws.Rows
            rowCount = row.RowNumber
            Using cmd As New SqlCommand
                With cmd
                    .Connection = Conn
                    .CommandType = Data.CommandType.Text
                    .CommandText = insert
                    .Parameters.AddWithValue("@Lab", ws.Cell(rowCount, 1).GetString)
                    .Parameters.AddWithValue("@DateReceived", ws.Cell(rowCount, 2).GetString)
                    .Parameters.AddWithValue("@TestDate", ws.Cell(rowCount, 3).GetString)
                    .Parameters.AddWithValue("@SampleNumber", ws.Cell(rowCount, 4).GetString)
                    .Parameters.AddWithValue("@LiquidLimit", ws.Cell(rowCount, 12).GetString)
                    .Parameters.AddWithValue("@PlasticLimit", ws.Cell(rowCount, 13).GetString)
                    .Parameters.AddWithValue("@PlasticityIndex", ws.Cell(rowCount, 14).GetString)
                    .Parameters.AddWithValue("@Classification", ws.Cell(rowCount, 8).GetString)
                    .Parameters.AddWithValue("@FileName", ws.Cell(rowCount, 9).GetString)
                    .Parameters.AddWithValue("@DateLoaded", ws.Cell(rowCount, 11).GetString)
                End With
                Try
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    errLog = errLog & "Row: " & rowCount & " encountered an error!" & vbNewLine
                End Try
            End Using
        Next
        Conn.Close()
        'Throw New NotImplementedException()
        Dim val As String
        If errLog <> "" Then
            lblError.Text = errLog
        End If
        val = DropDownList1.SelectedValue
        lblSuccess.Visible = True
        lblSuccess.Text = val & " Uploaded Successfully!"
    End Sub


End Class