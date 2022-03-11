Imports System.Data
Imports System.Data.OleDb
Imports System.Data.DataSet
Imports System.Text.RegularExpressions
Public Class Form1
    Dim con As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\Study\GUI Application Development using VB.net\Project\CO_Sem_IV.accdb")
    Dim con1 As New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=E:\Study\GUI Application Development using VB.net\Project\Clear.accdb")
    Dim cmd As OleDbCommand
    Dim dapt As OleDbDataAdapter
    Dim dread As OleDbDataReader
    Dim dr As DataRow
    Dim ds As New DataSet
    Dim cmb1, cmb2 As Integer
    Dim crc, sem, gndr As String
    Dim icount, total, avrg As Integer
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        cmb1 = ComboBox1.SelectedIndex
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        cmb2 = ComboBox2.SelectedIndex
        If cmb1 = 0 And cmb2 = 3 Then
            Sub1.Visible = True
            Sub1.Text = "GAD :"
            sub2.Visible = True
            sub2.Text = "JPR :"
            sub3.Visible = True
            sub3.Text = "SEN :"
            sub4.Visible = True
            sub4.Text = "DCC :"
            sub5.Visible = True
            sub5.Text = "MIC :"
            Tot.Visible = True
            Avg.Visible = True
            TSub1.Visible = True
            TSub2.Visible = True
            Tsub3.Visible = True
            Tsub4.Visible = True
            Tsub5.Visible = True
            Ttot.Visible = True
            Tavg.Visible = True
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        LoginForm.Close()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'TODO: This line of code loads data into the 'CO_Sem_IVDataSet.Comp_sem4' table. You can move, or remove it, as needed.
        Me.Comp_sem4TableAdapter.Fill(Me.CO_Sem_IVDataSet.Comp_sem4)
        con.Open()
        cmd = New OleDbCommand("Select * from Comp_sem4 ", con)
        dapt = New OleDbDataAdapter(cmd)
        dapt.Fill(ds, "Comp_sem4")
        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "Comp_sem4"
    End Sub

    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        End
        LoginForm.Close()
    End Sub

    Private Sub savebutton_Click(sender As Object, e As EventArgs) Handles savebutton.Click
        dr = ds.Tables("Comp_sem4").NewRow
        dr("Last_Name") = TLName.Text
        dr("First_Name") = TFName.Text
        dr("Middle_Name") = TMName.Text
        If Not Regex.IsMatch(TDOB.Text, "^[0-3]{1}[0-9]{1}[\/\-][0-9]{2}[\/\-][0-9]{4}") Or TDOB.Text = "" Then
            ErrorProvider1.SetError(TDOB, "Invalid Date")
            TDOB.Text = ""
            TDOB.Focus()
        Else
            TDOB.Focus()
            MsgBox("Validation Done.......")
        End If
        dr("DOB") = CDate(TDOB.Text)
        dr("Address") = TAddr.Text
        Select Case ComboBox3.SelectedIndex
            Case 0
                dr("Gender") = "Male"
                gndr = "Male"
            Case 1
                dr("Gender") = "Female"
                gndr = "Female"
        End Select

        dr("Roll_no") = CInt(TRNO.Text)
        dr("Enrollment_no") = CInt(TENO.Text)
        dr("Exam_seat_no") = CInt(TEXM.Text)
        Select Case ComboBox1.SelectedIndex
            Case 0
                dr("Cource") = "Computer Engineering"
                crc = "Computer Engineering"
            Case 1
                dr("Cource") = "Mechanical Engineering"
                crc = "Mechanical Engineering"
            Case 2
                dr("Cource") = "Civil Engineering"
                crc = "Civil Engineering"
            Case 3
                dr("Cource") = "Electronics Engineering"
                crc = "Electronics Engineering"
            Case 4
                dr("Cource") = "Electrical Engineering"
                crc = "Electrical Engineering"
            Case 5
                dr("Cource") = "Automobile Engineering"
                crc = "Automobile Engineering"
        End Select
        Select Case ComboBox2.SelectedIndex
            Case 0
                dr("Semester") = 1
            Case 1
                dr("Semester") = 2
            Case 2
                dr("Semester") = 3
            Case 3
                dr("Semester") = 4
            Case 4
                dr("Semester") = 5
            Case 5
                dr("Semester") = 6
        End Select
        dr("JPR") = CInt(TSub1.Text)
        dr("MIC") = CInt(TSub2.Text)
        dr("SEN") = CInt(Tsub3.Text)
        dr("DCC") = CInt(Tsub4.Text)
        dr("GAD") = CInt(Tsub5.Text)
        total = CInt(TSub1.Text) + CInt(TSub2.Text) + CInt(Tsub3.Text) + CInt(Tsub4.Text) + CInt(Tsub5.Text)
        dr("Total") = total
        avrg = (CInt(TSub1.Text) + CInt(TSub2.Text) + CInt(Tsub3.Text) + CInt(Tsub4.Text) + CInt(Tsub5.Text)) / 5
        dr("Average") = avrg
        ds.Tables("Comp_sem4").Rows.Add(dr)

        cmd = New OleDbCommand("insert into Comp_sem4 values( " & CInt(TRNO.Text) & ",'" & TFName.Text & "','" & TMName.Text & "','" & TLName.Text & "','" & TAddr.Text & "','" & TDOB.Text & "','" & gndr & "'," & CInt(TENO.Text) & ",'" & crc & "'," & cmb1 & "," & CInt(TEXM.Text) & "," & CInt(TSub1.Text) & "," & CInt(TSub2.Text) & "," & CInt(Tsub3.Text) & "," & CInt(Tsub4.Text) & "," & CInt(Tsub5.Text) & "," & CInt(total) & "," & CInt(avrg) & ")", con)
        icount = cmd.ExecuteNonQuery
        MsgBox(icount & " Record Added Successfully")
    End Sub

    Private Sub SerachButton_Click(sender As Object, e As EventArgs) Handles SerachButton.Click
        cmd = New OleDbCommand("Select * from Comp_sem4 where Roll_no=" & CInt(TSerach.Text), con)
        dread = cmd.ExecuteReader
        Sub1.Visible = True
        Sub1.Text = "GAD :"
        sub2.Visible = True
        sub2.Text = "JPR :"
        sub3.Visible = True
        sub3.Text = "SEN :"
        sub4.Visible = True
        sub4.Text = "DCC :"
        sub5.Visible = True
        sub5.Text = "MIC :"
        Tot.Visible = True
        Avg.Visible = True
        TSub1.Visible = True
        TSub2.Visible = True
        Tsub3.Visible = True
        Tsub4.Visible = True
        Tsub5.Visible = True
        Ttot.Visible = True
        Tavg.Visible = True
        If dread.Read() Then
            TRNO.Text = dread(0)
            TFName.Text = dread(1)
            TMName.Text = dread(2)
            TLName.Text = dread(3)
            TAddr.Text = dread(4)
            TDOB.Text = dread(5)
            gndr = dread(6)
            Select Case gndr
                Case "Male"
                    ComboBox3.SelectedIndex = 0
                Case "Female"
                    ComboBox3.SelectedIndex = 1
            End Select
            TENO.Text = dread(7)
            crc = dread(8)
            Select Case crc
                Case "Computer Engineering"
                    ComboBox1.SelectedIndex = 0
                Case "Mechanical Engineering"
                    ComboBox1.SelectedIndex = 1
                Case "Civil Engineering"
                    ComboBox1.SelectedIndex = 2
                Case "Electronics Engineering"
                    ComboBox1.SelectedIndex = 3
                Case "Electrical Engineering"
                    ComboBox1.SelectedIndex = 4
                Case "Automobile Engineering"
                    ComboBox1.SelectedIndex = 5
            End Select
            sem = dread(9)
            Select Case sem
                Case 1
                    ComboBox2.SelectedIndex = 0
                Case 2
                    ComboBox2.SelectedIndex = 1
                Case 3
                    ComboBox2.SelectedIndex = 2
                Case 4
                    ComboBox2.SelectedIndex = 3
                Case 5
                    ComboBox2.SelectedIndex = 4
                Case 6
                    ComboBox2.SelectedIndex = 5
            End Select
            TEXM.Text = dread(10)
            TSub1.Text = dread(11)
            TSub2.Text = dread(12)
            Tsub3.Text = dread(13)
            Tsub4.Text = dread(14)
            Tsub5.Text = dread(15)
            Ttot.Text = dread(16)
            Tavg.Text = dread(17)
        End If
    End Sub


    Private Sub DeleteButton_Click(sender As Object, e As EventArgs) Handles DeleteButton.Click
        Dim del As Integer
        del = CInt(TSerach.Text)
        cmd = New OleDbCommand("delete from Comp_sem4 where Roll_no =" & del, con)
        icount = cmd.ExecuteNonQuery
        MsgBox(icount & " Record Deleted ")
        TRNO.Text = ""
        TFName.Text = ""
        TMName.Text = ""
        TLName.Text = ""
        TAddr.Text = ""
        TDOB.Text = ""
        ComboBox1.Text = "Choose Cource"
        ComboBox2.Text = "Choose Semester"
        ComboBox3.Text = "Choose Gender"
        TENO.Text = ""
        TEXM.Text = ""
        TSub1.Text = ""
        TSub2.Text = ""
        Tsub3.Text = ""
        Tsub4.Text = ""
        Tsub5.Text = ""
        Ttot.Text = ""
        Tavg.Text = ""
    End Sub


    Private Sub UpdateButton_Click(sender As Object, e As EventArgs) Handles UpdateButton.Click
        Select Case ComboBox1.SelectedIndex
            Case 0                
                crc = "Computer Engineering"
            Case 1                
                crc = "Mechanical Engineering"
            Case 2                
                crc = "Civil Engineering"
            Case 3            
                crc = "Electronics Engineering"
            Case 4                
                crc = "Electrical Engineering"
            Case 5
                crc = "Automobile Engineering"
        End Select
        Select Case ComboBox3.SelectedIndex
            Case 0
                gndr = "Male"
            Case 1
                gndr = "Female"
        End Select
        Select Case ComboBox2.SelectedIndex
            Case 0
                sem = 1
            Case 1
                sem = 2
            Case 2
                sem = 3
            Case 3
                sem = 4
            Case 4
                sem = 5
            Case 5
                sem = 6
        End Select
        total = CInt(TSub1.Text) + CInt(TSub2.Text) + CInt(Tsub3.Text) + CInt(Tsub4.Text) + CInt(Tsub5.Text)
        avrg = (CInt(TSub1.Text) + CInt(TSub2.Text) + CInt(Tsub3.Text) + CInt(Tsub4.Text) + CInt(Tsub5.Text)) / 5
        cmd = New OleDbCommand("update Comp_sem4 set Roll_no=" & CInt(TRNO.Text) & ",First_Name='" & TFName.Text & "',Middle_Name='" & TMName.Text & "',Last_Name='" & TLName.Text & "',Address='" & TAddr.Text & "',DOB='" & TDOB.Text & "',Gender='" & gndr & "',Enrollment_no=" & CInt(TENO.Text) & ",Cource='" & crc & "',Semester=" & sem & ",Exam_seat_no=" & CInt(TEXM.Text) & ",JPR=" & CInt(TSub1.Text) & ",MIC=" & CInt(TSub2.Text) & ",SEN=" & CInt(Tsub3.Text) & ",DCC=" & CInt(Tsub4.Text) & ",GAD=" & CInt(Tsub5.Text) & ",Total=" & CInt(Ttot.Text) & ",Average=" & CInt(Tavg.Text) & " where Roll_no=" & TSerach.Text, con)
        icount = cmd.ExecuteNonQuery
        MsgBox(icount & " Record Updated")
    End Sub

    Private Sub Bcloseapp_Click(sender As Object, e As EventArgs) Handles Bcloseapp.Click
        End
        LoginForm.Close()
    End Sub
End Class

