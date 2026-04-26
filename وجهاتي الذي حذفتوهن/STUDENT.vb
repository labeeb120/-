Imports System
Imports System.Data
Imports System.Data.OleDb


Public Class STUDENT
    Public Sub ClearFormControls(frm As Form)
        For Each ctrl As Control In frm.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.Text = ""
            ElseIf TypeOf ctrl Is ComboBox Then
                CType(ctrl, ComboBox).SelectedIndex = -1
            ElseIf TypeOf ctrl Is DateTimePicker Then
                CType(ctrl, DateTimePicker).Value = Date.Now
            End If

            If ctrl.HasChildren Then
                For Each child As Control In ctrl.Controls
                    If TypeOf child Is TextBox Then child.Text = ""
                    If TypeOf child Is ComboBox Then CType(child, ComboBox).SelectedIndex = -1
                Next
            End If
        Next
    End Sub

    Public Const AccessDBPath As String = "C:\Users\ENJAZ\Desktop\مشروع التخرج\الخنساء\الخنساء\users.accdb"

    ' جملة الاتصال
    Public Const ConnectionString As String =
        "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & AccessDBPath & ";Persist Security Info=False;"

    ' إنشاء الاتصال
    Public Conn As New OleDbConnection(ConnectionString)



    Private Sub STUDENT_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadStudyTypes()     ' من جدول StudyTypes
        LoadDepartments()    ' من جدول Specializations أو Departments
        LoadLevels()         ' من جدول Levels
        Loadstayis()      ' قائمة ثابتة (الفصول)
        LoadStudents()  ' قائمة ثابتة (حالة الطالب)
        TextBox1.ReadOnly = True

    End Sub

    Private Sub LoadStudyTypes()
        Try
            Dim dt As New DataTable()
            Dim cmd As New OleDbCommand("SELECT StudyTypeID, StudyTypeName FROM StudyTypes", Conn)
            Dim da As New OleDbDataAdapter(cmd)

            Conn.Open() ' فتح الاتصال يدوياً
            da.Fill(dt)
            Conn.Close() ' إغلاق الاتصال بعد الجلب

            ComboBox1.DataSource = dt
            ComboBox1.DisplayMember = "StudyTypeName" ' ما يظهر للمستخدم
            ComboBox1.ValueMember = "StudyTypeID"     ' القيمة البرمجية (ID)
            ComboBox1.SelectedIndex = -1 ' جعل القائمة فارغة في البداية
        Catch ex As Exception
            If Conn.State = ConnectionState.Open Then Conn.Close()
            MsgBox("خطأ في جلب أنواع الدراسة: " & ex.Message)
        End Try
    End Sub
    Private Sub LoadDepartments()
        Try
            Dim dt As New DataTable()
            ' تأكدي من اسم الجدول والحقول في Access
            Dim cmd As New OleDbCommand("SELECT SpecializationID, SpecializationName FROM Specializations", Conn)
            Dim da As New OleDbDataAdapter(cmd)

            Conn.Open()
            da.Fill(dt)
            Conn.Close()

            ComboBox3.DataSource = dt
            ComboBox3.DisplayMember = "DepartmentName"
            ComboBox3.ValueMember = "DepartmentID"
            ComboBox3.SelectedIndex = -1
        Catch ex As Exception
            If Conn.State = ConnectionState.Open Then Conn.Close()
            MsgBox("خطأ في جلب التخصصات: " & ex.Message)
        End Try
    End Sub
    Private Sub LoadLevels()
        Try
            Dim dt As New DataTable()
            Dim cmd As New OleDbCommand("SELECT LevelID, LevelName FROM Levels", Conn)
            Dim da As New OleDbDataAdapter(cmd)

            Conn.Open()
            da.Fill(dt)
            Conn.Close()

            ComboBox5.DataSource = dt
            ComboBox5.DisplayMember = "LevelName"
            ComboBox5.ValueMember = "LevelID"
            ComboBox5.SelectedIndex = -1
        Catch ex As Exception
            If Conn.State = ConnectionState.Open Then Conn.Close()
            MsgBox("خطأ في جلب المستويات: " & ex.Message)
        End Try
    End Sub
    Private Sub Loadstayis()
        ComboBox2.Items.Clear()
        ComboBox2.Items.Add("الفصل الأول")
        ComboBox2.Items.Add("الفصل الثاني")
        ComboBox2.Items.Add("الفصل الثالث")
        ComboBox2.Items.Add("الفصل الرابع")
        ComboBox2.SelectedIndex = -1
    End Sub
    Private Sub LoadStudents()
        ComboBox4.Items.Clear()
        ComboBox4.Items.Add("نشط")
        ComboBox4.Items.Add("متوقف")
        ComboBox4.Items.Add("متخرج")
        ComboBox4.Items.Add("منقول")
        ComboBox4.SelectedIndex = -1
    End Sub


    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click


        TextBox1.Clear() ' السنة الدراسية
        TextBox2.Clear() ' مبلغ الرسوم
        TextBox3.Clear() ' المبلغ المدفوع
        TextBox8.Clear() ' الملاحضات
        TextBox10.Clear() ' اسم الطالبة
        TextBox6.Clear() ' هاتف الطالبة
        TextBox7.Clear() ' الملاحظات
        ComboBox1.SelectedIndex = -1 ' نوع الدراسة
        ComboBox2.SelectedIndex = -1 ' الفصل الدراسي
        ComboBox3.SelectedIndex = -1 ' التخصص
        ComboBox4.SelectedIndex = -1 ' المستوى
        ComboBox5.SelectedIndex = -1 ' حالة الطالبة
        DateTimePicker1.Value = Date.Now ' تاريخ الالتحاق
        DateTimePicker2.Value = Date.Now
    End Sub


    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click

        If String.IsNullOrWhiteSpace(TextBox2.Text) Then
            MsgBox("الرجاء إدخال اسم الطالبة.")
            Return
        End If

        Try
            ' استعلام الإدخال - تأكدي من ترتيب الحقول وعددها
            ' ملاحظة: الحقول التي تحتوي على مسافات مثل [Student Status] يجب وضعها بين أقواس مربعة
            Dim sql As String = "INSERT INTO Students (StudentName, Phone, [Department ID], LevelID, EnrollmentDate, StudyTypeID, [Student Status], AcademicYear, [chapter NO], Notes, FeeAmount, PaidAmount, PaymentDate) " &
                               "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

            Dim cmd As New OleDbCommand(sql, Conn)

            ' إضافة البارامترات بنفس ترتيب علامات الاستفهام (?)
            cmd.Parameters.AddWithValue("@name", TextBox2.Text)        ' اسم الطالبة
            cmd.Parameters.AddWithValue("@phone", TextBox6.Text)       ' هاتف الطالبة
            cmd.Parameters.AddWithValue("@dept", ComboBox3.SelectedValue) ' رقم التخصص
            cmd.Parameters.AddWithValue("@lvl", ComboBox5.SelectedValue)  ' رقم المستوى
            cmd.Parameters.AddWithValue("@enroll", DateTimePicker2.Value) ' تاريخ الالتحاق
            cmd.Parameters.AddWithValue("@study", ComboBox1.SelectedValue) ' رقم نوع الدراسة
            cmd.Parameters.AddWithValue("@status", ComboBox4.SelectedItem.ToString()) ' حالة الطالبة
            cmd.Parameters.AddWithValue("@year", TextBox10.Text)        ' السنة الدراسية (تأكدي من رقم التكست بوكس)
            cmd.Parameters.AddWithValue("@chap", ComboBox2.SelectedItem.ToString()) ' الفصل
            cmd.Parameters.AddWithValue("@note", TextBox8.Text)        ' الملاحظات
            cmd.Parameters.AddWithValue("@fee", If(IsNumeric(TextBox3.Text), CDec(TextBox3.Text), 0)) ' مبلغ الرسوم
            cmd.Parameters.AddWithValue("@paid", If(IsNumeric(TextBox7.Text), CDec(TextBox7.Text), 0)) ' المبلغ المدفوع
            cmd.Parameters.AddWithValue("@pdate", DateTimePicker1.Value) ' تاريخ الدفع

            Conn.Open()
            cmd.ExecuteNonQuery()
            Conn.Close()

            MsgBox("تم الحفظ بنجاح")
            ClearFormControls(Me) ' مسح الحقول بعد الحفظ
            Dim f As New DATE_STUDENT()
            f.Show()
            f.LoadStudents()


        Catch ex As Exception
            If Conn.State = ConnectionState.Open Then Conn.Close()
            MsgBox("خطأ أثناء الحفظ: " & ex.Message)
        End Try
    End Sub


    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        Dim f As New DATE_STUDENT()
        f.Show()

    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        Dim f As New DATE_STUDENT
        f.ShowDialog()
        Dim sql As String = "INSERT INTO Students (StudentName, Phone, [Department ID], LevelID, EnrollmentDate, StudyTypeID, [Student Status], AcademicYear, [chapter NO], Notes, FeeAmount, PaidAmount, PaymentDate) " &
                      "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

        Dim parameters As New List(Of OleDbParameter) From {
            New OleDbParameter With {.Value = TextBox2.Text.Trim()},
            New OleDbParameter With {.Value = TextBox6.Text.Trim()},
            New OleDbParameter With {.Value = If(ComboBox3.SelectedValue, DBNull.Value)},
            New OleDbParameter With {.Value = If(ComboBox5.SelectedValue, DBNull.Value)},
            New OleDbParameter With {.Value = DateTimePicker2.Value},
            New OleDbParameter With {.Value = If(ComboBox1.SelectedValue, DBNull.Value)},
            New OleDbParameter With {.Value = If(ComboBox4.SelectedItem IsNot Nothing, ComboBox4.SelectedItem.ToString(), "")},
            New OleDbParameter With {.Value = TextBox10.Text.Trim()},
            New OleDbParameter With {.Value = If(ComboBox2.SelectedItem IsNot Nothing, ComboBox2.SelectedItem.ToString(), "")},
            New OleDbParameter With {.Value = TextBox8.Text.Trim()},
            New OleDbParameter With {.Value = If(IsNumeric(TextBox3.Text), CDec(TextBox3.Text), 0)},
            New OleDbParameter With {.Value = If(IsNumeric(TextBox7.Text), CDec(TextBox7.Text), 0)},
            New OleDbParameter With {.Value = DateTimePicker1.Value}
        }

        Dim result = DataAccessHelper.ExecuteNonQuery(sql, parameters)

        If result > 0 Then
            MsgBox("تمت الإضافة بنجاح ✅")
            Me.Close()
        Else
            MsgBox("فشل في الإضافة ❌")
        End If
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Application.Exit()
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click


        If SelectedID = 0 Then
            MsgBox("اختار سجل أولاً")
            Exit Sub
        End If

        Dim sql As String = "UPDATE Students SET StudentName=?, Phone=?, [Department ID]=?, LevelID=?, EnrollmentDate=?, StudyTypeID=?, [Student Status]=?, AcademicYear=?, [chapter NO]=?, Notes=?, FeeAmount=?, PaidAmount=?, PaymentDate=? WHERE StudentID=?"

        Dim parameters As New List(Of OleDbParameter) From {
        New OleDbParameter With {.Value = TextBox2.Text.Trim()},
        New OleDbParameter With {.Value = TextBox6.Text.Trim()},
        New OleDbParameter With {.Value = If(ComboBox3.SelectedValue, DBNull.Value)},
        New OleDbParameter With {.Value = If(ComboBox5.SelectedValue, DBNull.Value)},
        New OleDbParameter With {.Value = DateTimePicker2.Value},
        New OleDbParameter With {.Value = If(ComboBox1.SelectedValue, DBNull.Value)},
        New OleDbParameter With {.Value = If(ComboBox4.SelectedItem IsNot Nothing, ComboBox4.SelectedItem.ToString(), "")},
        New OleDbParameter With {.Value = TextBox10.Text.Trim()},
        New OleDbParameter With {.Value = If(ComboBox2.SelectedItem IsNot Nothing, ComboBox2.SelectedItem.ToString(), "")},
        New OleDbParameter With {.Value = TextBox8.Text.Trim()},
        New OleDbParameter With {.Value = If(IsNumeric(TextBox3.Text), CDec(TextBox3.Text), 0)},
        New OleDbParameter With {.Value = If(IsNumeric(TextBox7.Text), CDec(TextBox7.Text), 0)},
        New OleDbParameter With {.Value = DateTimePicker1.Value},
        New OleDbParameter With {.Value = SelectedID} ' 🔥 مهم جدًا (WHERE)
    }

        Dim result = DataAccessHelper.ExecuteNonQuery(sql, parameters)

        If result > 0 Then
            MsgBox("تم التعديل بنجاح ✅")
            LoadData() ' تحديث الجدول
        Else
            MsgBox("فشل التعديل ❌")
        End If

    End Sub

End Class