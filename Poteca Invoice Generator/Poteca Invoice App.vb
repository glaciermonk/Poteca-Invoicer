'Database name POTECA_INVOICE
'Table Name invoice
'ODBC name POTECA_INVOICE_APP_ODBC_BRIDGE
'user root
'password root

'create database POTECA_INVOICE;
'use POTECA_INVOICE;
'create table items_info(serial int(3), item_desc varchar(255), qty varchar(7), uom varchar(7), price varchar(7), tax varchar(7), discount varchar(7), amount varchar(7), PRIMARY KEY(serial));
'create table client_info(btline1 varchar(255), btline2 varchar(255),  btline3 varchar(255),  btline4 varchar(255), stline1 varchar(255), stline2 varchar(255), issue_date DATE, due_date DATE, subtotal varchar(7), total_taxable_value varchar(7), total varchar(7));
'commit;

'mysql> desc client_info;
'+---------------------+--------------+------+-----+---------+-------+
'| Field               | Type         | Null | Key | Default | Extra |
'+---------------------+--------------+------+-----+---------+-------+
'| btline1             | varchar(255) | YES  |     | NULL    |       |
'| btline2             | varchar(255) | YES  |     | NULL    |       |
'| btline3             | varchar(255) | YES  |     | NULL    |       |
'| btline4             | varchar(255) | YES  |     | NULL    |       |
'| stline1             | varchar(255) | YES  |     | NULL    |       |
'| stline2             | varchar(255) | YES  |     | NULL    |       |
'| issue_date          | date         | YES  |     | NULL    |       |
'| due_date            | date         | YES  |     | NULL    |       |
'| subtotal            | varchar(7)   | YES  |     | NULL    |       |
'| total_taxable_value | varchar(7)   | YES  |     | NULL    |       |
'| total               | varchar(7)   | YES  |     | NULL    |       |
'+---------------------+--------------+------+-----+---------+-------+
'11 rows in set (0.01 sec)

'mysql> desc items_info;
'+-----------+--------------+------+-----+---------+-------+
'| Field     | Type         | Null | Key | Default | Extra |
'+-----------+--------------+------+-----+---------+-------+
'| serial    | int(3)       | NO   | PRI | 0       |       |
'| item_desc | varchar(255) | YES  |     | NULL    |       |
'| qty       | varchar(7)   | YES  |     | NULL    |       |
'| uom       | varchar(7)   | YES  |     | NULL    |       |
'| price     | varchar(7)   | YES  |     | NULL    |       |
'| tax       | varchar(7)   | YES  |     | NULL    |       |
'| discount  | varchar(7)   | YES  |     | NULL    |       |
'| amount    | varchar(7)   | YES  |     | NULL    |       |
'+-----------+--------------+------+-----+---------+-------+
'8 rows in set (0.01 sec)


Imports System.Data.Odbc
Public Class Form1
    Dim SERIAL, CLEAR_PREVIOUS_RECORD, ENTRY_COUNT As Integer 'serial number of the items
    Dim afterqtyprice, afterqtytax, afterqtydiscount, afterqtyamount As Decimal
    Dim subtotal, total_taxable_value, total As Decimal
    'ENTRY_COUNT IS TO LIMIT ENTRY INPUTS TO PREVENT CRYSTAL REPORT FROM BEING OFF SIZE


    'print button
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PRINT.Click
        'if any textbox is left empty, MYSQL will throw an error so to prevent that empty textboxes are filled with an empty space

        If btline1.Text = "" Then
            btline1.Text = " "
        End If
        If btline2.Text = "" Then
            btline2.Text = " "
        End If
        If btline3.Text = "" Then
            btline3.Text = " "
        End If
        If btline4.Text = "" Then
            btline4.Text = " "
        End If
        If stline1.Text = "" Then
            stline1.Text = " "
        End If
        If stline2.Text = "" Then
            stline2.Text = " "
        End If

        Dim command As New OdbcCommand
        Dim connection As New OdbcConnection
        connection = New OdbcConnection("dsn=POTECA_INVOICE_APP_ODBC_BRIDGE;user=root;pwd=root")
        connection.Open()

        'assuming the user selects date and time at the end after entering all item infos
        command = New OdbcCommand("update client_info set issue_date=' " & Format(DateTimePicker1.Value, "yyyy-MM-dd") & " ',due_date=' " & Format(DateTimePicker2.Value, "yyyy-MM-dd") & " 'where serial=' " & SERIAL.ToString & " '", connection)
        connection.Close()
        Invoice_Print_Preview.Show()
    End Sub


    'add button
    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If ENTRY_COUNT > 11 Then
            MsgBox("Sorry, you can only enter 12 at once!", MsgBoxStyle.Information)
        Else
            Dim connection As New OdbcConnection
            Dim command As New OdbcCommand
            ENTRY_COUNT = ENTRY_COUNT + 1
            afterqtyprice = Convert.ToDecimal(price.Text) * Convert.ToDecimal(qty.Text) 'calculate price according to quantity
            subtotal = subtotal + afterqtyprice
            afterqtytax = Convert.ToDecimal(tax.Text) * Convert.ToDecimal(qty.Text) 'calculate tax according to quantity
            total_taxable_value = total_taxable_value + afterqtytax
            afterqtydiscount = Convert.ToDecimal(discount.Text) * Convert.ToDecimal(qty.Text) 'calculate discount according to quantity
            afterqtyamount = (afterqtyprice + afterqtytax) - afterqtydiscount
            total = total + afterqtyamount
            SERIAL = SERIAL + 1 'serial number of the items

            connection = New OdbcConnection("dsn=POTECA_INVOICE_APP_ODBC_BRIDGE;user=root;pwd=root")
            connection.Open()
            If CLEAR_PREVIOUS_RECORD = 0 Then '0 meaning true
                command = New OdbcCommand("delete from items_info", connection)
                command.ExecuteNonQuery()
                command = New OdbcCommand("delete from client_info", connection)
                command.ExecuteNonQuery()
                CLEAR_PREVIOUS_RECORD = +1
            End If

            command = New OdbcCommand("insert into items_info values(' " & SERIAL.ToString & " ',' " & item_desc.Text & " ',' " & qty.Text & " ',' " & uom.Text & " ',' " & afterqtyprice.ToString & " ',' " & afterqtytax.ToString & " ',' " & afterqtydiscount.ToString & " ',' " & afterqtyamount.ToString & " ')", connection)
            command.ExecuteNonQuery()

            command = New OdbcCommand("delete from client_info", connection)
            command.ExecuteNonQuery()
            command = New OdbcCommand("insert into client_info values(' " & btline1.Text & " ',' " & btline2.Text & " ',' " & btline3.Text & " ',' " & btline4.Text & " ',' " & stline1.Text & " ',' " & stline2.Text & " ', ' " & Format(DateTimePicker1.Value, "yyyy-MM-dd") & " ', ' " & Format(DateTimePicker2.Value, "yyyy-MM-dd") & " ',' " & subtotal.ToString & " ',' " & total_taxable_value.ToString & " ',' " & total.ToString & " ')", connection)
            command.ExecuteNonQuery()

            'display data in datagridviewer
            Dim adapter As New OdbcDataAdapter
            Dim data_table As New DataTable
            connection = New OdbcConnection("dsn=POTECA_INVOICE_APP_ODBC_BRIDGE;user=root;pwd=root")
            adapter = New OdbcDataAdapter("select * from items_info", connection)
            adapter.Fill(data_table)
            DataGridView1.DataSource = data_table

            connection.Close()
        End If
    End Sub


    'clear all data from sql tables
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ENTRY_COUNT = 0
        Dim connection As New OdbcConnection
        Dim command As New OdbcCommand
        SERIAL = 0
        connection = New OdbcConnection("dsn=POTECA_INVOICE_APP_ODBC_BRIDGE;user=root;pwd=root")
        connection.Open()
        command = New OdbcCommand("delete from items_info", connection)
        command.ExecuteNonQuery()
        command = New OdbcCommand("delete from client_info", connection)
        command.ExecuteNonQuery()
        connection.Close()

        'reload data into datagridviewer
        Dim adapter As New OdbcDataAdapter
        Dim data_table As New DataTable
        connection = New OdbcConnection("dsn=POTECA_INVOICE_APP_ODBC_BRIDGE;user=root;pwd=root")
        adapter = New OdbcDataAdapter("select * from items_info", connection)
        adapter.Fill(data_table)
        DataGridView1.DataSource = data_table

        afterqtyprice = 0
        afterqtytax = 0
        afterqtydiscount = 0
        afterqtyamount = 0
        subtotal = 0
        total_taxable_value = 0
        total = 0
    End Sub


    'clear the whole form
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        btline1.Clear()
        btline2.Clear()
        btline3.Clear()
        btline4.Clear()
        item_desc.Clear()
        qty.Clear()
        uom.Clear()
        price.Clear()
        discount.Clear()
        tax.Clear()
        stline1.Clear()
        stline2.Clear()
        btline1.Focus()
    End Sub


    'address preset using combobox
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        If ComboBox1.Text = 1 Then
            btline1.Text = "Six Mile, VIP Road Guwahati, Assam, Zip 781022, India, Guwahati,"
            btline2.Text = "AS (18) 781022, IN"
            btline3.Text = "+91 8876 052 074"
            btline4.Text = "hello@potecaservices.com"
            qty.Text = "1"
            item_desc.Text = "1"
            uom.Text = "1"
            price.Text = "1"
            tax.Text = "1"
            discount.Text = "1"
            stline1.Text = "1"
            stline2.Text = "1"
        End If
    End Sub


    'clear button
    Private Sub BTN_CLEAR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BTN_CLEAR.Click
        item_desc.Clear()
        qty.Clear()
        uom.Clear()
        price.Clear()
        discount.Clear()
        tax.Clear()
        stline1.Clear()
        stline2.Clear()
        ComboBox1.Items.Clear()
        item_desc.Focus()
    End Sub


    'selecting issue date in datetimepicker automatically sets due date/datetimepicker2 to 15 days apart
    Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged
        Dim currentDate, currentDate15 As DateTime
        currentDate = DateTimePicker1.Value()
        currentDate15 = DateTimePicker1.Value.AddDays(15)
        currentDate15.AddDays(15)
        DateTimePicker2.Value() = currentDate15
        DateTimePicker1.Value() = currentDate
    End Sub


End Class