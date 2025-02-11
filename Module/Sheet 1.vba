Sub UploadToSQLServer()
    Dim conn As Object
    Dim cmd As Object
    Dim sql As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim cell As Range
    Dim hasError As Boolean
    Dim isValid As Boolean
    Dim dataUploaded As Boolean ' Flag untuk menandakan apakah ada data yang berhasil diupload

    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("DBM_QualityMonitoring_Daily") ' Ganti dengan nama sheet Anda

    ' Cek apakah ada sel yang berisi "error"
    hasError = False
    For Each cell In ws.UsedRange
        If LCase(cell.Value) = "error" Then
            hasError = True
            Exit For
        End If
    Next cell

    If hasError Then
        MsgBox "Terdapat kesalahan dalam data: salah satu sel berisi 'error'.", vbExclamation
        Exit Sub
    End If

    ' Buat koneksi ke SQL Server
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionString = "DSN=.........;UID=.........;PWD=.........;DATABASE=.........;LANGUAGE=us_english;"
    
    On Error GoTo ErrorHandler
    conn.Open

    ' Dapatkan baris terakhir yang diisi
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row ' Asumsikan data mulai dari kolom A

    ' Loop melalui setiap baris dan masukkan data ke SQL Server
    For i = 2 To lastRow ' Mulai dari baris kedua untuk menghindari header
        ' Validasi kolom 1-5 dan 9-10
        isValid = True
        For j = 1 To 5
            If IsEmpty(ws.Cells(i, j).Value) Then
                isValid = False
                Exit For
            End If
        Next j
        
        If isValid Then
            For j = 9 To 10
                If IsEmpty(ws.Cells(i, j).Value) Then
                    isValid = False
                    Exit For
                End If
            Next j
        End If

        If Not isValid Then
            MsgBox "Data pada baris " & i & " tidak valid: Kolom 1-5 dan 9-10 tidak boleh kosong.", vbExclamation
            GoTo NextIteration
        End If

        sql = "INSERT INTO DBM_QualityMonitoring_Daily (Date, Departemen, Group_criteria_qm, Criteria_qm, Value_type, Value_number, Percentage, Time, Insert_Date, Insert_By) VALUES ("

        ' Tambahkan nilai dengan memeriksa apakah sel kosong
        For j = 1 To 10 ' Ganti dengan jumlah kolom yang sesuai
            If IsEmpty(ws.Cells(i, j).Value) Then
                sql = sql & "NULL"
            Else
                sql = sql & "'" & ws.Cells(i, j).Value & "'"
            End If
            
            ' Tambahkan koma jika bukan kolom terakhir
            If j < 10 Then
                sql = sql & ", "
            End If
        Next j

        sql = sql & ");" ' Akhiri pernyataan SQL

        Set cmd = CreateObject("ADODB.Command")
        cmd.ActiveConnection = conn
        cmd.CommandText = sql
        cmd.Execute
        
        ' Reset nilai sel yang telah diproses, kecuali kolom 1, 5, dan 9
        For j = 1 To 10 ' Ganti dengan jumlah kolom yang sesuai
            If j <> 1 And j <> 5 And j <> 9 Then ' Jika kolom bukan kolom 1, 5, atau 9
                ws.Cells(i, j).Value = "" ' Mengosongkan nilai tetapi mempertahankan picklist
            End If
        Next j
        
        dataUploaded = True ' Tandai bahwa ada data yang berhasil diupload

NextIteration:
    Next i

    If dataUploaded Then
        MsgBox "Data berhasil diupload ke SQL Server!"
    Else
        MsgBox "Tidak ada data yang berhasil diupload."
    End If

    ' Tutup koneksi
    conn.Close
    Set conn = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Terjadi kesalahan: " & Err.Description
    If Not conn Is Nothing Then conn.Close
End Sub
