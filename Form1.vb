Imports System
Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports System.Globalization
Imports System.Threading.Tasks
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Windows.Forms
Imports OfficeOpenXml
Imports QRCoder

Public Class Form1
    Inherits Form

    ' Controles
    Private txtExcel As TextBox
    Private btnBrowseExcel As Button
    Private txtOutput As TextBox
    Private btnBrowseOutput As Button
    Private txtQrCols As TextBox
    Private txtTextCols As TextBox
    Private nudLabelW As NumericUpDown
    Private nudLabelH As NumericUpDown
    Private nudQrSize As NumericUpDown
    Private nudDpi As NumericUpDown
    Private nudPerRow As NumericUpDown
    Private nudPerCol As NumericUpDown
    Private chkIndividual As CheckBox
    Private chkSheets As CheckBox
    Private btnPreview As Button
    Private btnGenerate As Button
    Private pbPreview As PictureBox
    Private progressBar As ProgressBar
    Private lblStatus As Label

    Public Sub New()
        Me.Text = "QR Label Maker (VB WinForms)"
        Me.ClientSize = New Size(920, 520)
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        ' Inicializar controles y ubicaciones
        txtExcel = New TextBox() With {.Left = 12, .Top = 12, .Width = 700}
        btnBrowseExcel = New Button() With {.Left = 720, .Top = 10, .Width = 180, .Text = "Seleccionar Excel..."}
        AddHandler btnBrowseExcel.Click, AddressOf BtnBrowseExcel_Click

        txtOutput = New TextBox() With {.Left = 12, .Top = 44, .Width = 700}
        btnBrowseOutput = New Button() With {.Left = 720, .Top = 42, .Width = 180, .Text = "Carpeta salida..."}
        AddHandler btnBrowseOutput.Click, AddressOf BtnBrowseOutput_Click

        Dim lbl1 = New Label() With {.Left = 12, .Top = 80, .Width = 200, .Text = "Columnas para QR (coma-sep):"}
        txtQrCols = New TextBox() With {.Left = 12, .Top = 100, .Width = 430, .Text = "codigo_qr"}

        Dim lbl2 = New Label() With {.Left = 460, .Top = 80, .Width = 200, .Text = "Columnas para texto (coma-sep):"}
        txtTextCols = New TextBox() With {.Left = 460, .Top = 100, .Width = 440, .Text = "nombre,direccion"}

        Dim lbl3 = New Label() With {.Left = 12, .Top = 140, .Width = 220, .Text = "Etiqueta ancho (mm) / alto (mm):"}
        nudLabelW = New NumericUpDown() With {.Left = 12, .Top = 160, .Width = 90, .Minimum = 10, .Maximum = 300, .Value = 70}
        nudLabelH = New NumericUpDown() With {.Left = 110, .Top = 160, .Width = 90, .Minimum = 5, .Maximum = 300, .Value = 25}

        Dim lbl4 = New Label() With {.Left = 220, .Top = 140, .Width = 200, .Text = "Tamaño QR (mm) / DPI:"}
        nudQrSize = New NumericUpDown() With {.Left = 220, .Top = 160, .Width = 90, .Minimum = 5, .Maximum = 200, .Value = 20}
        nudDpi = New NumericUpDown() With {.Left = 320, .Top = 160, .Width = 90, .Minimum = 72, .Maximum = 1200, .Value = 300}

        Dim lbl5 = New Label() With {.Left = 420, .Top = 140, .Width = 200, .Text = "Por fila / Por columna (hoja):"}
        nudPerRow = New NumericUpDown() With {.Left = 420, .Top = 160, .Width = 90, .Minimum = 1, .Maximum = 20, .Value = 3}
        nudPerCol = New NumericUpDown() With {.Left = 520, .Top = 160, .Width = 90, .Minimum = 1, .Maximum = 50, .Value = 8}

        chkIndividual = New CheckBox() With {.Left = 620, .Top = 158, .Width = 140, .Text = "Guardar individuales", .Checked = True}
        chkSheets = New CheckBox() With {.Left = 760, .Top = 158, .Width = 140, .Text = "Guardar hojas", .Checked = True}

        btnPreview = New Button() With {.Left = 12, .Top = 200, .Width = 140, .Text = "Vista previa"}
        AddHandler btnPreview.Click, AddressOf BtnPreview_Click

        btnGenerate = New Button() With {.Left = 170, .Top = 200, .Width = 140, .Text = "Generar etiquetas"}
        AddHandler btnGenerate.Click, AddressOf BtnGenerate_Click

        progressBar = New ProgressBar() With {.Left = 12, .Top = 240, .Width = 520, .Height = 24}
        lblStatus = New Label() With {.Left = 12, .Top = 270, .Width = 880, .Height = 24, .Text = "Estado: listo"}

        pbPreview = New PictureBox() With {.Left = 560, .Top = 200, .Width = 340, .Height = 260, .BorderStyle = BorderStyle.FixedSingle, .SizeMode = PictureBoxSizeMode.Zoom}

        ' Añadir controles
        Me.Controls.AddRange(New Control() {
            txtExcel, btnBrowseExcel, txtOutput, btnBrowseOutput,
            lbl1, txtQrCols, lbl2, txtTextCols,
            lbl3, nudLabelW, nudLabelH, lbl4, nudQrSize, nudDpi,
            lbl5, nudPerRow, nudPerCol, chkIndividual, chkSheets,
            btnPreview, btnGenerate, progressBar, lblStatus, pbPreview
        })
    End Sub

    Private Sub BtnBrowseExcel_Click(sender As Object, e As EventArgs)
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Excel files|*.xlsx;*.xlsm;*.xls"
            If ofd.ShowDialog() = DialogResult.OK Then
                txtExcel.Text = ofd.FileName
            End If
        End Using
    End Sub

    Private Sub BtnBrowseOutput_Click(sender As Object, e As EventArgs)
        Using fbd As New FolderBrowserDialog()
            If fbd.ShowDialog() = DialogResult.OK Then
                txtOutput.Text = fbd.SelectedPath
            End If
        End Using
    End Sub

    Private Sub BtnPreview_Click(sender As Object, e As EventArgs)
        Try
            If Not File.Exists(txtExcel.Text) Then
                MessageBox.Show("Selecciona un archivo Excel válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
            lblStatus.Text = "Estado: generando vista previa..."
            Application.DoEvents()
            Dim img = GeneratePreviewImage(txtExcel.Text, GetQrCols(), GetTextCols())
            If img IsNot Nothing Then
                pbPreview.Image = img
                lblStatus.Text = "Estado: vista previa generada"
            Else
                lblStatus.Text = "Estado: no se pudo generar vista previa (posible falta de datos)."
                MessageBox.Show("No se encontró una fila con datos de QR.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show("Error en vista previa: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            lblStatus.Text = "Estado: error en vista previa"
        End Try
    End Sub

    Private Sub BtnGenerate_Click(sender As Object, e As EventArgs)
        If Not File.Exists(txtExcel.Text) Then
            MessageBox.Show("Selecciona un archivo Excel válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If
        If String.IsNullOrWhiteSpace(txtOutput.Text) Then
            MessageBox.Show("Selecciona una carpeta de salida.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Recolectar opciones
        Dim excelPath = txtExcel.Text
        Dim outDir = txtOutput.Text
        Dim qrCols = GetQrCols()
        Dim textCols = GetTextCols()
        Dim labelWmm = Convert.ToDouble(nudLabelW.Value)
        Dim labelHmm = Convert.ToDouble(nudLabelH.Value)
        Dim qrMm = Convert.ToDouble(nudQrSize.Value)
        Dim dpi = Convert.ToInt32(nudDpi.Value)
        Dim perRow = Convert.ToInt32(nudPerRow.Value)
        Dim perCol = Convert.ToInt32(nudPerCol.Value)
        Dim saveIndividual = chkIndividual.Checked
        Dim saveSheets = chkSheets.Checked

        btnGenerate.Enabled = False
        btnPreview.Enabled = False
        lblStatus.Text = "Estado: generando..."
        progressBar.Value = 0

        Task.Run(Sub()
                     Try
                         ProcessExcelAndGenerate(
                             excelPath, outDir, qrCols, textCols,
                             labelWmm, labelHmm, qrMm, dpi,
                             saveIndividual, saveSheets, perRow, perCol)
                         Me.Invoke(Sub()
                                       lblStatus.Text = "Estado: terminado"
                                       progressBar.Value = 100
                                       MessageBox.Show("Generación completada.", "Listo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                   End Sub)
                     Catch ex As Exception
                         Me.Invoke(Sub()
                                       lblStatus.Text = "Estado: error"
                                       MessageBox.Show("Error durante generación: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                   End Sub)
                     Finally
                         Me.Invoke(Sub()
                                       btnGenerate.Enabled = True
                                       btnPreview.Enabled = True
                                   End Sub)
                     End Try
                 End Sub)
    End Sub

    ' ---------------------------
    ' Lógica: lectura, QR y composición
    ' ---------------------------
    Private Function GetQrCols() As String()
        Return txtQrCols.Text.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries).[Select](Function(s) s.Trim()).ToArray()
    End Function

    Private Function GetTextCols() As String()
        Return txtTextCols.Text.Split(New Char() {","c}, StringSplitOptions.RemoveEmptyEntries).[Select](Function(s) s.Trim()).ToArray()
    End Function

    Private Function GeneratePreviewImage(excelPath As String, qrCols As String(), textCols As String()) As Image
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Using p As New ExcelPackage(New FileInfo(excelPath))
            Dim ws = p.Workbook.Worksheets.FirstOrDefault()
            If ws Is Nothing OrElse ws.Dimension Is Nothing Then Return Nothing
            Dim startRow = ws.Dimension.Start.Row
            Dim endRow = ws.Dimension.End.Row
            Dim startCol = ws.Dimension.Start.Column
            Dim endCol = ws.Dimension.End.Column

            ' headers
            Dim headers = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For c = startCol To endCol
                Dim h = Convert.ToString(ws.Cells(startRow, c).Value)
                If Not String.IsNullOrWhiteSpace(h) Then headers(h.Trim()) = c
            Next

            For r = startRow + 1 To endRow
                Dim qrData = ConcatRowValues(ws, r, qrCols, headers)
                If String.IsNullOrWhiteSpace(qrData) Then Continue For
                Dim qrPx = MmToPx(Convert.ToDouble(nudQrSize.Value), Convert.ToInt32(nudDpi.Value))
                Using qrBmp = MakeQrBitmap(qrData, qrPx)
                    Dim textLines = BuildTextLines(ws, r, textCols, headers)
                    Dim labelW = MmToPx(Convert.ToDouble(nudLabelW.Value), Convert.ToInt32(nudDpi.Value))
                    Dim labelH = MmToPx(Convert.ToDouble(nudLabelH.Value), Convert.ToInt32(nudDpi.Value))
                    Dim img = DrawLabel(qrBmp, textLines, labelW, labelH, MmToPx(6, Convert.ToInt32(nudDpi.Value)))
                    Return img
                End Using
            Next
        End Using
        Return Nothing
    End Function

    Private Sub ProcessExcelAndGenerate(
        excelPath As String,
        outDir As String,
        qrCols As String(),
        textCols As String(),
        labelWmm As Double,
        labelHmm As Double,
        qrMm As Double,
        dpi As Integer,
        saveIndividual As Boolean,
        saveSheets As Boolean,
        perRow As Integer,
        perCol As Integer)

        Directory.CreateDirectory(outDir)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim labelW = MmToPx(labelWmm, dpi)
        Dim labelH = MmToPx(labelHmm, dpi)
        Dim qrPx = MmToPx(qrMm, dpi)
        Dim qrMargin = MmToPx(6, dpi) ' fijo; puedes exponerlo si quieres
        Dim spacingX = MmToPx(6, dpi)
        Dim marginPx = MmToPx(10, dpi)

        Dim labelsForSheet As New System.Collections.Generic.List(Of Bitmap)()
        Dim totalCount As Integer = 0

        Using p As New ExcelPackage(New FileInfo(excelPath))
            Dim ws = p.Workbook.Worksheets.FirstOrDefault()
            If ws Is Nothing OrElse ws.Dimension Is Nothing Then Throw New Exception("La hoja Excel está vacía o no tiene dimensiones.")
            Dim startRow = ws.Dimension.Start.Row
            Dim endRow = ws.Dimension.End.Row
            Dim startCol = ws.Dimension.Start.Column
            Dim endCol = ws.Dimension.End.Column
            Dim headers = New Dictionary(Of String, Integer)(StringComparer.OrdinalIgnoreCase)
            For c = startCol To endCol
                Dim h = Convert.ToString(ws.Cells(startRow, c).Value)
                If Not String.IsNullOrWhiteSpace(h) Then headers(h.Trim()) = c
            Next

            Dim rowsWithQr = New List(Of Integer)
            For r = startRow + 1 To endRow
                Dim q = ConcatRowValues(ws, r, qrCols, headers)
                If Not String.IsNullOrWhiteSpace(q) Then rowsWithQr.Add(r)
            Next

            Dim totalRows = rowsWithQr.Count
            Dim processed = 0
            For Each r In rowsWithQr
                Dim qrData = ConcatRowValues(ws, r, qrCols, headers)
                Using qrBmp = MakeQrBitmap(qrData, qrPx)
                    Dim textLines = BuildTextLines(ws, r, textCols, headers)
                    Dim labelBmp = DrawLabel(qrBmp, textLines, labelW, labelH, qrMargin)
                    Dim baseName = SafeFilename(ConcatRowValues(ws, r, qrCols, headers).Replace("|", "_"))
                    If String.IsNullOrEmpty(baseName) Then baseName = "row_" & r.ToString()
                    If saveIndividual Then
                        Dim outPath = Path.Combine(outDir, baseName & ".png")
                        labelBmp.Save(outPath, ImageFormat.Png)
                    End If
                    labelsForSheet.Add(labelBmp)
                    totalCount += 1
                    processed += 1
                End Using

                ' actualizar progreso
                Dim perc = CInt(Math.Round(100.0 * processed / Math.Max(1, totalRows)))
                Me.Invoke(Sub() progressBar.Value = Math.Min(100, Math.Max(0, perc)))
            Next

            ' crear hojas
            If saveSheets AndAlso labelsForSheet.Count > 0 Then
                Dim pageSize = perRow * perCol
                Dim pageIndex = 0
                For i = 0 To labelsForSheet.Count - 1 Step pageSize
                    pageIndex += 1
                    Dim chunk = labelsForSheet.Skip(i).Take(pageSize).ToList()
                    Dim sheetBmp = MakeSheet(chunk, perRow, perCol, spacingX, spacingX, marginPx)
                    Dim sheetPath = Path.Combine(outDir, $"sheet_{pageIndex}.png")
                    sheetBmp.Save(sheetPath, ImageFormat.Png)
                    sheetBmp.Dispose()
                Next
            End If
        End Using
    End Sub

    ' ---------------------------
    ' Helpers: lectura y texto
    ' ---------------------------
    Private Function ConcatRowValues(ws As OfficeOpenXml.ExcelWorksheet, row As Integer, cols As String(), headers As Dictionary(Of String, Integer)) As String
        If cols Is Nothing OrElse cols.Length = 0 Then Return ""
        Dim list As New List(Of String)
        For Each c In cols
            If String.IsNullOrWhiteSpace(c) Then Continue For
            Dim name = c.Trim()
            If headers.ContainsKey(name) Then
                Dim v = ws.Cells(row, headers(name)).Value
                If v IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Convert.ToString(v)) Then
                    list.Add(Convert.ToString(v))
                End If
            End If
        Next
        Return String.Join(" | ", list)
    End Function

    Private Function BuildTextLines(ws As OfficeOpenXml.ExcelWorksheet, row As Integer, cols As String(), headers As Dictionary(Of String, Integer)) As List(Of String)
        Dim lines As New List(Of String)
        If cols Is Nothing OrElse cols.Length = 0 Then Return lines
        For Each c In cols
            If String.IsNullOrWhiteSpace(c) Then Continue For
            Dim name = c.Trim()
            If headers.ContainsKey(name) Then
                Dim v = ws.Cells(row, headers(name)).Value
                If v IsNot Nothing Then lines.Add($"{name}: {Convert.ToString(v)}")
            End If
        Next
        Return lines
    End Function

    Private Function SafeFilename(s As String) As String
        If s Is Nothing Then s = ""
        s = Regex.Replace(s, "[^\w\-_\. ]", "_")
        s = Regex.Replace(s, "\s+", "_")
        If s.Length > 200 Then s = s.Substring(0, 200)
        Return s
    End Function

    ' ---------------------------
    ' QR + dibujo
    ' ---------------------------
    Private Function MakeQrBitmap(data As String, qrSizePx As Integer) As Bitmap
        If data Is Nothing Then data = ""
        Using generator As New QRCodeGenerator()
            Using qrData = generator.CreateQrCode(data, QRCodeGenerator.ECCLevel.M)
                Using qrcode = New QRCode(qrData)
                    Dim raw As Bitmap = qrcode.GetGraphic(10, Color.Black, Color.White, True)
                    Dim resized As New Bitmap(qrSizePx, qrSizePx)
                    Using g As Graphics = Graphics.FromImage(resized)
                        g.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
                        g.CompositingQuality = Drawing2D.CompositingQuality.HighSpeed
                        g.SmoothingMode = Drawing2D.SmoothingMode.None
                        g.Clear(Color.White)
                        g.DrawImage(raw, 0, 0, qrSizePx, qrSizePx)
                    End Using
                    Return resized
                End Using
            End Using
        End Using
    End Function

    Private Function DrawLabel(qrImg As Bitmap, textLines As List(Of String), labelW As Integer, labelH As Integer, qrMargin As Integer) As Bitmap
        Dim bmp As New Bitmap(labelW, labelH)
        Using g As Graphics = Graphics.FromImage(bmp)
            g.Clear(Color.White)
            g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit

            Dim padding As Integer = CInt(Math.Round(labelW * 0.02)) ' small padding
            Dim qrW As Integer = qrImg.Width
            Dim qrH As Integer = qrImg.Height
            Dim xQr As Integer = padding
            Dim yQr As Integer = (labelH - qrH) \ 2
            g.DrawImage(qrImg, xQr, yQr, qrW, qrH)

            Dim xText As Integer = xQr + qrW + qrMargin
            Dim textW As Integer = labelW - xText - padding
            Dim y As Integer = padding

            Dim font As New Font(SystemFonts.DefaultFont.FontFamily, 12.0F, FontStyle.Regular, GraphicsUnit.Point)
            Dim brush As New SolidBrush(Color.Black)

            ' Wrap lines manually
            For Each line In textLines
                If String.IsNullOrEmpty(line) Then
                    y += CInt(font.GetHeight(g) * 0.6)
                    Continue For
                End If
                Dim words = line.Split(New Char() {" "c}, StringSplitOptions.RemoveEmptyEntries)
                Dim cur As String = ""
                For Each w In words
                    Dim test = If(String.IsNullOrEmpty(cur), w, cur & " " & w)
                    Dim size = g.MeasureString(test, font)
                    If size.Width <= textW OrElse String.IsNullOrEmpty(cur) Then
                        cur = test
                    Else
                        g.DrawString(cur, font, brush, New RectangleF(xText, y, textW, labelH - y))
                        y += CInt(g.MeasureString(cur, font).Height) + 2
                        cur = w
                    End If
                Next
                If Not String.IsNullOrEmpty(cur) Then
                    g.DrawString(cur, font, brush, New RectangleF(xText, y, textW, labelH - y))
                    y += CInt(g.MeasureString(cur, font).Height) + 4
                End If
            Next
        End Using
        Return bmp
    End Function

    Private Function MakeSheet(labels As List(Of Bitmap), perRow As Integer, perCol As Integer, spacingX As Integer, spacingY As Integer, marginPx As Integer) As Bitmap
        If labels Is Nothing OrElse labels.Count = 0 Then Throw New ArgumentException("No hay etiquetas para crear la hoja.")
        Dim lw = labels(0).Width
        Dim lh = labels(0).Height
        Dim sheetW = marginPx * 2 + perRow * lw + (perRow - 1) * spacingX
        Dim sheetH = marginPx * 2 + perCol * lh + (perCol - 1) * spacingY
        Dim sheet As New Bitmap(sheetW, sheetH)
        Using g As Graphics = Graphics.FromImage(sheet)
            g.Clear(Color.White)
            Dim i As Integer = 0
            For r As Integer = 0 To perCol - 1
                For c As Integer = 0 To perRow - 1
                    If i >= labels.Count Then Exit For
                    Dim x = marginPx + c * (lw + spacingX)
                    Dim y = marginPx + r * (lh + spacingY)
                    g.DrawImage(labels(i), x, y, lw, lh)
                    i += 1
                Next
            Next
        End Using
        Return sheet
    End Function

    ' ---------------------------
    ' Util
    ' ---------------------------
    Private Function MmToPx(mm As Double, dpi As Integer) As Integer
        Return CInt((mm / 25.4) * dpi + 0.5)
    End Function

End Class
