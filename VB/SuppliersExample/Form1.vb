#Region "#Namespaces"
Imports System
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports SuppliersExample.NWindDataSetTableAdapters

' ...
#End Region  ' #Namespaces
Namespace SuppliersExample

    Public Partial Class Form1
        Inherits DevExpress.XtraBars.Ribbon.RibbonForm

        Private applyChangesOnRowsRemoved As Boolean = False

#Region "#BindToData"
        Private dataSet As NWindDataSet

        Private adapter As SuppliersTableAdapter

        Public Sub New()
            InitializeComponent()
            dataSet = New NWindDataSet()
            adapter = New SuppliersTableAdapter()
        End Sub

        Private Sub Form1_Load(ByVal sender As Object, ByVal e As EventArgs)
            ' Populate the "Suppliers" data table with data.
            adapter.Fill(dataSet.Suppliers)
            Dim workbook As IWorkbook = spreadsheetControl1.Document
            ' Load the template document into the SpreadsheetControl.
            workbook.LoadDocument("Suppliers_template.xlsx")
            Dim sheet As Worksheet = workbook.Worksheets(0)
            ' Load data from the "Suppliers" data table into the worksheet starting from the cell "B12".
            sheet.DataBindings.BindToDataSource(dataSet.Suppliers, 11, 1)
        End Sub

#End Region  ' #BindToData
#Region "#UpdateData"
        Private Sub spreadsheetControl1_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs)
            Dim cell As Cell = spreadsheetControl1.GetCellFromPoint(e.Location)
            If cell Is Nothing Then Return
            Dim sheet As Worksheet = spreadsheetControl1.ActiveWorksheet
            Dim cellReference As String = cell.GetReferenceA1()
            ' If the "Save" cell is clicked in the data entry form, 
            ' add a row containing the entered values to the database table.
            If Equals(cellReference, "I4") Then
                AddRow(sheet)
                HideDataEntryForm(sheet)
                ApplyChanges()
            ' If the "Cancel" cell is clicked in the data entry form, 
            ' cancel adding new data and hide the data entry form.
            ElseIf Equals(cellReference, "I6") Then
                HideDataEntryForm(sheet)
            End If
        End Sub

        Private Sub AddRow(ByVal sheet As Worksheet)
            Try
                ' Append a new row to the "Suppliers" data table.
                dataSet.Suppliers.AddSuppliersRow(sheet("C4").Value.TextValue, sheet("C6").Value.TextValue, sheet("C8").Value.TextValue, sheet("E4").Value.TextValue, sheet("E6").Value.TextValue, sheet("E8").Value.TextValue, sheet.Cells("G4").DisplayText, sheet.Cells("G6").DisplayText)
            Catch ex As Exception
                Dim message As String = String.Format("Cannot add a row to a database table." & Microsoft.VisualBasic.Constants.vbLf & "{0}", ex.Message)
                MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub HideDataEntryForm(ByVal sheet As Worksheet)
            Dim range As CellRange = sheet.Range.Parse("C4,C6,C8,E4,E6,E8,G4,G6")
            range.ClearContents()
            sheet.Rows.Hide(2, 9)
        End Sub

        Private Sub ApplyChanges()
            Try
                ' Send the updated data back to the database.
                adapter.Update(dataSet.Suppliers)
            Catch ex As Exception
                Dim message As String = String.Format("Cannot update data in a database table." & Microsoft.VisualBasic.Constants.vbLf & "{0}", ex.Message)
                MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Sub

        Private Sub spreadsheetControl1_RowsRemoving(ByVal sender As Object, ByVal e As RowsChangingEventArgs)
            Dim sheet As Worksheet = spreadsheetControl1.ActiveWorksheet
            Dim rowRange As CellRange = sheet.Range.FromLTRB(0, e.StartIndex, 16383, e.StartIndex + e.Count - 1)
            Dim boundRange As CellRange = sheet.DataBindings(0).Range
            ' If the rows to be removed belong to the data-bound range,
            ' display a dialog requesting the user to confirm the deletion of records. 
            If boundRange.IsIntersecting(rowRange) Then
                Dim result As DialogResult = MessageBox.Show("Want to delete the selected supplier(s)?", "Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                applyChangesOnRowsRemoved = result = DialogResult.Yes
                e.Cancel = result = DialogResult.No
                Return
            End If
        End Sub

        Private Sub spreadsheetControl1_RowsRemoved(ByVal sender As Object, ByVal e As RowsChangedEventArgs)
            If applyChangesOnRowsRemoved Then
                applyChangesOnRowsRemoved = False
                ' Update data in the database.
                ApplyChanges()
            End If
        End Sub

        Private Sub buttonAddRecord_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs)
            CloseInplaceEditor()
            Dim sheet As Worksheet = spreadsheetControl1.ActiveWorksheet
            ' Display the data entry form on the worksheet to add a new record to the "Suppliers" data table.
            If Not sheet.Rows(4).Visible Then sheet.Rows.Unhide(2, 9)
            spreadsheetControl1.SelectedCell = sheet("C4")
        End Sub

        Private Sub buttonRemoveRecord_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs)
            CloseInplaceEditor()
            Dim sheet As Worksheet = spreadsheetControl1.ActiveWorksheet
            Dim selectedRange As CellRange = spreadsheetControl1.Selection
            Dim boundRange As CellRange = sheet.DataBindings(0).Range
            ' Verify that the selected cell range belongs to the data-bound range.
            If Not boundRange.IsIntersecting(selectedRange) OrElse selectedRange.TopRowIndex < boundRange.TopRowIndex Then
                MessageBox.Show("Select a record first!", "Remove Record", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' Remove the topmost row of the selected cell range.
            sheet.Rows.Remove(selectedRange.TopRowIndex)
        End Sub

        Private Sub buttonApplyChanges_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs)
            CloseInplaceEditor()
            ' Update data in the database.
            ApplyChanges()
        End Sub

        Private Sub buttonCancelChanges_ItemClick(ByVal sender As Object, ByVal e As DevExpress.XtraBars.ItemClickEventArgs)
            ' Close the cell in-place editor if it's currently active. 
            CloseInplaceEditor()
            ' Load the latest saved data into the "Suppliers" data table.
            adapter.Fill(dataSet.Suppliers)
        End Sub

        Private Sub CloseInplaceEditor()
            If spreadsheetControl1.IsCellEditorActive Then spreadsheetControl1.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default)
        End Sub
#End Region  ' #UpdateData
    End Class
End Namespace
