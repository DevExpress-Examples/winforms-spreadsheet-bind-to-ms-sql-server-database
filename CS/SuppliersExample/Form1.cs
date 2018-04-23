#region #Namespaces
using System;
using System.Windows.Forms;
using DevExpress.Spreadsheet;
using SuppliersExample.NWindDataSetTableAdapters;
// ...
#endregion #Namespaces

namespace SuppliersExample
{
    public partial class Form1 : DevExpress.XtraBars.Ribbon.RibbonForm
    {
        bool applyChangesOnRowsRemoved = false;
        #region #BindToData
        NWindDataSet dataSet;
        SuppliersTableAdapter adapter;

        public Form1() {
            InitializeComponent();
            dataSet = new NWindDataSet();
            adapter = new SuppliersTableAdapter();
        }

        private void Form1_Load(object sender, EventArgs e) {
            // Populate the "Suppliers" data table with data.
            adapter.Fill(dataSet.Suppliers);
            IWorkbook workbook = spreadsheetControl1.Document;
            // Load the template document into the SpreadsheetControl.
            workbook.LoadDocument("Suppliers_template.xlsx");
            Worksheet sheet = workbook.Worksheets[0];
            // Load data from the "Suppliers" data table into the worksheet starting from the cell "B12".
            sheet.DataBindings.BindToDataSource(dataSet.Suppliers, 11, 1);
        }
        #endregion #BindToData

        #region #UpdateData
        private void spreadsheetControl1_MouseClick(object sender, MouseEventArgs e) {
            Cell cell = spreadsheetControl1.GetCellFromPoint(e.Location);
            if (cell == null)
                return;
            Worksheet sheet = spreadsheetControl1.ActiveWorksheet;
            string cellReference = cell.GetReferenceA1();
            // If the "Save" cell is clicked in the data entry form, 
            // add a row containing the entered values to the database table.
            if (cellReference == "I4") {
                AddRow(sheet);
                HideDataEntryForm(sheet);
                ApplyChanges();
            }
            // If the "Cancel" cell is clicked in the data entry form, 
            // cancel adding new data and hide the data entry form.
            else if (cellReference == "I6") {
                HideDataEntryForm(sheet);
            }
        }

        void AddRow(Worksheet sheet) {
            try {
                // Append a new row to the "Suppliers" data table.
                dataSet.Suppliers.AddSuppliersRow(
                    sheet["C4"].Value.TextValue, sheet["C6"].Value.TextValue, sheet["C8"].Value.TextValue,
                    sheet["E4"].Value.TextValue, sheet["E6"].Value.TextValue, sheet["E8"].Value.TextValue,
                    sheet.Cells["G4"].DisplayText, sheet.Cells["G6"].DisplayText);
            }
            catch (Exception ex) {
                string message = string.Format("Cannot add a row to a database table.\n{0}", ex.Message);
                MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        void HideDataEntryForm(Worksheet sheet) {
            Range range = sheet.Range.Parse("C4,C6,C8,E4,E6,E8,G4,G6");
            range.ClearContents();
            sheet.Rows.Hide(2, 9);
        }

        void ApplyChanges() {
            try {
                // Send the updated data back to the database.
                adapter.Update(dataSet.Suppliers);
            }
            catch (Exception ex) {
                string message = string.Format("Cannot update data in a database table.\n{0}", ex.Message);
                MessageBox.Show(message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void spreadsheetControl1_RowsRemoving(object sender, RowsChangingEventArgs e) {
            Worksheet sheet = spreadsheetControl1.ActiveWorksheet;
            Range rowRange = sheet.Range.FromLTRB(0, e.StartIndex, 16383, e.StartIndex + e.Count - 1);
            Range boundRange = sheet.DataBindings[0].Range;
            // If the rows to be removed belong to the data-bound range,
            // display a dialog requesting the user to confirm the deletion of records. 
            if (boundRange.IsIntersecting(rowRange)) {
                DialogResult result = MessageBox.Show("Want to delete the selected supplier(s)?", "Delete", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                applyChangesOnRowsRemoved = result == DialogResult.Yes;
                e.Cancel = result == DialogResult.No;
                return;
            }
        }

        private void spreadsheetControl1_RowsRemoved(object sender, RowsChangedEventArgs e) {
            if (applyChangesOnRowsRemoved) {
                applyChangesOnRowsRemoved = false;
                // Update data in the database.
                ApplyChanges();
            }
        }

        private void buttonAddRecord_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            CloseInplaceEditor();
            Worksheet sheet = spreadsheetControl1.ActiveWorksheet;
            // Display the data entry form on the worksheet to add a new record to the "Suppliers" data table.
            if (!sheet.Rows[4].Visible)
                sheet.Rows.Unhide(2, 9);
            spreadsheetControl1.SelectedCell = sheet["C4"];
        }

        private void buttonRemoveRecord_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            CloseInplaceEditor();
            Worksheet sheet = spreadsheetControl1.ActiveWorksheet;
            Range selectedRange = spreadsheetControl1.Selection;
            Range boundRange = sheet.DataBindings[0].Range;
            // Verify that the selected cell range belongs to the data-bound range.
            if (!boundRange.IsIntersecting(selectedRange) || selectedRange.TopRowIndex < boundRange.TopRowIndex) {
                MessageBox.Show("Select a record first!", "Remove Record", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            } 
            // Remove the topmost row of the selected cell range.
            sheet.Rows.Remove(selectedRange.TopRowIndex);
        }

        private void buttonApplyChanges_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            CloseInplaceEditor();
            // Update data in the database.
            ApplyChanges();
        }

        private void buttonCancelChanges_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e) {
            // Close the cell in-place editor if it's currently active. 
            CloseInplaceEditor();
            // Load the latest saved data into the "Suppliers" data table.
            adapter.Fill(dataSet.Suppliers);
        }

        void CloseInplaceEditor() {
            if (spreadsheetControl1.IsCellEditorActive)
                spreadsheetControl1.CloseCellEditor(DevExpress.XtraSpreadsheet.CellEditorEnterValueMode.Default);
        }
        #endregion #UpdateData
    }
}
