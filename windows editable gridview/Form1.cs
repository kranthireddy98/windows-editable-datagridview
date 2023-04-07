using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;
using OfficeOpenXml;
using System.IO;

namespace windows_editable_gridview
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public bool IncludeHeaders { get; set; } = true;
        public int columnIndex { get; set; } = -1;

        
        private void LoadData()
        {
            string connectionString = "Data Source=.;Initial Catalog=sample;Integrated Security=True;";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();

                string query = "SELECT * FROM dbo.Employees";

                SqlDataAdapter adapter = new SqlDataAdapter(query, connection);
                DataTable data = new DataTable();
                adapter.Fill(data);

                dataGridView1.DataSource = data;
            }
        }

        private void Form1_Load(object sender, System.EventArgs e)
        {
            LoadData();
        }
        private void DeleteColumn()
        {
            if (dataGridView1.SelectedCells.Count > 0)
            {
                int columnIndex = dataGridView1.SelectedCells[0].ColumnIndex;
                dataGridView1.Columns.RemoveAt(columnIndex);
            }
        }

        private void button1_Click(object sender, System.EventArgs e)
        {
            DeleteColumn();
        }

        private void ExportToExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Create a new Excel package
            // Create a new Excel package

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                // Add a new worksheet
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // Set the column headers if IncludeHeaders option is enabled
                if (IncludeHeaders)
                {
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        worksheet.Cells[1, i + 1].Value = dataGridView1.Columns[i].HeaderText;
                    }
                }

                // Fill the data
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataGridView1.Rows[i].Cells[j].Value;
                    }
                }

                // Save the Excel file
                FileInfo excelFile = new FileInfo("output.xlsx");
                excelPackage.SaveAs(excelFile);
            }


        }

        private void button2_Click(object sender, System.EventArgs e)

        {
            dataGridView2.DataSource = null;
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                if (column.Index != columnIndex)
                {
                    dataGridView2.Columns.Add(column.Clone() as DataGridViewColumn);
                }

            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                // Create a new row in the DataGridView
                DataGridViewRow row = new DataGridViewRow();

                // Copy the cell values from the first DataGridView to the new one, except for the removed column
                
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    if (j != columnIndex)
                    {
                        row.Cells.Add(new DataGridViewTextBoxCell { Value = dataGridView1.Rows[i].Cells[j].Value });
                    }
                }

                // Add the new row to the new DataGridView
                dataGridView2.Rows.Add(row);
                
            }
            //dataGridView2.DataSource = dataGridView1;
        }
        private void dataGridView1_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            // Get the clicked column index
            int columnIndex = e.ColumnIndex;

            // Set the Selected property of all the cells in the clicked column to true
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCell cell = row.Cells[columnIndex];
                cell.Selected = true;
            }
        }
        private void checkBox1_CheckedChanged(object sender, System.EventArgs e)
        {
            IncludeHeaders = false;
            
        }

       
        // private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        //    {
        //        // Check if the click is on a header cell
        //        if (e.RowIndex == -1)
        //        {
        //            // Get the index of the clicked column
        //            int columnIndex = e.ColumnIndex;

        //            // Remove the selected column from the DataGridView
        //            dataGridView1.Columns.RemoveAt(columnIndex);

        //            // Clone the DataGridView to create a new instance with the updated data
        //            DataGridView dataGridView2 = (DataGridView)dataGridView1.Clone();

        //            // Populate the new DataGridView with the updated data
        //            for (int i = 0; i < dataGridView1.Rows.Count; i++)
        //            {
        //                // Create a new row in the DataGridView
        //                DataGridViewRow row = new DataGridViewRow();

        //                // Copy the cell values from the first DataGridView to the new one, except for the removed column
        //                for (int j = 0; j < dataGridView1.Columns.Count; j++)
        //                {
        //                    if (j != columnIndex)
        //                    {
        //                        row.Cells.Add(new DataGridViewTextBoxCell { Value = dataGridView1.Rows[i].Cells[j].Value });
        //                    }
        //                }

        //                // Add the new row to the new DataGridView
        //                dataGridView2.Rows.Add(row);
        //            }

        //            // Display the new DataGridView
        //            dataGridView2.Dock = DockStyle.Fill;
        //            panel2.Controls.Add(dataGridView2);
        //        }
        //    }

        //}

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e, Form1 form1)
        {
            // Check if the click is on a header cell
            if (e.RowIndex == -1)
            {
                // Get the index of the clicked column
                int columnIndex = e.ColumnIndex;

                // Remove the selected column from the DataGridView
                dataGridView1.Columns.RemoveAt(columnIndex);

                // Create a new DataGridView control
                //DataGridView dataGridView2 = new DataGridView();

                //// Set the properties of the new DataGridView control to match those of the original DataGridView control
                //dataGridView2.AllowUserToAddRows = false;
                //dataGridView2.AllowUserToDeleteRows = false;
                //dataGridView2.AllowUserToResizeRows = false;
                //dataGridView2.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
                //dataGridView2.Dock = DockStyle.Fill;
                //dataGridView2.ReadOnly = true;
                //dataGridView2.RowHeadersVisible = false;
                //form1.Controls.Add(dataGridView2);

                // Populate the new DataGridView control with the updated data
                //for (int i = 0; i < dataGridView1.Rows.Count; i++)
                //{
                //    // Create a new row in the DataGridView
                //    DataGridViewRow row = new DataGridViewRow();

                //    // Copy the cell values from the first DataGridView to the new one, except for the removed column
                //    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                //    {
                //        if (j != columnIndex)
                //        {
                //            row.Cells.Add(new DataGridViewTextBoxCell { Value = dataGridView1.Rows[i].Cells[j].Value });
                //        }
                //    }

                //    // Add the new row to the new DataGridView
                //    dataGridView2.Rows.Add(row);
                //}
            }
        }

        private void button3_Click(object sender, System.EventArgs e)
        {
           
                dataGridView1.ColumnHeadersVisible = false;
            

        }
    }
}

