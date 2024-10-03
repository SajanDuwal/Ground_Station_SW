using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace S2S_Ground_Station
{
    public partial class commandList : Form
    {
        public operationMain? formMain;

        public commandList()
        {
            InitializeComponent();
        }

        private void CommandList_Load(object sender, EventArgs e)
        {
            tabControl1.SelectedIndex = 0;
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            insert_button_Click(sender, e);
        }

        private void insert_button_Click(object sender, EventArgs e)
        {
            // Declare the relevant DataGridView and check for CurrentRow
            DataGridView? currentGridView = null;
            switch (tabControl1.SelectedIndex)
            {
                case 0:
                    currentGridView = dataGridView1;
                    break;
                case 1:
                    currentGridView = dataGridView2;
                    break;
                case 2:
                    currentGridView = dataGridView3;
                    break;
            }

            // Ensure the CurrentRow is not null and check the first cell's value
            if (currentGridView?.CurrentRow == null)
            {
                return;
            }

            if (currentGridView.CurrentRow.Cells[0].Value == null && currentGridView.CurrentRow.Cells[1].Value == null)
            {
                return;
            }

            // Assign the CommandString and CommentString
            formMain!.CommandString = currentGridView.CurrentRow.Cells[0].Value.ToString() ?? string.Empty;
            formMain.CommentString = currentGridView.CurrentRow.Cells[1].Value.ToString() ?? string.Empty;

            // Move to the next row if it exists
            int nextRowIndex = currentGridView.CurrentCell.RowIndex + 1;
            if (nextRowIndex < currentGridView.Rows.Count) // Check if the next row exists
            {
                currentGridView.CurrentCell = currentGridView.Rows[nextRowIndex].Cells[0];
            }
        }

        private void open_button_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "csv file(*.csv)|*.csv|all file(*.*)|*.*";
            ofd.FilterIndex = 1;
            ofd.Title = "Select command file";
            ofd.RestoreDirectory = true;
            ofd.CheckFileExists = true;

            if (ofd.ShowDialog() == DialogResult.OK)
            {

                Console.WriteLine(ofd.FileName);

                // Register the encoding provider
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                // Now create the TextFieldParser with Shift_JIS encoding
                TextFieldParser parser = new(ofd.FileName, Encoding.GetEncoding("Shift_JIS"))
                {
                    TextFieldType = FieldType.Delimited
                };
                parser.SetDelimiters(",");

                parser.ReadFields();
                while (!parser.EndOfData)
                {
                    String[]? row = parser.ReadFields();

                    if (tabControl1.SelectedIndex == 0)
                    {
                        dataGridView1.Rows.Add(row);
                    }
                    else if (tabControl1.SelectedIndex == 1)
                    {
                        dataGridView2.Rows.Add(row);
                    }
                    else
                    {
                        dataGridView3.Rows.Add(row);
                    }
                }
            }
        }

        private void clear_button_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                dataGridView1.Rows.Clear();
            }
            else if (tabControl1.SelectedIndex == 1)
            {
                dataGridView2.Rows.Clear();
            }
            else
            {
                dataGridView3.Rows.Clear();
            }
        }

        private void clear_row_button_Click(object sender, EventArgs e)
        {
            // Handle DataGridView1 if Tab 0 is selected
            if (tabControl1.SelectedIndex == 0)
            {
                removeSelectedRows(dataGridView1);
            }
            // Handle DataGridView2 if Tab 1 is selected
            else if (tabControl1.SelectedIndex == 1)
            {
                removeSelectedRows(dataGridView2);
            }
            // Handle DataGridView3 for any other tab
            else
            {
                removeSelectedRows(dataGridView3);
            }
        }

        // Function to remove selected rows based on whether the DataGridView is bound or unbound
        private void removeSelectedRows(DataGridView dataGridView)
        {
            if (dataGridView.SelectedRows.Count > 0)
            {
                // Loop through selected rows backwards to avoid index issues
                for (int i = dataGridView.SelectedRows.Count - 1; i >= 0; i--)
                {
                    var row = dataGridView.SelectedRows[i];

                    // Check if the row is a new row (uncommitted)
                    if (row.IsNewRow)
                    {
                        continue; // Skip this row
                    }

                    // Ensure the row is not empty before attempting to remove
                    if (!row.IsNewRow && row.Cells.Cast<DataGridViewCell>().Any(cell => !cell.Value.Equals(string.Empty)))
                    {
                        // Remove the selected row
                        dataGridView.Rows.RemoveAt(row.Index);
                    }
                }
            }
        }

        private void dataGridView2_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            insert_button_Click(sender, e);
        }

        private void dataGridView3_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            insert_button_Click(sender, e);
        }

        private void bindingSource1_CurrentChanged(object sender, EventArgs e)
        {

        }
    }
}