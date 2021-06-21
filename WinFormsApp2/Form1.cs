using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;

namespace WinFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            List<string> fileList = new List<string>();
            foreach(string file in files)
            {
                fileList.Add(file);
            }
            var table = GetCsvToDataTable(fileList.First());
            dataGridView1.DataSource = table;
        }

        private DataTable GetCsvToDataTable(string path) {

            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            using (var parser = new TextFieldParser(path,Encoding.GetEncoding(932)))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.Delimiters = new string[] { "," };
                List<string[]> rows = new List<string[]>();

                while(!parser.EndOfData)
                {
                    rows.Add(parser.ReadFields());
                }

                DataTable table = new DataTable();

                table.Columns.AddRange(rows.First().Select(s => new DataColumn(s)).ToArray());

                foreach(var row in rows.Skip(1))
                {
                    table.Rows.Add(row);
                }

                return table;
            }

        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.All;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
    }
}
