using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace mapapprox
{
    public partial class Form1 : Form
    {
        private string currentStatusMessage = "";
        private bool isMappingMode = true;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            char[] EscapeChars = "~!@#$%^&*()_+`-=[]\\{}|;':\",./<>? ".ToCharArray();
            for (int i = 0; i < EscapeChars.Length; i++) comparableString.EscapeWords.Add(EscapeChars[i].ToString());
            //comparableString.EscapeWords.AddRange("~!@#$%^&*()_+`-=[]\\{}|;':\",./<>?".ToCharArray()); // new string[] { "-", ".", "/", "\\", "'", "\"", " ", ",", "(", ")", "+", "`", "~", "!", "@", "#", "$", "%", "^",  }

            //comparableString str1 = new comparableString("anshita private limited");
            //comparableString str2 = new comparableString("anshita pvt ltd");
            comparableString str1 = new comparableString("A.");
            comparableString str2 = new comparableString("A..");

            //MessageBox.Show(compare_string(str1, str2).ToString());

            //this.Close();
        }

        private int compare_string(comparableString str1, comparableString str2)
        {
            // non comparable // match full string are same
            if (str1.length == 0) return 0;

            List<string> match_string = new List<string>();

            for (int i = 0; i < str1.length; i++)
            {
                int ix = -1;
                int startIndex = 0;

                while (startIndex < str2.length && (ix = str2.IndexOf(str1[i, 1], startIndex)) != -1)
                {
                    int j = 1;
                    string s = str1[i, 1];
                    while (i + j < str1.length && ix + j < str2.length && str1[i + j] == str2[ix + j])
                    {
                        s += str1[i + j];
                        j++;
                    }

                    startIndex += ix + j;

                    match_string.Add(s);
                }
            }

            // sorting by length
            for (int i = 0; i < match_string.Count - 1; i++)
            {
                for (int j = i + 1; j < match_string.Count; j++)
                {
                    if (match_string[i].Length < match_string[j].Length)
                    {
                        string tempString = match_string[i];
                        match_string[i] = match_string[j];
                        match_string[j] = tempString;
                    }
                }
            }

            int[] str1_cleaned = new int[str1.length];
            int[] str2_cleaned = new int[str2.length];

            int match_count = 0;

            for (int i = 0; i < match_string.Count; i++)
            {
                int ix1 = 0;
                int ix2 = 0;

                while (ix1 < str1.length && (ix1 = str1.IndexOf(match_string[i], ix1)) != -1)
                {
                    if (str1_cleaned[ix1] == 0) break;
                    ix1++;
                }

                while (ix2 < str2.length && (ix2 = str2.IndexOf(match_string[i], ix2)) != -1)
                {
                    if (str2_cleaned[ix2] == 0) break;
                    ix2++;
                }

                if (ix1 == -1 || ix2 == -1 || ix1 >= str1.length || ix2 >= str2.length) continue;

                match_count += (int)Math.Pow(match_string[i].Length, 2); //match_string[i].Length * 2 - 1; //
                for (int j = 0; j < match_string[i].Length; j++)
                {
                    str1_cleaned[ix1 + j] = 1;
                    str2_cleaned[ix2 + j] = 1;
                }
            }

            //bool isContinues = false;
            //int continues_match_count = 0;
            //for (int i = 0; i < str1_cleaned.Length; i++)
            //{
            //    if (str1_cleaned[i] == 1)
            //    {
            //        match_count++;
            //        continues_match_count++;
            //        if (isContinues)
            //        {
            //            //match_count++;
            //            match_count += (continues_match_count - 1) * 2;
            //        }
            //        isContinues = true;
            //    }
            //    else
            //    {
            //        isContinues = false;
            //        continues_match_count = 0;
            //    }
            //}

            int result = (match_count * 100) / (int)Math.Pow(Math.Max(str1.length, str2.length), 2); // (Math.Max(str1.length, str2.length)*2-1); //
            return result;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel File|*.xlsx";
                if (ofd.ShowDialog() != System.Windows.Forms.DialogResult.OK) return;

                textBox4.Text = ofd.FileName;
            }

            comboBox1.Items.Clear();
            comboBox2.Items.Clear();

            string FileName = textBox4.Text;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wbexcel = excel.Workbooks.Open(FileName);
            foreach (Microsoft.Office.Interop.Excel.Worksheet ws in wbexcel.Worksheets)
            {
                comboBox1.Items.Add(ws.Name);
                comboBox2.Items.Add(ws.Name);
            }

            wbexcel.Close();

            excel.Quit();
        }

        List<comparableString> primaryList = new List<comparableString>();
        List<comparableString> secondryList = new List<comparableString>();

        private void button2_Click(object sender, EventArgs e)
        {
            string FileName = textBox4.Text;
            string columnIndexing = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            int primaryColumnIndex = columnIndexing.IndexOf(textBox1.Text) + 1;
            int secondaryColumnIndex = columnIndexing.IndexOf(textBox2.Text) + 1;


            if (comboBox1.Text.Length == 0 || textBox1.TextLength != 1 || primaryColumnIndex == 0) return;
            if (radioButton1.Checked && (comboBox2.Text.Length == 0 || textBox2.TextLength != 1 || secondaryColumnIndex == 0)) return;

            comparableString.EscapeWords2.Clear();
            primaryList.Clear();
            secondryList.Clear();

            if (textBox3.TextLength != 0)
            {
                string[] es = textBox3.Text.Replace(Environment.NewLine, "").Split(',');
                for (int k = 0; k < es.Length - 1; k++)
                {
                    for (int j = k + 1; j < es.Length; j++)
                    {
                        if (es[k].Length < es[j].Length)
                        {
                            string ts = es[k];
                            es[k] = es[j];
                            es[j] = ts;
                        }
                    }
                }

                for (int k = 0; k < es.Length; k++)
                {
                    es[k] = es[k].ToLower().Trim();
                }

                comparableString.EscapeWords2.AddRange(es);
            }

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wbexcel = excel.Workbooks.Open(FileName);
            Microsoft.Office.Interop.Excel.Worksheet ws = wbexcel.Worksheets[comboBox1.Text] as Microsoft.Office.Interop.Excel.Worksheet;

            int i = 1;
            if (checkBox1.Checked) i = 2;
            while (true)
            {
                string s = Convert.ToString((ws.Cells[i, primaryColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value);
                if (s == null || s.Length == 0) break;

                primaryList.Add(new comparableString(s));
                i++;
            }

            if (radioButton1.Checked)
            {
                ws = wbexcel.Worksheets[comboBox2.Text] as Microsoft.Office.Interop.Excel.Worksheet;

                i = 1;
                if (checkBox2.Checked) i = 2;
                while (true)
                {
                    string s = Convert.ToString((ws.Cells[i, secondaryColumnIndex] as Microsoft.Office.Interop.Excel.Range).Value);
                    if (s == null || s.Length == 0) break;

                    secondryList.Add(new comparableString(s));
                    i++;
                }
            }

            wbexcel.Close();

            excel.Quit();

            isMappingMode = radioButton1.Checked;

        }

        List<mappedString> mappedList = new List<mappedString>();
        private void button3_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = false;
            button3.Enabled = false;
            backgroundWorker1.RunWorkerAsync();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            currentStatusMessage = "Mapping...";
            if (isMappingMode)
            {
                for (int i = 0; i < primaryList.Count; i++)
                {
                    int maxMatchResult = 0;
                    int result = 0;
                    mappedList.Add(new mappedString(primaryList[i].orginalString, ""));
                    for (int j = 0; j < secondryList.Count; j++)
                    {
                        if (maxMatchResult < (result = compare_string(primaryList[i], secondryList[j])))
                        {
                            mappedList[i].secondaryString = secondryList[j].orginalString;
                            mappedList[i].percent_match = result;
                            maxMatchResult = result;
                        }
                    }

                    backgroundWorker1.ReportProgress((i * 100) / primaryList.Count);
                }
            }
            else
            {
                for (int i = 0; i < primaryList.Count - 1; i++)
                {
                    int maxMatchResult = 0;
                    int result = 0;
                    mappedList.Add(new mappedString(primaryList[i].orginalString, ""));
                    for (int j = i + 1; j < primaryList.Count; j++)
                    {
                        if (maxMatchResult < (result = compare_string(primaryList[i], primaryList[j])))
                        {
                            mappedList[i].secondaryString = primaryList[j].orginalString;
                            mappedList[i].percent_match = result;
                            maxMatchResult = result;
                        }
                    }

                    backgroundWorker1.ReportProgress((i * 100) / primaryList.Count);
                }
            }

            backgroundWorker1.ReportProgress(0);
            currentStatusMessage = "Saving...";

            string FileName = textBox4.Text;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel.Workbook wbexcel = excel.Workbooks.Open(FileName);
            Microsoft.Office.Interop.Excel.Worksheet ws = wbexcel.Sheets.Add() as Microsoft.Office.Interop.Excel.Worksheet;
            ws.Name = "mapped" + DateTime.Now.ToString("yyyyMMddHHmmss");

            ws.Cells[1, 1] = "primary list";
            ws.Cells[1, 2] = "Second list";
            ws.Cells[1, 3] = "Match level";
            for (int i = 0; i < mappedList.Count; i++)
            {
                ws.Cells[i + 2, 1] = mappedList[i].primaryString;
                ws.Cells[i + 2, 2] = mappedList[i].secondaryString;
                ws.Cells[i + 2, 3] = mappedList[i].percent_match;

                backgroundWorker1.ReportProgress((i * 100) / mappedList.Count);
            }

            mappedList.Clear();

            wbexcel.Save();

            wbexcel.Close();

            excel.Quit();

        }

        int resultshows = 0;
        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            int p = e.ProgressPercentage;
            p = Math.Min(100, Math.Max(0, p));
            p *= 10;

            progressBar1.Value = p;

            label4.Text = currentStatusMessage;

            //while (mappedList.Count - 1 > resultshows)
            //{
            //    if (resultshows > 13 && resultshows < 18)
            //    {
            //        int xx = 0;
            //    }
            //    dataGridView1.Rows.Add(mappedList[resultshows].primaryString, mappedList[resultshows].secondaryString, mappedList[resultshows].percent_match);
            //    resultshows++;
            //}
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            label4.Text = "";
            if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            button1.Enabled = true;
            button2.Enabled = true;
            button3.Enabled = true;

            progressBar1.Value = 1000;

            mappedList.Clear();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.facebok.com/mundhaSoft");
        }
    }

    public class mappedString
    {
        public string primaryString { get; private set; }
        public string secondaryString;
        public int percent_match = 0;

        public mappedString(string primary, string secondary)
        {
            this.primaryString = primary;
            this.secondaryString = secondary;
        }
    }

    public class comparableString
    {
        public string orginalString;
        string comparable_string;

        // property
        public int length { get { return this.comparable_string.Length; } }
        public char this[int i] { get { return this.comparable_string[i]; } }
        public string this[int i, int l] { get { return this.comparable_string.Substring(i, l); } }

        public int IndexOf(string str, int startIndex)
        {
            return this.comparable_string.IndexOf(str, startIndex);
        }

        // sort by length desc required
        // must be all letter in lower
        public static List<string> EscapeWords = new List<string>();
        public static List<string> EscapeWords2 = new List<string>();

        public comparableString(string comString)
        {
            this.orginalString = comString;
            this.comparable_string = comString.ToLower();

            for (int i = 0; i < comparableString.EscapeWords.Count; i++)
            {
                this.comparable_string = this.comparable_string.Replace(comparableString.EscapeWords[i], "");
            }
            for (int i = 0; i < comparableString.EscapeWords2.Count; i++)
            {
                if (comparableString.EscapeWords2[i].Length == 0) continue;
                this.comparable_string = this.comparable_string.Replace(comparableString.EscapeWords2[i], "");
            }
        }
    }
}
