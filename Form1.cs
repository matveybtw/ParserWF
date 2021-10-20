using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Fizzler.Systems.HtmlAgilityPack;
using HtmlAgilityPack;
using System.Drawing;
using System.Threading.Tasks;
using System.IO;
using System.Data;
using Newtonsoft.Json;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Threading;

namespace ProjectParserWF
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            button1.BackColor = Color.Aqua;
            dataGridView1.Columns[4].Width += 50;
            dataGridView1.Columns[6].Width = 250;
            progressBar1.Visible = false;
            label2.Visible = false;
        }
        public class AnnouncementInfo
        {
            public string Id { get; set; }
            public string Name { get; set; }
            public string Place { get; set; }
            public string Price { get; set; }
            public string Description { get; set; }
            public string TimePublished { get; set; }
            public string Url { get; set; }
            public AnnouncementInfo()
            {

            }
        }
        List<AnnouncementInfo> ans = new List<AnnouncementInfo>();
        private void AddToTable(AnnouncementInfo ai)
        {
            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridView1);
            row.Cells[0].Value = ai.Id;
            row.Cells[1].Value = ai.Name;
            row.Cells[2].Value = ai.Place;
            row.Cells[3].Value = ai.Price;
            row.Cells[4].Value = ai.Description;
            row.Cells[5].Value = ai.TimePublished;
            row.Cells[6].Value = ai.Url;
            dataGridView1.Rows.Add(row);

        }
        void Write(string name)
        {
            List<string> names = new List<string>() { "Id", "Название", "Место", "Цена", "Описание", "Опубликовано", "Ссылка на объявление" };
            var memoryStream = new MemoryStream();
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(ans), (typeof(DataTable)));
            progressBar1.Value = 0;
            using (var fs = new FileStream(name + ".xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheesh");
                List<String> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                for (int i = 0; i < names.Count; i++)
                {
                    columns.Add(table.Columns[i].ColumnName);
                    row.CreateCell(i).SetCellValue(names[i]);
                }
                int rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    int cellIndex = 0;
                    row = excelSheet.CreateRow(rowIndex);
                    foreach (String col in columns)
                    {
                        row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        cellIndex++;
                    }
                    rowIndex++;
                    progressBar1.Value++;
                }
                workbook.Write(fs);
            }

        }
        private void button1_Click(object sender, EventArgs e)
        {
            progressBar1.Visible = false;
            label2.Visible = false;
            dataGridView1.Rows.Clear();
            ans.Clear();
            button1.Enabled = false;
            Task.Run(new Action(() =>
            {
                for (int i = 0; i < numericUpDown1.Value; i++)
                {
                    var url = @"https://www.kijiji.ca/b-canada/iphone/page-" + (i + 1).ToString() + @"/k0l0?rb=true&dc=true";
                    var client = new HtmlWeb();
                    var html = client.Load(url);
                    var nodes = html.DocumentNode.QuerySelectorAll("div[data-listing-id]").ToList();
                    foreach (var item in nodes)
                    {
                        AnnouncementInfo ai = new AnnouncementInfo();
                        ai.Id = Regex.Replace(item.GetAttributeValue("data-listing-id", "").Trim(), @"\s{2}", "");
                        ai.Name = Regex.Replace(item.QuerySelector("div.title").InnerText.Trim(), @"\s{2}", "");
                        ai.Place = Regex.Replace(item.QuerySelector("div.location").InnerText.Trim(), @"\s{2}", "");
                        ai.Price = Regex.Replace(item.QuerySelector("div.price").InnerText.Trim(), @"\s{2}", "");
                        ai.Description = Regex.Replace(item.QuerySelector("div.description").InnerText.Trim(), @"\s{2}", "");
                        var d = item.QuerySelector("span.date-posted");
                        if (d != null)
                        {
                            ai.TimePublished = Regex.Replace(d.InnerText.Trim(), @"\s{2}", "");
                        }
                        ai.Url = Regex.Replace(@"https://www.kijiji.ca" + item.GetAttributeValue("data-vip-url", ""), @"\s{2}", "");
                        ans.Add(ai);
                        AddToTable(ai);
                    }
                }
            }));
            button1.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ans.Count==0)
            {
                MessageBox.Show("Вы не заполнили таблицу!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                progressBar1.Value = 0;
                progressBar1.Maximum = ans.Count;
                progressBar1.Visible = true;
                label2.Visible = true;
                Write("Announcments");
                Task.Run(new Action(() =>
                {
                    Thread.Sleep(2000);
                    System.Diagnostics.Process.Start("Announcments.xlsx");
                }));
                
            }
        }
    }
}
