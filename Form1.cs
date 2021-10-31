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
using QRCoder;
using SelectPdf;
using System.Text;
using System.Net;
using ProjectParserWF;
namespace ProjectParserWF
{
    public partial class Form1 : Form
    {
        private BindingSource dataSource = new BindingSource();
        List<AnnouncementInfo> ans = new List<AnnouncementInfo>();
        public Form1()
        {
            InitializeComponent();
            dataSource.DataSource = ans;
            dataGridView1.DataSource = dataSource;
            button1.BackColor = Color.Aqua;
            dataGridView1.Columns[4].Width += 50;
            dataGridView1.Columns[6].Width = 250;
            progressBar1.Visible = false;
            label2.Visible = false;
            AnnouncementInfo.html = File.ReadAllText("base.html", encoding: Encoding.UTF8);
            pictureBox1.Enabled = false;
        }
        public class AnnouncementInfo
        {
            public static string html { set; private get; }
            public string Id { get; set; }
            public string Name { get; set; }
            public string Place { get; set; }
            public string Price { get; set; }
            public string Description { get; set; }
            public string TimePublished { get; set; }
            public string Url { get; set; }
            public string Image { get; set; }
            public string UrlToQr(string name)
            {
                PayloadGenerator.Url urlPayload = new PayloadGenerator.Url(Url);
                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(urlPayload);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20);
                qrCodeImage.Save(name);
                return name;
            }
            public AnnouncementInfo Clone()
            {
                return new AnnouncementInfo()
                {
                    Id = this.Id,
                    Name = this.Name,
                    Place = this.Place,
                    Price = this.Price,
                    Description = this.Description,
                    TimePublished = this.TimePublished,
                    Url = this.Url,
                    Image = this.Image,
                };
            }
            public void SaveAsPdf(string name)
            {
                FileInfo qr = new FileInfo(UrlToQr("qr.png"));
                string ht = html.Replace("#IMG", Image);
                ht = ht.Replace("#ID", Id);
                ht = ht.Replace("#NAME", Name);
                ht = ht.Replace("#QR", qr.FullName);
                ht = ht.Replace("#DESC", Description);
                ht = ht.Replace("#PLACE", Place);
                ht = ht.Replace("#PRICE", Price);
                ht = ht.Replace("#TP", TimePublished);

                //Console.WriteLine(ht);
                HtmlToPdf htmlToPdf = new HtmlToPdf();
                htmlToPdf.Options.PdfPageSize = PdfPageSize.A4;
                htmlToPdf.Options.PdfPageOrientation = PdfPageOrientation.Portrait;
                PdfDocument pdf = htmlToPdf.ConvertHtmlString(ht, ".");
                pdf.Save(name + ".pdf");
                File.Delete(qr.FullName);
                pdf.Close();
                Console.WriteLine(name + " saved");
            }
        }
        void Write(string name, List<AnnouncementInfo> l)
        {
            List<string> names = new List<string>() { "Id", "Название", "Место", "Цена", "Описание", "Опубликовано", "Ссылка на объявление" };
            var memoryStream = new MemoryStream();
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(l), (typeof(DataTable)));
            progressBar1.Value = 0;
            progressBar1.Maximum = l.Count;
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
            Task.Run(new Action(() =>
            {
                Invoke(new Action(() =>
                {
                    button1.Enabled = false;
                }));
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
                        ai.Image = "";
                        var img = item.QuerySelector("img[data-src]");
                        if (img != null)
                        {
                            ai.Image = Regex.Replace(img.GetAttributeValue("data-src", "").Trim(), @"\s{2}", "");
                        }
                        var d = item.QuerySelector("span.date-posted");
                        if (d != null)
                        {
                            ai.TimePublished = Regex.Replace(d.InnerText.Trim(), @"\s{2}", "");
                        }
                        ai.Url = Regex.Replace(@"https://www.kijiji.ca" + item.GetAttributeValue("data-vip-url", ""), @"\s{2}", "");
                        ans.Add(ai);
                    }
                    Invoke(new Action(() =>
                    {
                        dataSource.ResetBindings(true);
                        dataGridView1.Refresh();
                    }));

                }
                Invoke(new Action(() =>
                {
                    dataGridView1.Rows.RemoveAt(0);
                    button1.Enabled = true;
                }));
            }));

        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ans.Count == 0)
            {
                MessageBox.Show("Вы не заполнили таблицу!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Вы не выбрали ни одного элемента!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                saveFileDialog1.DefaultExt = "xlsx";
                saveFileDialog1.Filter = "Excel file(*.xlsx;*xls)|*.xlsx;*.xls";
                string name = "";
                //do
                //{
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    button2.Enabled = false;
                    name = saveFileDialog1.FileName;
                    progressBar1.Value = 0;
                    progressBar1.Maximum = ans.Count;
                    progressBar1.Visible = true;
                    label2.Visible = true;
                    List<AnnouncementInfo> newl = new List<AnnouncementInfo>();
                    foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                    {
                        if (!newl.Contains(ans[cell.RowIndex]))
                        {
                            newl.Add(ans[cell.RowIndex]);
                        }
                    }
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        if (!newl.Contains(ans[row.Index]))
                        {
                            newl.Add(ans[row.Index]);
                        }
                    }
                    Write(name.Replace(".xlsx", ""), newl);
                    Task.Run(new Action(() =>
                    {
                        Thread.Sleep(2000);
                        if (name.Contains(".xlsx"))
                        {
                            System.Diagnostics.Process.Start(name);
                        }
                        else
                        {
                            System.Diagnostics.Process.Start(name + ".xlsx");
                        }

                    }));
                    button2.Enabled = true;
                }
                //break;
                //    }
                //} while (true);
                //Interaction.InputBox("Question?", "Title", "Default Text");


            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (ans.Count == 0)
            {
                MessageBox.Show("Вы не заполнили таблицу!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Вы не выбрали ни одного элемента!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
                {
                    List<AnnouncementInfo> newl = new List<AnnouncementInfo>();
                    foreach (DataGridViewCell cell in dataGridView1.SelectedCells)
                    {
                        if (!newl.Contains(ans[cell.RowIndex]))
                        {
                            newl.Add(ans[cell.RowIndex]);
                        }
                    }
                    foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                    {
                        if (!newl.Contains(ans[row.Index]))
                        {
                            newl.Add(ans[row.Index]);
                        }
                    }

                    progressBar1.Value = 0;
                    progressBar1.Maximum = newl.Count;
                    progressBar1.Visible = true;
                    label2.Visible = true;
                    Task.Run(() =>
                    {
                        Invoke(new Action(() =>
                        {
                            button3.Enabled = false;
                        }));
                        foreach (var an in newl)
                        {
                            an.SaveAsPdf(folderBrowserDialog1.SelectedPath + "\\"+ (newl.IndexOf(an)+1).ToString() + ".pdf");
                            Invoke(new Action(() =>
                            {
                                progressBar1.Value++;
                            }));
                        }
                        System.Diagnostics.Process.Start(folderBrowserDialog1.SelectedPath);
                        Invoke(new Action(() =>
                        {
                            button3.Enabled = true;
                        }));

                    });


                }
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            pictureBox1.Enabled = true;
            var request = WebRequest.Create(ans[e.RowIndex].Image);
            using (var response = request.GetResponse())
            using (var stream = response.GetResponseStream())
            {
                pictureBox1.Image = Bitmap.FromStream(stream);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form2 form;
            
            if (dataGridView1.SelectedRows.Count>0)
            {
               form= new Form2(ans[dataGridView1.SelectedRows[0].Index]);
                form.ShowDialog();
                if (form.save)
                {
                    ans[dataGridView1.SelectedRows[0].Index] = form.selected;
                    dataGridView1.Refresh();
                }
                
            }
            else  if(dataGridView1.SelectedCells.Count > 0)
            {
                form = new Form2(ans[dataGridView1.SelectedCells[0].RowIndex]);
                form.ShowDialog();
                if (form.save)
                {
                    ans[dataGridView1.SelectedCells[0].RowIndex] = form.selected;
                    dataGridView1.Refresh();
                }

            }

        }
    }
}
