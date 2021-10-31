using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ProjectParserWF;
namespace ProjectParserWF
{
    public partial class Form2 : Form
    {
        public Form2(Form1.AnnouncementInfo s)
        {
            InitializeComponent();
            selected = s.Clone();
            bind.Add(selected);
            textBox1.DataBindings.Clear();
            textBox1.DataBindings.Add("Text", bind[0], "Id", false, DataSourceUpdateMode.OnPropertyChanged);
            textBox2.DataBindings.Clear();
            textBox2.DataBindings.Add("Text", bind[0], "Name", false, DataSourceUpdateMode.OnPropertyChanged);
            textBox3.DataBindings.Clear();
            textBox3.DataBindings.Add("Text", bind[0], "Place", false, DataSourceUpdateMode.OnPropertyChanged);
            textBox4.DataBindings.Clear();
            textBox4.DataBindings.Add("Text", bind[0], "Description", false, DataSourceUpdateMode.OnPropertyChanged);
            textBox5.DataBindings.Clear();
            textBox5.DataBindings.Add("Text", bind[0], "TimePublished", false, DataSourceUpdateMode.OnPropertyChanged);
            textBox6.DataBindings.Clear();
            textBox6.DataBindings.Add("Text", bind[0], "Url", false, DataSourceUpdateMode.OnPropertyChanged);
            textBox7.DataBindings.Clear();
            textBox7.DataBindings.Add("Text", bind[0], "Image", false, DataSourceUpdateMode.OnPropertyChanged);
        }
        public Form1.AnnouncementInfo selected;
        BindingList<Form1.AnnouncementInfo> bind=new BindingList<Form1.AnnouncementInfo>();
        public bool save { get; set; }
        private void button1_Click(object sender, EventArgs e)
        {
            save = true;
            this.Close();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            save = false;
            this.Close();
        }
    }
}
