using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Creating_Writing_DGV_TO_EXCEL
{
    public partial class FormSetting : Form
    {
        private Config config = new Config();

        public FormSetting()
        {
            InitializeComponent();
            readSetting();
        }

        private void readSetting()
        {
            try
            {
                string path = "config.txt";
                string[] lines = File.ReadAllLines(path, Encoding.UTF8);

                config.servername = lines[0];
                config.userid = lines[1];
                config.password = lines[2];

                txtServer.Text = config.servername;
                txtUserID.Text = config.userid;
                txtPassword.Text = config.password;
            }
            catch
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = "config.txt";
            StreamWriter sw = new StreamWriter(path);
            sw.WriteLine(txtServer.Text);
            sw.WriteLine(txtUserID.Text);
            sw.WriteLine(txtPassword.Text);
            sw.Close();
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (txtServer.Text == "")
            {
                MessageBox.Show("ตั้งค่าก่อนครับ");
                return;
            }
            try
            {
                string con = "server=" + config.servername + ";userid=" + config.userid + ";password=" + config.password + ";";
                MySqlConnection conn = new MySqlConnection(con);
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    MessageBox.Show("Connection is OK!!!");
                    return;
                }
                else
                {
                    MessageBox.Show("Please check connection carefully..");
                    return;
                }
            }
            catch
            {
            }
        }
    }
}