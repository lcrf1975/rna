using System;
using System.Configuration;
using System.Diagnostics;
using System.ServiceModel.Configuration;
using System.Windows.Forms;

namespace RNA_Tools
{
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        // Loads the application configuration file
        private void Form3_Load(object sender, EventArgs e)
        {
            textBox1.Text = ConfigurationManager.AppSettings["Login"];
            textBox2.Text = ConfigurationManager.AppSettings["Password"];

            Configuration configuration = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.PerUserRoamingAndLocal);
            ServiceModelSectionGroup serviceModelSectionGroup = ServiceModelSectionGroup.GetSectionGroup(configuration);
            ClientSection clientSection = serviceModelSectionGroup.Client;
            var E1 = clientSection.Endpoints[0];
            textBox3.Text = E1.Address.ToString();
            var E2 = clientSection.Endpoints[1];
            textBox4.Text = E2.Address.ToString();
            var E3 = clientSection.Endpoints[2];
            textBox5.Text = E3.Address.ToString();
            var E4 = clientSection.Endpoints[3];
            textBox6.Text = E4.Address.ToString();
            var E5 = clientSection.Endpoints[4];
            textBox7.Text = E5.Address.ToString();
        }

        // Opens the application configuration file in the Notepad.exe editor
        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                Process ExternalProcess = new Process();
                ExternalProcess.StartInfo.FileName = "Notepad.exe";
                ExternalProcess.StartInfo.Arguments = AppDomain.CurrentDomain.BaseDirectory.ToString() + "\\RNA_Tools.exe.config";
                ExternalProcess.StartInfo.WindowStyle = ProcessWindowStyle.Maximized;
                ExternalProcess.Start();
                ExternalProcess.WaitForExit();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Error");
            }
        }
    }
}
