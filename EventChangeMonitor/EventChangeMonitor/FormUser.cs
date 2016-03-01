using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ActivityMonitor
{
    public partial class FormUser : Form
    {

        public FormUser(string app)
        {
            InitializeComponent();
            this.appName = app;
            getUserSelection();
        }

        private void FormUser_Load(object sender, EventArgs e)
        {

        }

        private void getUserSelection()
        {
            lblHeader.Text = "The process " + appName + " belongs to ?";
            Dictionary<string, List<string>> buckets = DBConnector.getInstance().getBuckets();
            String[] bucketList = new String[buckets.Count]; 
            
            int i = 0;
            foreach (KeyValuePair<string, List<string>> pair in buckets)
            {
                bucketList[i++] = pair.Key;
            }

            radioButtons = new System.Windows.Forms.RadioButton[buckets.Count];

            for (int j = 0; j < buckets.Count; j++)
            {
                radioButtons[j] = new RadioButton();
                radioButtons[j].Text = bucketList[j];
                radioButtons[j].Location = new System.Drawing.Point(20, 50 + j * 20);
                this.pnlMain.Controls.Add(radioButtons[j]);
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void lblHeader_Click(object sender, EventArgs e)
        {
            lblHeader.Text = "sfgsdgf";
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < radioButtons.Length; i++)
            {
                if (radioButtons[i].Checked)
                {
                    this.bucketName = radioButtons[i].Text;
                    DBConnector.getInstance().addAppToBucket(bucketName, appName);
                    break;
                }
            }
        }
    }
}
