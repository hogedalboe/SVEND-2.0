using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace SVEND_2._0
{
    public partial class Form2_Choose_template : Form
    {
        public Form2_Choose_template()
        {
            InitializeComponent();
        }

        // Global variables
        string files_certificate_template_matches = Directory.GetCurrentDirectory() + @"\settings\files_certificate_template_matches"; // To store potential matches for a specialization's certificate templates
        string folder_certificate_templates = Directory.GetCurrentDirectory() + @"\settings\folder_certificate_templates";
        string folder_temp = Directory.GetCurrentDirectory() + @"\temp\" + Environment.UserName + @"\";

        private void Form2_Choose_template_Load(object sender, EventArgs e)
        {
            this.TopMost = true;

            // Make size constant
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;

            // Reading current specialization to label
            label2.Text = File.ReadAllText(folder_temp + "dialog.tmp", Encoding.GetEncoding(1252));

            // Reading possible template matches to listbox
            string[] matches = File.ReadAllLines(files_certificate_template_matches, Encoding.GetEncoding(1252));
            if (matches.Length > 0)
            {
                foreach (string match in matches)
                {
                    listBox1.Items.Add(Path.GetFileNameWithoutExtension(match).ToString());
                }
            }
            // If no possible template matches have been found, all docx files in the directory are loaded to listBox1 and files_certificate_template_matches
            else
            {
                string tmp_folder_path = File.ReadAllText(folder_certificate_templates, Encoding.GetEncoding(1252)).Replace(@"C:\Users\%USERPROFILE%", @"C:\Users\" + Environment.UserName);
                DirectoryInfo tmp_directory_info = new DirectoryInfo(tmp_folder_path);
                FileInfo[] tmp_potential_templates = tmp_directory_info.GetFiles("*.docx");

                foreach (FileInfo match in tmp_potential_templates)
                {
                    // Add full template path to files_certificate_template_matches
                    File.AppendAllText(files_certificate_template_matches, match.FullName + Environment.NewLine, Encoding.GetEncoding(1252)); // [CL:1]

                    // Add filename to listbox
                    listBox1.Items.Add(match.Name);
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem != null)
            {
                // Reading possible template matches to array
                string[] matches = File.ReadAllLines(files_certificate_template_matches, Encoding.GetEncoding(1252));

                // Clearing the possible templates
                File.WriteAllText(files_certificate_template_matches, string.Empty, Encoding.GetEncoding(1252));

                // Finding the full path of the chosen file
                foreach (string match in matches)
                {
                    if (match.Contains(listBox1.SelectedItem.ToString()))
                    {
                        File.WriteAllText(files_certificate_template_matches, match, Encoding.GetEncoding(1252));
                    }
                }

                // Indicates that the form has been closed properly
                File.Delete(folder_temp + "dialog.tmp");

                // Close form
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Clear the matches in files_certificate_tenplate_matches and close form
            File.WriteAllText(files_certificate_template_matches, string.Empty, Encoding.GetEncoding(1252));
            File.Delete(folder_temp + "dialog.tmp"); // Indicates that the form has been closed properly
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }
    }
}
