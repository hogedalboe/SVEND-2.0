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
using System.Drawing.Printing;
using Word = Microsoft.Office.Interop.Word;
using System.Diagnostics;

/* ABOUT:
 *      Developed by: Nikolaj Høgedal Boe
 *      For: Industriens Uddannelser
 *      https://iu.dk/om-os/iu-organisation/about-us/
 * 
 * 
 * /

/* CHANGELOG:
 * 
 * 2019-03-05: 
 *      - More user friendly review of errors in the certificates after printing.
 *      - User oriented handling of missing or void student data via log_report() and richTextBox: [CL:2]
 *      - Adding bookmarks to all methods in Form1.
 *
 * 2019-02-28: 
 *      - Optimization of data exchange between Form1 and Form2 upon Form2 close: [CL:1]
 */

/* TO DO:
 *      - Fjernelse af karakter1 ved Hanne Doe medfører ikke en fejl, som det ellers er forventet jf. [CL:2]
 *      - Fill file_documentation with some pretty html.
 *      - Add automatic school address if the student's practical training company is a known school.
 */

namespace SVEND_2._0
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        // Global variables

        string userprofile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

        string file_dictionary_specializations = Directory.GetCurrentDirectory() + @"\data\dictionary_specializations";
        Dictionary<string, string> dictionary_specializations = new Dictionary<string, string> { }; // Specializations along with indication of printer paper

        string folder_certificate_templates = Directory.GetCurrentDirectory() + @"\settings\folder_certificate_templates";
        string file_letter_template = Directory.GetCurrentDirectory() + @"\settings\file_letter_template";
        string folder_csv_students = Directory.GetCurrentDirectory() + @"\settings\folder_csv_students";
        string folder_save_certificates = Directory.GetCurrentDirectory() + @"\settings\folder_save_certificates";

        string file_certificate_template = ""; // To be set with find_certificate_template()

        bool bool_print_letters = true; // Whether there should be made a letter to be attached with each certificate

        string file_mergefield_specialization = Directory.GetCurrentDirectory() + @"\data\mergefield_specialization";
        string mergefield_specialization = ""; // To be used for setting the specialization variable for the main loop's iterations (iterating through all specializations and matching them with current student's specialization)

        string file_mergefield_student_name = Directory.GetCurrentDirectory() + @"\data\mergefield_student_name";
        string mergefield_student_name = ""; // Used to fill datagridview3 (letters) and to save certificates with the students' names in the files

        string file_user_setup_teamsvendeprover = Directory.GetCurrentDirectory() + @"\settings\user_setup_teamsvendeprover";
        bool bool_user_setup_teamsvendeprover = false;

        string file_functionality_framework = Directory.GetCurrentDirectory() + @"\settings\functionality_framework";
        string functionality_framework = "";

        string folder_backups = Directory.GetCurrentDirectory() + @"\backups\";

        string file_print_paper = Directory.GetCurrentDirectory() + @"\settings\print_paper";
        string file_mergefields = Directory.GetCurrentDirectory() + @"\data\mergefields";

        string save_certificates = Directory.GetCurrentDirectory() + @"\settings\bool_save_certificates";
        bool bool_save_certificates = false;

        bool form1_fully_loaded = false; // Used to deactivate textBox_textChanged event until form is fully loaded

        string folder_logs = Directory.GetCurrentDirectory() + @"\log\";
        string file_log_error = Directory.GetCurrentDirectory() + @"\log\" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + " (" + Environment.UserName + ") Errors.log";
        string file_log_report = Directory.GetCurrentDirectory() + @"\log\" + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + " (" + Environment.UserName + ") Report.log";

        string folder_temp = Directory.GetCurrentDirectory() + @"\temp\" + Environment.UserName + @"\";

        bool bool_teamsvendeprover = false;

        string files_certificate_template_matches = Directory.GetCurrentDirectory() + @"\settings\files_certificate_template_matches"; // To store potential matches for a specialization's certificate templates

        Dictionary<string, List<string>> dictionary_print = new Dictionary<string, List<string>> { }; // Create dictionary to store which paper to print each certificate on
        int int_files_to_print = 0; // Simply to set a maximum for the progress bar when printing

        string file_documentation = Directory.GetCurrentDirectory() + @"\documentation\documentation.html";

        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        private void Form1_Load(object sender, EventArgs e)
        {   
            try
            {
                load_all();

                // Create backup from current setting if form i loaded succesfully
                if (form1_fully_loaded)
                {
                    create_backup();
                }
            }
            catch (Exception ex)
            {
                log_error(ex.ToString());

                // If loading of form1 fails, choose a backup and try again
                MessageBox.Show("Noget gik galt i indlæsningen af tidligere indstillinger og opsætning. Tryk OK og vælg en af mapperne med angivelsen 'BACKUP YYYY-MM-DD HHMMSS'.\n" +
                    "\n" +
                    "Fejlbeskrivelse:\n" +
                    ex.ToString());

                choose_backup();

                load_all();

                // Create backup from current setting if form i loaded succesfully
                if (form1_fully_loaded)
                {
                    log_error("SVEND 2.0: Form was loaded succesfully after restoring with a backup.");

                    create_backup();
                }
            }
        }

        private void load_all()
        {
            // PLACE FIRST
            form1_fully_loaded = false;

            // Set back color of tabs
            var backcolor = Color.Snow;
            tabPage1.BackColor = backcolor;
            tabPage2.BackColor = backcolor;
            tabPage3.BackColor = backcolor;
            tabPage5.BackColor = backcolor;
            tabPage7.BackColor = backcolor;
            tabPage7.BackColor = backcolor;

            // Make size constant
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            MaximumSize = new Size(int.MaxValue, 1000);

            // Hide
            dataGridView3.Hide(); // Letters
            label34.Hide(); // Info about letter datagridviews
            button11.Hide();
            progressBar1.Hide();
            label22.Hide();
            label22.Text = "";
            label1.Hide();
            richTextBox3.Hide();
            button12.Hide();
            button13.Hide();

            // Load documentation
            webBrowser1.Url = new Uri(String.Format("file:///{0}", file_documentation));

            // Get user specific settings
            bool_user_setup_teamsvendeprover = get_binary_file_setting(file_user_setup_teamsvendeprover);
            if (bool_user_setup_teamsvendeprover)
            {
                checkBox1.Checked = true;
            }
            else
            {
                checkBox1.Checked = false;
            }

            // Read standard printer
            PrinterSettings printerSettings = new PrinterSettings();
            label21.Text = printerSettings.PrinterName;

            // Read bool_save_certificates: Determines if the final certificates should be saved in a folder
            string str_save_certificates = File.ReadAllText(save_certificates, Encoding.GetEncoding(1252));
            if (str_save_certificates == "Y")
            {
                bool_save_certificates = true;
                button7.Text = "Ja";
            }
            else if (str_save_certificates == "N")
            {
                bool_save_certificates = false;
                button7.Text = "Nej";
            }

            // Read functionality_framework
            functionality_framework = File.ReadAllText(file_functionality_framework, Encoding.GetEncoding(1252));
            button9.Text = functionality_framework;

            // Read mergefield_specialization (essential string variable for iterating specializations)
            mergefield_specialization = File.ReadAllText(file_mergefield_specialization, Encoding.GetEncoding(1252));
            textBox5.Text = mergefield_specialization;

            // Read mergefield_student_name
            mergefield_student_name = File.ReadAllText(file_mergefield_student_name, Encoding.GetEncoding(1252));
            textBox6.Text = mergefield_student_name;

            // Read specializations from file_dictionary_specializations to dictionary_specialization
            string[] array_specializations_print = File.ReadAllLines(file_dictionary_specializations, Encoding.GetEncoding(1252));
            dictionary_specializations.Clear();
            for (int i = 0; i < array_specializations_print.Length; i++)
            {
                string[] tmp = array_specializations_print[i].Split(';');
                dictionary_specializations[tmp[0]] = tmp[1];
            }

            // Read specializations to tabpage with specializations
            foreach (var entry in dictionary_specializations)
            {
                richTextBox1.AppendText(entry.Key + Environment.NewLine);
            }

            // Read specializations to printer settings grid
            load_printer_settings();

            // Read mergefields to mergefield grid
            load_mergefield_settings();

            // Read csv folder
            read_settings(label6, folder_csv_students);

            // Read certificate template folder
            read_settings(label7, folder_certificate_templates);

            // Read letter template file
            string path_letter_template = File.ReadAllText(file_letter_template, Encoding.GetEncoding(1252));
            path_letter_template = path_letter_template.Replace(@"C:\Users\%USERPROFILE%", @"C:\Users\" + Environment.UserName);
            if (path_letter_template == "")
            {
                // If file is empty, then set the checkbox to true an delete text in label8
                checkBox2.Checked = true;
                label8.Text = "";
                bool_print_letters = false;
            }
            else
            {
                if (File.Exists(path_letter_template))
                {
                    label8.Text = path_letter_template;
                }
                else
                {
                    label8.Text = userprofile + @"\Desktop";
                }
                bool_print_letters = true;
            }

            // Read folder for saving certificates
            read_settings(label10, folder_save_certificates);

            // Get paper designations
            get_paper_designation(file_print_paper, textBox1);
            get_paper_designation(file_print_paper, textBox2);
            get_paper_designation(file_print_paper, textBox3);
            get_paper_designation(file_print_paper, textBox4);

            // Remove the oldest logs
            int max_logs = 50;
            int number_of_logs = Directory.GetFiles(folder_logs).Length;
            if (number_of_logs > max_logs)
            {
                for (int i = 0; i < number_of_logs - max_logs; i++)
                {
                    FileSystemInfo fileInfo = new DirectoryInfo(folder_logs).GetFileSystemInfos().OrderBy(fi => fi.CreationTime).First(); // https://stackoverflow.com/questions/44690815/how-to-delete-oldest-folder-created-from-local-disk-using-c-sharp
                    File.Delete(fileInfo.FullName);
                }
            }

            // Create user specific temp folder
            Directory.CreateDirectory(folder_temp);

            // Clear the temp folder (it might already exist from previous run, if files wasn't closed properly)
            System.IO.DirectoryInfo dirInfo_folder_temp = new DirectoryInfo(folder_temp);
            try
            {
                foreach (FileInfo file in dirInfo_folder_temp.GetFiles())
                {
                    file.Delete();
                }
            }
            catch
            {
                // If the file is open, try to kill its process and close it again
                foreach (FileInfo file in dirInfo_folder_temp.GetFiles())
                {
                    DateTime dt = File.GetLastAccessTime(file.FullName);
                    kill_process("WINWORD", dt);
                    try
                    {
                        file.Delete();
                    }
                    catch
                    {
                        //
                    }
                }
            }

            //PLACE LAST
            form1_fully_loaded = true;
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        //////////////////// .NET  ////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        private void dotNET()
        {
            button12.Hide();
            
            // Progressbar
            progressBar1.Show();
            label22.Show();
            
            // Clear the temp folder if it exists, otherwise create the directory
            if (Directory.Exists(folder_temp))
            {
                System.IO.DirectoryInfo dirInfo_folder_temp = new DirectoryInfo(folder_temp);
                try
                {
                    foreach (FileInfo file in dirInfo_folder_temp.GetFiles())
                    {
                        file.Delete();
                    }
                }
                catch
                {
                    foreach (FileInfo file in dirInfo_folder_temp.GetFiles())
                    {
                        // If the file is open, try to kill its process and close it again
                        DateTime dt = File.GetLastAccessTime(file.FullName);
                        kill_process("WINWORD", dt);
                        try
                        {
                            file.Delete();
                        }
                        catch
                        {
                            //
                        }
                    }
                }
            }
            else
            {
                // Create user specific temp folder
                Directory.CreateDirectory(folder_temp);
            }

            // Get tailored settings for Team Svendeprøver
            string string_teamsvendeprover = File.ReadAllText(file_user_setup_teamsvendeprover, Encoding.GetEncoding(1252));
            if (string_teamsvendeprover == "1")
            {
                bool_teamsvendeprover = true;
            }

            // Clear current file_log_report from previous run report
            File.WriteAllText(file_log_report, string.Empty);

            // Read csv files to datatable
            //
            //// See if header_csv_specialization can be replaced with mergefield_specialization (global variable)
            // 
            string header_csv_specialization = File.ReadAllText(file_mergefield_specialization, Encoding.GetEncoding(1252)); // Used for avoiding rows with empty specialization
            int column_header_csv_specialization = 0; // Used for avoiding rows with empty specialization, as well as finding the specialization column in the main loop
            string csv_folder = File.ReadAllText(folder_csv_students, Encoding.GetEncoding(1252)).Replace(@"C:\Users\%USERPROFILE%", @"C:\Users\" + Environment.UserName); // The directory path as string with files containing students and their data
            string[] csv_files = Directory.GetFiles(csv_folder);
            DataTable datatable_csv = new DataTable(); // https://stackoverflow.com/questions/1050112/how-to-read-a-csv-file-into-a-net-datatable
            try
            {
                // Add header once
                using (StreamReader sr = new StreamReader(csv_files[0], Encoding.GetEncoding(1252)))
                {
                    string[] headers = sr.ReadLine().Split(';');
                    for (int i = 0; i < headers.Length; i++)
                    {
                        datatable_csv.Columns.Add(headers[i]);

                        // Get the column number of the column with specializations, so that rows with empty specializations can be remowed
                        if (headers[i] == header_csv_specialization)
                        {
                            column_header_csv_specialization = i;
                        }
                    }

                    // Add rows from first file
                    while (!sr.EndOfStream)
                    {
                        string[] cells = sr.ReadLine().Split(';');
                        DataRow dr = datatable_csv.NewRow();
                        for (int i = 0; i < headers.Length; i++)
                        {
                            dr[i] = cells[i];
                        }
                        datatable_csv.Rows.Add(dr);
                    }
                }
                // Add rows from remaining rows
                for (int j = 1; j < csv_files.Length; j++)
                {
                    using (StreamReader sr = new StreamReader(csv_files[j], Encoding.GetEncoding(1252)))
                    {
                        bool ignore_header_row = true;
                        while (!sr.EndOfStream)
                        {
                            string[] cells = sr.ReadLine().Split(';');
                            DataRow dr = datatable_csv.NewRow();
                            for (int i = 0; i < cells.Length; i++)
                            {
                                dr[i] = cells[i];
                            }

                            // Ignore header row
                            if (ignore_header_row)
                            {
                                ignore_header_row = false;
                            }
                            else
                            {
                                datatable_csv.Rows.Add(dr); // Add row
                            }
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                log_error(exception.ToString());
                exit_and_cleanup();
                return;
            }

            // Length of datatable with students determines the maximum range of the progresbar
            progressBar1.Maximum = datatable_csv.Rows.Count;

            // Fill students from datatable to datagridview3, so that user can define which certificates should be printed with a letter
            if (bool_print_letters)
            {
                dataGridView3.Rows.Clear();

                try
                {
                    for (int i = 0; i < datatable_csv.Rows.Count; i++)
                    {
                        int name_column = 0;
                        int specialization_column = 0; // Double declaration: column_header_csv_specialization

                        foreach (DataColumn datatable_column in datatable_csv.Columns)
                        {
                            // Get the column with student name
                            if (datatable_column.ColumnName == mergefield_student_name)
                            {
                                name_column = datatable_column.Ordinal; // https://stackoverflow.com/questions/11340264/get-index-of-datatable-column-with-name
                            }
                            // Get the column with the specialization
                            else if (datatable_column.ColumnName == mergefield_specialization)
                            {
                                specialization_column = datatable_column.Ordinal;
                            }
                        }

                        DataGridViewRow grid_row = (DataGridViewRow)dataGridView3.Rows[i].Clone();

                        // See if datagridview contains a student with the same name and specialization
                        string dublicate_indicator = "";
                        foreach (DataGridViewRow row in dataGridView3.Rows)
                        {
                            if (row.Cells["dataGridViewTextBoxColumn1"].Value != null)
                            {
                                if (row.Cells["dataGridViewTextBoxColumn1"].Value.ToString() == datatable_csv.Rows[i].ItemArray[name_column].ToString())
                                {
                                    dublicate_indicator = "(1)";

                                    // See if a dublicate already exists
                                    foreach (DataGridViewRow row_inner in dataGridView3.Rows)
                                    {
                                        if (row_inner.Cells["dataGridViewTextBoxColumn1"].Value != null)
                                        {
                                            if (row_inner.Cells["dataGridViewTextBoxColumn1"].Value.ToString() == datatable_csv.Rows[i].ItemArray[name_column].ToString() + dublicate_indicator)
                                            {
                                                dublicate_indicator = "(2)";
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // Student column
                        grid_row.Cells[0].Value = datatable_csv.Rows[i].ItemArray[name_column] + dublicate_indicator;

                        // Specialization column
                        grid_row.Cells[1].Value = datatable_csv.Rows[i].ItemArray[specialization_column];

                        dataGridView3.Rows.Add(grid_row);

                        // Check all checkboxes as default
                        grid_row.Cells[2].Value = true;
                    }
                }
                catch (Exception exception)
                {
                    MessageBox.Show(exception.ToString());
                    log_error(exception.ToString());
                    exit_and_cleanup();
                    return;
                }
            }

            // Remove rows from datatable, if they don't contain a specialization in the column dedicated for the mergefield specialization
            for (int i = 0; i < datatable_csv.Rows.Count; i++)
            {
                if (datatable_csv.Rows[i].ItemArray[column_header_csv_specialization].ToString() == "")
                {
                    datatable_csv.Rows[i].Delete();
                }
            }

            // Get defined specializations from file
            string[] defined_specializations = File.ReadAllLines(file_dictionary_specializations, Encoding.GetEncoding(1252));
            for (int i = 0; i < defined_specializations.Length; i++)
            {
                defined_specializations[i] = defined_specializations[i].Replace(defined_specializations[i].Substring(defined_specializations[i].Length - 2), "");
            }

            // Read specializations to array and check if any of the students have specializations which are not specified
            string[] specializations_datatable_csv = datatable_csv.Rows.OfType<DataRow>().Select(k => k[column_header_csv_specialization].ToString()).ToArray(); // Array of the student's specializations in the csv files (datatable)
            List<string> nondefined_specializations = new List<string>(); // The list to read non-defined specializations to
            foreach (string student_specialization in specializations_datatable_csv)
            {
                bool exist_in_defined_specializations = false;

                foreach (string defined_specialization in defined_specializations)
                {
                    if (student_specialization == defined_specialization)
                    {
                        exist_in_defined_specializations = true;
                    }
                }

                if (exist_in_defined_specializations == false)
                {
                    if (student_specialization != "")
                    {
                        nondefined_specializations.Add(student_specialization);
                    }
                }
            }

            // Write the undefined specializations to report
            if (nondefined_specializations.Count > 0)
            {
                log_report("Følgende specialer er ikke definerede under fanen 'Specialer', og der kan derfor ikke genereres uddannelsesbeviser for de pågældende elever:" + Environment.NewLine);
                foreach (string tmp_specialization in nondefined_specializations)
                {
                    log_report("\t- " + tmp_specialization + Environment.NewLine);
                }
                log_report(Environment.NewLine);
            }

            ///////////////////////////////////////
            // MAIN LOOP //////////////////////////
            ///////////////////////////////////////
            // The magic happens here /////////////
            ///////////////////////////////////////

            // Prepare Word app
            DateTime dtWordStart = DateTime.Now;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            app.Visible = false;
            /*
            app.Selection.Find.ClearFormatting();
            app.Selection.Find.Replacement.ClearFormatting();
            */

            // Iterate each defined specialization
            foreach (string defined_specialization in defined_specializations)
            {
                // If the defined specialization exists in the datatable, do stuff
                if (specializations_datatable_csv.Contains(defined_specialization))
                {
                    // Add the specialization as key in dictionary_print
                    dictionary_print.Add(defined_specialization, new List<string>());

                    // Team Svendeprover: Check if 'plastmager' is among the specializations
                    if (bool_teamsvendeprover)
                    {
                        if (defined_specialization.Contains("Plastmager"))
                        {
                            MessageBox.Show("Du skal være opmærksom på, at der er en eller flere plastmager-elever i dine flettefiler. " + Environment.NewLine +
                                "Nogle af plastmager-uddannelsens specialer eksisterer ikke i EASY, og du vil derfor kun kunne se disse specialer på den afsluttende skole- og praktikerklæring. " +
                                "Hvis du har mere end én type plastmager-profiler i dine flettefiler, er det derfor meget vigtigt, " +
                                "at du afbryder SVEND og derefter kun kører programmet med én type plastmager-profil ad gangen. " +
                                "Ellers vil den bevisskabelon, du vælger, bliver anvendt for alle plastmager-profilerne, selvom disse kan være forskellige.");
                        }
                    }

                    // Find matching certificate template file for current specialization
                    file_certificate_template = find_certificate_template(defined_specialization);
                    if (file_certificate_template == null)
                    {
                        exit_and_cleanup();
                        return;
                    }

                    // Copy the letter template to the folder_temp
                    string path_letter_template = File.ReadAllText(file_letter_template, Encoding.GetEncoding(1252)).Replace(@"C:\Users\%USERPROFILE%", @"C:\Users\" + Environment.UserName);
                    string path_letter_template_copy = folder_temp + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + " LETTER TEMPLATE " + Environment.UserName + " " + defined_specialization + ".docx";
                    if (bool_print_letters)
                    {
                        File.Copy(path_letter_template, path_letter_template_copy);
                    }

                    // Iterate each student row in the datatable
                    for (int i = 0; i < datatable_csv.Rows.Count; i++)
                    {
                        // If the student specialization equals the defined specialization in the current iteration, proceed
                        if (datatable_csv.Rows[i].ItemArray[column_header_csv_specialization].ToString() == defined_specialization)
                        {
                            // Team Svendeprover: Variable to take action if a student has not received a passing grade
                            bool should_print = true;
                            
                            // Open the matching certificate template
                            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(file_certificate_template);

                            // Show progress
                            progressBar1.Value++;

                            // Write relevant datatable data to assigned mergefields in document
                            try
                            {
                                ///////////////////////////////////////////////////////////////////////

                                // Read mergefields
                                string[] mergefield_pairs = File.ReadAllLines(file_mergefields, Encoding.GetEncoding(1252));

                                // Iterate mergefield pairs and insert them into document
                                Object missing = Type.Missing;
                                foreach (string mergefield_pair in mergefield_pairs)
                                {
                                    string[] mergefields = mergefield_pair.Split(';');

                                    try
                                    {
                                        // If there is a complete mergefield pair
                                        if (mergefields.Length > 1)
                                        {
                                            //-------------------------------------------------------------------------------------------------------
                                            ///---------------------------------------------------------------------------------------------------------------
                                            app.Selection.Find.Execute(mergefields[1], missing, missing, missing, missing, missing, missing, missing, missing, datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefields[0]].Ordinal].ToString(), 2); // https://stackoverflow.com/questions/11340264/get-index-of-datatable-column-with-name
                                            ///---------------------------------------------------------------------------------------------------------------
                                            //-------------------------------------------------------------------------------------------------------

                                            // Specific certificate formats for Team Svendeprover
                                            if (bool_teamsvendeprover)
                                            {
                                                // Removing textboxes with "FREMRAGENDE PRÆSTATION ***12***" from template, if the student has a grade lower than 12
                                                if (datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Karakter1"].Ordinal].ToString() != "12")
                                                {
                                                    // Delete all textboxes in the document
                                                    foreach (Word.Shape figure in doc.Shapes)
                                                    {
                                                        if (figure.Type == Microsoft.Office.Core.MsoShapeType.msoTextBox)
                                                        {
                                                            figure.Delete();
                                                        }
                                                    }
                                                }

                                                // Check if the student has received a passing grade
                                                if (Convert.ToInt32(datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Karakter1"].Ordinal]) < 2)
                                                {
                                                    log_report("Følgende elev ser ikke ud til at have bestået den afsluttende prøve og er derfor ikke blevet printet:" + Environment.NewLine + "\t- " +
                                                        datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString() +
                                                        " (" + defined_specialization + ")" + Environment.NewLine + Environment.NewLine);

                                                    // Avoid the files being adding to dictionary_print
                                                    should_print = false;

                                                    break;
                                                }

                                                // Replacing "«Elevtype»" with acknowledgement if the student data contains "T" in column "Elevtype"
                                                if (datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Elevtype"].Ordinal].ToString().Contains("T"))
                                                {
                                                    app.Selection.Find.Execute("«Elevtype»", missing, missing, missing, missing, missing, missing, missing, missing, "\v\vUddannelsen er gennemført med talentspor", 2);
                                                }

                                                // Check if student data is missing or void [CL:2]
                                                //
                                                // Karakter1
                                                if (datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Karakter1"].Ordinal].ToString() == "")
                                                {
                                                    log_report("Der ser ud til at mangle karakter for følgende elev:" + Environment.NewLine + "\t- " +
                                                        datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString() +
                                                        " (" + defined_specialization + ")" + Environment.NewLine + Environment.NewLine);

                                                    // Avoid the files being adding to dictionary_print
                                                    should_print = false;

                                                    break;
                                                }
                                                //
                                                // Student name
                                                if (datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString() == "")
                                                {
                                                    log_report("Der mangler et navn på en elev. Du bør inspicere de uprintede filer for at finde ud af, hvorfor der mangler et navn på eleven." + Environment.NewLine + Environment.NewLine);

                                                    // Avoid the files being adding to dictionary_print
                                                    should_print = false;

                                                    break;
                                                }
                                                //
                                                // CPR-nr
                                                if (datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["CPR-nr."].Ordinal].ToString() == "")
                                                {
                                                    log_report("Der mangler CPR-nr for eleven:" + Environment.NewLine + "\t- " +
                                                        datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString() +
                                                        " (" + defined_specialization + ")" + Environment.NewLine + Environment.NewLine);

                                                    // Avoid the files being adding to dictionary_print
                                                    should_print = false;

                                                    break;
                                                }
                                                else if (datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["CPR-nr."].Ordinal].ToString().Length < 10)
                                                {
                                                    log_report("Formatet på følgende elevs CPR-nr er for kort:" + Environment.NewLine + "\t- " +
                                                        datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString() +
                                                        " (" + defined_specialization + ")" + Environment.NewLine + Environment.NewLine);

                                                    // Avoid the files being adding to dictionary_print
                                                    should_print = false;

                                                    break;
                                                }
                                                //
                                                // Aftaleperiode slut
                                                if (datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Aftaleperiode slut"].Ordinal].ToString() == "")
                                                {
                                                    log_report("Der mangler en gyldig slutdato for følgende elevs aftaleperiode:" + Environment.NewLine + "\t- " +
                                                        datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString() +
                                                        " (" + defined_specialization + ")" + Environment.NewLine + Environment.NewLine);

                                                    // Avoid the files being adding to dictionary_print
                                                    should_print = false;

                                                    break;
                                                }
                                                //
                                                // Praktiksted
                                                if (datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Praktiksted navn"].Ordinal].ToString() == "" ||
                                                    datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Praktiksted adr."].Ordinal].ToString() == "" ||
                                                    datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Praktiksted postnr."].Ordinal].ToString() == "" ||
                                                    datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns["Praktiksted postdistrikt"].Ordinal].ToString() == "")
                                                {
                                                    log_report("Der mangler en eller flere dele af praktikstedsadressen for:" + Environment.NewLine + "\t- " +
                                                        datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString() +
                                                        " (" + defined_specialization + ")" + Environment.NewLine + Environment.NewLine);

                                                    // Avoid the files being adding to dictionary_print
                                                    should_print = false;

                                                    break;
                                                }



                                                //
                                                //// Skoleadresser
                                                //
                                                // Læg mærke til at manglende praktikstedsoplysninger for øjeblikket medfører, at beviset ikke printes. 
                                                //
                                                //
                                                //

                                            }
                                        }
                                        else
                                        {
                                            // If no source is defined for the mergefield, merge it with nothing
                                            try
                                            {
                                                app.Selection.Find.Execute(mergefields[0], missing, missing, missing, missing, missing, missing, missing, missing, "", 2);
                                            }
                                            catch
                                            {
                                                // Might fail if the source (datatable cell) is empty
                                            }
                                        }
                                    }
                                    catch (Exception exception)
                                    {
                                        log_error(exception.ToString());
                                    }
                                }

                                // Remove remaining mergefields
                                app.Selection.Find.Execute("«*»", missing, missing, true, missing, missing, missing, missing, missing, "", 2);

                                // Show progress on label
                                label22.Text = Convert.ToString(progressBar1.Value) + " / " + datatable_csv.Rows.Count.ToString() + ": " + "(" + defined_specialization + ") " + datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString();

                                // Save certificate
                                string save_as = folder_temp + "(" + defined_specialization + ") " + datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefield_student_name].Ordinal].ToString();
                                if (!File.Exists(save_as + ".docx"))
                                {
                                    doc.SaveAs2(save_as);
                                }
                                else
                                {
                                    // If the file already exists...
                                    save_as = save_as + "(1)";
                                    if (!File.Exists(save_as + ".docx"))
                                    {
                                        doc.SaveAs2(save_as);
                                    }
                                    else
                                    {
                                        save_as = save_as.Replace("1", "2");
                                        doc.SaveAs2(save_as);

                                        // If more than three students exist with the same name and specializations, it will have to be handled in this code (currently unhandled)
                                    }
                                }

                                // Add 'save_as' to a dictionary with specialization as key, so that the correct files can be printed
                                if (should_print)
                                {
                                    dictionary_print[defined_specialization].Add(save_as + ".docx");
                                }

                                int_files_to_print++;

                                // Create letter in temp folder and add its path to dictionary_print (if the bool indicates so)
                                if (bool_print_letters)
                                {
                                    // Open the letter template
                                    Microsoft.Office.Interop.Word.Document doc_letter = app.Documents.Open(path_letter_template_copy);

                                    // Merge
                                    foreach (string mergefield_pair in mergefield_pairs)
                                    {
                                        string[] mergefields = mergefield_pair.Split(';');

                                        try
                                        {
                                            // If there is a complete mergefield pair
                                            if (mergefields.Length > 1)
                                            {
                                                //-------------------------------------------------------------------------------------------------------
                                                app.Selection.Find.Execute(mergefields[1], missing, missing, missing, missing, missing, missing, missing, missing, datatable_csv.Rows[i].ItemArray[datatable_csv.Rows[i].Table.Columns[mergefields[0]].Ordinal].ToString(), 2); // https://stackoverflow.com/questions/11340264/get-index-of-datatable-column-with-name                                                                                                                                                                                                                            ///---------------------------------------------------------------------------------------------------------------
                                                //-------------------------------------------------------------------------------------------------------
                                            }
                                            else
                                            {
                                                // If no source is defined for the mergefield, merge it with nothing
                                                try
                                                {
                                                    app.Selection.Find.Execute(mergefields[0], missing, missing, missing, missing, missing, missing, missing, missing, "", 2);
                                                }
                                                catch
                                                {
                                                    // Might fail if the source (datatable cell) is empty
                                                }
                                            }
                                        }
                                        catch (Exception exception)
                                        {
                                            log_error(exception.ToString());
                                        }
                                    }

                                    // Remove remaining mergefields
                                    app.Selection.Find.Execute("«*»", missing, missing, true, missing, missing, missing, missing, missing, "", 2);

                                    // Save letter
                                    string save_as_letter = save_as + " - Følgebrev";
                                    doc_letter.SaveAs2(save_as_letter);

                                    // Close doc_letter
                                    doc_letter.Close();

                                    // Add letter to dictionary, so that it can be printed
                                    if (should_print)
                                    {
                                        dictionary_print[defined_specialization].Add(save_as_letter + ".docx");
                                    }

                                    // Increment
                                    int_files_to_print++;
                                }

                                // Copy certificate to save folder (if bool indicates so)
                                if (bool_save_certificates)
                                {
                                    string save_as_filename = Path.GetFileName(save_as);
                                    string save_as_folder = File.ReadAllText(folder_save_certificates, Encoding.GetEncoding(1252));
                                    save_as_folder = save_as_folder.Replace(@"C:\Users\%USERPROFILE%", @"C:\Users\" + Environment.UserName);

                                    string save_as_pdf = save_as_folder + @"\" + save_as_filename;

                                    try
                                    {
                                        doc.SaveAs2(save_as_pdf, Word.WdSaveFormat.wdFormatPDF);
                                    }
                                    catch
                                    {
                                        log_report("Could not save certificate: " + Environment.NewLine + "\t" + save_as_pdf + Environment.NewLine);
                                    }
                                }
                                //////////////////////////////////////////////////////////////////////
                            }
                            catch (Exception exception)
                            {
                                MessageBox.Show(exception.ToString());
                                log_error(exception.ToString());
                                exit_and_cleanup();
                                return;
                            }

                            // Try to close Word scope
                            try
                            {
                                doc.Close(false);
                            }
                            catch
                            {
                                //
                            }
                        }
                    }

                    // Delete the matched certificate template
                    File.Delete(file_certificate_template);

                    // Delete the temporary letter template
                    File.Delete(path_letter_template_copy);
                }
            }

            // Make sure all Word activity created by this app is terminated
            app.Quit(false, false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            kill_process("WINWORD", dtWordStart);

            // Hide progressbar
            progressBar1.Hide();
            label22.Hide();

            // If letters are to be printed with the certificates, show menu for determining which certificates should be printed with a letter
            if (bool_print_letters)
            {
                dataGridView3.Show();
                label34.Text = "Angiv om der skal dannes følgebrev for hver elev:";
                label34.Show();
                button11.Show(); // Calls print_and_finish()
                button13.Show(); // Calls exit_and_cleanup()
            }
            // Otherwise just give the option of printing or BORTING
            else
            {
                button11.Show(); // Calls print_and_finish()
                button13.Show(); // Calls exit_and_cleanup()
            }
        }
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

        void print_and_finish()
        {
            progressBar1.Show();
            progressBar1.Maximum = int_files_to_print;
            progressBar1.Value = 0;
            label22.Show();
            label22.Text = "Printer...";

            // Exclude unchecked letters (datagridview3)
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (row.Cells["dataGridViewTextBoxColumn1"].Value != null)
                {
                    if (row.Cells["dataGridViewTextBoxColumn2"].Value != null)
                    {
                        if (Convert.ToBoolean(row.Cells["dataGridViewCheckBoxColumn3"].Value) == false)
                        {
                            for (int i = 0; i < dictionary_print.Count; i++)
                            {
                                for (int j = 0; j < dictionary_print[row.Cells["dataGridViewTextboxColumn2"].Value.ToString()].Count; j++)
                                {
                                    string file = dictionary_print[row.Cells["dataGridViewTextboxColumn2"].Value.ToString()][j];

                                    if (file.Contains(row.Cells["dataGridViewTextBoxColumn1"].Value.ToString()))
                                    {
                                        if (file.Contains(row.Cells["dataGridViewTextboxColumn2"].Value.ToString()))
                                        {
                                            if (file.Contains("Følgebrev"))
                                            {
                                                dictionary_print[row.Cells["dataGridViewTextboxColumn2"].Value.ToString()].Remove(file);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            // Word app to print from
            DateTime dtWordStart = DateTime.Now;
            Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
            app.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
            app.Visible = false;

            // Detect first instance of new paper
            bool paper_change = false;

            // Iterate dictionary print and match with printer paper
            for (int i = 1; i < 6; i++)
            {
                foreach (var entry_print in dictionary_specializations)
                {
                    foreach (var entry_files in dictionary_print)
                    {
                        foreach (string entry_file in entry_files.Value)
                        {
                            if (entry_print.Value == i.ToString())
                            {
                                try
                                {
                                    // Print on paper 1 without informing user first
                                    if (i == 1)
                                    {
                                        if (entry_print.Key == entry_files.Key)
                                        {
                                            // Print
                                            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(entry_file);

                                            //doc.PrintOut();
                                            //
                                            //
                                            //
                                            //

                                            doc.Close(false);

                                            progressBar1.Value++;

                                            // Delete the file from folder_temp
                                            try
                                            {
                                                File.Delete(entry_file);
                                            }
                                            catch
                                            {
                                                //
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (entry_print.Key == entry_files.Key)
                                        {
                                            // By the first instance of a certificate on the printer paper, inform the user about paper change
                                            if (paper_change == false)
                                            {
                                                // Read names of paper types
                                                string[] papers = File.ReadAllLines(file_print_paper, Encoding.GetEncoding(1252));

                                                string current_paper = "";

                                                // Try to get current paper name (might need to be handled because it might not be named by user)
                                                try
                                                {
                                                    current_paper = papers[i - 2];
                                                }
                                                catch
                                                {
                                                    // If paper name cannot be determined, use the current specialization name
                                                    current_paper = entry_print.Key;
                                                }

                                                MessageBox.Show("Der mangler at blive printet beviser på " +
                                                    current_paper.ToUpper() +
                                                    "-papir. Skift papir, når printeren er færdig med dets igangværende job, og tryk derefter OK for at printe disse.");

                                                paper_change = true;
                                            }

                                            // Print
                                            Microsoft.Office.Interop.Word.Document doc = app.Documents.Open(entry_file);
                                            //doc.PrintOut();
                                            doc.Close(false);

                                            progressBar1.Value++;

                                            // Delete the file from folder_temp
                                            try
                                            {
                                                File.Delete(entry_file);
                                            }
                                            catch
                                            {
                                                //
                                            }
                                        }
                                    }
                                }
                                catch (Exception exception)
                                {
                                    MessageBox.Show(exception.ToString());
                                    log_error(exception.ToString());
                                }
                            }
                        }
                    }
                }

                // Reset bool
                paper_change = false;
            }

            // Make sure all Word activity created by this app is terminated
            app.Quit(false, false, false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            kill_process("WINWORD", dtWordStart);

            // No longer possible to abort
            button13.Hide();

            // Show report if it is not empty
            string report = File.ReadAllText(file_log_report, Encoding.GetEncoding(1252));
            if (report.Length > 0)
            {
                button11.Text = "OK";
                button11.Show();
                richTextBox3.Text = report;
                richTextBox3.Show();
            }

            // If there are files in the temp folder, allow the user to view the content of the folder
            if (Directory.GetFiles(folder_temp).Length > 0)
            {
                button12.Show();
            }
            else
            {
                button12.Hide();
            }
        }

        void kill_process(string process_name, DateTime aprox_start_time)
        {
            DateTime time_before = aprox_start_time.AddSeconds(-7);
            DateTime time_after = aprox_start_time.AddSeconds(7);

            foreach (Process process in Process.GetProcesses())
            {
                if (process.ProcessName == process_name)
                {
                    DateTime current_process_start_time = process.StartTime;

                    if (time_before < current_process_start_time && time_after > current_process_start_time)
                    {
                        try
                        {
                            process.Kill();
                        }
                        catch
                        {
                            // 
                        }
                    }
                }
            }
        }

        public void choose_backup()
        {
            // Choose the specific folder
            FolderBrowserDialog backup_chooser = new FolderBrowserDialog();
            backup_chooser.SelectedPath = folder_backups; ; //Initial search directory
            if (backup_chooser.ShowDialog() == DialogResult.OK)
            {
                foreach (string directory in Directory.GetDirectories(backup_chooser.SelectedPath))
                {
                    string mutual_directory_name = new DirectoryInfo(directory).Name;

                    // Replace each folder in current directory with a corresponding backup directory
                    try
                    {
                        // Delete current folder
                        Directory.Delete(Directory.GetCurrentDirectory() + @"\" + mutual_directory_name, true);

                        // Replace with backup folder
                        Directory.Move(directory, Directory.GetCurrentDirectory() + @"\" + mutual_directory_name);
                    }
                    catch (Exception exception)
                    {
                        MessageBox.Show(exception.ToString());
                        log_error(exception.ToString());
                        exit_and_cleanup();
                        return;
                    }
                }
            }
        }

        public void create_backup()
        {
            // Create new backup folder with timestamp
            string folder_new_backup = folder_backups + "BACKUP " + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + @"\";
            Directory.CreateDirectory(folder_new_backup);

            foreach (string directory in Directory.GetDirectories(Directory.GetCurrentDirectory()))
            {
                if (directory.Contains("setting") || directory.Contains("data"))
                {
                    // Create the backup subfolder
                    string subfolder_new_backup = folder_new_backup + directory.Replace(Directory.GetCurrentDirectory(), "") + @"\";
                    Directory.CreateDirectory(subfolder_new_backup);

                    // Copy files to the backup subfolder
                    foreach (string file in Directory.GetFiles(directory))
                    {
                        string file_new_backup = subfolder_new_backup + file.Replace(directory, "");
                        File.Copy(file, file_new_backup);
                    }
                }
            }

            // Remove the oldest backup
            if (Directory.GetDirectories(folder_backups).Length > 10)
            {
                FileSystemInfo fileInfo = new DirectoryInfo(folder_backups).GetFileSystemInfos().OrderBy(fi => fi.CreationTime).First(); // https://stackoverflow.com/questions/44690815/how-to-delete-oldest-folder-created-from-local-disk-using-c-sharp
                Directory.Delete(fileInfo.FullName, true);
            }
        }

        public void exit_and_cleanup()
        {
            dataGridView3.Hide();
            label34.Hide();
            button11.Hide();
            progressBar1.Value = 0;
            progressBar1.Hide();
            label22.Hide();
            richTextBox3.Hide();
            button12.Hide();
            button13.Hide();

            // Reset variables and objects
            dictionary_print.Clear();
            int_files_to_print = 0;
            label22.Text = "";

            // Reset from previous run
            label34.Text = "Angiv om der skal dannes følgebrev for hver elev:";
            button11.Text = "Print";

            // Delete the user specific temp folder
            try
            {
                Directory.Delete(folder_temp, true);
            }
            catch
            {
                // If this for some reason is impossible (due to a file not having been closed properly): Try to delete as much in the folder as possible
                System.IO.DirectoryInfo dirInfo_folder_temp = new DirectoryInfo(folder_temp);
                try
                {
                    foreach (FileInfo file in dirInfo_folder_temp.GetFiles())
                    {
                        //file.Delete();
                    }
                }
                catch
                {
                    //
                }
            }
        }

        public string find_certificate_template(string specialization)
        {
            // Find potential certificate templates by specialization (filename must contain specialization)
            string tmp_folder_path = File.ReadAllText(folder_certificate_templates, Encoding.GetEncoding(1252)).Replace(@"C:\Users\%USERPROFILE%", @"C:\Users\" + Environment.UserName);
            DirectoryInfo tmp_directory_info = new DirectoryInfo(tmp_folder_path);
            FileInfo[] tmp_potential_templates = tmp_directory_info.GetFiles("*" + specialization + "*.docx");

            // Integer to count number of matches
            int kopier = 0;

            // Clear the settings file to store template matches
            File.WriteAllText(files_certificate_template_matches, string.Empty, Encoding.GetEncoding(1252));

            // Finding matching certificate template files where the filename contains the specialization name
            if (!bool_teamsvendeprover)
            {
                foreach (FileInfo match in tmp_potential_templates)
                {
                    // Must not be a temporary file
                    if (!match.Name.StartsWith("~$"))
                    {
                        string filename = match.FullName;
                        kopier++; // Count number of matches

                        // Registering path to matching certificate template in settings file
                        File.AppendAllText(files_certificate_template_matches, filename + Environment.NewLine, Encoding.GetEncoding(1252));
                    }
                }
            }
            // If specific settings for Team Svendeprover has been activated, narrow the search
            else
            {
                foreach (FileInfo match in tmp_potential_templates)
                {
                    // Must not be a temporary file
                    if (!match.Name.StartsWith("~$"))
                    {
                        if (!match.Name.Contains("historisk"))
                        {
                            if (!match.Name.Contains("kopi"))
                            {
                                string filename = match.FullName;
                                kopier++; // Count number of matches

                                // Registering path to matching certificate template in settings file
                                File.AppendAllText(files_certificate_template_matches, filename + Environment.NewLine, Encoding.GetEncoding(1252));
                            }
                        }
                    }
                }
            }

            // Filename to copy to and return
            string tmp_certificate_template_copy = folder_temp + DateTime.Now.ToString("yyyy-MM-dd HHmmss") + " TEMPLATE " + Environment.UserName + " " + specialization + ".docx";

            // If only one match has been found, copy this to temp directory, return this match copy path and clear files_certificate_template_matches
            if (kopier == 1)
            {
                string tmp_certificate_template_original = File.ReadAllText(files_certificate_template_matches, Encoding.GetEncoding(1252));
                File.WriteAllText(files_certificate_template_matches, string.Empty, Encoding.GetEncoding(1252));

                File.Copy(tmp_certificate_template_original, tmp_certificate_template_copy);

                return tmp_certificate_template_copy;
            }
            // If more than one file has been found, narrow it down
            else if (kopier > 0)
            {
                // Team Svendeprover might have different specializations for 'Plastmager'
                if (!bool_teamsvendeprover)
                {
                    //
                }
                else
                {
                    // PLASTMAGER
                }

                // Create file to determine whether below form has been closed properly (the file is closed from form2)
                FileStream str = File.Create(folder_temp + "dialog.tmp");
                str.Close();

                // Write current specialization to the tmp file
                File.WriteAllText(folder_temp + "dialog.tmp", specialization, Encoding.GetEncoding(1252));

                // Open form to choose template
                Form2_Choose_template form2 = new Form2_Choose_template();
                form2.ShowDialog();

                // Check if the form has been closed properly
                if (!File.Exists(folder_temp + "dialog.tmp"))
                {
                    // After the closing of Form2_Choose_temple, either a file has been designated in files_certificate_templates_matches or no file was designated
                    string designated_template = File.ReadAllText(files_certificate_template_matches, Encoding.GetEncoding(1252));
                    if (designated_template.Length > 0)
                    {
                        // If a file has been designated as template, return it after clearing the file 
                        File.WriteAllText(files_certificate_template_matches, string.Empty, Encoding.GetEncoding(1252));
                        File.Copy(designated_template, tmp_certificate_template_copy);
                        return tmp_certificate_template_copy;
                    }
                    else
                    {
                        // Otherwise return null
                        return null;
                    }
                }
                else
                {
                    File.Delete(folder_temp + "dialog.tmp");
                    return null;
                }
            }
            // If no matchin file has been found, prompt the user to input a file path
            else
            {
                // Create file to determine whether below form has been closed properly (the file is closed from form2)
                FileStream str = File.Create(folder_temp + "dialog.tmp");
                str.Close();

                // Write current specialization to the tmp file
                File.WriteAllText(folder_temp + "dialog.tmp", specialization, Encoding.GetEncoding(1252));

                // Open form to choose template
                Form2_Choose_template form2 = new Form2_Choose_template();
                form2.ShowDialog();

                // Check if the form has been closed properly
                if (!File.Exists(folder_temp + "dialog.tmp"))
                {
                    // After the closing of Form2_Choose_temple, either a file has been designated in files_certificate_templates_matches or no file was designated
                    string designated_template = File.ReadAllText(files_certificate_template_matches, Encoding.GetEncoding(1252));
                    if (designated_template.Length > 0)
                    {
                        // If a file has been designated as template, return it after clearing the file 
                        File.WriteAllText(files_certificate_template_matches, string.Empty, Encoding.GetEncoding(1252));
                        File.Copy(designated_template, tmp_certificate_template_copy);
                        return tmp_certificate_template_copy;
                    }
                    else
                    {
                        // Otherwise return null
                        return null;
                    }
                }
                else
                {
                    File.Delete(folder_temp + "dialog.tmp");
                    return null;
                }
            }
        }

        public void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        public bool get_binary_file_setting(string path)
        {
            if (File.ReadAllText(path, Encoding.GetEncoding(1252)).Contains("1"))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public void dict_to_file()
        {
            // Write dictionary_specializations to file
            File.WriteAllText(file_dictionary_specializations, string.Empty, Encoding.GetEncoding(1252)); // Clear file
            foreach (var entry in dictionary_specializations)
            {
                File.AppendAllText(file_dictionary_specializations, entry.Key + ";" + entry.Value + Environment.NewLine, Encoding.GetEncoding(1252));

                //
                // How to avoid specialization dublication (an issue because printer settings may vary!)
                //
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Save changes in specializations to dictionary_specializations
            Dictionary<string, string> copy_dictionary_specializations = new Dictionary<string, string>(dictionary_specializations);
            dictionary_specializations.Clear();
            string str_specializations = richTextBox1.Text;
            string[] array_specializations = str_specializations.Split('\n');
            Array.Sort(array_specializations);

            for (int i = 0; i < array_specializations.Length; i++)
            {
                // Ignore empty lines
                if (array_specializations[i] != "")
                {
                    // If the specialization is not already in the file
                    if (!dictionary_specializations.ContainsKey(array_specializations[i]))
                    {
                        // Write printer_paper_setting if it is defined
                        if (copy_dictionary_specializations.ContainsKey(array_specializations[i]))
                        {
                            dictionary_specializations[array_specializations[i]] = copy_dictionary_specializations[array_specializations[i]];
                        }
                        // Otherwise set default printer_paper_setting: "1"
                        else
                        {
                            dictionary_specializations[array_specializations[i]] = "1";
                        }
                    }
                }
            }
            copy_dictionary_specializations.Clear();

            // Write dictionary_specializations to file
            dict_to_file();

            // Reload
            load_printer_settings();

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Set folder_csv_students
            set_folder(label6, folder_csv_students);
        }

        public void set_folder(Label label, string folder)
        {
            // Set standard folder
            FolderBrowserDialog dialog_folder = new FolderBrowserDialog();
            dialog_folder.SelectedPath = @"C:\"; // Starting folder for search
            if (dialog_folder.ShowDialog() == DialogResult.OK)
            {
                string selected_folder = dialog_folder.SelectedPath;
                label.Text = selected_folder;
                File.WriteAllText(folder, selected_folder.Replace(@"C:\Users\" + Environment.UserName, @"C:\Users\%USERPROFILE%"), Encoding.GetEncoding(1252));
            }
        }

        public void read_settings(Label label, string path)
        {
            string content = File.ReadAllText(path, Encoding.GetEncoding(1252));
            content = content.Replace(@"C:\Users\%USERPROFILE%", @"C:\Users\" + Environment.UserName);
            if (Directory.Exists(content))
            {
                label.Text = content;
            }
            else
            {
                label.Text = userprofile + @"\Desktop";
            }
        }

        public void load_printer_settings()
        {
            // Clear rows in printer settings
            dataGridView1.Rows.Clear();

            int count = 0;
            foreach (var entry in dictionary_specializations)
            {
                if (entry.Key != "")
                {
                    DataGridViewRow row = (DataGridViewRow)dataGridView1.Rows[count].Clone();

                    // Specialization column
                    row.Cells[0].Value = entry.Key;
                    dataGridView1.Rows.Add(row);

                    // Set printer settings
                    int column = Convert.ToInt32(entry.Value);
                    row.Cells[column].Value = true;

                    count++;
                }
            }
            count = 0;
        }

        public void load_mergefield_settings()
        {
            // Clear rows in mergefield grid
            dataGridView2.Rows.Clear();

            string[] mergefield_sets = File.ReadAllLines(file_mergefields, Encoding.GetEncoding(1252));

            for (int i=0; i<mergefield_sets.Length; i++)
            {
                if (mergefield_sets[i] != "")
                {
                    // Split mergefields into source and destination
                    string[] mergefields = mergefield_sets[i].Split(';');

                    DataGridViewRow row = (DataGridViewRow)dataGridView2.Rows[i].Clone();

                    // Write mergefield_source to data grid
                    row.Cells[0].Value = mergefields[0];
                    row.Cells[1].Value = mergefields[1];
                    dataGridView2.Rows.Add(row);
                }
            }
        }

        public void get_paper_designation(string path, TextBox textbox)
        {
            // Get the int number of the textbox from its name so that it can be used to find the specified line in the settings file
            int paper_index = Convert.ToInt32(textbox.Name.Substring(textbox.Name.Length - 1)) - 1;

            try
            {
                string[] paper_designation = File.ReadAllLines(path, Encoding.GetEncoding(1252));
                textbox.Text = paper_designation[paper_index];
            }
            catch
            {
                // Line is empty and therefore the textbox should be empty
            }
        }

        public void set_paper_designations()
        {
            if (form1_fully_loaded)
            {
                File.WriteAllText(file_print_paper, string.Empty);
                File.AppendAllText(file_print_paper, textBox1.Text + Environment.NewLine, Encoding.GetEncoding(1252));
                File.AppendAllText(file_print_paper, textBox2.Text + Environment.NewLine, Encoding.GetEncoding(1252));
                File.AppendAllText(file_print_paper, textBox3.Text + Environment.NewLine, Encoding.GetEncoding(1252));
                File.AppendAllText(file_print_paper, textBox4.Text, Encoding.GetEncoding(1252));
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            // Set folder_certificate_templates
            set_folder(label7, folder_certificate_templates);
        }

        private void button5_Click(object sender, EventArgs e)
        {            
            // Set file_letter_template
            OpenFileDialog dialog_file = new OpenFileDialog();
            dialog_file.InitialDirectory = @"C:\";
            if (dialog_file.ShowDialog() == DialogResult.OK)
            {
                string selected_file = dialog_file.FileName;
                label8.Text = selected_file;
                File.WriteAllText(file_letter_template, selected_file.Replace(@"C:\Users\" + Environment.UserName, @"C:\Users\%USERPROFILE%"), Encoding.GetEncoding(1252));

                checkBox2.Checked = false;
                bool_print_letters = true;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            // Set folder_save_certificates
            set_folder(label10, folder_save_certificates);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                if (functionality_framework == "UiPath")
                {
                    // Run method for UiPath
                    UiPath();
                }
                else if (functionality_framework == ".NET")
                {
                    // Run method for .NET
                    dotNET();
                }
            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.ToString());
                log_error(exception.ToString());
                exit_and_cleanup();
                return;
            }
        }

        private void UiPath()
        {
            MessageBox.Show("Kører SVEND via UiPath (xaml)");

            MessageBox.Show("Der er endnu ikke udviklet en funktionalitet i Svend, som kan køre via UiPath." + Environment.NewLine + "Ændr framework under fanen 'Indstillinger' for at køre Svend via .NET.");
        }

        public void log_error(string error)
        {
            File.AppendAllText(file_log_error, DateTime.Now.ToString("yyyy-MM-dd HHmmss") + ": " + Environment.NewLine + error + Environment.NewLine + Environment.NewLine, Encoding.GetEncoding(1252));
        }

        public void log_report(string report)
        {
            File.AppendAllText(file_log_report, report, Encoding.GetEncoding(1252));
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (button7.Text == "Nej")
            {
                button7.Text = "Ja";
                File.WriteAllText(save_certificates, "Y");
                bool_save_certificates = true;
            }
            else if (button7.Text == "Ja")
            {
                button7.Text = "Nej";
                File.WriteAllText(save_certificates, "N");
                bool_save_certificates = false;
            }
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // When a checkbox is chosen to indicate printer paper, then uncheck all other checkboxes in the row
                DataGridViewCheckBoxCell checked_cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex] as DataGridViewCheckBoxCell;

                for (int i = 1; i < 6; i++)
                {
                    // Sets each cell in the row (iteration)
                    DataGridViewCheckBoxCell cell = dataGridView1.Rows[e.RowIndex].Cells[i] as DataGridViewCheckBoxCell;

                    if (cell.ColumnIndex > 0)
                    {
                        if (Convert.ToBoolean(cell.Value) == true)
                        {
                            if (checked_cell.ColumnIndex != cell.ColumnIndex)
                            {
                                // Uncheck other cells
                                cell.Value = false;
                            }
                        }
                    }
                }

                // Set the new printer setting in dictionary specialization and write the change to file_dictionary_specializations
                dictionary_specializations[checked_cell.OwningRow.Cells[0].Value.ToString()] = checked_cell.ColumnIndex.ToString();
                dict_to_file();
            }
            catch
            {
                //
                // Load backups (this catch is important, since checking a checkbox in the last empty row creates an exception)
                //
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            set_paper_designations();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            set_paper_designations();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            set_paper_designations();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            set_paper_designations();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // Clear file_mergefields
            File.WriteAllText(file_mergefields, string.Empty);

            string mergefield_source = "";
            string mergefield_destination = "";

            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                try
                {
                    mergefield_source = row.Cells[0].Value.ToString();
                }
                catch
                {
                    // Empty object
                }

                try
                {
                    mergefield_destination = row.Cells[1].Value.ToString();
                }
                catch
                {
                    // Empty object
                }

                if (mergefield_source != "" || mergefield_destination != "") // Or
                {
                    File.AppendAllText(file_mergefields, mergefield_source + ";" + mergefield_destination + Environment.NewLine, Encoding.GetEncoding(1252));
                }

                // Reset foreach loop
                mergefield_source = "";
                mergefield_destination = "";
            }

            // Remove empty lines in file_mergefield
            string[] content = File.ReadAllLines(file_mergefields, Encoding.GetEncoding(1252));
            File.WriteAllText(file_mergefields, string.Empty);
            foreach (string line in content)
            {
                if (line != "")
                {
                    File.AppendAllText(file_mergefields, line + Environment.NewLine, Encoding.GetEncoding(1252));
                }

                //
                // Remove the last line!
                //
            }

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MessageBox.Show("Denne brugertilpassede opsætning indeholder:\n" +
                "-   Kontrol af karakter (>= 02)\n" +
                "-   Automatisk påførelse af skoleadresse, hvis denne mangler\n" +
                "-   Advarsel om manglende plastmagerspecialer i EASY-P\n" +
                "-   Udelader figuren 'Fremragende præstation' på beviser med karakter under 12");
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            string binary_bool = "";

            if (checkBox1.Checked == true)
            {
                binary_bool = "1";
                bool_teamsvendeprover = true;
            }
            else
            {
                binary_bool = "0";
                bool_teamsvendeprover = false;
            }

            File.WriteAllText(file_user_setup_teamsvendeprover, binary_bool);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(file_mergefield_specialization, textBox5.Text, Encoding.GetEncoding(1252));
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (button9.Text == ".NET")
            {
                button9.Text = "UiPath";
                File.WriteAllText(file_functionality_framework, "UiPath", Encoding.GetEncoding(1252));
                functionality_framework = "UiPath";
            }
            else if (button9.Text == "UiPath")
            {
                button9.Text = ".NET";
                File.WriteAllText(file_functionality_framework, ".NET", Encoding.GetEncoding(1252));
                functionality_framework = ".NET";
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            choose_backup();
            load_all();
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            File.WriteAllText(file_mergefield_student_name, textBox6.Text, Encoding.GetEncoding(1252));
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox2.Checked == true)
            {
                File.WriteAllText(file_letter_template, string.Empty, Encoding.GetEncoding(1252));
                label8.Text = "";
                bool_print_letters = false;
            }
            else
            {
                button5.PerformClick();
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (button11.Text == "Print")
            {
                dataGridView3.Hide();
                label34.Hide();
                button11.Hide();
                button13.Hide();

                print_and_finish();
            }
            else if (button11.Text == "OK")
            {
                richTextBox3.Hide();
                button11.Hide();
                label34.Hide();
                button13.Hide();

                exit_and_cleanup();
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            // Delete the user specific temp folder
            try
            {
                Directory.Delete(folder_temp, true);
            }
            catch
            {
                // If this for some reason is impossible (due to a file not having been closed properly): Try to delete as much in the folder as possible
                System.IO.DirectoryInfo dirInfo_folder_temp = new DirectoryInfo(folder_temp);
                try
                {
                    foreach (FileInfo file in dirInfo_folder_temp.GetFiles())
                    {
                        file.Delete();
                    }
                }
                catch
                {
                    //
                }
            }
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            Process.Start(folder_temp);
        }

        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            exit_and_cleanup();
        }
    }
}
