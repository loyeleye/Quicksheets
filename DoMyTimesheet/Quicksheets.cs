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
using System.Xml;

namespace DoMyTimesheet
{
    public partial class Quicksheets : Form
    {
        private SubtaskHistory subtaskHistory = new SubtaskHistory();
        private TaskHistory taskHistory = new TaskHistory();
        private JiraHistory jiraHistory = new JiraHistory();
        public DateTime dateSelected = DateTime.Now;
        public static string lastName = String.Empty;
        public Dictionary<int,string> timesheetEntries = new Dictionary <int,string>();

        public Quicksheets()
        {
            InitializeComponent();
            ProcessSettings();
            DrawSubtaskCheckboxes();
        }

        public void ProcessSettings()
        {
            if (!File.Exists("TimesheetData/Settings.xml"))
            {
                StreamWriter settingsXML = new StreamWriter("TimesheetData/Settings.xml");
                settingsXML.Write("<Settings>\n\t<Required>\n\t\t<LastName></LastName>\n\t</Required>\n\t<Default>\n\t\t<Customer>AIM</Customer>\n\t\t<Project></Project>\n\t\t<Name></Name>\n\t\t<Task></Task>\n\t\t<Jira></Jira>\n\t\t<FMCCPLR></FMCCPLR>\n\t\t<RFC></RFC>\n\t\t<Hours></Hours>\n\t</Default>\n</Settings>");
                settingsXML.Close();
            }

            XmlTextReader sReader = new XmlTextReader("TimesheetData/Settings.xml");

            string setting = "";
            while (sReader.Read())
            {
                if (sReader.NodeType == XmlNodeType.Element)
                {
                    setting = sReader.Name;
                }
                else if (sReader.NodeType == XmlNodeType.Text)
                {
                    switch (setting)
                    {
                        case "LastName":
                            lastName = sReader.Value;
                            break;
                        case "Customer":
                            this.customerField.Text = sReader.Value;
                            break;
                        case "Project":
                            this.projectField.Text = sReader.Value;
                            break;
                        case "Name":
                            this.nameField.Text = sReader.Value;
                            break;
                        case "Task":
                            this.taskField.Text = sReader.Value;
                            break;
                        case "Jira":
                            this.jiraField.Text = sReader.Value;
                            break;
                        case "FMCCPLR":
                            int fmcc;
                            if (!Int32.TryParse(sReader.Value, out fmcc)) fmcc = 0;
                            this.fmccUpDown.Value = fmcc;
                            break;
                        case "RFC":
                            int rfc;
                            if (!Int32.TryParse(sReader.Value, out rfc)) rfc = 0;
                            this.rfcUpDown.Value = rfc;
                            break;
                        case "Hours":
                            int hours;
                            if (!Int32.TryParse(sReader.Value, out hours)) hours = 0;
                            this.rfcUpDown.Value = hours;
                            break;
                    }
                }
            }
        }

        public void LoadTimesheetEntry(string csv)
        {
            String[] values = csv.Split(',');

            // Date
            string date = values[0];
            if (!DateTime.TryParse(date, out dateSelected))
            {
                dateSelected = DateTime.Now;
            }
            this.dateButton.Text = "Date: " + dateSelected.ToShortDateString();
            // Customer
            string customer = values[1];
            this.customerField.Text = customer;
            // Project
            string project = values[2];
            this.projectField.Text = project;
            // Resource
            string name = values[3];
            this.nameField.Text = name;
            // Task
            string task = values[4];
            this.taskField.Text = task;
            // Jira
            string jira = values[5];
            this.jiraField.Text = jira;
            // Fmcc
            int fmcc;
            if (!Int32.TryParse(values[6], out fmcc)) fmcc = 0;
            this.fmccUpDown.Value = fmcc;
            // Rfc
            int rfc;
            if (!Int32.TryParse(values[7], out rfc)) rfc = 0;
            this.rfcUpDown.Value = rfc;
            // Subtask
            string subtasks = values[8];
            MarkSubtasksIfExists(subtasks);
            // Hours
            int hours;
            if (!Int32.TryParse(values[9], out hours)) hours = 0;
            this.hourUpDown.Value = hours;
            // Completion Bar and Event Log
            SetCompletionBarAndEventLog();
        }

        public string SaveTimesheetEntry()
        {
            string csv = "";
            TrimTextBoxFields();

            taskHistory.AddIfNewItem(taskField.Text);
            jiraHistory.AddIfNewItem(jiraField.Text);
            subtaskHistory.AddIfNewItem(subtaskField.Text);

            // Date
            csv += dateSelected.ToShortDateString().Replace(",", ";");
            // Customer
            csv += "," + this.customerField.Text.Replace(",", ";");
            // Project
            csv += "," + this.projectField.Text.Replace(",", ";");
            // Resource
            csv += "," + this.nameField.Text.Replace(",", ";");
            // Task
            csv += "," + this.taskField.Text.Replace(",", ";");
            // Jira
            csv += "," + this.jiraField.Text.Replace(",", ";");
            // Fmcc
            string fmccValue = (this.fmccUpDown.Value == 0) ? "N/A" : this.fmccUpDown.Value.ToString();
            csv += "," + fmccValue;
            // Rfc
            string rfcValue = (this.rfcUpDown.Value == 0) ? "N/A" : this.rfcUpDown.Value.ToString();
            csv += "," + rfcValue;
            // Subtask
            csv += "," + GetSubtasksToStringIfMarked().Replace(",", ";");
            // Hours
            csv += "," + this.hourUpDown.Value.ToString().Replace(",", ";");
            
            DrawSubtaskCheckboxes();
            return csv;
        }

        /// <summary>
        /// This function will check if the checkbox is marked. If it is marked,
        /// the subtask will be appended to the string. It also checks the subtask
        /// field to see if a subtask that isn't 'Other', empty string, or duplicate is
        /// filled in. If so, adds it to the list too.
        /// </summary>
        /// <returns>List of marked subtasks</returns>
        public string GetSubtasksToStringIfMarked()
        {
            string rtnSubtasks = "";
            bool firstItem = true;

            foreach (Control child in this.Main.TabPages[2].Controls)
            {
                if ((child is CheckBox) 
                    && (child as CheckBox).Checked
                    && rtnSubtasks.Has((child as CheckBox).Text) == false)
                {
                    if (!firstItem) rtnSubtasks += "/";
                    else firstItem = false;
                        
                    rtnSubtasks += (child as CheckBox).Text;
                }
            }
            if (this.subtaskField.Text.Trim() != "Other" 
                && this.subtaskField.Text.Trim() != ""
                && !rtnSubtasks.Has(this.subtaskField.Text))
            {
                if (!firstItem) rtnSubtasks += "/";
                else firstItem = false;
                
                rtnSubtasks += this.subtaskField.Text;
            }
            return rtnSubtasks;
        }

        private void DisposeOfSubtaskCheckboxes()
        {
            foreach (Control child in this.Main.TabPages[2].Controls)
            {
                if (child is CheckBox)
                {
                    child.Dispose();
                }
            }
            return;
        }

        private CheckBox CreateCheckbox_Subtask(string text, int id, int xPos, int yPos)
        {
            CheckBox box = new CheckBox();

            box.Tag = "cbSubtask" + id.ToString();
            box.Enabled = true;
            box.Text = text; // Since indexes start at 0
            box.ForeColor = Color.White;
            box.Location = new Point(xPos, yPos);
            box.Width = GetCBWidth(box);
            box.Height = 30;
            this.Main.TabPages[2].Controls.Add(box);

            return box;
        }

        private void LayoutSTCheckboxes(string[] contentList, int cSize, int xPos=5)
        {
            CheckBox box;
            string currText;
            int numCheckBoxes = contentList.Length;
            int currentBox = 0;
            int rSize = (numCheckBoxes - 1) / cSize + 1;
            int yPos;

            for (int row = 0; row < rSize; row++)
            {
                yPos = 40 * (row + 1);
                for (int col = 0; col < cSize; col++)
                {
                    currentBox = (row * cSize) + (col + 1);
                    currText = contentList[currentBox - 1];

                    box = CreateCheckbox_Subtask(currText, currentBox, xPos, yPos);
                    xPos += box.Width;
                    if (currentBox >= numCheckBoxes) break;
                }
                if (currentBox >= numCheckBoxes) break;
                xPos = 5;
            }
            this.tabPage3.Refresh();
        }

        private void DrawSubtaskCheckboxes()
        {
            string subtasks = GetSubtasksToStringIfMarked();
            DisposeOfSubtaskCheckboxes();

            string[] subtaskList = subtaskHistory.UseAsStringArray();

            if (subtaskList.Length < 1) return;

            int checkboxesPerLine = 4;
            LayoutSTCheckboxes(subtaskList, checkboxesPerLine);
            MarkSubtasksIfExists(subtasks);
        }

        private void MarkSubtasksIfExists(string markList)
        {
            string[] subtasks = markList.Split('/');

            foreach (Control child in this.Main.TabPages[2].Controls)
            {
                if (child is CheckBox)
                {
                    if (subtasks.Contains(child.Text))
                        (child as CheckBox).Checked = true;
                    else (child as CheckBox).Checked = false;
                }
            }
            return;
        }

        private int GetCBWidth(CheckBox cb)
        {
            int rtnWidth = 0;
            int incWidthBy = 50;

            Font font = new Font(cb.Font.Name, cb.Font.Size);
            Size cbTextSize = TextRenderer.MeasureText(cb.Text, font);

            while ((cbTextSize.Width + 40) > rtnWidth)
            {
                rtnWidth += incWidthBy;
            }

            return rtnWidth;
        }
    }

    abstract public class History
    {
        protected History()
        {
            SetDirectoryPath();

            LoadFromFile();
        }

        virtual public void LoadFromFile()
        {
            historyList = new List<string>();

            if (!Directory.Exists(dataFolder))
            {
                DirectoryInfo newDir = Directory.CreateDirectory(dataFolder);
            }
            if (!File.Exists(dataFolder + "/" + dataFile))
            {
                StreamWriter sw = File.CreateText(dataFolder + "/" + dataFile);
                sw.Close();
            }

            var reader = new StreamReader(File.OpenRead(dataFolder + "/" + dataFile));
            while (!reader.EndOfStream)
            {
                historyList.Add(reader.ReadLine());
            }
        }

        virtual public void AddIfNewItem(string item)
        {
            item = item.Trim().Replace(",", ";");

            if (historyList.Contains(item)) return;
            if (item == String.Empty) return;

            var appender = new FileStream(dataFolder + "/" + dataFile, FileMode.Append);
            var sw = new StreamWriter(appender);
            sw.WriteLine(item);
            historyList.Add(item);
            sw.Close();
            appender.Close();
        }

        abstract protected void SetDirectoryPath();

        protected List<string> historyList = new List<string>();
        protected string dataFolder;
        protected string dataFile;
    }

    public class TaskHistory : History
    {
        public TaskHistory() : base() { }

        protected override void SetDirectoryPath()
        {
            dataFolder = @"TimesheetData";
            dataFile = @"TaskHistory.txt";
        }

        public AutoCompleteStringCollection UseForAutoComplete()
        {
            AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
            autoComplete.AddRange(historyList.ToArray());

            return autoComplete;
        }
    }

    public class JiraHistory : History
    {
        public JiraHistory() : base() { }

        protected override void SetDirectoryPath()
        {
            dataFolder = @"TimesheetData";
            dataFile = @"JiraHistory.txt";
        }

        public AutoCompleteStringCollection UseForAutoComplete()
        {
            AutoCompleteStringCollection autoComplete = new AutoCompleteStringCollection();
            autoComplete.AddRange(historyList.ToArray());

            return autoComplete;
        }
    }

    public class SubtaskHistory : History
    {
        public SubtaskHistory() : base() { }

        protected override void SetDirectoryPath()
        {
            dataFolder = @"TimesheetData";
            dataFile = @"SubtaskHistory.txt";
        }

        public override void AddIfNewItem(string item)
        {
            if (item == "Other") return;
            base.AddIfNewItem(item);
        }

        public string[] UseAsStringArray()
        {
            return historyList.ToArray();
        }
    }

    public static class StringExt
    {
        // Cool little snippet I found on stack overflow
        public static string Truncate(this string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return value;
            return value.Length <= maxLength ? value : value.Substring(0, maxLength) + "...";
        }

        /// <summary>
        /// Treats the string as a list, delimited by a character. Looks through
        /// the list comparing the given item with every other item and returns true
        /// if it is a duplicate.
        /// </summary>
        /// <param name="value">Given string being treated as a list</param>
        /// <param name="thisItem">Item attempted to be added</param>
        /// <param name="delim">Default delimiter is '/'</param>
        /// <returns>String with our without the last item</returns>
        public static bool Has(this string value, string thisItem, char delim = '/')
        {
            string[] list = value.Split(delim);

            if (list.Contains(thisItem)) return true;

            return false;
        }
    }
}
