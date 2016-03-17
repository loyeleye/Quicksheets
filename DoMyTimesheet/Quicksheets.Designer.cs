using OfficeOpenXml.Style;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace DoMyTimesheet
{
    partial class Quicksheets
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Quicksheets));
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.Main = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.loadCommonButton = new System.Windows.Forms.Button();
            this.saveCommonButton = new System.Windows.Forms.Button();
            this.addEntryButton = new System.Windows.Forms.Button();
            this.timesheetCompletionBar = new System.Windows.Forms.ProgressBar();
            this.panel1 = new System.Windows.Forms.Panel();
            this.eventLog = new System.Windows.Forms.Label();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.dateButton = new System.Windows.Forms.Button();
            this.hourButton = new System.Windows.Forms.Button();
            this.halfDayButton = new System.Windows.Forms.Button();
            this.hoursLabel = new System.Windows.Forms.Label();
            this.allDayButton = new System.Windows.Forms.Button();
            this.hourUpDown = new System.Windows.Forms.NumericUpDown();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.subtaskField = new System.Windows.Forms.TextBox();
            this.subtaskLabel = new System.Windows.Forms.Label();
            this.refreshSubtasksButton = new System.Windows.Forms.Button();
            this.tabPage4 = new System.Windows.Forms.TabPage();
            this.rfcUpDown = new System.Windows.Forms.NumericUpDown();
            this.rfcLabel = new System.Windows.Forms.Label();
            this.tabPage5 = new System.Windows.Forms.TabPage();
            this.fmccUpDown = new System.Windows.Forms.NumericUpDown();
            this.fmccLabel = new System.Windows.Forms.Label();
            this.tabPage6 = new System.Windows.Forms.TabPage();
            this.jiraLabel = new System.Windows.Forms.Label();
            this.tabPage7 = new System.Windows.Forms.TabPage();
            this.taskField = new System.Windows.Forms.TextBox();
            this.taskLabel = new System.Windows.Forms.Label();
            this.tabPage8 = new System.Windows.Forms.TabPage();
            this.nameField = new System.Windows.Forms.TextBox();
            this.nameLabel = new System.Windows.Forms.Label();
            this.tabPage9 = new System.Windows.Forms.TabPage();
            this.projectField = new System.Windows.Forms.TextBox();
            this.projectLabel = new System.Windows.Forms.Label();
            this.tabPage10 = new System.Windows.Forms.TabPage();
            this.customerField = new System.Windows.Forms.TextBox();
            this.customerLabel = new System.Windows.Forms.Label();
            this.editButton = new System.Windows.Forms.Button();
            this.deleteButton = new System.Windows.Forms.Button();
            this.saveButton = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.jiraUpDown = new System.Windows.Forms.NumericUpDown();
            this.jiraDropDown = new System.Windows.Forms.ComboBox();
            this.jiraReadField = new System.Windows.Forms.TextBox();
            this.bindRFCButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.Main.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.panel1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.hourUpDown)).BeginInit();
            this.tabPage3.SuspendLayout();
            this.tabPage4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rfcUpDown)).BeginInit();
            this.tabPage5.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fmccUpDown)).BeginInit();
            this.tabPage6.SuspendLayout();
            this.tabPage7.SuspendLayout();
            this.tabPage8.SuspendLayout();
            this.tabPage9.SuspendLayout();
            this.tabPage10.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.jiraUpDown)).BeginInit();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.Main);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.editButton);
            this.splitContainer1.Panel2.Controls.Add(this.deleteButton);
            this.splitContainer1.Panel2.Controls.Add(this.saveButton);
            this.splitContainer1.Panel2.Controls.Add(this.listBox1);
            this.splitContainer1.Size = new System.Drawing.Size(597, 362);
            this.splitContainer1.SplitterDistance = 199;
            this.splitContainer1.TabIndex = 0;
            // 
            // Main
            // 
            this.Main.Controls.Add(this.tabPage1);
            this.Main.Controls.Add(this.tabPage2);
            this.Main.Controls.Add(this.tabPage3);
            this.Main.Controls.Add(this.tabPage4);
            this.Main.Controls.Add(this.tabPage5);
            this.Main.Controls.Add(this.tabPage6);
            this.Main.Controls.Add(this.tabPage7);
            this.Main.Controls.Add(this.tabPage8);
            this.Main.Controls.Add(this.tabPage9);
            this.Main.Controls.Add(this.tabPage10);
            this.Main.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Main.Location = new System.Drawing.Point(3, 3);
            this.Main.Name = "Main";
            this.Main.SelectedIndex = 0;
            this.Main.Size = new System.Drawing.Size(591, 193);
            this.Main.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight;
            this.Main.TabIndex = 0;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.LightCoral;
            this.tabPage1.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage1.BackgroundImage")));
            this.tabPage1.Controls.Add(this.loadCommonButton);
            this.tabPage1.Controls.Add(this.saveCommonButton);
            this.tabPage1.Controls.Add(this.addEntryButton);
            this.tabPage1.Controls.Add(this.timesheetCompletionBar);
            this.tabPage1.Controls.Add(this.panel1);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(583, 164);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Timesheet";
            // 
            // loadCommonButton
            // 
            this.loadCommonButton.BackColor = System.Drawing.Color.DimGray;
            this.loadCommonButton.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.loadCommonButton.ForeColor = System.Drawing.Color.Cornsilk;
            this.loadCommonButton.Location = new System.Drawing.Point(344, 102);
            this.loadCommonButton.Name = "loadCommonButton";
            this.loadCommonButton.Size = new System.Drawing.Size(224, 36);
            this.loadCommonButton.TabIndex = 4;
            this.loadCommonButton.Text = "Load Common Entry";
            this.loadCommonButton.UseVisualStyleBackColor = false;
            // 
            // saveCommonButton
            // 
            this.saveCommonButton.BackColor = System.Drawing.Color.DimGray;
            this.saveCommonButton.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.saveCommonButton.ForeColor = System.Drawing.Color.Cornsilk;
            this.saveCommonButton.Location = new System.Drawing.Point(344, 60);
            this.saveCommonButton.Name = "saveCommonButton";
            this.saveCommonButton.Size = new System.Drawing.Size(224, 36);
            this.saveCommonButton.TabIndex = 3;
            this.saveCommonButton.Text = "Save As Common Entry";
            this.saveCommonButton.UseVisualStyleBackColor = false;
            // 
            // addEntryButton
            // 
            this.addEntryButton.BackColor = System.Drawing.Color.Black;
            this.addEntryButton.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.addEntryButton.ForeColor = System.Drawing.Color.Cornsilk;
            this.addEntryButton.Location = new System.Drawing.Point(344, 7);
            this.addEntryButton.Name = "addEntryButton";
            this.addEntryButton.Size = new System.Drawing.Size(224, 36);
            this.addEntryButton.TabIndex = 2;
            this.addEntryButton.Text = "Add Timesheet Entry";
            this.addEntryButton.UseVisualStyleBackColor = false;
            // 
            // timesheetCompletionBar
            // 
            this.timesheetCompletionBar.Location = new System.Drawing.Point(6, 148);
            this.timesheetCompletionBar.Name = "timesheetCompletionBar";
            this.timesheetCompletionBar.Size = new System.Drawing.Size(562, 10);
            this.timesheetCompletionBar.TabIndex = 1;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.eventLog);
            this.panel1.Location = new System.Drawing.Point(7, 7);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(296, 135);
            this.panel1.TabIndex = 0;
            // 
            // eventLog
            // 
            this.eventLog.AutoSize = true;
            this.eventLog.Font = new System.Drawing.Font("Monotype Corsiva", 11.75F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.eventLog.ForeColor = System.Drawing.Color.White;
            this.eventLog.Location = new System.Drawing.Point(4, 4);
            this.eventLog.Name = "eventLog";
            this.eventLog.Size = new System.Drawing.Size(0, 18);
            this.eventLog.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage2.BackgroundImage")));
            this.tabPage2.Controls.Add(this.dateButton);
            this.tabPage2.Controls.Add(this.hourButton);
            this.tabPage2.Controls.Add(this.halfDayButton);
            this.tabPage2.Controls.Add(this.hoursLabel);
            this.tabPage2.Controls.Add(this.allDayButton);
            this.tabPage2.Controls.Add(this.hourUpDown);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(583, 164);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Hours";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // dateButton
            // 
            this.dateButton.BackColor = System.Drawing.Color.Navy;
            this.dateButton.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dateButton.ForeColor = System.Drawing.Color.AliceBlue;
            this.dateButton.Location = new System.Drawing.Point(6, 125);
            this.dateButton.Name = "dateButton";
            this.dateButton.Size = new System.Drawing.Size(156, 33);
            this.dateButton.TabIndex = 5;
            this.dateButton.UseVisualStyleBackColor = false;
            // 
            // hourButton
            // 
            this.hourButton.BackColor = System.Drawing.Color.DimGray;
            this.hourButton.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.hourButton.ForeColor = System.Drawing.Color.AliceBlue;
            this.hourButton.Location = new System.Drawing.Point(374, 58);
            this.hourButton.Name = "hourButton";
            this.hourButton.Size = new System.Drawing.Size(177, 44);
            this.hourButton.TabIndex = 4;
            this.hourButton.Text = "One Hour";
            this.hourButton.UseVisualStyleBackColor = false;
            // 
            // halfDayButton
            // 
            this.halfDayButton.BackColor = System.Drawing.Color.DimGray;
            this.halfDayButton.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.halfDayButton.ForeColor = System.Drawing.Color.AliceBlue;
            this.halfDayButton.Location = new System.Drawing.Point(191, 58);
            this.halfDayButton.Name = "halfDayButton";
            this.halfDayButton.Size = new System.Drawing.Size(177, 44);
            this.halfDayButton.TabIndex = 3;
            this.halfDayButton.Text = "Half Day";
            this.halfDayButton.UseVisualStyleBackColor = false;
            // 
            // hoursLabel
            // 
            this.hoursLabel.AutoSize = true;
            this.hoursLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.hoursLabel.ForeColor = System.Drawing.Color.Black;
            this.hoursLabel.Location = new System.Drawing.Point(3, 3);
            this.hoursLabel.Name = "hoursLabel";
            this.hoursLabel.Size = new System.Drawing.Size(386, 20);
            this.hoursLabel.TabIndex = 2;
            this.hoursLabel.Text = "How many hours have you been working on this task?";
            // 
            // allDayButton
            // 
            this.allDayButton.BackColor = System.Drawing.Color.DimGray;
            this.allDayButton.Font = new System.Drawing.Font("Trebuchet MS", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.allDayButton.ForeColor = System.Drawing.Color.AliceBlue;
            this.allDayButton.Location = new System.Drawing.Point(10, 58);
            this.allDayButton.Name = "allDayButton";
            this.allDayButton.Size = new System.Drawing.Size(175, 44);
            this.allDayButton.TabIndex = 1;
            this.allDayButton.Text = "All Day";
            this.allDayButton.UseVisualStyleBackColor = false;
            // 
            // hourUpDown
            // 
            this.hourUpDown.BackColor = System.Drawing.SystemColors.Window;
            this.hourUpDown.DecimalPlaces = 2;
            this.hourUpDown.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.hourUpDown.ForeColor = System.Drawing.SystemColors.InfoText;
            this.hourUpDown.Increment = new decimal(new int[] {
            25,
            0,
            0,
            131072});
            this.hourUpDown.Location = new System.Drawing.Point(464, 127);
            this.hourUpDown.Maximum = new decimal(new int[] {
            24,
            0,
            0,
            0});
            this.hourUpDown.Name = "hourUpDown";
            this.hourUpDown.Size = new System.Drawing.Size(113, 31);
            this.hourUpDown.TabIndex = 0;
            // 
            // tabPage3
            // 
            this.tabPage3.AutoScroll = true;
            this.tabPage3.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage3.BackgroundImage")));
            this.tabPage3.Controls.Add(this.subtaskField);
            this.tabPage3.Controls.Add(this.subtaskLabel);
            this.tabPage3.Controls.Add(this.refreshSubtasksButton);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(583, 164);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Subtask";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // subtaskField
            // 
            this.subtaskField.Location = new System.Drawing.Point(365, 4);
            this.subtaskField.MaxLength = 30;
            this.subtaskField.Name = "subtaskField";
            this.subtaskField.Size = new System.Drawing.Size(212, 22);
            this.subtaskField.TabIndex = 3;
            this.subtaskField.Text = "Other";
            // 
            // subtaskLabel
            // 
            this.subtaskLabel.AutoSize = true;
            this.subtaskLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.subtaskLabel.ForeColor = System.Drawing.Color.Black;
            this.subtaskLabel.Location = new System.Drawing.Point(3, 3);
            this.subtaskLabel.Name = "subtaskLabel";
            this.subtaskLabel.Size = new System.Drawing.Size(355, 20);
            this.subtaskLabel.TabIndex = 2;
            this.subtaskLabel.Text = "What kind of task?  Testing?  Development, etc.?";
            // 
            // refreshSubtasksButton
            // 
            this.refreshSubtasksButton.Image = global::DoMyTimesheet.Properties.Resources.refresh;
            this.refreshSubtasksButton.Location = new System.Drawing.Point(549, 132);
            this.refreshSubtasksButton.Name = "refreshSubtasksButton";
            this.refreshSubtasksButton.Size = new System.Drawing.Size(27, 26);
            this.refreshSubtasksButton.TabIndex = 4;
            this.refreshSubtasksButton.UseVisualStyleBackColor = true;
            // 
            // tabPage4
            // 
            this.tabPage4.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage4.BackgroundImage")));
            this.tabPage4.Controls.Add(this.bindRFCButton);
            this.tabPage4.Controls.Add(this.rfcUpDown);
            this.tabPage4.Controls.Add(this.rfcLabel);
            this.tabPage4.Location = new System.Drawing.Point(4, 25);
            this.tabPage4.Name = "tabPage4";
            this.tabPage4.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage4.Size = new System.Drawing.Size(583, 164);
            this.tabPage4.TabIndex = 3;
            this.tabPage4.Text = "RFC#";
            this.tabPage4.UseVisualStyleBackColor = true;
            // 
            // rfcUpDown
            // 
            this.rfcUpDown.Font = new System.Drawing.Font("Trebuchet MS", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rfcUpDown.ImeMode = System.Windows.Forms.ImeMode.On;
            this.rfcUpDown.Location = new System.Drawing.Point(142, 59);
            this.rfcUpDown.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.rfcUpDown.Name = "rfcUpDown";
            this.rfcUpDown.Size = new System.Drawing.Size(278, 35);
            this.rfcUpDown.TabIndex = 4;
            // 
            // rfcLabel
            // 
            this.rfcLabel.AutoSize = true;
            this.rfcLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rfcLabel.ForeColor = System.Drawing.Color.Black;
            this.rfcLabel.Location = new System.Drawing.Point(3, 3);
            this.rfcLabel.Name = "rfcLabel";
            this.rfcLabel.Size = new System.Drawing.Size(161, 20);
            this.rfcLabel.TabIndex = 3;
            this.rfcLabel.Text = "What was the RFC#?";
            // 
            // tabPage5
            // 
            this.tabPage5.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage5.BackgroundImage")));
            this.tabPage5.Controls.Add(this.fmccUpDown);
            this.tabPage5.Controls.Add(this.fmccLabel);
            this.tabPage5.Location = new System.Drawing.Point(4, 25);
            this.tabPage5.Name = "tabPage5";
            this.tabPage5.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage5.Size = new System.Drawing.Size(583, 164);
            this.tabPage5.TabIndex = 4;
            this.tabPage5.Text = "FMCC PLR#";
            this.tabPage5.UseVisualStyleBackColor = true;
            // 
            // fmccUpDown
            // 
            this.fmccUpDown.Font = new System.Drawing.Font("Trebuchet MS", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fmccUpDown.ImeMode = System.Windows.Forms.ImeMode.On;
            this.fmccUpDown.Location = new System.Drawing.Point(142, 59);
            this.fmccUpDown.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.fmccUpDown.Name = "fmccUpDown";
            this.fmccUpDown.Size = new System.Drawing.Size(278, 35);
            this.fmccUpDown.TabIndex = 6;
            // 
            // fmccLabel
            // 
            this.fmccLabel.AutoSize = true;
            this.fmccLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.fmccLabel.ForeColor = System.Drawing.Color.Black;
            this.fmccLabel.Location = new System.Drawing.Point(3, 3);
            this.fmccLabel.Name = "fmccLabel";
            this.fmccLabel.Size = new System.Drawing.Size(208, 20);
            this.fmccLabel.TabIndex = 5;
            this.fmccLabel.Text = "What was the FMCC PLR#?";
            // 
            // tabPage6
            // 
            this.tabPage6.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage6.BackgroundImage")));
            this.tabPage6.Controls.Add(this.jiraReadField);
            this.tabPage6.Controls.Add(this.jiraDropDown);
            this.tabPage6.Controls.Add(this.jiraUpDown);
            this.tabPage6.Controls.Add(this.jiraLabel);
            this.tabPage6.Location = new System.Drawing.Point(4, 25);
            this.tabPage6.Name = "tabPage6";
            this.tabPage6.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage6.Size = new System.Drawing.Size(583, 164);
            this.tabPage6.TabIndex = 5;
            this.tabPage6.Text = "JIRA#";
            this.tabPage6.UseVisualStyleBackColor = true;
            // 
            // jiraLabel
            // 
            this.jiraLabel.AutoSize = true;
            this.jiraLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.jiraLabel.ForeColor = System.Drawing.Color.Black;
            this.jiraLabel.Location = new System.Drawing.Point(3, 3);
            this.jiraLabel.Name = "jiraLabel";
            this.jiraLabel.Size = new System.Drawing.Size(164, 20);
            this.jiraLabel.TabIndex = 7;
            this.jiraLabel.Text = "What was the JIRA#?";
            // 
            // tabPage7
            // 
            this.tabPage7.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage7.BackgroundImage")));
            this.tabPage7.Controls.Add(this.taskField);
            this.tabPage7.Controls.Add(this.taskLabel);
            this.tabPage7.Location = new System.Drawing.Point(4, 25);
            this.tabPage7.Name = "tabPage7";
            this.tabPage7.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage7.Size = new System.Drawing.Size(583, 164);
            this.tabPage7.TabIndex = 6;
            this.tabPage7.Text = "Task";
            this.tabPage7.UseVisualStyleBackColor = true;
            // 
            // taskField
            // 
            this.taskField.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest;
            this.taskField.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.CustomSource;
            this.taskField.Font = new System.Drawing.Font("Trebuchet MS", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.taskField.Location = new System.Drawing.Point(17, 66);
            this.taskField.Name = "taskField";
            this.taskField.Size = new System.Drawing.Size(550, 32);
            this.taskField.TabIndex = 7;
            // 
            // taskLabel
            // 
            this.taskLabel.AutoSize = true;
            this.taskLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.taskLabel.ForeColor = System.Drawing.Color.Black;
            this.taskLabel.Location = new System.Drawing.Point(3, 3);
            this.taskLabel.Name = "taskLabel";
            this.taskLabel.Size = new System.Drawing.Size(233, 20);
            this.taskLabel.TabIndex = 3;
            this.taskLabel.Text = "What did you do on the project?";
            // 
            // tabPage8
            // 
            this.tabPage8.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage8.BackgroundImage")));
            this.tabPage8.Controls.Add(this.nameField);
            this.tabPage8.Controls.Add(this.nameLabel);
            this.tabPage8.Location = new System.Drawing.Point(4, 25);
            this.tabPage8.Name = "tabPage8";
            this.tabPage8.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage8.Size = new System.Drawing.Size(583, 164);
            this.tabPage8.TabIndex = 7;
            this.tabPage8.Text = "Name";
            this.tabPage8.UseVisualStyleBackColor = true;
            // 
            // nameField
            // 
            this.nameField.Font = new System.Drawing.Font("Trebuchet MS", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameField.Location = new System.Drawing.Point(116, 66);
            this.nameField.Name = "nameField";
            this.nameField.Size = new System.Drawing.Size(351, 32);
            this.nameField.TabIndex = 6;
            // 
            // nameLabel
            // 
            this.nameLabel.AutoSize = true;
            this.nameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.nameLabel.ForeColor = System.Drawing.Color.Black;
            this.nameLabel.Location = new System.Drawing.Point(3, 3);
            this.nameLabel.Name = "nameLabel";
            this.nameLabel.Size = new System.Drawing.Size(136, 20);
            this.nameLabel.TabIndex = 4;
            this.nameLabel.Text = "Enter Your Name:";
            // 
            // tabPage9
            // 
            this.tabPage9.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage9.BackgroundImage")));
            this.tabPage9.Controls.Add(this.projectField);
            this.tabPage9.Controls.Add(this.projectLabel);
            this.tabPage9.Location = new System.Drawing.Point(4, 25);
            this.tabPage9.Name = "tabPage9";
            this.tabPage9.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage9.Size = new System.Drawing.Size(583, 164);
            this.tabPage9.TabIndex = 8;
            this.tabPage9.Text = "Project";
            this.tabPage9.UseVisualStyleBackColor = true;
            // 
            // projectField
            // 
            this.projectField.Font = new System.Drawing.Font("Trebuchet MS", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.projectField.Location = new System.Drawing.Point(116, 66);
            this.projectField.Name = "projectField";
            this.projectField.Size = new System.Drawing.Size(351, 32);
            this.projectField.TabIndex = 5;
            // 
            // projectLabel
            // 
            this.projectLabel.AutoSize = true;
            this.projectLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.projectLabel.ForeColor = System.Drawing.Color.Black;
            this.projectLabel.Location = new System.Drawing.Point(3, 3);
            this.projectLabel.Name = "projectLabel";
            this.projectLabel.Size = new System.Drawing.Size(180, 20);
            this.projectLabel.TabIndex = 4;
            this.projectLabel.Text = "Fill in project name here:";
            // 
            // tabPage10
            // 
            this.tabPage10.BackgroundImage = ((System.Drawing.Image)(resources.GetObject("tabPage10.BackgroundImage")));
            this.tabPage10.Controls.Add(this.customerField);
            this.tabPage10.Controls.Add(this.customerLabel);
            this.tabPage10.Location = new System.Drawing.Point(4, 25);
            this.tabPage10.Name = "tabPage10";
            this.tabPage10.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage10.Size = new System.Drawing.Size(583, 164);
            this.tabPage10.TabIndex = 9;
            this.tabPage10.Text = "Customer";
            this.tabPage10.UseVisualStyleBackColor = true;
            // 
            // customerField
            // 
            this.customerField.Font = new System.Drawing.Font("Trebuchet MS", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.customerField.Location = new System.Drawing.Point(116, 66);
            this.customerField.Name = "customerField";
            this.customerField.Size = new System.Drawing.Size(351, 32);
            this.customerField.TabIndex = 6;
            this.customerField.Text = "AiM";
            // 
            // customerLabel
            // 
            this.customerLabel.AutoSize = true;
            this.customerLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.customerLabel.ForeColor = System.Drawing.Color.Black;
            this.customerLabel.Location = new System.Drawing.Point(3, 3);
            this.customerLabel.Name = "customerLabel";
            this.customerLabel.Size = new System.Drawing.Size(290, 20);
            this.customerLabel.TabIndex = 5;
            this.customerLabel.Text = "Who is the work for?  Aim?  Client?  Etc.";
            // 
            // editButton
            // 
            this.editButton.BackColor = System.Drawing.SystemColors.ButtonFace;
            this.editButton.Image = ((System.Drawing.Image)(resources.GetObject("editButton.Image")));
            this.editButton.Location = new System.Drawing.Point(545, 63);
            this.editButton.Name = "editButton";
            this.editButton.Size = new System.Drawing.Size(30, 29);
            this.editButton.TabIndex = 3;
            this.editButton.UseVisualStyleBackColor = false;
            // 
            // deleteButton
            // 
            this.deleteButton.BackColor = System.Drawing.SystemColors.ControlDarkDark;
            this.deleteButton.Image = ((System.Drawing.Image)(resources.GetObject("deleteButton.Image")));
            this.deleteButton.Location = new System.Drawing.Point(545, 98);
            this.deleteButton.Name = "deleteButton";
            this.deleteButton.Size = new System.Drawing.Size(30, 29);
            this.deleteButton.TabIndex = 2;
            this.deleteButton.UseVisualStyleBackColor = false;
            // 
            // saveButton
            // 
            this.saveButton.BackColor = System.Drawing.SystemColors.ControlDark;
            this.saveButton.Image = ((System.Drawing.Image)(resources.GetObject("saveButton.Image")));
            this.saveButton.Location = new System.Drawing.Point(545, 30);
            this.saveButton.Name = "saveButton";
            this.saveButton.Size = new System.Drawing.Size(30, 27);
            this.saveButton.TabIndex = 1;
            this.saveButton.UseVisualStyleBackColor = false;
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(6, 6);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(533, 147);
            this.listBox1.TabIndex = 0;
            // 
            // jiraUpDown
            // 
            this.jiraUpDown.Font = new System.Drawing.Font("Trebuchet MS", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.jiraUpDown.Location = new System.Drawing.Point(393, 102);
            this.jiraUpDown.Maximum = new decimal(new int[] {
            9999,
            0,
            0,
            0});
            this.jiraUpDown.Name = "jiraUpDown";
            this.jiraUpDown.Size = new System.Drawing.Size(175, 35);
            this.jiraUpDown.TabIndex = 8;
            // 
            // jiraDropDown
            // 
            this.jiraDropDown.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.jiraDropDown.Font = new System.Drawing.Font("Trebuchet MS", 18F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.jiraDropDown.FormattingEnabled = true;
            this.jiraDropDown.Location = new System.Drawing.Point(393, 40);
            this.jiraDropDown.Name = "jiraDropDown";
            this.jiraDropDown.Size = new System.Drawing.Size(175, 37);
            this.jiraDropDown.TabIndex = 9;
            // 
            // jiraReadField
            // 
            this.jiraReadField.Enabled = false;
            this.jiraReadField.Font = new System.Drawing.Font("Trebuchet MS", 24F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.jiraReadField.ForeColor = System.Drawing.SystemColors.WindowText;
            this.jiraReadField.Location = new System.Drawing.Point(48, 62);
            this.jiraReadField.Name = "jiraReadField";
            this.jiraReadField.ReadOnly = true;
            this.jiraReadField.Size = new System.Drawing.Size(273, 45);
            this.jiraReadField.TabIndex = 10;
            // 
            // bindRFCButton
            // 
            this.bindRFCButton.BackColor = System.Drawing.Color.Black;
            this.bindRFCButton.Font = new System.Drawing.Font("Trebuchet MS", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.bindRFCButton.ForeColor = System.Drawing.Color.LightCyan;
            this.bindRFCButton.Location = new System.Drawing.Point(142, 114);
            this.bindRFCButton.Name = "bindRFCButton";
            this.bindRFCButton.Size = new System.Drawing.Size(278, 33);
            this.bindRFCButton.TabIndex = 5;
            this.bindRFCButton.Text = "Bind RFC# to JIRA#";
            this.bindRFCButton.UseVisualStyleBackColor = false;
            this.bindRFCButton.Visible = false;
            // 
            // Quicksheets
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLight;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(597, 362);
            this.Controls.Add(this.splitContainer1);
            this.ForeColor = System.Drawing.Color.CornflowerBlue;
            this.Name = "Quicksheets";
            this.Text = "Quicksheets";
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.Main.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.hourUpDown)).EndInit();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage4.ResumeLayout(false);
            this.tabPage4.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rfcUpDown)).EndInit();
            this.tabPage5.ResumeLayout(false);
            this.tabPage5.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.fmccUpDown)).EndInit();
            this.tabPage6.ResumeLayout(false);
            this.tabPage6.PerformLayout();
            this.tabPage7.ResumeLayout(false);
            this.tabPage7.PerformLayout();
            this.tabPage8.ResumeLayout(false);
            this.tabPage8.PerformLayout();
            this.tabPage9.ResumeLayout(false);
            this.tabPage9.PerformLayout();
            this.tabPage10.ResumeLayout(false);
            this.tabPage10.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.jiraUpDown)).EndInit();
            this.ResumeLayout(false);

        }

        private void RefreshSubtasksButton_Click(object sender, EventArgs e)
        {
            subtaskHistory.LoadFromFile();
            DrawSubtaskCheckboxes();
        }

        private void EditButton_Click(object sender, EventArgs e)
        {
            EditSelectedEntry();
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            int key = this.listBox1.SelectedIndex;

            if (timesheetEntries.ContainsKey(key))
            {
                TryDelete(key);
            }
            else
            {
                TryDeleteAll();
            }
        }

        private enum DeleteType
        {
            None = 0,
            Selected = 1,
            All = 2
        }

        // Got this cool snippet from stack overflow
        private DeleteType ConfirmDelete()
        {
            Form f = new DeleteDialog();
            var result = f.ShowDialog(this);

            if (result == DialogResult.Yes)
            {
                return DeleteType.Selected;
            }
            else if (result == DialogResult.Abort)
            {
                return DeleteType.All;
            }
            else
            {
                return DeleteType.None;
            }
        }

        private void TryDeleteAll()
        {
            if (timesheetEntries.Count > 0)
            {
                if (ConfirmDeleteAll())
                {
                    this.listBox1.Items.Clear();
                    timesheetEntries.Clear();
                }
            }
        }

        private void DeleteAtKey(int selectedIndex)
        {
            this.listBox1.Items.RemoveAt(selectedIndex);
            timesheetEntries.Remove(selectedIndex);
            Dictionary<int, string> moveKeys = new Dictionary<int, string>();

            foreach (KeyValuePair<int,string> entry in timesheetEntries)
            {
                if (entry.Key > selectedIndex) moveKeys.Add(entry.Key - 1, entry.Value);
                else moveKeys.Add(entry.Key, entry.Value);
            }

            timesheetEntries = moveKeys;
        }

        private void TryDelete(int selectedIndex)
        {
            switch (ConfirmDelete())
            {
                case DeleteType.None:
                    return;

                case DeleteType.Selected:
                    DeleteAtKey(selectedIndex);
                    return;

                case DeleteType.All:
                    this.listBox1.Items.Clear();
                    timesheetEntries.Clear();
                    return;
            }
        }

        private bool ConfirmDeleteAll()
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure you would like to delete all timesheet entries?", "Delete Entries", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                return true;
            }
            else if (dialogResult == DialogResult.No)
            {
                return false;
            }
            return false;
        }

        private bool ConfirmSave()
        {
            DialogResult dialogResult = MessageBox.Show("You are about to move all your entries to your timesheet(s).\nWould you like to save?", "Save Entries", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                return true;
            }
            else if (dialogResult == DialogResult.No)
            {
                return false;
            }
            return false;
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {
            if (timesheetEntries.Count < 1) return;
            if (!ConfirmSave()) return;
            List<int> successfulKeys = new List<int>();
            // Iterate through list of entries
            foreach (KeyValuePair<int, string> entry in timesheetEntries)
            {
                string[] entryValues = entry.Value.Split(',');

                ExcelCriteria saveCriteria = new ExcelCriteria(entryValues);

                if (!Directory.Exists("TimesheetData/Timesheets"))
                {
                    DirectoryInfo newDir = Directory.CreateDirectory("TimesheetData/Timesheets");
                }

                try
                {
                    SaveToExcel(saveCriteria);
                    successfulKeys.Add(entry.Key);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(string.Format("Error: Entry index {0} failed.\nOriginal error: {1}.", entry.Key, ex.Message));
                }
            }

            // TODO: This is a very bad way to do this save. It is basically rebuilding the timesheet 
            // entries dictionary every time you delete a key. Need to refactor this.

            // Remove successful keys from timesheet list
            if (successfulKeys.Count == timesheetEntries.Count)
            {
                this.listBox1.Items.Clear();
                timesheetEntries.Clear();
            }
            else
            {
                foreach (int key in successfulKeys)
                {
                    DeleteAtKey(key);
                }
            }
        }

        private void AddEntryButton_Click(object sender, EventArgs e)
        {
            string listBoxEntry = dateSelected.ToShortDateString()
                + "\t" + this.taskField.Text.Truncate(20)
                + "\t" + GetSubtasksToStringIfMarked().Truncate(30)
                + "\t" + this.hourUpDown.Value.ToString();
            this.listBox1.Items.Add(listBoxEntry);
            int k = this.listBox1.Items.Count - 1;
            string v = SaveTimesheetEntry();
            timesheetEntries.Add(k, v);
        }

        private void EditSelectedEntry()
        {
            int key = this.listBox1.SelectedIndex;
            string csv;

            if (!timesheetEntries.TryGetValue(key, out csv)) return;

            ShowEntry f = new ShowEntry();
            f.TitleLabel = listBox1.SelectedItem.ToString();
            f.TextField = csv;

            var result = f.ShowDialog(this);
            if (result == DialogResult.Cancel)
            {
                f.Close();
            }
            if (result == DialogResult.Retry)
            {
                csv = f.TextField;
            }
            if (result == DialogResult.OK)
            {
                csv = f.TextField;
                f.Close();
            }

            timesheetEntries[key] = csv;
            string[] values = csv.Split(',');
            this.listBox1.Items[key] = values[0] +
                "\t" + values[4].Truncate(20) +
                "\t" + values[8].Truncate(30) +
                "\t" + values[9];
        }

        private void ListBox1_DoubleClick(object sender, EventArgs e)
        {
            EditSelectedEntry();
        }

        private class ExcelCriteria
        {
            public ExcelCriteria(string[] entryValues)
            {
                string date = entryValues[0];
                Date = date;
                Name = TimesheetNameForDate(date);
                DateFormula = GetDateFormula(date);
                NameWithExt = Name + ".xlsx";
                Path = "TimesheetData/Timesheets/" + NameWithExt;
                BackupPath = "TimesheetData/Timesheets/Backups/" + NameWithExt;
                WorksheetName = GetWorksheetName(date);
                EntryValues = entryValues;
            }

            public string GetWorksheetName(string date)
            {
                string rtnName = "";

                string[] MDY = date.Split('/');
                int month;
                if (!Int32.TryParse(MDY[0], out month)) return "Timesheet";
                rtnName += System.Globalization
                    .CultureInfo.CurrentCulture
                    .DateTimeFormat.GetAbbreviatedMonthName(month);
                rtnName += " 01 " + MDY[2];

                return rtnName;
            }

            /// <summary>
            /// Returns timesheet name for a given date (always first of month)
            /// ex) Oyeleye 2015-03-01_timesheet
            /// </summary>
            /// <param name="date">Date that will determine timesheet name</param>
            private string TimesheetNameForDate(string date)
            {
                string[] dateValues = date.Split('/');
                string year = dateValues[2].PadLeft(4, '0');
                string month = dateValues[0].PadLeft(2, '0');
                string dateSubstring = year + '-' + month + '-' + "01";
                string ifLastName = (lastName == String.Empty) ? "" : lastName + " ";

                return ifLastName + dateSubstring + "_timesheet";
            }

            private string GetDateFormula(string date)
            {
                string[] dateValues = date.Split('/');
                string year = dateValues[2].PadLeft(4, '0');
                string day = dateValues[1].PadLeft(2, '0');
                string month = dateValues[0].PadLeft(2, '0');

                return "=DATE(" + year + ',' + month + ',' + day + ')';
            }

            public string Date { get; private set; }
            public string Name { get; private set; }
            public string NameWithExt { get; private set; }
            public string Path { get; private set; }
            public string BackupPath { get; private set; }
            public string WorksheetName { get; private set; }
            public string[] EntryValues { get; private set; }
            public string DateFormula { get; private set; }
        }

        private void BackupExcel(ExcelCriteria criteria)
        {
            // If the original file doesn't exist, there's nothing to backup.
            if (!File.Exists(criteria.Path)) return;

            if (!Directory.Exists("TimesheetData/Timesheets/Backups"))
            {
                Directory.CreateDirectory("TimesheetData/Timesheets/Backups");
            }

            // If a backup doesn't exist yet, create one.
            if (!File.Exists(criteria.BackupPath))
            {
                File.WriteAllText(criteria.BackupPath, "");
            }

            File.Copy(criteria.Path, criteria.BackupPath, true);
        }

        private void SaveToExcel(ExcelCriteria criteria)
        {
            FileInfo file;
            bool isNewFile = !File.Exists(criteria.Path);

            if (!isNewFile) BackupExcel(criteria);

            file = new FileInfo(criteria.Path);
            using (var package = new ExcelPackage(file))
            {
                ExcelWorksheet timesheet;
                package.Workbook.Properties.Author = lastName;
                package.Workbook.Properties.Title = criteria.Name;

                if (isNewFile || package.Workbook.Worksheets.Count < 1)
                {
                    package.Workbook.Properties.Comments = "Auto-generated by Quicksheets";
                    package.Workbook.Worksheets.Add(criteria.WorksheetName);
                    timesheet = package.Workbook.Worksheets[1];
                    timesheet.DefaultColWidth = 15;
                    timesheet.Column(3).Width = 30;
                    timesheet.Column(4).Width = 20;
                    timesheet.Column(5).Width = 50;
                    timesheet.Column(6).Width = 20;
                    timesheet.Column(7).Width = 20;
                    timesheet.Column(8).Width = 20;
                    timesheet.Column(9).Width = 50;
                    timesheet.Column(10).Width = 10;
                    timesheet.Cells["A1:J1"].Style.Font.SetFromFont(new Font("Calibri", 11));
                    timesheet.Cells["A1:J1"].Style.Font.Bold = true;
                    timesheet.Cells["A1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["B1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["C1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["D1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["E1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["F1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["G1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["H1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["I1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["J1"].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    timesheet.Cells["A1:J1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.CenterContinuous;
                    // Headers
                    timesheet.Cells["A1"].Value = "Date";
                    timesheet.Cells["B1"].Value = "Customer";
                    timesheet.Cells["C1"].Value = "Project";
                    timesheet.Cells["D1"].Value = "Resource";
                    timesheet.Cells["E1"].Value = "TASK";
                    timesheet.Cells["F1"].Value = "Jira Ticket #";
                    timesheet.Cells["G1"].Value = "FMCC PLR #";
                    timesheet.Cells["H1"].Value = "RFC#";
                    timesheet.Cells["I1"].Value = "SUBTASK";
                    timesheet.Cells["J1"].Value = "HOURS";
                    // Subheaders
                    timesheet.Cells["B2"].Value = "Who is the work for?  Aim?  Client?  Etc.";
                    timesheet.Cells["B2"].Style.WrapText = true;
                    timesheet.Cells["C2"].Value = "Fill in project name here";
                    timesheet.Cells["D2"].Value = "Your name";
                    timesheet.Cells["E2"].Value = "What did you do on the project?";
                    timesheet.Cells["E2"].Style.WrapText = true;
                    timesheet.Cells["F2"].Value = "Add Jira # here";
                    timesheet.Cells["G2"].Value = "Add FMCC PLR # here";
                    timesheet.Cells["H2"].Value = "What was the RFC#";
                    timesheet.Cells["I2"].Value = "What kind of task?  Testing?  Development, etc.?";
                    timesheet.Cells["J2"].Value = "1";
                }
                timesheet = package.Workbook.Worksheets[1];
                int rowNum = timesheet.Dimension.End.Row + 1;

                timesheet.Cells[rowNum, 1].Style.Numberformat.Format = "d-MMM";
                timesheet.Cells[rowNum, 1].Formula = criteria.DateFormula;

                for (int colNum = 1; colNum < criteria.EntryValues.Length; colNum++)
                {
                    timesheet.Cells[rowNum, colNum + 1].Value = criteria.EntryValues[colNum];
                }

                package.Save();
            }
        }

        private void SetCompletionBarAndEventLog()
        {
            TrimTextBoxFields();

            int completed_items = 0;
            string event_log = "";

            // Hours
            if (this.hourUpDown.Value > 0) completed_items++;
            else event_log += "Your hours have not been set.\n";
            // Subtasks
            if (GetSubtasksToStringIfMarked() != "") completed_items++;
            else event_log += "You must choose at least one subtask.\n";
            // RFC JIRA FMCC
            if (this.rfcUpDown.Value > 0
                || this.jiraField.Text != String.Empty
                || this.fmccUpDown.Value > 0
                )
                completed_items++;
            else event_log += "You must have either a RFC, JIRA, or FMCC#.\n";
            // Task
            if (this.taskField.Text != "") completed_items++;
            else event_log += "You have not entered a task.\n";
            // Name
            if (this.nameField.Text != "") completed_items++;
            else event_log += "You must write your name.\n";
            // Project
            if (this.projectField.Text != "") completed_items++;
            else event_log += "You must write your project name.\n";
            // Customer
            if (this.customerField.Text != "") completed_items++;
            else event_log += "Please write the customer you did work for.\n";

            if (event_log == String.Empty)
            {
                event_log = "Your timesheet entry is filled out.\n";
            }
            this.timesheetCompletionBar.Value = Convert.ToInt32((100.0 / 7) * completed_items);
            this.eventLog.Text = event_log;
        }

        private void Main_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.Main.SelectedIndex != 0)
            {
                return;
            }

            SetCompletionBarAndEventLog();
        }

        private void HourButton_Click(object sender, EventArgs e)
        {
            this.hourUpDown.Value = 1;
        }

        private void AllDayButton_Click(object sender, EventArgs e)
        {
            this.hourUpDown.Value = 8;
        }

        private void HalfDayButton_Click(object sender, EventArgs e)
        {
            this.hourUpDown.Value = 4;
        }

        private void DateButton_Click(object sender, EventArgs e)
        {
            SetDateForm f = new SetDateForm(dateSelected);

            var result = f.ShowDialog(this);
            if (result == DialogResult.Cancel)
            {
                f.Close();
            }
            if (result == DialogResult.OK)
            {
                dateSelected = f.MonthCalendar.SelectionStart;
                this.dateButton.Text = "Date: " + dateSelected.ToShortDateString();
                f.Close();
            }
        }

        private void TrimTextBoxFields()
        {
            foreach (Control child in this.Controls)
            {
                if (child is TextBox)
                {
                    child.Text = child.Text.Trim();
                }
            }
        }

        private void SaveCommonButton_Click(object sender, EventArgs e)
        {
            Stream fileStream;
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            string saveDir = Environment.CurrentDirectory + "/TimesheetData/Common";

            saveFileDialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            saveFileDialog.FilterIndex = 1;
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.FileName = "CommonEntry.csv";

            if (!Directory.Exists(saveDir))
            {
                DirectoryInfo newDir = Directory.CreateDirectory(saveDir);
            }

            saveFileDialog.InitialDirectory = Path.GetFullPath(saveDir);

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if ((fileStream = saveFileDialog.OpenFile()) != null)
                {
                    // Code to write the stream goes here.
                    StreamWriter sw = new StreamWriter(fileStream);
                    sw.WriteLine(SaveTimesheetEntry());
                    sw.Close();
                    fileStream.Close();
                }
            }
        }

        private void LoadCommonButton_Click(object sender, EventArgs e)
        {
            Stream fileStream = null;
            OpenFileDialog openFileDialog = new OpenFileDialog();

            string loadDir = Environment.CurrentDirectory + "/TimesheetData/Common";

            if (!Directory.Exists(loadDir))
            {
                loadDir = Environment.CurrentDirectory;
            }

            openFileDialog.InitialDirectory = Path.GetFullPath(loadDir);
            openFileDialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            openFileDialog.FilterIndex = 1;
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if ((fileStream = openFileDialog.OpenFile()) != null)
                    {
                        using (fileStream)
                        {
                            // Insert code to read the stream here.
                            StreamReader sr = new StreamReader(fileStream);
                            string csvLine = sr.ReadLine();
                            LoadTimesheetEntry(csvLine);
                            sr.Close();
                            fileStream.Close();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }
        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TabControl Main;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.TabPage tabPage4;
        private System.Windows.Forms.TabPage tabPage5;
        private System.Windows.Forms.TabPage tabPage6;
        private System.Windows.Forms.TabPage tabPage7;
        private System.Windows.Forms.TabPage tabPage8;
        private System.Windows.Forms.TabPage tabPage9;
        private System.Windows.Forms.TabPage tabPage10;
        private System.Windows.Forms.Button addEntryButton;
        private System.Windows.Forms.ProgressBar timesheetCompletionBar;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button saveCommonButton;
        private System.Windows.Forms.Button loadCommonButton;
        private System.Windows.Forms.Label hoursLabel;
        private System.Windows.Forms.Button allDayButton;
        private System.Windows.Forms.NumericUpDown hourUpDown;
        private System.Windows.Forms.Button hourButton;
        private System.Windows.Forms.Button halfDayButton;
        private System.Windows.Forms.Label subtaskLabel;
        private System.Windows.Forms.NumericUpDown rfcUpDown;
        private System.Windows.Forms.Label rfcLabel;
        private System.Windows.Forms.NumericUpDown fmccUpDown;
        private System.Windows.Forms.Label fmccLabel;
        private System.Windows.Forms.Label jiraLabel;
        private System.Windows.Forms.Label taskLabel;
        private System.Windows.Forms.TextBox nameField;
        private System.Windows.Forms.Label nameLabel;
        private System.Windows.Forms.TextBox projectField;
        private System.Windows.Forms.Label projectLabel;
        private System.Windows.Forms.TextBox customerField;
        private System.Windows.Forms.Label customerLabel;
        private System.Windows.Forms.TextBox taskField;
        private TextBox subtaskField;
        private Button dateButton;
        private ListBox listBox1;
        private Label eventLog;
        private Button editButton;
        private Button deleteButton;
        private Button saveButton;
        private Button refreshSubtasksButton;
        private Button bindRFCButton;
        private TextBox jiraReadField;
        private ComboBox jiraDropDown;
        private NumericUpDown jiraUpDown;
    }
}