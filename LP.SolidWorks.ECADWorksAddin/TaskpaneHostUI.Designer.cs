namespace LP.SolidWorks.ECADWorksAddin
{
    partial class TaskpaneHostUI
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.buttonBrowseJson = new System.Windows.Forms.Button();
            this.colorPresets = new System.Windows.Forms.ComboBox();
            this.buttonApplyStyle = new System.Windows.Forms.Button();
            this.comboComponentList = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.checkLoadSelected = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.checkBoxBottomOnly = new System.Windows.Forms.CheckBox();
            this.buttonLoadLibrary = new System.Windows.Forms.Button();
            this.checkBoxTopOnly = new System.Windows.Forms.CheckBox();
            this.buttonClearModel = new System.Windows.Forms.Button();
            this.buttonBrowseModel = new System.Windows.Forms.Button();
            this.labelInstances = new System.Windows.Forms.Label();
            this.label31 = new System.Windows.Forms.Label();
            this.labelAssignedFile = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.chkIsSMD = new System.Windows.Forms.CheckBox();
            this.buttonAutoAlign = new System.Windows.Forms.Button();
            this.checkRotateSelected = new System.Windows.Forms.CheckBox();
            this.label6 = new System.Windows.Forms.Label();
            this.comboAlongAxis = new System.Windows.Forms.ComboBox();
            this.numericUpDownAngle = new System.Windows.Forms.NumericUpDown();
            this.btnRotate = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label7 = new System.Windows.Forms.Label();
            this.lblHolesColor = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.lblBoardColor = new System.Windows.Forms.Label();
            this.lbl2 = new System.Windows.Forms.Label();
            this.lblSilkColor = new System.Windows.Forms.Label();
            this.lbl3 = new System.Windows.Forms.Label();
            this.lblPadColor = new System.Windows.Forms.Label();
            this.lb = new System.Windows.Forms.Label();
            this.lblTraceColor = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.lblMaskColor = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.buttonSaveJSON = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label13 = new System.Windows.Forms.Label();
            this.numericUpDownIncrement = new System.Windows.Forms.NumericUpDown();
            this.buttonZMinus = new System.Windows.Forms.Button();
            this.buttonZPlus = new System.Windows.Forms.Button();
            this.buttonYMinus = new System.Windows.Forms.Button();
            this.buttonYPlus = new System.Windows.Forms.Button();
            this.buttonXMinus = new System.Windows.Forms.Button();
            this.buttonXPlus = new System.Windows.Forms.Button();
            this.label12 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.checkBoxCollisionDetection = new System.Windows.Forms.CheckBox();
            this.buttonSaveLibrary = new System.Windows.Forms.Button();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.buttonDecalImages = new System.Windows.Forms.Button();
            this.buttonGerberFiles = new System.Windows.Forms.Button();
            this.buttonChangeHeight = new System.Windows.Forms.Button();
            this.label15 = new System.Windows.Forms.Label();
            this.numericUpDownBoardHeight = new System.Windows.Forms.NumericUpDown();
            this.groupBox5 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownAngle)).BeginInit();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownIncrement)).BeginInit();
            this.groupBox4.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownBoardHeight)).BeginInit();
            this.groupBox5.SuspendLayout();
            this.SuspendLayout();
            // 
            // buttonBrowseJson
            // 
            this.buttonBrowseJson.Location = new System.Drawing.Point(6, 19);
            this.buttonBrowseJson.Name = "buttonBrowseJson";
            this.buttonBrowseJson.Size = new System.Drawing.Size(101, 23);
            this.buttonBrowseJson.TabIndex = 7;
            this.buttonBrowseJson.Text = "Open Board";
            this.buttonBrowseJson.UseVisualStyleBackColor = true;
            this.buttonBrowseJson.Click += new System.EventHandler(this.button8_Click);
            // 
            // colorPresets
            // 
            this.colorPresets.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.colorPresets.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.colorPresets.FormattingEnabled = true;
            this.colorPresets.Location = new System.Drawing.Point(6, 40);
            this.colorPresets.Name = "colorPresets";
            this.colorPresets.Size = new System.Drawing.Size(231, 21);
            this.colorPresets.TabIndex = 8;
            this.colorPresets.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.colorPresets_DrawItem);
            this.colorPresets.SelectedIndexChanged += new System.EventHandler(this.colorPresets_SelectedIndexChanged);
            // 
            // buttonApplyStyle
            // 
            this.buttonApplyStyle.Location = new System.Drawing.Point(135, 109);
            this.buttonApplyStyle.Name = "buttonApplyStyle";
            this.buttonApplyStyle.Size = new System.Drawing.Size(101, 23);
            this.buttonApplyStyle.TabIndex = 9;
            this.buttonApplyStyle.Text = "Apply Style";
            this.buttonApplyStyle.UseVisualStyleBackColor = true;
            this.buttonApplyStyle.Click += new System.EventHandler(this.buttonApplyStyle_Click);
            // 
            // comboComponentList
            // 
            this.comboComponentList.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.comboComponentList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboComponentList.FormattingEnabled = true;
            this.comboComponentList.Location = new System.Drawing.Point(6, 40);
            this.comboComponentList.Name = "comboComponentList";
            this.comboComponentList.Size = new System.Drawing.Size(231, 21);
            this.comboComponentList.TabIndex = 10;
            this.comboComponentList.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.comboComponentList_DrawItem);
            this.comboComponentList.SelectedIndexChanged += new System.EventHandler(this.comboComponentList_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.checkLoadSelected);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.checkBoxBottomOnly);
            this.groupBox1.Controls.Add(this.buttonLoadLibrary);
            this.groupBox1.Controls.Add(this.checkBoxTopOnly);
            this.groupBox1.Controls.Add(this.buttonClearModel);
            this.groupBox1.Controls.Add(this.buttonBrowseModel);
            this.groupBox1.Controls.Add(this.labelInstances);
            this.groupBox1.Controls.Add(this.label31);
            this.groupBox1.Controls.Add(this.labelAssignedFile);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Controls.Add(this.comboComponentList);
            this.groupBox1.Location = new System.Drawing.Point(0, 296);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(244, 221);
            this.groupBox1.TabIndex = 11;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Component Settings";
            // 
            // checkLoadSelected
            // 
            this.checkLoadSelected.AutoSize = true;
            this.checkLoadSelected.Location = new System.Drawing.Point(7, 195);
            this.checkLoadSelected.Name = "checkLoadSelected";
            this.checkLoadSelected.Size = new System.Drawing.Size(95, 17);
            this.checkLoadSelected.TabIndex = 43;
            this.checkLoadSelected.Text = "Load Selected";
            this.checkLoadSelected.UseVisualStyleBackColor = true;
            this.checkLoadSelected.CheckedChanged += new System.EventHandler(this.checkLoadSelected_CheckedChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(4, 64);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(112, 13);
            this.label3.TabIndex = 42;
            this.label3.Text = "Highlight selected on :";
            // 
            // checkBoxBottomOnly
            // 
            this.checkBoxBottomOnly.AutoSize = true;
            this.checkBoxBottomOnly.Location = new System.Drawing.Point(176, 64);
            this.checkBoxBottomOnly.Name = "checkBoxBottomOnly";
            this.checkBoxBottomOnly.Size = new System.Drawing.Size(59, 17);
            this.checkBoxBottomOnly.TabIndex = 41;
            this.checkBoxBottomOnly.Text = "Bottom";
            this.checkBoxBottomOnly.UseVisualStyleBackColor = true;
            this.checkBoxBottomOnly.CheckedChanged += new System.EventHandler(this.checkBoxBottomOnly_CheckedChanged);
            // 
            // buttonLoadLibrary
            // 
            this.buttonLoadLibrary.Location = new System.Drawing.Point(138, 192);
            this.buttonLoadLibrary.Name = "buttonLoadLibrary";
            this.buttonLoadLibrary.Size = new System.Drawing.Size(101, 23);
            this.buttonLoadLibrary.TabIndex = 17;
            this.buttonLoadLibrary.Text = "Load Library";
            this.buttonLoadLibrary.UseVisualStyleBackColor = true;
            this.buttonLoadLibrary.Click += new System.EventHandler(this.buttonLoadLibrary_Click);
            // 
            // checkBoxTopOnly
            // 
            this.checkBoxTopOnly.AutoSize = true;
            this.checkBoxTopOnly.Location = new System.Drawing.Point(126, 64);
            this.checkBoxTopOnly.Name = "checkBoxTopOnly";
            this.checkBoxTopOnly.Size = new System.Drawing.Size(45, 17);
            this.checkBoxTopOnly.TabIndex = 40;
            this.checkBoxTopOnly.Text = "Top";
            this.checkBoxTopOnly.UseVisualStyleBackColor = true;
            this.checkBoxTopOnly.CheckedChanged += new System.EventHandler(this.checkBoxTopOnly_CheckedChanged);
            // 
            // buttonClearModel
            // 
            this.buttonClearModel.Enabled = false;
            this.buttonClearModel.Location = new System.Drawing.Point(6, 163);
            this.buttonClearModel.Name = "buttonClearModel";
            this.buttonClearModel.Size = new System.Drawing.Size(101, 23);
            this.buttonClearModel.TabIndex = 16;
            this.buttonClearModel.Text = "Clear Model";
            this.buttonClearModel.UseVisualStyleBackColor = true;
            this.buttonClearModel.Click += new System.EventHandler(this.buttonClearModel_Click);
            // 
            // buttonBrowseModel
            // 
            this.buttonBrowseModel.Enabled = false;
            this.buttonBrowseModel.Location = new System.Drawing.Point(138, 163);
            this.buttonBrowseModel.Name = "buttonBrowseModel";
            this.buttonBrowseModel.Size = new System.Drawing.Size(101, 23);
            this.buttonBrowseModel.TabIndex = 15;
            this.buttonBrowseModel.Text = "Browse Model";
            this.buttonBrowseModel.UseVisualStyleBackColor = true;
            this.buttonBrowseModel.Click += new System.EventHandler(this.buttonBrowseModel_Click);
            // 
            // labelInstances
            // 
            this.labelInstances.AutoSize = true;
            this.labelInstances.Location = new System.Drawing.Point(6, 145);
            this.labelInstances.Name = "labelInstances";
            this.labelInstances.Size = new System.Drawing.Size(33, 13);
            this.labelInstances.TabIndex = 14;
            this.labelInstances.Text = "None";
            // 
            // label31
            // 
            this.label31.AutoSize = true;
            this.label31.Location = new System.Drawing.Point(6, 128);
            this.label31.Name = "label31";
            this.label31.Size = new System.Drawing.Size(105, 13);
            this.label31.TabIndex = 13;
            this.label31.Text = "Number of Instances";
            // 
            // labelAssignedFile
            // 
            this.labelAssignedFile.AutoSize = true;
            this.labelAssignedFile.Location = new System.Drawing.Point(6, 106);
            this.labelAssignedFile.Name = "labelAssignedFile";
            this.labelAssignedFile.Size = new System.Drawing.Size(33, 13);
            this.labelAssignedFile.TabIndex = 12;
            this.labelAssignedFile.Text = "None";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(6, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(69, 13);
            this.label2.TabIndex = 11;
            this.label2.Text = "Assigned File";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Component List";
            // 
            // chkIsSMD
            // 
            this.chkIsSMD.AutoSize = true;
            this.chkIsSMD.Checked = true;
            this.chkIsSMD.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkIsSMD.Location = new System.Drawing.Point(83, 147);
            this.chkIsSMD.Name = "chkIsSMD";
            this.chkIsSMD.Size = new System.Drawing.Size(96, 17);
            this.chkIsSMD.TabIndex = 38;
            this.chkIsSMD.Text = "SMD Package";
            this.chkIsSMD.UseVisualStyleBackColor = true;
            this.chkIsSMD.CheckedChanged += new System.EventHandler(this.chkIsSMD_CheckedChanged);
            // 
            // buttonAutoAlign
            // 
            this.buttonAutoAlign.Location = new System.Drawing.Point(3, 145);
            this.buttonAutoAlign.Name = "buttonAutoAlign";
            this.buttonAutoAlign.Size = new System.Drawing.Size(75, 23);
            this.buttonAutoAlign.TabIndex = 37;
            this.buttonAutoAlign.Text = "Auto Align";
            this.buttonAutoAlign.UseVisualStyleBackColor = true;
            this.buttonAutoAlign.Click += new System.EventHandler(this.buttonAutoAlign_Click_1);
            // 
            // checkRotateSelected
            // 
            this.checkRotateSelected.AutoSize = true;
            this.checkRotateSelected.Location = new System.Drawing.Point(6, 178);
            this.checkRotateSelected.Name = "checkRotateSelected";
            this.checkRotateSelected.Size = new System.Drawing.Size(182, 17);
            this.checkRotateSelected.TabIndex = 36;
            this.checkRotateSelected.Text = "Transform only the selected ones";
            this.checkRotateSelected.UseVisualStyleBackColor = true;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(151, 22);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(57, 13);
            this.label6.TabIndex = 35;
            this.label6.Text = "Deg Along";
            // 
            // comboAlongAxis
            // 
            this.comboAlongAxis.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboAlongAxis.FormattingEnabled = true;
            this.comboAlongAxis.Items.AddRange(new object[] {
            "X",
            "Y",
            "Z"});
            this.comboAlongAxis.Location = new System.Drawing.Point(210, 18);
            this.comboAlongAxis.Name = "comboAlongAxis";
            this.comboAlongAxis.Size = new System.Drawing.Size(30, 21);
            this.comboAlongAxis.TabIndex = 34;
            this.comboAlongAxis.SelectedIndexChanged += new System.EventHandler(this.comboAlongAxis_SelectedIndexChanged);
            // 
            // numericUpDownAngle
            // 
            this.numericUpDownAngle.Increment = new decimal(new int[] {
            45,
            0,
            0,
            0});
            this.numericUpDownAngle.Location = new System.Drawing.Point(89, 19);
            this.numericUpDownAngle.Maximum = new decimal(new int[] {
            360,
            0,
            0,
            0});
            this.numericUpDownAngle.Name = "numericUpDownAngle";
            this.numericUpDownAngle.Size = new System.Drawing.Size(62, 20);
            this.numericUpDownAngle.TabIndex = 33;
            // 
            // btnRotate
            // 
            this.btnRotate.Location = new System.Drawing.Point(6, 18);
            this.btnRotate.Name = "btnRotate";
            this.btnRotate.Size = new System.Drawing.Size(75, 23);
            this.btnRotate.TabIndex = 32;
            this.btnRotate.Text = "Rotate";
            this.btnRotate.UseVisualStyleBackColor = true;
            this.btnRotate.Click += new System.EventHandler(this.button2_Click_1);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.lblHolesColor);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.lblBoardColor);
            this.groupBox2.Controls.Add(this.lbl2);
            this.groupBox2.Controls.Add(this.lblSilkColor);
            this.groupBox2.Controls.Add(this.lbl3);
            this.groupBox2.Controls.Add(this.lblPadColor);
            this.groupBox2.Controls.Add(this.lb);
            this.groupBox2.Controls.Add(this.lblTraceColor);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.lblMaskColor);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.colorPresets);
            this.groupBox2.Controls.Add(this.buttonApplyStyle);
            this.groupBox2.Location = new System.Drawing.Point(0, 148);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(244, 142);
            this.groupBox2.TabIndex = 12;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Traces and Silkscreen";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(200, 86);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(34, 13);
            this.label7.TabIndex = 33;
            this.label7.Text = "Holes";
            // 
            // lblHolesColor
            // 
            this.lblHolesColor.BackColor = System.Drawing.Color.Green;
            this.lblHolesColor.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblHolesColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHolesColor.Location = new System.Drawing.Point(180, 86);
            this.lblHolesColor.Name = "lblHolesColor";
            this.lblHolesColor.Size = new System.Drawing.Size(13, 13);
            this.lblHolesColor.TabIndex = 32;
            this.lblHolesColor.Text = "  ";
            this.lblHolesColor.Click += new System.EventHandler(this.lblHolesColor_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(200, 68);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(35, 13);
            this.label9.TabIndex = 31;
            this.label9.Text = "Board";
            // 
            // lblBoardColor
            // 
            this.lblBoardColor.BackColor = System.Drawing.Color.Green;
            this.lblBoardColor.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblBoardColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBoardColor.Location = new System.Drawing.Point(180, 68);
            this.lblBoardColor.Name = "lblBoardColor";
            this.lblBoardColor.Size = new System.Drawing.Size(13, 13);
            this.lblBoardColor.TabIndex = 30;
            this.lblBoardColor.Text = "  ";
            this.lblBoardColor.Click += new System.EventHandler(this.lblBoardColor_Click);
            // 
            // lbl2
            // 
            this.lbl2.AutoSize = true;
            this.lbl2.Location = new System.Drawing.Point(105, 86);
            this.lbl2.Name = "lbl2";
            this.lbl2.Size = new System.Drawing.Size(56, 13);
            this.lbl2.TabIndex = 29;
            this.lbl2.Text = "Silkscreen";
            // 
            // lblSilkColor
            // 
            this.lblSilkColor.BackColor = System.Drawing.Color.Green;
            this.lblSilkColor.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblSilkColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSilkColor.Location = new System.Drawing.Point(85, 86);
            this.lblSilkColor.Name = "lblSilkColor";
            this.lblSilkColor.Size = new System.Drawing.Size(13, 13);
            this.lblSilkColor.TabIndex = 28;
            this.lblSilkColor.Text = "  ";
            this.lblSilkColor.Click += new System.EventHandler(this.lblSilkColor_Click);
            // 
            // lbl3
            // 
            this.lbl3.AutoSize = true;
            this.lbl3.Location = new System.Drawing.Point(28, 86);
            this.lbl3.Name = "lbl3";
            this.lbl3.Size = new System.Drawing.Size(31, 13);
            this.lbl3.TabIndex = 27;
            this.lbl3.Text = "Pads";
            // 
            // lblPadColor
            // 
            this.lblPadColor.BackColor = System.Drawing.Color.Green;
            this.lblPadColor.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblPadColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPadColor.Location = new System.Drawing.Point(8, 86);
            this.lblPadColor.Name = "lblPadColor";
            this.lblPadColor.Size = new System.Drawing.Size(13, 13);
            this.lblPadColor.TabIndex = 26;
            this.lblPadColor.Text = "  ";
            this.lblPadColor.Click += new System.EventHandler(this.lblPadColor_Click);
            // 
            // lb
            // 
            this.lb.AutoSize = true;
            this.lb.Location = new System.Drawing.Point(105, 68);
            this.lb.Name = "lb";
            this.lb.Size = new System.Drawing.Size(35, 13);
            this.lb.TabIndex = 25;
            this.lb.Text = "Trace";
            // 
            // lblTraceColor
            // 
            this.lblTraceColor.BackColor = System.Drawing.Color.Green;
            this.lblTraceColor.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblTraceColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTraceColor.Location = new System.Drawing.Point(85, 68);
            this.lblTraceColor.Name = "lblTraceColor";
            this.lblTraceColor.Size = new System.Drawing.Size(13, 13);
            this.lblTraceColor.TabIndex = 24;
            this.lblTraceColor.Text = "  ";
            this.lblTraceColor.Click += new System.EventHandler(this.lblTraceColor_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(28, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(33, 13);
            this.label5.TabIndex = 23;
            this.label5.Text = "Mask";
            // 
            // lblMaskColor
            // 
            this.lblMaskColor.BackColor = System.Drawing.Color.Green;
            this.lblMaskColor.Cursor = System.Windows.Forms.Cursors.Hand;
            this.lblMaskColor.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMaskColor.Location = new System.Drawing.Point(8, 68);
            this.lblMaskColor.Name = "lblMaskColor";
            this.lblMaskColor.Size = new System.Drawing.Size(13, 13);
            this.lblMaskColor.TabIndex = 22;
            this.lblMaskColor.Text = "  ";
            this.lblMaskColor.Click += new System.EventHandler(this.lblMaskColor_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(6, 17);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Style Preset";
            // 
            // buttonSaveJSON
            // 
            this.buttonSaveJSON.Location = new System.Drawing.Point(135, 48);
            this.buttonSaveJSON.Name = "buttonSaveJSON";
            this.buttonSaveJSON.Size = new System.Drawing.Size(101, 23);
            this.buttonSaveJSON.TabIndex = 14;
            this.buttonSaveJSON.Text = "Save Board";
            this.buttonSaveJSON.UseVisualStyleBackColor = true;
            this.buttonSaveJSON.Click += new System.EventHandler(this.buttonSaveJSON_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label14);
            this.groupBox3.Controls.Add(this.label13);
            this.groupBox3.Controls.Add(this.numericUpDownIncrement);
            this.groupBox3.Controls.Add(this.buttonZMinus);
            this.groupBox3.Controls.Add(this.buttonZPlus);
            this.groupBox3.Controls.Add(this.buttonYMinus);
            this.groupBox3.Controls.Add(this.buttonYPlus);
            this.groupBox3.Controls.Add(this.buttonXMinus);
            this.groupBox3.Controls.Add(this.buttonXPlus);
            this.groupBox3.Controls.Add(this.label12);
            this.groupBox3.Controls.Add(this.label11);
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.label8);
            this.groupBox3.Controls.Add(this.btnRotate);
            this.groupBox3.Controls.Add(this.numericUpDownAngle);
            this.groupBox3.Controls.Add(this.comboAlongAxis);
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Controls.Add(this.checkRotateSelected);
            this.groupBox3.Controls.Add(this.chkIsSMD);
            this.groupBox3.Controls.Add(this.buttonAutoAlign);
            this.groupBox3.Location = new System.Drawing.Point(0, 523);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(244, 206);
            this.groupBox3.TabIndex = 15;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Transformations";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(154, 73);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(23, 13);
            this.label14.TabIndex = 58;
            this.label14.Text = "mm";
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(6, 72);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(77, 13);
            this.label13.TabIndex = 57;
            this.label13.Text = "Increment step";
            // 
            // numericUpDownIncrement
            // 
            this.numericUpDownIncrement.DecimalPlaces = 3;
            this.numericUpDownIncrement.Increment = new decimal(new int[] {
            5,
            0,
            0,
            65536});
            this.numericUpDownIncrement.Location = new System.Drawing.Point(89, 71);
            this.numericUpDownIncrement.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownIncrement.Name = "numericUpDownIncrement";
            this.numericUpDownIncrement.Size = new System.Drawing.Size(62, 20);
            this.numericUpDownIncrement.TabIndex = 56;
            // 
            // buttonZMinus
            // 
            this.buttonZMinus.Location = new System.Drawing.Point(217, 114);
            this.buttonZMinus.Name = "buttonZMinus";
            this.buttonZMinus.Size = new System.Drawing.Size(23, 23);
            this.buttonZMinus.TabIndex = 55;
            this.buttonZMinus.Text = "-";
            this.buttonZMinus.UseVisualStyleBackColor = true;
            this.buttonZMinus.Click += new System.EventHandler(this.buttonZMinus_Click);
            // 
            // buttonZPlus
            // 
            this.buttonZPlus.Location = new System.Drawing.Point(193, 114);
            this.buttonZPlus.Name = "buttonZPlus";
            this.buttonZPlus.Size = new System.Drawing.Size(23, 23);
            this.buttonZPlus.TabIndex = 54;
            this.buttonZPlus.Text = "+";
            this.buttonZPlus.UseVisualStyleBackColor = true;
            this.buttonZPlus.Click += new System.EventHandler(this.buttonZPlus_Click);
            // 
            // buttonYMinus
            // 
            this.buttonYMinus.Location = new System.Drawing.Point(125, 114);
            this.buttonYMinus.Name = "buttonYMinus";
            this.buttonYMinus.Size = new System.Drawing.Size(23, 23);
            this.buttonYMinus.TabIndex = 53;
            this.buttonYMinus.Text = "-";
            this.buttonYMinus.UseVisualStyleBackColor = true;
            this.buttonYMinus.Click += new System.EventHandler(this.buttonYMinus_Click);
            // 
            // buttonYPlus
            // 
            this.buttonYPlus.Location = new System.Drawing.Point(101, 114);
            this.buttonYPlus.Name = "buttonYPlus";
            this.buttonYPlus.Size = new System.Drawing.Size(23, 23);
            this.buttonYPlus.TabIndex = 52;
            this.buttonYPlus.Text = "+";
            this.buttonYPlus.UseVisualStyleBackColor = true;
            this.buttonYPlus.Click += new System.EventHandler(this.buttonYPlus_Click);
            // 
            // buttonXMinus
            // 
            this.buttonXMinus.Location = new System.Drawing.Point(28, 114);
            this.buttonXMinus.Name = "buttonXMinus";
            this.buttonXMinus.Size = new System.Drawing.Size(23, 23);
            this.buttonXMinus.TabIndex = 51;
            this.buttonXMinus.Text = "-";
            this.buttonXMinus.UseVisualStyleBackColor = true;
            this.buttonXMinus.Click += new System.EventHandler(this.buttonXMinus_Click);
            // 
            // buttonXPlus
            // 
            this.buttonXPlus.Location = new System.Drawing.Point(4, 114);
            this.buttonXPlus.Name = "buttonXPlus";
            this.buttonXPlus.Size = new System.Drawing.Size(23, 23);
            this.buttonXPlus.TabIndex = 50;
            this.buttonXPlus.Text = "+";
            this.buttonXPlus.UseVisualStyleBackColor = true;
            this.buttonXPlus.Click += new System.EventHandler(this.button1_Click);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(210, 97);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(14, 13);
            this.label12.TabIndex = 49;
            this.label12.Text = "Z";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(118, 97);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(14, 13);
            this.label11.TabIndex = 47;
            this.label11.Text = "Y";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(22, 97);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(14, 13);
            this.label10.TabIndex = 45;
            this.label10.Text = "X";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(6, 49);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(91, 13);
            this.label8.TabIndex = 44;
            this.label8.Text = "Move Component";
            // 
            // checkBoxCollisionDetection
            // 
            this.checkBoxCollisionDetection.AutoSize = true;
            this.checkBoxCollisionDetection.Location = new System.Drawing.Point(4, 748);
            this.checkBoxCollisionDetection.Name = "checkBoxCollisionDetection";
            this.checkBoxCollisionDetection.Size = new System.Drawing.Size(150, 17);
            this.checkBoxCollisionDetection.TabIndex = 59;
            this.checkBoxCollisionDetection.Text = "Stop at colision with board";
            this.checkBoxCollisionDetection.UseVisualStyleBackColor = true;
            this.checkBoxCollisionDetection.Visible = false;
            // 
            // buttonSaveLibrary
            // 
            this.buttonSaveLibrary.Location = new System.Drawing.Point(144, 742);
            this.buttonSaveLibrary.Name = "buttonSaveLibrary";
            this.buttonSaveLibrary.Size = new System.Drawing.Size(101, 23);
            this.buttonSaveLibrary.TabIndex = 16;
            this.buttonSaveLibrary.Text = "test";
            this.buttonSaveLibrary.UseVisualStyleBackColor = true;
            this.buttonSaveLibrary.Visible = false;
            this.buttonSaveLibrary.Click += new System.EventHandler(this.buttonSaveLibrary_Click);
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.buttonDecalImages);
            this.groupBox4.Controls.Add(this.buttonGerberFiles);
            this.groupBox4.Controls.Add(this.buttonBrowseJson);
            this.groupBox4.Controls.Add(this.buttonSaveJSON);
            this.groupBox4.Location = new System.Drawing.Point(0, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(244, 79);
            this.groupBox4.TabIndex = 60;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Board Settings";
            // 
            // buttonDecalImages
            // 
            this.buttonDecalImages.Location = new System.Drawing.Point(135, 19);
            this.buttonDecalImages.Name = "buttonDecalImages";
            this.buttonDecalImages.Size = new System.Drawing.Size(101, 23);
            this.buttonDecalImages.TabIndex = 64;
            this.buttonDecalImages.Text = "Open Decals";
            this.buttonDecalImages.UseVisualStyleBackColor = true;
            this.buttonDecalImages.Click += new System.EventHandler(this.buttonDecalImages_Click);
            // 
            // buttonGerberFiles
            // 
            this.buttonGerberFiles.Location = new System.Drawing.Point(6, 48);
            this.buttonGerberFiles.Name = "buttonGerberFiles";
            this.buttonGerberFiles.Size = new System.Drawing.Size(101, 23);
            this.buttonGerberFiles.TabIndex = 63;
            this.buttonGerberFiles.Text = "Open Gerber";
            this.buttonGerberFiles.UseVisualStyleBackColor = true;
            this.buttonGerberFiles.Click += new System.EventHandler(this.buttonGerberFiles_Click);
            // 
            // buttonChangeHeight
            // 
            this.buttonChangeHeight.Location = new System.Drawing.Point(135, 16);
            this.buttonChangeHeight.Name = "buttonChangeHeight";
            this.buttonChangeHeight.Size = new System.Drawing.Size(101, 23);
            this.buttonChangeHeight.TabIndex = 62;
            this.buttonChangeHeight.Text = "Apply Change";
            this.buttonChangeHeight.UseVisualStyleBackColor = true;
            this.buttonChangeHeight.Click += new System.EventHandler(this.buttonChangeHeight_Click);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(80, 23);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(29, 13);
            this.label15.TabIndex = 61;
            this.label15.Text = "mm  ";
            // 
            // numericUpDownBoardHeight
            // 
            this.numericUpDownBoardHeight.DecimalPlaces = 3;
            this.numericUpDownBoardHeight.Increment = new decimal(new int[] {
            5,
            0,
            0,
            65536});
            this.numericUpDownBoardHeight.Location = new System.Drawing.Point(12, 19);
            this.numericUpDownBoardHeight.Maximum = new decimal(new int[] {
            10,
            0,
            0,
            0});
            this.numericUpDownBoardHeight.Name = "numericUpDownBoardHeight";
            this.numericUpDownBoardHeight.Size = new System.Drawing.Size(56, 20);
            this.numericUpDownBoardHeight.TabIndex = 59;
            this.numericUpDownBoardHeight.ValueChanged += new System.EventHandler(this.numericUpDownBoardHeight_ValueChanged);
            // 
            // groupBox5
            // 
            this.groupBox5.Controls.Add(this.buttonChangeHeight);
            this.groupBox5.Controls.Add(this.numericUpDownBoardHeight);
            this.groupBox5.Controls.Add(this.label15);
            this.groupBox5.Location = new System.Drawing.Point(0, 88);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new System.Drawing.Size(244, 54);
            this.groupBox5.TabIndex = 61;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Board Thickness";
            // 
            // TaskpaneHostUI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.groupBox5);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.checkBoxCollisionDetection);
            this.Controls.Add(this.buttonSaveLibrary);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Name = "TaskpaneHostUI";
            this.Size = new System.Drawing.Size(247, 782);
            this.Load += new System.EventHandler(this.TaskpaneHostUI_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownAngle)).EndInit();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownIncrement)).EndInit();
            this.groupBox4.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDownBoardHeight)).EndInit();
            this.groupBox5.ResumeLayout(false);
            this.groupBox5.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button buttonBrowseJson;
        private System.Windows.Forms.ComboBox colorPresets;
        private System.Windows.Forms.Button buttonApplyStyle;
        private System.Windows.Forms.ComboBox comboComponentList;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button buttonClearModel;
        private System.Windows.Forms.Button buttonBrowseModel;
        private System.Windows.Forms.Label labelInstances;
        private System.Windows.Forms.Label label31;
        private System.Windows.Forms.Label labelAssignedFile;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label lb;
        private System.Windows.Forms.Label lblTraceColor;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblMaskColor;
        private System.Windows.Forms.Label lbl2;
        private System.Windows.Forms.Label lblSilkColor;
        private System.Windows.Forms.Label lbl3;
        private System.Windows.Forms.Label lblPadColor;
        private System.Windows.Forms.Button buttonSaveJSON;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox comboAlongAxis;
        private System.Windows.Forms.NumericUpDown numericUpDownAngle;
        private System.Windows.Forms.Button btnRotate;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblHolesColor;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label lblBoardColor;
        private System.Windows.Forms.CheckBox checkRotateSelected;
        private System.Windows.Forms.Button buttonAutoAlign;
        private System.Windows.Forms.CheckBox chkIsSMD;
        private System.Windows.Forms.CheckBox checkBoxBottomOnly;
        private System.Windows.Forms.CheckBox checkBoxTopOnly;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.NumericUpDown numericUpDownIncrement;
        private System.Windows.Forms.Button buttonZMinus;
        private System.Windows.Forms.Button buttonZPlus;
        private System.Windows.Forms.Button buttonYMinus;
        private System.Windows.Forms.Button buttonYPlus;
        private System.Windows.Forms.Button buttonXMinus;
        private System.Windows.Forms.Button buttonXPlus;
        private System.Windows.Forms.CheckBox checkBoxCollisionDetection;
        private System.Windows.Forms.Button buttonSaveLibrary;
        private System.Windows.Forms.Button buttonLoadLibrary;
        private System.Windows.Forms.GroupBox groupBox4;
        private System.Windows.Forms.CheckBox checkLoadSelected;
        private System.Windows.Forms.NumericUpDown numericUpDownBoardHeight;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Button buttonChangeHeight;
        private System.Windows.Forms.Button buttonDecalImages;
        private System.Windows.Forms.Button buttonGerberFiles;
        private System.Windows.Forms.GroupBox groupBox5;
        //private CCombobox cb;

    }
}
