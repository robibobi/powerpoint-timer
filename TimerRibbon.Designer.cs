namespace PowerPointTimer
{
    partial class TimerRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TimerRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.TimerGroup = this.Factory.CreateRibbonGroup();
            this.AddTimerButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.TimerGroup.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.TimerGroup);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // TimerGroup
            // 
            this.TimerGroup.Items.Add(this.AddTimerButton);
            this.TimerGroup.Label = "Timers";
            this.TimerGroup.Name = "TimerGroup";
            // 
            // AddTimerButton
            // 
            this.AddTimerButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.AddTimerButton.Label = "Add a Timer";
            this.AddTimerButton.Name = "AddTimerButton";
            this.AddTimerButton.ShowImage = true;
            this.AddTimerButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.AddTimerButton_Click);
            // 
            // TimerRibbon
            // 
            this.Name = "TimerRibbon";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TimerRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.TimerGroup.ResumeLayout(false);
            this.TimerGroup.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup TimerGroup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton AddTimerButton;
    }

    partial class ThisRibbonCollection
    {
        internal TimerRibbon TimerRibbon
        {
            get { return this.GetRibbon<TimerRibbon>(); }
        }
    }
}
