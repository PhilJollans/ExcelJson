namespace ExcelJson
{
  partial class ExcelJsonRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
  {
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    public ExcelJsonRibbon()
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
      this.JsonGroup = this.Factory.CreateRibbonGroup();
      this.AngularGroup = this.Factory.CreateRibbonGroup();
      this.ReadButton = this.Factory.CreateRibbonButton();
      this.WriteButton = this.Factory.CreateRibbonButton();
      this.ReadAngularI18nFiles = this.Factory.CreateRibbonButton();
      this.WriteAngularI18nFiles = this.Factory.CreateRibbonButton();
      this.tab1.SuspendLayout();
      this.JsonGroup.SuspendLayout();
      this.AngularGroup.SuspendLayout();
      this.SuspendLayout();
      // 
      // tab1
      // 
      this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
      this.tab1.Groups.Add(this.JsonGroup);
      this.tab1.Groups.Add(this.AngularGroup);
      this.tab1.Label = "TabAddIns";
      this.tab1.Name = "tab1";
      // 
      // JsonGroup
      // 
      this.JsonGroup.Items.Add(this.ReadButton);
      this.JsonGroup.Items.Add(this.WriteButton);
      this.JsonGroup.Label = "json";
      this.JsonGroup.Name = "JsonGroup";
      // 
      // AngularGroup
      // 
      this.AngularGroup.Items.Add(this.ReadAngularI18nFiles);
      this.AngularGroup.Items.Add(this.WriteAngularI18nFiles);
      this.AngularGroup.Label = "Angular";
      this.AngularGroup.Name = "AngularGroup";
      // 
      // ReadButton
      // 
      this.ReadButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
      this.ReadButton.Image = global::ExcelJson.Properties.Resources.json;
      this.ReadButton.Label = "Read Json Array";
      this.ReadButton.Name = "ReadButton";
      this.ReadButton.ShowImage = true;
      this.ReadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReadButton_Click);
      // 
      // WriteButton
      // 
      this.WriteButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
      this.WriteButton.Image = global::ExcelJson.Properties.Resources.json;
      this.WriteButton.Label = "Write Json Array";
      this.WriteButton.Name = "WriteButton";
      this.WriteButton.ShowImage = true;
      this.WriteButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WriteButton_Click);
      // 
      // ReadAngularI18nFiles
      // 
      this.ReadAngularI18nFiles.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
      this.ReadAngularI18nFiles.Image = global::ExcelJson.Properties.Resources.angular_black;
      this.ReadAngularI18nFiles.Label = "Read i18n";
      this.ReadAngularI18nFiles.Name = "ReadAngularI18nFiles";
      this.ReadAngularI18nFiles.ShowImage = true;
      this.ReadAngularI18nFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReadAngularI18nFiles_Click);
      // 
      // WriteAngularI18nFiles
      // 
      this.WriteAngularI18nFiles.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
      this.WriteAngularI18nFiles.Image = global::ExcelJson.Properties.Resources.angular_black;
      this.WriteAngularI18nFiles.Label = "Write i18n";
      this.WriteAngularI18nFiles.Name = "WriteAngularI18nFiles";
      this.WriteAngularI18nFiles.ShowImage = true;
      this.WriteAngularI18nFiles.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WriteAngularI18nFiles_Click);
      // 
      // ExcelJsonRibbon
      // 
      this.Name = "ExcelJsonRibbon";
      this.RibbonType = "Microsoft.Excel.Workbook";
      this.Tabs.Add(this.tab1);
      this.tab1.ResumeLayout(false);
      this.tab1.PerformLayout();
      this.JsonGroup.ResumeLayout(false);
      this.JsonGroup.PerformLayout();
      this.AngularGroup.ResumeLayout(false);
      this.AngularGroup.PerformLayout();
      this.ResumeLayout(false);

    }

    #endregion

    internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
    internal Microsoft.Office.Tools.Ribbon.RibbonGroup JsonGroup;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton ReadButton;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton WriteButton;
    internal Microsoft.Office.Tools.Ribbon.RibbonGroup AngularGroup;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton ReadAngularI18nFiles;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton WriteAngularI18nFiles;
  }

  partial class ThisRibbonCollection
  {
    internal ExcelJsonRibbon ExcelJsonRibbon
    {
      get { return this.GetRibbon<ExcelJsonRibbon>(); }
    }
  }
}
