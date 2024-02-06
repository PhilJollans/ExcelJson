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
      this.ReadButton = this.Factory.CreateRibbonButton();
      this.WriteButton = this.Factory.CreateRibbonButton();
      this.tab1.SuspendLayout();
      this.JsonGroup.SuspendLayout();
      this.SuspendLayout();
      // 
      // tab1
      // 
      this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
      this.tab1.Groups.Add(this.JsonGroup);
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
      // ReadButton
      // 
      this.ReadButton.Image = global::ExcelJson.Properties.Resources.json;
      this.ReadButton.Label = "Read";
      this.ReadButton.Name = "ReadButton";
      this.ReadButton.ShowImage = true;
      this.ReadButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ReadButton_Click);
      // 
      // WriteButton
      // 
      this.WriteButton.Image = global::ExcelJson.Properties.Resources.json;
      this.WriteButton.Label = "Write";
      this.WriteButton.Name = "WriteButton";
      this.WriteButton.ShowImage = true;
      this.WriteButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.WriteButton_Click);
      // 
      // ExcelJsonRibbon
      // 
      this.Name = "ExcelJsonRibbon";
      this.RibbonType = "Microsoft.Excel.Workbook";
      this.Tabs.Add(this.tab1);
      this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExcelJsonRibbon_Load);
      this.tab1.ResumeLayout(false);
      this.tab1.PerformLayout();
      this.JsonGroup.ResumeLayout(false);
      this.JsonGroup.PerformLayout();
      this.ResumeLayout(false);

    }

    #endregion

    internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
    internal Microsoft.Office.Tools.Ribbon.RibbonGroup JsonGroup;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton ReadButton;
    internal Microsoft.Office.Tools.Ribbon.RibbonButton WriteButton;
  }

  partial class ThisRibbonCollection
  {
    internal ExcelJsonRibbon ExcelJsonRibbon
    {
      get { return this.GetRibbon<ExcelJsonRibbon>(); }
    }
  }
}
