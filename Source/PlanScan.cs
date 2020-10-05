using System;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using System.Linq;

using VMS.TPS.Common.Model.API;

namespace VMS.TPS
{
  class Script
  {
    public Script()
    {
    }

    public void Execute(ScriptContext context, Window window)
    {
      // Check for patient loaded
      if(context.Patient == null)
      {
        MessageBox.Show("Error: No patient selected!");
        return;
      }
      // Check for patient plan loaded
      PlanSetup plan = context.PlanSetup;
      if(plan == null)
      {
        MessageBox.Show("Error: No plan loaded!");
        return;
      }

      // Grab and evaluate plan info
      Grid planInfoGrid = MakeInfoGrid(context, plan);
      Grid generalCheckGrid = MakeGeneralCheckGrid(context, plan);
      Grid densityOverrideGrid = MakeDensityOverrideGrid(context, plan);
      Grid beamGrid = MakeBeamGrid(context, plan);
      Grid miscellaneousGrid = MakeMiscGrid(context, plan);

      // Initialize window
      window.Title = "PlanScan";
      window.Height = 900;
      window.Width = 700;
      ScrollViewer mainView = new ScrollViewer();
      StackPanel rootPanel = new StackPanel();
      //rootPanel.Orientation = Orientation.Vertical;
      rootPanel.Children.Add(planInfoGrid);
      rootPanel.Children.Add(generalCheckGrid);
      rootPanel.Children.Add(densityOverrideGrid);
      rootPanel.Children.Add(beamGrid);
      rootPanel.Children.Add(miscellaneousGrid);
      mainView.Content = rootPanel;
      window.Content = mainView;
    }

    public Grid MakeInfoGrid(ScriptContext context, PlanSetup plan)
    {
      string[] infoText = new string[10] {
        "Patient (ID)", context.Patient.LastName+", "+context.Patient.FirstName+" ("+context.Patient.Id+")",
        "Course", plan.Course.Id,
        "Plan", plan.Id,
        "Prescription Dose (%)", plan.TotalPrescribedDose.ValueAsString+' '+plan.TotalPrescribedDose.UnitAsString+" ("+((100.0*plan.PrescribedPercentage).ToString())+" %)",
        "Fractionation", plan.UniqueFractionation.PrescribedDosePerFraction+" x "+plan.UniqueFractionation.NumberOfFractions+" fractions"
      };

      // Grid for basic patient information
      Grid infoGrid = new Grid();
      //infoGrid.HorizontalAlignment = HorizontalAlignment.Center;
      infoGrid.VerticalAlignment = VerticalAlignment.Top;
      infoGrid.Margin = new Thickness(16.0);
      //infoGrid.ShowGridLines = true;
      double[] ig_widths = new double[2] {200.0, 400.0};
      for(int n = 0; n < 2; n++)
      {
        ColumnDefinition cDef = new ColumnDefinition();
        cDef.Width = new GridLength(ig_widths[n]);
        infoGrid.ColumnDefinitions.Add(cDef);
      }
      for(int n = 0; n < (infoText.Length/2); n++)
      {
        RowDefinition rDef = new RowDefinition();
        rDef.Height = new GridLength(20.0);
        infoGrid.RowDefinitions.Add(rDef);
      }
      for(int n = 0; n < infoText.Length; n++)
      {
        TextBlock txtblk = new TextBlock();
        txtblk.Text = infoText[n];
        txtblk.VerticalAlignment = VerticalAlignment.Center;
        txtblk.FontSize = 14;
        txtblk.FontWeight = FontWeights.Bold;
        Grid.SetColumn(txtblk, n%2);
        Grid.SetRow(txtblk, n/2);
        infoGrid.Children.Add(txtblk);
      }

      return infoGrid;
    }

    public Grid MakeGeneralCheckGrid(ScriptContext context, PlanSetup plan)
    {
      // Make list of checks
      List<string> planScanText = new List<string>();
      List<bool> PassFail = new List<bool>();
      // Column Headers
      planScanText.Add("Parameter");
      planScanText.Add("Expected Value");
      planScanText.Add("Plan Value");
      /////////////////
      // Course Info //
      /////////////////
      // Course has diagnosis attached?
      planScanText.Add("Course: Attached Diagnoses");
      planScanText.Add("> 0");
      planScanText.Add(plan.Course.Diagnoses.Count().ToString());
      if(plan.Course.Diagnoses.Count() > 0) PassFail.Add(true);
      else PassFail.Add(false);
      ///////////////
      // Plan Info //
      ///////////////
      // Plan intent versus course
      planScanText.Add("Plan: Intent (vs. Course)");
      planScanText.Add(plan.Course.Intent);
      planScanText.Add(plan.PlanIntent);
      if(plan.PlanIntent.Equals(plan.Course.Intent, StringComparison.OrdinalIgnoreCase)) PassFail.Add(true);
      else PassFail.Add(false);
      // Plan Approval Status (PlanningApproved)
      planScanText.Add("Plan: Approval Status");
      planScanText.Add("PlanningApproved");
      planScanText.Add(plan.ApprovalStatus.ToString());
      if(plan.ApprovalStatus.ToString() == "PlanningApproved") PassFail.Add(true);
      else PassFail.Add(false);
      // Plan-Imaging Orientation
      planScanText.Add("Plan: Orientation (vs. Image)");
      planScanText.Add(plan.StructureSet.Image.ImagingOrientation.ToString());
      planScanText.Add(plan.TreatmentOrientation.ToString());
      if(plan.StructureSet.Image.ImagingOrientation == plan.TreatmentOrientation) PassFail.Add(true);
      else PassFail.Add(false);
      ////////////////////
      // Structure Info //
      ////////////////////
      // Check for couch model
      var couchModel = new bool[2]{false, false};
      var structures = plan.StructureSet.Structures;
      foreach(Structure s in structures)
      {
        if(s.Id.Contains("CouchSurface")) couchModel[0] = true;
        if(s.Id.Contains("CouchInterior")) couchModel[1] = true;
      }
      planScanText.Add("Structures: Couch Model");
      planScanText.Add("Yes");
      if(couchModel[0] && couchModel[1])
      {
        planScanText.Add("Yes");
        PassFail.Add(true);
      }
      else
      {
        planScanText.Add("No");
        PassFail.Add(false);
      }
      ///////////////
      // Dose Info //
      ///////////////
      // Dose Calculation Models
      planScanText.Add("Dose: Calculation Model (Photon)");
      planScanText.Add("AAA_13623");
      planScanText.Add(plan.PhotonCalculationModel);
      if(plan.PhotonCalculationModel == "AAA_13623") PassFail.Add(true);
      else PassFail.Add(false);
      // Dose calculation options
      var calcOptions = plan.PhotonCalculationOptions;
      var calcLabels = new List<string>{"Calculation Grid (cm)",
        "Field Normalization", "Heterogeneity Correction"};
      var preferredOptions = new List<string>{"0.25", "100% to isocenter", "ON"};
      int op_id = 0;
      foreach(KeyValuePair<string, string> kvp in calcOptions)
      {
        planScanText.Add("Dose: "+calcLabels[op_id]);
        planScanText.Add(kvp.Value);
        planScanText.Add(preferredOptions[op_id]);
        if(kvp.Value == preferredOptions[op_id]) PassFail.Add(true);
        else PassFail.Add(false);
        op_id++;
      }
      // Valid dose
      planScanText.Add("Dose: Valid Dose");
      planScanText.Add("Yes");
      if(plan.IsDoseValid)
      {
        planScanText.Add("Yes");
        PassFail.Add(true);
      }
      else
      {
        planScanText.Add("No");
        PassFail.Add(false);
      }
      // Maximum dose
      planScanText.Add("Dose: Maximum Dose");
      planScanText.Add("110 %");
      double max_dose = plan.Dose.DoseMax3D.Dose;
      planScanText.Add(string.Format("{0:0.0}",max_dose)+" %");
      if(max_dose <= 110.0) PassFail.Add(true);
      else PassFail.Add(false);
      /////////////////////////////
      // Make grid to store text //
      /////////////////////////////
      Grid mGrid = new Grid();
      //mGrid.ShowGridLines = true;
      //mGrid.HorizontalAlignment = HorizontalAlignment.Center;
      mGrid.VerticalAlignment = VerticalAlignment.Top;
      mGrid.Margin = new Thickness(16.0);
      double[] mg_widths = new double[3] {200.0, 200.0, 240.0};
      for(int n = 0; n < 3; n++)
      {
        ColumnDefinition cDef = new ColumnDefinition();
        cDef.Width = new GridLength(mg_widths[n]);
        mGrid.ColumnDefinitions.Add(cDef);
      }
      for(int n = 0; n < planScanText.Count/3 + 1; n++)
      {
        RowDefinition rDef = new RowDefinition();
        if(n==0) rDef.Height = new GridLength(36.0);
        else rDef.Height = new GridLength(20.0);
        mGrid.RowDefinitions.Add(rDef);
      }
      // Grid Title
      TextBlock title = new TextBlock();
      title.Text = "General Checks";
      title.VerticalAlignment = VerticalAlignment.Center;
      //title.HorizontalAlignment = HorizontalAlignment.Center;
      title.Height = 36.0;
      title.FontSize = 16;
      title.FontWeight = FontWeights.Bold;
      Grid.SetColumn(title, 0);
      Grid.SetRow(title, 0);
      Grid.SetColumnSpan(title, 3);
      mGrid.Children.Add(title);
      // Grid Table elements
      for(int n = 0; n < planScanText.Count; n++)
      {
        TextBlock txtblk = new TextBlock();
        txtblk.Text = planScanText[n];
        txtblk.VerticalAlignment = VerticalAlignment.Center;
        //txtblk.FontSize = 14;
        if(n < 3) txtblk.FontWeight = FontWeights.Bold;
        if(n >= 3 && (n%3==2))
        {
          if(PassFail[((n/3)-1)]) txtblk.Foreground = System.Windows.Media.Brushes.Green;
          else txtblk.Foreground = System.Windows.Media.Brushes.Orange;
        }
        Grid.SetRow(txtblk, n/3 + 1);
        Grid.SetColumn(txtblk, n%3);
        mGrid.Children.Add(txtblk);
      }

      return mGrid;
    }

    public Grid MakeDensityOverrideGrid(ScriptContext context, PlanSetup plan)
    {
      // Find all boluses
      int n_beams = 0;
      List<string> bolusList = new List<string>();
      foreach(Beam b in plan.Beams)
      {
        if(!b.IsSetupField) n_beams++;
        if(b.Boluses.Count() > 0)
        {
          foreach(Bolus bol in b.Boluses)
          {
            bolusList.Add(bol.Id);
          }
        }
      }
      List<string> bolusUnique = new List<string>();
      List<string> bolusNote = new List<string>();
      var g = bolusList.GroupBy(i => i);
      foreach(var grp in g)
      {
        bolusUnique.Add(grp.Key);
        bolusNote.Add(String.Format("Bolus attached to {0}/{0} beams", grp.Count(), n_beams));
      }
      // Load density override information and check bolus attachment
      double huValue;
      List<string> overrideText = new List<string>();
      overrideText.Add("Structure");
      overrideText.Add("Assigned HU");
      overrideText.Add("Bolus Note");
      foreach(Structure s in plan.StructureSet.Structures)
      {
        if(s.GetAssignedHU(out huValue))
        {
          overrideText.Add(s.Id);
          overrideText.Add(huValue.ToString("0"));
          if(bolusUnique.Contains(s.Id))
          {
            overrideText.Add(bolusNote[bolusUnique.FindIndex(x => x == s.Id)]);
          }
          else overrideText.Add("NA");
        }
      }

      /////////////////////////////
      // Make grid to store text //
      /////////////////////////////
      Grid dGrid = new Grid();
      //mGrid.ShowGridLines = true;
      //mGrid.HorizontalAlignment = HorizontalAlignment.Center;
      dGrid.VerticalAlignment = VerticalAlignment.Top;
      dGrid.Margin = new Thickness(16.0);
      double[] dg_widths = new double[3] {150.0, 100.0, 200.0};
      for(int n = 0; n < 3; n++)
      {
        ColumnDefinition cDef = new ColumnDefinition();
        cDef.Width = new GridLength(dg_widths[n]);
        dGrid.ColumnDefinitions.Add(cDef);
      }
      for(int n = 0; n < overrideText.Count/3 + 1; n++)
      {
        RowDefinition rDef = new RowDefinition();
        if(n==0) rDef.Height = new GridLength(36.0);
        else rDef.Height = new GridLength(20.0);
        dGrid.RowDefinitions.Add(rDef);
      }
      // Grid Title
      TextBlock title = new TextBlock();
      title.Text = "Density Overrides & Bolus";
      title.VerticalAlignment = VerticalAlignment.Center;
      //title.HorizontalAlignment = HorizontalAlignment.Center;
      title.Height = 36.0;
      title.FontSize = 16;
      title.FontWeight = FontWeights.Bold;
      Grid.SetColumn(title, 0);
      Grid.SetRow(title, 0);
      Grid.SetColumnSpan(title, 3);
      dGrid.Children.Add(title);
      // Grid Table elements
      for(int n = 0; n < overrideText.Count; n++)
      {
        TextBlock txtblk = new TextBlock();
        txtblk.Text = overrideText[n];
        txtblk.VerticalAlignment = VerticalAlignment.Center;
        //txtblk.FontSize = 14;
        if(n < 3) txtblk.FontWeight = FontWeights.Bold;
        Grid.SetRow(txtblk, n/3 + 1);
        Grid.SetColumn(txtblk, n%3);
        dGrid.Children.Add(txtblk);
      }

      return dGrid;
    }

    public Grid MakeBeamGrid(ScriptContext context, PlanSetup plan)
    {
      // Load with beam data
      List<string> beamText = new List<string>();
      foreach(Beam b in plan.Beams)
      {
        string beamName = b.Id;
        beamText.Add(beamName+": Beam MU/cGy");
        //beamText.Add(b.Meterset.Value.ToString()+' '+b.Meterset.Unit);
        //beamText.Add(b.Technique.Id);
        //beamText.Add(b.GantryDirection.ToString());
        if(!b.IsSetupField) beamText.Add((b.MetersetPerGy/100.0).ToString("0.00"));
        else beamText.Add("NA");
        //beamText.Add(b.ControlPoints.Count.ToString());
        //beamText.Add(b.EnergyModeDisplayName);
        //beamText.Add((b.AverageSSD/10.0).ToString());
        //beamText.Add(b.DoseRate.ToString());
        //beamText.Add(b.IsSetupField.ToString());
        //beamText.Add(b.MLCPlanType.ToString());
        //beamText.Add(b.MLCTransmissionFactor.ToString());
        //beamText.Add((b.PlannedSSD/10.0).ToString());
        //beamText.Add((b.SSD/10.0).ToString());
        //beamText.Add(b.Technique.Id);
        //beamText.Add(b.ToleranceTableLabel);
        //beamText.Add(b.TreatmentUnit.Id);
        //beamText.Add(b.ArcLength.ToString());
      }
      // Grid for basic patient information
      Grid bGrid = new Grid();
      //bGrid.HorizontalAlignment = HorizontalAlignment.Center;
      bGrid.VerticalAlignment = VerticalAlignment.Top;
      bGrid.Margin = new Thickness(16.0);
      //bGrid.ShowGridLines = true;
      // Add list items
      double[] bg_widths = new double[2] {300.0, 300.0};
      for(int n = 0; n < 2; n++)
      {
        ColumnDefinition cDef = new ColumnDefinition();
        cDef.Width = new GridLength(bg_widths[n]);
        bGrid.ColumnDefinitions.Add(cDef);
      }
      for(int n = 0; n < (beamText.Count/2)+1; n++)
      {
        RowDefinition rDef = new RowDefinition();
        if(n==0) rDef.Height = new GridLength(36.0);
        else rDef.Height = new GridLength(20.0);
        bGrid.RowDefinitions.Add(rDef);
      }
      // Grid Title
      TextBlock title = new TextBlock();
      title.Text = "Beam Checks";
      title.VerticalAlignment = VerticalAlignment.Center;
      //title.HorizontalAlignment = HorizontalAlignment.Center;
      title.Height = 36.0;
      title.FontSize = 16;
      title.FontWeight = FontWeights.Bold;
      Grid.SetColumn(title, 0);
      Grid.SetRow(title, 0);
      Grid.SetColumnSpan(title, 3);
      bGrid.Children.Add(title);
      for(int n = 0; n < beamText.Count; n++)
      {
        TextBlock txtblk = new TextBlock();
        txtblk.Text = beamText[n];
        txtblk.VerticalAlignment = VerticalAlignment.Center;
        //txtblk.FontSize = 14;
        //txtblk.FontWeight = FontWeights.Bold;
        Grid.SetColumn(txtblk, n%2);
        Grid.SetRow(txtblk, n/2 + 1);
        bGrid.Children.Add(txtblk);
      }

      return bGrid;
    }

    public Grid MakeMiscGrid(ScriptContext context, PlanSetup plan)
    {
      string[] miscText = new string[22] {
        "Creation Date (User)", plan.CreationDateTime+" ("+plan.CreationUserName+")",
        "Dose Value Presentation", plan.DoseValuePresentation.ToString(),
        "Dose Calculation Model (Electron)", plan.ElectronCalculationModel,
        "Last Edit Date (User)", plan.HistoryDateTime+" ("+plan.HistoryUserName+")",
        "Plan Treated", plan.IsTreated.ToString(),
        "Planning Approval Date (User)", plan.PlanningApprovalDate+" ("+plan.PlanningApprover+")",
        "Plan Normalization Method", plan.PlanNormalizationMethod,
        "Plan Normalization Value", plan.PlanNormalizationValue.ToString(),
        "Plan Type", plan.PlanType.ToString(),
        "Target Volume ID", plan.TargetVolumeID,
        "Treatment Approval Date (User)", plan.TreatmentApprovalDate+" ("+plan.TreatmentApprover+")",
      };

      // Grid for basic patient information
      Grid miscGrid = new Grid();
      //miscGrid.HorizontalAlignment = HorizontalAlignment.Center;
      miscGrid.VerticalAlignment = VerticalAlignment.Top;
      miscGrid.Margin = new Thickness(16.0);
      //miscGrid.ShowGridLines = true;
      double[] mg_widths = new double[2] {300.0, 300.0};
      for(int n = 0; n < 2; n++)
      {
        ColumnDefinition cDef = new ColumnDefinition();
        cDef.Width = new GridLength(mg_widths[n]);
        miscGrid.ColumnDefinitions.Add(cDef);
      }
      for(int n = 0; n < (miscText.Length/2)+1; n++)
      {
        RowDefinition rDef = new RowDefinition();
        if(n==0) rDef.Height = new GridLength(36.0);
        else rDef.Height = new GridLength(20.0);
        miscGrid.RowDefinitions.Add(rDef);
      }
      // Grid Title
      TextBlock title = new TextBlock();
      title.Text = "Miscellaneous Information";
      title.VerticalAlignment = VerticalAlignment.Center;
      //title.HorizontalAlignment = HorizontalAlignment.Center;
      title.Height = 36.0;
      title.FontSize = 16;
      title.FontWeight = FontWeights.Bold;
      Grid.SetColumn(title, 0);
      Grid.SetRow(title, 0);
      Grid.SetColumnSpan(title, 2);
      miscGrid.Children.Add(title);
      // Add list items
      for(int n = 0; n < miscText.Length; n++)
      {
        TextBlock txtblk = new TextBlock();
        txtblk.Text = miscText[n];
        txtblk.VerticalAlignment = VerticalAlignment.Center;
        //txtblk.FontSize = 14;
        //txtblk.FontWeight = FontWeights.Bold;
        Grid.SetColumn(txtblk, n%2);
        Grid.SetRow(txtblk, n/2 + 1);
        miscGrid.Children.Add(txtblk);
      }

      return miscGrid;
    }
  }
}
