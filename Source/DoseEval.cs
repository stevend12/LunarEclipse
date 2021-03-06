/******************************************************************************/
//
// DoseEval.cs
// Eclipse Plan Dose Evaluation Script
// Steven Dolly
// Created: August 12, 2019
//
// This script evaluates the dose volume histograms from the current plan or
// plan sum against a user-defined constraint list.
//
/******************************************************************************/

using System;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Data;
using System.IO;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;
using System.Windows.Controls;

namespace VMS.TPS
{
  public class Script
  {
    public Script()
    {
    }

    PlanningItem SelectedPlanningItem { get; set; }
    StructureSet SelectedStructureSet { get; set; }

    Button btn1 = new Button();
    Button btn2 = new Button();
    DataGrid doseDataGrid = new DataGrid();
    string missingText;
    TextBlock scriptNotes = new TextBlock();
    List<string> constraintFiles = new List<string>();
    List<string> groupList = new List<string>();
    List< List<string> > protocolList = new List< List<string> >();
    List<string> planList = new List<string>();

    ComboBox groupSelectorMenu = new ComboBox();
    ComboBox protocolSelectorMenu = new ComboBox();
    ComboBox planSelectorMenu = new ComboBox();

    string mainFolder = System.IO.Path.GetDirectoryName(new System.Uri(System.Reflection.Assembly.GetExecutingAssembly().CodeBase).LocalPath);

    public class PassFailBrushConverter : IValueConverter
    {
      public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
      {
        string input = value as string;
        switch(input)
        {
          case "Pass":
            return System.Windows.Media.Brushes.Green;
          case "Fail":
            return System.Windows.Media.Brushes.Red;
          default:
            return System.Windows.Media.Brushes.Black;
        }
      }
      public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
      {
        throw new NotSupportedException();
      }
    }

    public void Execute(ScriptContext context, Window window)
    {
      ////////////////////////////////////////////////////////////////
      // 1. Check for valid plan or plan sum, then make list of all //
      ////////////////////////////////////////////////////////////////
      PlanSetup plan = context.PlanSetup;
      PlanSum psum = context.PlanSumsInScope.FirstOrDefault();
      if(plan == null && psum == null)
      {
        MessageBox.Show("Error: No plan or plan sum loaded");
        return;
      }
      // Make list of all plans and plan sums
      foreach(PlanSetup p in context.PlansInScope)
      {
        planList.Add(p.Id);
      }
      foreach(PlanSum s in context.PlanSumsInScope)
      {
        planList.Add(s.Id);
      }

      ///////////////////////////////////////////////////////////////////
      // 2. Load constraints from CSV file and compare to current plan //
      ///////////////////////////////////////////////////////////////////

      // 2A. Have user choose constraint CSV file, then load file
      string constraintFolder = mainFolder+"\\Constraint Data";
      if(!Directory.Exists(constraintFolder)) constraintFolder = "Constraint Data";
      constraintFiles.AddRange(Directory.GetFiles(constraintFolder,"*.csv").Select(Path.GetFileName));

      for(int n = 0; n < constraintFiles.Count(); n++)
      {
        string[] words = constraintFiles[n].Split('_');
        if(!groupList.Contains(words[0]))
        {
          groupList.Add(words[0]);
        }
      }

      for(int g = 0; g < groupList.Count(); g++)
      {
        List<string> subList = new List<string>();
        for(int n = 0; n < constraintFiles.Count(); n++)
        {
          string[] words = constraintFiles[n].Split('_');
          if(groupList[g] == (words[0]))
          {
            subList.Add(words[1].Remove(words[1].Count()-4));
          }
        }
        protocolList.Add(subList);
      }

      // 2B. For each structure in constraint file, find mathing plan structure
      // and compare constraints; return results as data grid
      doseDataGrid.AutoGeneratingColumn += DDG_AutoGeneratingColumn;
      string missingText = "Choose a constraint";
      doseDataGrid.IsReadOnly = true;
      doseDataGrid.CanUserAddRows = false;
      // Display
      doseDataGrid.MaxHeight = 700;
      doseDataGrid.Width = 880;
      doseDataGrid.AlternationCount = 2;
      doseDataGrid.RowBackground = System.Windows.Media.Brushes.Ivory;
      doseDataGrid.AlternatingRowBackground = System.Windows.Media.Brushes.PowderBlue;

      scriptNotes.TextWrapping = TextWrapping.Wrap;
      scriptNotes.Padding = new Thickness(20, 10, 10, 10);
      scriptNotes.Text = missingText;

      var allSelector = new Grid();
      ColumnDefinition colDef1 = new ColumnDefinition();
      ColumnDefinition colDef2 = new ColumnDefinition();
      ColumnDefinition colDef3 = new ColumnDefinition();
      allSelector.ColumnDefinitions.Add(colDef1);
      allSelector.ColumnDefinitions.Add(colDef2);
      allSelector.ColumnDefinitions.Add(colDef3);
      RowDefinition rowDef1 = new RowDefinition();
      RowDefinition rowDef2 = new RowDefinition();
      RowDefinition rowDef3 = new RowDefinition();
      allSelector.RowDefinitions.Add(rowDef1);
      allSelector.RowDefinitions.Add(rowDef2);
      allSelector.RowDefinitions.Add(rowDef3);
      //allSelector.ShowGridLines = true;
      //allSelector.Width = 400;

      ////////////////////////////////////////////////////
      // Dropdown and label to select constraint group. //
      ////////////////////////////////////////////////////
      TextBlock groupSelectorLabel = new TextBlock();
      groupSelectorLabel.Text = "Physician/Protocol:";
      groupSelectorLabel.TextAlignment = TextAlignment.Right;
      groupSelectorLabel.VerticalAlignment = VerticalAlignment.Center;
      groupSelectorMenu.Margin = new Thickness(10, 10, 10, 10);
      //groupSelectorMenu.Width = 200;
      groupSelectorMenu.SelectionChanged += OnSelectionChanged;
      groupSelectorMenu.ItemsSource = groupList;

      Grid.SetRow(groupSelectorLabel, 0);
      Grid.SetColumn(groupSelectorLabel, 0);
      Grid.SetRow(groupSelectorMenu, 0);
      Grid.SetColumn(groupSelectorMenu, 1);
      allSelector.Children.Add(groupSelectorLabel);
      allSelector.Children.Add(groupSelectorMenu);

      ///////////////////////////////////////////////////////
      // Dropdown and label to select constraint protocol. //
      ///////////////////////////////////////////////////////
      TextBlock protocolSelectorLabel = new TextBlock();
      protocolSelectorLabel.Text = "Constraint Template:";
      protocolSelectorLabel.TextAlignment = TextAlignment.Right;
      protocolSelectorLabel.VerticalAlignment = VerticalAlignment.Center;
      protocolSelectorMenu.Margin = new Thickness(10, 10, 10, 10);

      Grid.SetRow(protocolSelectorLabel, 1);
      Grid.SetColumn(protocolSelectorLabel, 0);
      Grid.SetRow(protocolSelectorMenu, 1);
      Grid.SetColumn(protocolSelectorMenu, 1);
      allSelector.Children.Add(protocolSelectorLabel);
      allSelector.Children.Add(protocolSelectorMenu);

      /////////////////////////////////////////////////
      // Dropdown and label to select plan/plan sum. //
      /////////////////////////////////////////////////
      TextBlock planSelectorLabel = new TextBlock();
      planSelectorLabel.Text = "Selected Plan/Sum:";
      planSelectorLabel.TextAlignment = TextAlignment.Right;
      planSelectorLabel.VerticalAlignment = VerticalAlignment.Center;
      planSelectorMenu.Margin = new Thickness(10, 10, 10, 10);
      planSelectorMenu.ItemsSource = planList;
      planSelectorMenu.SelectedIndex = 0;

      Grid.SetRow(planSelectorLabel, 2);
      Grid.SetColumn(planSelectorLabel, 0);
      Grid.SetRow(planSelectorMenu, 2);
      Grid.SetColumn(planSelectorMenu, 1);
      allSelector.Children.Add(planSelectorLabel);
      allSelector.Children.Add(planSelectorMenu);

      ////////////////////////////////////////////////////////////
      // Button to evaluate using selected constraint template. //
      ////////////////////////////////////////////////////////////
      btn1.Content = "Evaluate";
      btn1.Margin = new Thickness(10, 10, 10, 10);
      btn1.Width = 200;
      btn1.Click += delegate(object sender, RoutedEventArgs e){
          OnClick1(sender, e, context); };
      Grid.SetRow(btn1, 0);
      Grid.SetColumn(btn1, 2);
      allSelector.Children.Add(btn1);

      /////////////////////////////////////////////////////////////
      // Button to evaluate using constraint template from file. //
      /////////////////////////////////////////////////////////////
      btn2.Content = "Evaluate From File";
      btn2.Margin = new Thickness(10, 10, 10, 10);
      btn2.Width = 200;
      btn2.Click += delegate(object sender, RoutedEventArgs e){
          OnClick2(sender, e, context); };
      Grid.SetRow(btn2, 1);
      Grid.SetColumn(btn2, 2);
      allSelector.Children.Add(btn2);

      //////////////////////////////////////
      // 3. Display results in new window //
      //////////////////////////////////////
      window.Title = "DoseEval";
      //window.Closing += new System.ComponentModel.CancelEventHandler(OnWindowClosing);
      //window.Background = System.Windows.Media.Brushes.Cornsilk;
      window.Height = 800;
      window.Width = 900;
      StackPanel rootPanel = new StackPanel();
      rootPanel.Orientation = Orientation.Vertical;
      rootPanel.Children.Add(allSelector);
      rootPanel.Children.Add(doseDataGrid);
      rootPanel.Children.Add(scriptNotes);
      window.Content = rootPanel;
    }

    private void OnSelectionChanged(object sender, EventArgs e)
    {
      protocolSelectorMenu.ItemsSource = protocolList[groupSelectorMenu.SelectedIndex];
    }

    private void OnClick1(object sender, RoutedEventArgs e,
        ScriptContext context)
    {
      string folder = mainFolder+"\\Constraint Data\\";
      if(!Directory.Exists(folder)) folder = "Constraint Data\\";
      string file = folder + groupSelectorMenu.SelectedValue + '_' +
          protocolSelectorMenu.SelectedValue + ".csv";
      if(File.Exists(file))
      {
        btn1.Background = System.Windows.Media.Brushes.LightBlue;
        missingText = EvaluateConstraints(context, file, doseDataGrid);
        scriptNotes.Text = missingText;
      }
      else
      {
        MessageBox.Show("Error: constraint file not found!");
      }
    }

    private void OnClick2(object sender, RoutedEventArgs e,
        ScriptContext context)
    {
      var fileDialog = new Microsoft.Win32.OpenFileDialog();
      fileDialog.DefaultExt = "csv";
      fileDialog.Multiselect = false;
      fileDialog.Title = "Choose Constraint File";
      fileDialog.ShowReadOnly = true;
      fileDialog.Filter = "CSV files (*.csv)|*.csv";
      fileDialog.FilterIndex = 0;
      fileDialog.CheckFileExists = true;
      if(fileDialog.ShowDialog() == false)
      {
        MessageBox.Show("Error: No constraint file selected");
        return;
      }
      var file = fileDialog.FileName;

      btn2.Background = System.Windows.Media.Brushes.LightBlue;
      missingText = EvaluateConstraints(context, file, doseDataGrid);
      scriptNotes.Text = missingText;
    }

    private void DDG_AutoGeneratingColumn(object sender, DataGridAutoGeneratingColumnEventArgs e)
    {
      string[] column_headers = {"Organ Name", "Structure Name", "Constraint",
          "Constraint Value", "Plan Value", "Pass/Fail", "Notes"};
      string headername = e.Column.Header.ToString();
      if (headername == "OrganName")
      {
        e.Column.Header = "Organ Name";
        e.Column.Width = 110;
      }
      if (headername == "StructureName")
      {
        e.Column.Header = "Structure Name";
        e.Column.Width = 110;
      }
      if (headername == "ConstraintName")
      {
        e.Column.Header = "Constraint";
        e.Column.Width = 75;
      }
      if (headername == "ConstraintValue")
      {
        e.Column.Header = "Value";
        e.Column.Width = 75;
      }
      if (headername == "PlanValue")
      {
        e.Column.Header = "Plan Value";
        e.Column.Width = 75;
      }
      if (headername == "PassFail")
      {
        e.Column.Header = "Pass/Fail";
        e.Column.Width = 75;

        DataGridTextColumn tCol = (DataGridTextColumn)e.Column;
        Style myStyle = new Style(typeof(TextBlock));
        Binding b = new Binding();
        b.RelativeSource = RelativeSource.Self;
        b.Path = new PropertyPath(TextBlock.TextProperty);
        b.Converter = new PassFailBrushConverter();
        myStyle.Setters.Add(new Setter(TextBlock.ForegroundProperty, b));
        tCol.ElementStyle = myStyle;

        e.Column = tCol;
      }
      if (headername == "ConstraintComplication")
      {
        e.Column.Header = "Notes";
        e.Column.Width = 340;
      }
    }

    public class ConstraintData
    {
      public string OrganName { get; set; }
      public bool Bilateral;
      public List<string> ConstraintList { get; set; }
      public List<string> ConstraintComparators { get; set; }
      public List<double> ConstraintValues { get; set; }
      public List<string> ConstraintUnits { get; set; }
      public List<string> ConstraintComplications { get; set; }

      public ConstraintData() {
        ConstraintList = new List<string>();
        ConstraintComparators = new List<string>();
        ConstraintValues = new List<double>();
        ConstraintUnits = new List<string>();
        ConstraintComplications = new List<string>();
      }
    }

    private List<ConstraintData> LoadConstraintData(string filename)
    {
      List<ConstraintData> data = new List<ConstraintData>();
      string[] fields;

      try
      {
        var reader = new StreamReader(File.OpenRead(filename));
        while (!reader.EndOfStream)
        {
          ConstraintData tempConstraint = new ConstraintData();
          fields = reader.ReadLine().Split(',');
          tempConstraint.OrganName = fields[0];
          tempConstraint.Bilateral = fields[1].Equals("1");
          fields = reader.ReadLine().Split(',');
          while(fields[0] != "end")
          {
            tempConstraint.ConstraintList.Add(fields[0]);
            tempConstraint.ConstraintComparators.Add(fields[1]);
            tempConstraint.ConstraintValues.Add(Convert.ToDouble(fields[2]));
            tempConstraint.ConstraintUnits.Add(fields[3]);
            tempConstraint.ConstraintComplications.Add(fields[4]);
            fields = reader.ReadLine().Split(',');
          }
          data.Add(tempConstraint);
        }
        reader.Close();
      }
      catch(Exception e)
      {
        MessageBox.Show(e.Message);
      }

      return data;
    }

    private List< List<string> > LoadDictionary(string filename)
    {
      List< List<string> > dictionary = new List< List<string> >();
      string[] fields;

      try
      {
        var reader = new StreamReader(File.OpenRead(filename));
        while (!reader.EndOfStream)
        {
          fields = reader.ReadLine().Split(',');
          var tempList = new List<string>(fields);
          dictionary.Add(tempList);
        }
        reader.Close();
      }
      catch(Exception e)
      {
        MessageBox.Show(e.Message);
      }

      return dictionary;
    }

    private List<string> MatchOrganName(string oName, bool bilateral,
        List<string> pNames, List< List<string> > dictionary)
    {
      var tempList = new List<string>();
      // First, try exact match with constraint name (case-sensitive)
      if(pNames.Contains(oName)) tempList.Add(oName);
      // Then try dictionary (case-insensitive)
      else
      {
        int dictionary_id = -1;
        for(int d = 0; d < dictionary.Count(); d++)
        {
          if(dictionary[d].Contains(oName,
            StringComparer.OrdinalIgnoreCase)) dictionary_id = d;
        }
        // If organ name is not in dictionary add "None" (meaning no match),
        // otherwise try dictionary search (case-insensitive)
        if((dictionary_id == -1))
        {
          tempList.Add("None");
        }
        else
        {
          for(int o = 0; o < pNames.Count(); o++)
          {
            string nameToCheck = pNames[o].Replace('_',' ');
            var checkList = new List<string>();
            checkList.AddRange(dictionary[dictionary_id]);
            if(bilateral)
            {
              int oldLength = dictionary[dictionary_id].Count();
              for(int l = 0; l < oldLength; l++)
              {
                checkList.Add("L "+checkList[l]);
                checkList.Add("Lt "+checkList[l]);
                checkList.Add("Left "+checkList[l]);
                checkList.Add(checkList[l]+" L");
                checkList.Add(checkList[l]+" Lt");
                checkList.Add(checkList[l]+" Left");
                checkList.Add("R "+checkList[l]);
                checkList.Add("Rt "+checkList[l]);
                checkList.Add("Right "+checkList[l]);
                checkList.Add(checkList[l]+" R");
                checkList.Add(checkList[l]+" Rt");
                checkList.Add(checkList[l]+" Right");
              }
            }
            if(checkList.Contains(nameToCheck,
                StringComparer.OrdinalIgnoreCase)) tempList.Add(pNames[o]);
          }
        }
      }

      // If dictionary search fails, add "None" (meaning no match)
      if(tempList.Count() == 0) tempList.Add("None");

      return tempList;
    }

    public class PlanComparison
    {
      public string OrganName { get; set; }
      public string StructureName { get; set; }
      public string ConstraintName { get; set; }
      public string ConstraintValue { get; set; }
      public string PlanValue { get; set; }
      public string PassFail { get; set; }
      public string ConstraintComplication { get; set; }
    }

    private string EvaluateConstraints(ScriptContext context,
        string filename, DataGrid data)
    {
      // Load plan/sum based on user input; check for valid dose and structures
      PlanSetup plan = null;
      PlanSum psum = null;
      foreach(PlanSetup p in context.PlansInScope)
      {
        if(p.Id == planList[planSelectorMenu.SelectedIndex]) plan = p;
      }
      foreach(PlanSum s in context.PlanSumsInScope)
      {
        if(s.Id == planList[planSelectorMenu.SelectedIndex]) psum = s;
      }
      SelectedPlanningItem = plan != null ? (PlanningItem)plan : (PlanningItem)psum;
      // Plans in plansum can have different structuresets but here we only use structureset to allow chosing one structure
      //SelectedStructureSet = plan != null ? plan.StructureSet : psum.PlanSetups.First().StructureSet;
      SelectedStructureSet = plan != null ? plan.StructureSet : psum.StructureSet;
      if (SelectedPlanningItem.Dose == null)
      {
        MessageBox.Show("Error: No calculated dose");
        return "Error: No calculated dose";
      }
      if (SelectedStructureSet == null)
      {
        MessageBox.Show("Error: Could not find a structure set");
        return "Error: Could not find a structure set";
      }

      // Other calculation variables
      double cValue, pValue;
      string excludedList = "Constraints Excluded:\n";
      bool any_missing = false;

      // Load QUANTEC dose constraint data and dictionary
      List<ConstraintData> Constraints = LoadConstraintData(filename);
      string d_file = mainFolder+"\\Resources\\OrganDictionary.csv";
      if(!File.Exists(d_file)) d_file = "Resources\\OrganDictionary.csv";
      List< List<string> > Dictionary = LoadDictionary(d_file);

      // Make list of plan structure names
      List<string> PlanNames = new List<string>();
      foreach(var s in SelectedStructureSet.Structures)
      {
        if(!s.IsEmpty) PlanNames.Add(s.Id);
      }
      // For each organ constraint, see if any plan structure names match
      // If so, compare constraint and add to list
      List<PlanComparison> results = new List<PlanComparison>();
      for(int n = 0; n < Constraints.Count(); n++)
      {
        // Find matching plan structure name, if able
        List<string> matchNames = MatchOrganName(Constraints[n].OrganName,
            Constraints[n].Bilateral, PlanNames, Dictionary);
        if(matchNames[0] == "None")
        {
          excludedList += (Constraints[n].OrganName+", ");
          any_missing = true;
          continue;
        }
        // If matching plan structure name found, compare constraint(s)
        for(int m = 0; m < matchNames.Count(); m++)
        {
          for(int p = 0; p < Constraints[n].ConstraintList.Count(); p++)
          {
            // Temporary container for constraint row data
            PlanComparison tempResult = new PlanComparison();

            // Set name and load constraint name & value from file
            tempResult.OrganName = Constraints[n].OrganName;
            tempResult.StructureName = matchNames[m];
            tempResult.ConstraintName = Constraints[n].ConstraintList[p];
            cValue = Constraints[n].ConstraintValues[p];
            tempResult.ConstraintValue = cValue.ToString()+" "+Constraints[n].ConstraintUnits[p];
            tempResult.ConstraintComplication = Constraints[n].ConstraintComplications[p];
            // Load DVH for particular structure
            Structure oar = (from s in SelectedStructureSet.Structures where
              s.Id == tempResult.StructureName select s).FirstOrDefault();
            DVHData dvhData = SelectedPlanningItem.GetDVHCumulativeData(oar,
            DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
            // Calculate plan value from DVH
            if(tempResult.ConstraintName == "Max")
            {
              pValue = dvhData.MaxDose.Dose / 100.0;
              tempResult.PlanValue = pValue.ToString("F1")+" "+Constraints[n].ConstraintUnits[p];
            }
            else if(tempResult.ConstraintName == "Mean")
            {
              pValue = dvhData.MeanDose.Dose / 100.0;
              tempResult.PlanValue = pValue.ToString("F1")+" "+Constraints[n].ConstraintUnits[p];
            }
            else if(tempResult.ConstraintName == "Min")
            {
              pValue = dvhData.MinDose.Dose / 100.0;
              tempResult.PlanValue = pValue.ToString("F1")+" "+Constraints[n].ConstraintUnits[p];
            }
            else
            {
              string current = Constraints[n].ConstraintList[p];
              double cNumber = Convert.ToDouble(current.Substring(1));
              if(current[0] == 'V')
              {
                if(Constraints[n].ConstraintUnits[p] == "cc")
                {
                  pValue = MyDVH.VolumeAtDose(dvhData, cNumber*100.0, true);
                }
                else pValue = MyDVH.VolumeAtDose(dvhData, cNumber*100.0);
                tempResult.PlanValue = pValue.ToString("F1")+" "+Constraints[n].ConstraintUnits[p];
              }
              else
              {
                pValue = (MyDVH.DoseAtVolume(dvhData, cNumber)).Dose / 100.0;
                tempResult.PlanValue = pValue.ToString("F1")+" "+Constraints[n].ConstraintUnits[p];
              }
            }
            // Compare plan and constraint values
            if (Constraints[n].ConstraintComparators[p] == "<")
            {
              if(cValue >= pValue) tempResult.PassFail = "Pass";
              else tempResult.PassFail = "Fail";
            }
            else
            {
              if(cValue <= pValue) tempResult.PassFail = "Pass";
              else tempResult.PassFail = "Fail";
            }
            tempResult.ConstraintName += " (" + Constraints[n].ConstraintComparators[p] + ")";
            results.Add(tempResult);
          }
        }
      }

      data.ItemsSource = results;
      if(any_missing) excludedList = excludedList.Remove(excludedList.Count()-2);
      else excludedList += ("None");
      return excludedList;
    }
  }

  class MyDVH
  {
    public static DoseValue DoseAtVolume(DVHData dvhData, double volume)
    {
      if (dvhData == null || dvhData.CurveData.Count() == 0)
        return DoseValue.UndefinedDose();
      double absVolume = dvhData.CurveData[0].VolumeUnit == "%" ? volume * dvhData.Volume * 0.01 : volume;
      if (volume < 0.0 || absVolume > dvhData.Volume)
        return DoseValue.UndefinedDose();
      DVHPoint[] hist = dvhData.CurveData;
      for(int i = 0; i < hist.Length; i++)
      {
        if (hist[i].Volume < volume) return hist[i].DoseValue;
      }
      return DoseValue.UndefinedDose();
    }

    public static double VolumeAtDose(DVHData dvhData, double dose, bool absolute = false)
    {
      if (dvhData == null) return Double.NaN;

      DVHPoint[] hist = dvhData.CurveData;
      int index = (int)(hist.Length * dose / dvhData.MaxDose.Dose);
      if (index < 0 || index > hist.Length) return 0.0;
      else
      {
        double relVolume = hist[index].Volume;
        double absVolume = relVolume * dvhData.Volume * 0.01;
        if(absolute) return absVolume;
        else return relVolume;
      }
    }
  }
}
