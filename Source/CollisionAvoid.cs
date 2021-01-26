/******************************************************************************/
//
// CollisionAvoid.cs
// Eclipse Linac Collision Evaluation Script
// Steven Dolly
// Created: January 7, 2021
//
// This script evaluates the contours (i.e. structures) from the current plan or
// plan sum.
//
/******************************************************************************/

using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Media3D;
using System.Collections.Generic;

using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace VMS.TPS
{
  public class Script
  {
    public Script()
    {
    }

    Viewport3D myViewport3D = new Viewport3D();
    double cameraRadius, cameraTheta, cameraPhi;
    VVector cv;

    public void Execute(ScriptContext context, Window window)
    {
      ////////////////////////////////////////////////
      // 1. Check for valid plan with structure set //
      ////////////////////////////////////////////////
      PlanSetup plan = context.PlanSetup;
      if(plan == null)
      {
        MessageBox.Show("Error: No plan loaded");
        return;
      }
      StructureSet structureSet = plan.StructureSet;
      if(structureSet == null)
      {
        MessageBox.Show("Error: No structure set loaded");
        return;
      }

      ////////////////////////////////////////////////////////////
      // 2. Load plan/structure data; calculate collision zones //
      ////////////////////////////////////////////////////////////
      Structure bodyStructure = structureSet.Structures.FirstOrDefault(s => s.Id == "BODY");
      MeshGeometry3D bodyContour = bodyStructure.MeshGeometry;
      Color bodyColor = bodyStructure.Color;

      MeshGeometry3D couchContour = new MeshGeometry3D();
      Color couchColor = new Color();
      foreach(var s in structureSet.Structures)
      {
        if(s.Id.Contains("CouchSurface"))
        {
          Structure couchStructure = s;
          couchContour = couchStructure.MeshGeometry;
          couchColor = couchStructure.Color;
          break;
        }
      }

      VVector iso = plan.Beams.ElementAt(0).IsocenterPosition;

      double radius = 380.0;
      double r, ga;
      Point3D p = new Point3D();
      List<double> collisionAngles = new List<double>();
      for(int n = 0; n < bodyContour.Positions.Count(); n++)
      {
        p = bodyContour.Positions[n];
        r = Math.Sqrt(Math.Pow(p.X - iso.x, 2) + Math.Pow(p.Y - iso.y, 2));
        if(r >= radius)
        {
          ga = (Math.Atan2(p.Y - iso.y, p.X - iso.x))*(180.0/Math.PI) + 90.0;
          while(ga < 0.0) ga += 360.0;
          collisionAngles.Add(ga);
        }
      }
      for(int n = 0; n < couchContour.Positions.Count(); n++)
      {
        p = couchContour.Positions[n];
        r = Math.Sqrt(Math.Pow(p.X - iso.x, 2) + Math.Pow(p.Y - iso.y, 2));
        if(r >= radius)
        {
          ga = (Math.Atan2(p.Y - iso.y, p.X - iso.x))*(180.0/Math.PI) + 90.0;
          while(ga < 0.0) ga += 360.0;
          collisionAngles.Add(ga);
        }
      }
      double minVal = 0.0, maxVal = 0.0;
      if(collisionAngles.Count() > 0)
      {
        minVal = collisionAngles.Min();
        maxVal = collisionAngles.Max();
        string results = minVal.ToString("F1")+" - "+maxVal.ToString("F1");
        //MessageBox.Show(results);
      }

      ///////////////////////////////////
      // 3. Make collision zone visual //
      ///////////////////////////////////
      MeshGeometry3D linacZonePass = new MeshGeometry3D();
      Color linacColorPass = new Color();
      linacColorPass = Color.FromScRgb(0.6f, 0.5f, 0.5f, 0.5f);
      Point3DCollection vertices = new Point3DCollection();
      Vector3DCollection normals = new Vector3DCollection();
      Int32Collection indices = new Int32Collection();
      int numSpokes = 100;
      double halfLength = 150.0;
      double startAngle = (maxVal-90)*(Math.PI/180.0);
      double dA = (360 - (maxVal - minVal)) / (double)(numSpokes);
      vertices.Add(new Point3D(iso.x + radius*Math.Cos(startAngle),
        iso.y + radius*Math.Sin(startAngle), -halfLength));
      vertices.Add(new Point3D(iso.x + radius*Math.Cos(startAngle),
        iso.y + radius*Math.Sin(startAngle), halfLength));
      normals.Add(new Vector3D(Math.Cos(startAngle), Math.Sin(startAngle), 0.0));
      normals.Add(new Vector3D(Math.Cos(startAngle), Math.Sin(startAngle), 0.0));
      for(int n = 1; n <= numSpokes; n++)
      {
        double angle = startAngle + Convert.ToDouble(n)*dA*(Math.PI/180.0);
        vertices.Add(new Point3D(iso.x + radius*Math.Cos(angle),
          iso.y + radius*Math.Sin(angle), -halfLength));
        vertices.Add(new Point3D(iso.x + radius*Math.Cos(angle),
          iso.y + radius*Math.Sin(angle), halfLength));
        normals.Add(new Vector3D(Math.Cos(angle), Math.Sin(angle), 0.0));
        normals.Add(new Vector3D(Math.Cos(angle), Math.Sin(angle), 0.0));
        indices.Add(2*(n-1));
        indices.Add(2*(n-1) + 3);
        indices.Add(2*(n-1) + 1);
        indices.Add(2*(n-1));
        indices.Add(2*(n-1) + 2);
        indices.Add(2*(n-1) + 3);
      }
      //indices.Add(2*(numSpokes-1));
      //indices.Add(2*(numSpokes-1) + 1);
      //indices.Add(1);
      //indices.Add(2*(numSpokes-1));
      //indices.Add(1);
      //indices.Add(0);
      linacZonePass.Positions = vertices;
      linacZonePass.Normals = normals;
      linacZonePass.TriangleIndices = indices;

      MeshGeometry3D linacZoneFail = new MeshGeometry3D();
      Color linacColorFail = new Color();
      linacColorFail = Color.FromScRgb(1.0f, 1.0f, 0.0f, 0.0f);
      Point3DCollection verticesF = new Point3DCollection();
      Vector3DCollection normalsF = new Vector3DCollection();
      Int32Collection indicesF = new Int32Collection();
      //numSpokes = (maxVal - minVal) + 1;
      startAngle = (minVal-90)*(Math.PI/180.0);
      dA = (maxVal - minVal) / (double)(numSpokes);
      verticesF.Add(new Point3D(iso.x + radius*Math.Cos(startAngle),
        iso.y + radius*Math.Sin(startAngle), -halfLength));
      verticesF.Add(new Point3D(iso.x + radius*Math.Cos(startAngle),
        iso.y + radius*Math.Sin(startAngle), halfLength));
      normalsF.Add(new Vector3D(Math.Cos(startAngle), Math.Sin(startAngle), 0.0));
      normalsF.Add(new Vector3D(Math.Cos(startAngle), Math.Sin(startAngle), 0.0));
      for(int n = 1; n <= numSpokes; n++)
      {
        double angle = startAngle + Convert.ToDouble(n)*dA*(Math.PI/180.0);
        verticesF.Add(new Point3D(iso.x + radius*Math.Cos(angle),
          iso.y + radius*Math.Sin(angle), -halfLength));
        verticesF.Add(new Point3D(iso.x + radius*Math.Cos(angle),
          iso.y + radius*Math.Sin(angle), halfLength));
        normalsF.Add(new Vector3D(Math.Cos(angle), Math.Sin(angle), 0.0));
        normalsF.Add(new Vector3D(Math.Cos(angle), Math.Sin(angle), 0.0));
        indicesF.Add(2*(n-1));
        indicesF.Add(2*(n-1) + 3);
        indicesF.Add(2*(n-1) + 1);
        indicesF.Add(2*(n-1));
        indicesF.Add(2*(n-1) + 2);
        indicesF.Add(2*(n-1) + 3);
      }
      linacZoneFail.Positions = verticesF;
      linacZoneFail.Normals = normalsF;
      linacZoneFail.TriangleIndices = indicesF;

      //////////////////////////////////////
      // 4. Display results in new window //
      //////////////////////////////////////

      // 3D-view objects
      Model3DGroup myModel3DGroup = new Model3DGroup();
      GeometryModel3D bodyGeometryModel = new GeometryModel3D();
      GeometryModel3D couchGeometryModel = new GeometryModel3D();
      GeometryModel3D linacPassGeometryModel = new GeometryModel3D();
      GeometryModel3D linacFailGeometryModel = new GeometryModel3D();
      ModelVisual3D myModelVisual3D = new ModelVisual3D();

      // Defines the camera used to view the 3D object.
      PerspectiveCamera myPCamera = new PerspectiveCamera();
      cameraRadius = 1000.0;
      cameraTheta = 90.0*(Math.PI/180.0);
      cameraPhi = 0.0*(Math.PI/180.0);
      cv = bodyStructure.CenterPoint;
      double cameraX = cv.x + cameraRadius*Math.Sin(cameraTheta)*Math.Cos(cameraPhi - Math.PI/2.0);
      double cameraY = cv.y + cameraRadius*Math.Sin(cameraTheta)*Math.Sin(cameraPhi - Math.PI/2.0);
      double cameraZ = cv.z + (cameraRadius * Math.Cos(cameraTheta));
      myPCamera.Position = new Point3D(cameraX, cameraY, cameraZ);
      myPCamera.LookDirection = new Vector3D(cv.x-cameraX, cv.y-cameraY, cv.z-cameraZ);
      myPCamera.FieldOfView = 60; // in degrees
      myPCamera.UpDirection = new Vector3D(0.0,0.0,1.0);
      myViewport3D.Camera = myPCamera;

      // Define directional lighting (points from camera to origin)
      DirectionalLight myDirectionalLight = new DirectionalLight();
      myDirectionalLight.Color = Colors.White;
      myDirectionalLight.Direction = myPCamera.LookDirection;
      myModel3DGroup.Children.Add(myDirectionalLight);

      // Initialize body mesh
      bodyGeometryModel.Geometry = bodyContour;
      DiffuseMaterial bodyMaterial = new DiffuseMaterial(new SolidColorBrush(bodyColor));
      bodyGeometryModel.Material = bodyMaterial;
      bodyGeometryModel.BackMaterial = bodyMaterial;
      myModel3DGroup.Children.Add(bodyGeometryModel);

      // Initialize couch mesh
      couchGeometryModel.Geometry = couchContour;
      DiffuseMaterial couchMaterial = new DiffuseMaterial(new SolidColorBrush(couchColor));
      couchGeometryModel.Material = couchMaterial;
      couchGeometryModel.BackMaterial = couchMaterial;
      myModel3DGroup.Children.Add(couchGeometryModel);

      // Initialize linac mesh
      linacPassGeometryModel.Geometry = linacZonePass;
      DiffuseMaterial linacPassMaterial = new DiffuseMaterial(new SolidColorBrush(linacColorPass));
      linacPassGeometryModel.Material = linacPassMaterial;
      linacPassGeometryModel.BackMaterial = linacPassMaterial;
      myModel3DGroup.Children.Add(linacPassGeometryModel);

      linacFailGeometryModel.Geometry = linacZoneFail;
      DiffuseMaterial linacFailMaterial = new DiffuseMaterial(new SolidColorBrush(linacColorFail));
      linacFailGeometryModel.Material = linacFailMaterial;
      linacFailGeometryModel.BackMaterial = linacFailMaterial;
      myModel3DGroup.Children.Add(linacFailGeometryModel);

      myModelVisual3D.Content = myModel3DGroup;
      myViewport3D.Children.Add(myModelVisual3D);

      // Calculate beam collision table
      Grid beamGrid = MakeBeamGrid(context, plan, minVal, maxVal);

      var mainGrid = new Grid();
      ColumnDefinition colDef1 = new ColumnDefinition();
      mainGrid.ColumnDefinitions.Add(colDef1);
      RowDefinition rowDef1 = new RowDefinition();
      RowDefinition rowDef2 = new RowDefinition();
      RowDefinition rowDef3 = new RowDefinition();
      RowDefinition rowDef4 = new RowDefinition();
      rowDef1.Height = new System.Windows.GridLength(32);
      rowDef2.Height = new System.Windows.GridLength(500);
      mainGrid.RowDefinitions.Add(rowDef1);
      mainGrid.RowDefinitions.Add(rowDef2);
      mainGrid.RowDefinitions.Add(rowDef3);
      mainGrid.RowDefinitions.Add(rowDef4);
      //mainGrid.ShowGridLines = true;

      TextBlock keyCommands = new TextBlock();
      keyCommands.Text = "Q/W: Rotate\tZ/X: Zoom\tR: Reset View";
      keyCommands.FontSize = 16;
      keyCommands.HorizontalAlignment = HorizontalAlignment.Center;
      Grid.SetRow(keyCommands, 0);
      Grid.SetColumn(keyCommands, 0);
      mainGrid.Children.Add(keyCommands);

      Grid.SetRow(myViewport3D, 1);
      Grid.SetColumn(myViewport3D, 0);
      mainGrid.Children.Add(myViewport3D);

      Grid.SetRow(beamGrid, 2);
      Grid.SetColumn(beamGrid, 0);
      mainGrid.Children.Add(beamGrid);

      TextBlock appNotes = new TextBlock();
      appNotes.Text = "**Extended gantry angles are not included in this calculation and are treated as unextended (e.g. 179.0E -> 179.0)";
      appNotes.HorizontalAlignment = HorizontalAlignment.Center;
      Grid.SetRow(appNotes, 3);
      Grid.SetColumn(appNotes, 0);
      mainGrid.Children.Add(appNotes);

      // Create window
      window.Title = "CollisionAvoid";
      //window.Closing += new System.ComponentModel.CancelEventHandler(OnWindowClosing);
      window.Height = 800;
      window.Width = 800;
      window.Content = mainGrid;
      window.KeyDown += HandleKeyPressEvent;
    }

    public bool CheckBounds(double a1, double a2, double b1, double b2)
    {
      // Map to -180 -> 180 range
      if(a1 > 180.0) a1 -= 360.0;
      if(a2 > 180.0) a2 -= 360.0;
      if(b1 > 180.0) b1 -= 360.0;
      if(b2 > 180.0) b2 -= 360.0;
      // Sort ranges
      double x1 = Math.Min(a1, a2);
      double x2 = Math.Max(a1, a2);
      double y1 = Math.Min(b1, b2);
      double y2 = Math.Max(b1, b2);
      // Compare
      return ((x2 >= y1) && (y2 >= x1));
    }

    public Grid MakeBeamGrid(ScriptContext context, PlanSetup plan, double c1, double c2)
    {
      // Load with beam data
      const int beamColumns = 6;
      List<string> beamText = new List<string>();
      List<double> gantryPath = new List<double>();
      beamText.Add("Beam ID");
      beamText.Add("Type");
      beamText.Add("Start");
      beamText.Add("Stop");
      beamText.Add("Delivery");
      beamText.Add("Move");
      double b1, b2;
      for(int b = 0; b < plan.Beams.Count(); b++)
      {
        Beam bm = plan.Beams.ElementAt(b);
        if(bm.IsSetupField) continue;
        gantryPath.Add(bm.ControlPoints[0].GantryAngle);
        gantryPath.Add(bm.ControlPoints[(bm.ControlPoints.Count()-1)].GantryAngle);
      }
      for(int b = 0; b < plan.Beams.Count(); b++)
      {
        Beam bm = plan.Beams.ElementAt(b);
        if(bm.IsSetupField) continue;
        beamText.Add(bm.Id);
        beamText.Add(bm.Technique.Id);
        beamText.Add(gantryPath[(2*b)].ToString("F1"));
        beamText.Add(gantryPath[(2*b+1)].ToString("F1"));
        // Check delivery for collision
        bool no_collision_zone = (c1 == 0.0 && c2 == 0.0);
        bool collide_with_zone = CheckBounds(gantryPath[(2*b)],
          gantryPath[(2*b+1)], c1, c2);
        if(!no_collision_zone && collide_with_zone) beamText.Add("Collision");
        else beamText.Add("OK");
        // Check move to next field for collision
        b1 = gantryPath[(2*b+1)];
        if(2*b+2 >= gantryPath.Count()) b2 = gantryPath[(2*b+1)];
        else b2 = gantryPath[(2*b+2)];
        collide_with_zone = CheckBounds(b1, b2, c1, c2);
        if(!no_collision_zone && collide_with_zone) beamText.Add("Collision");
        else beamText.Add("OK");
      }
      // Make grid for beam collision information
      Grid bGrid = new Grid();
      bGrid.HorizontalAlignment = HorizontalAlignment.Center;
      bGrid.VerticalAlignment = VerticalAlignment.Top;
      bGrid.Margin = new Thickness(16.0);
      double[] bg_widths = new double[beamColumns]{180.0, 60.0, 60.0, 60.0, 60.0, 60.0};
      for(int n = 0; n < beamColumns; n++)
      {
        ColumnDefinition cDef = new ColumnDefinition();
        cDef.Width = new GridLength(bg_widths[n]);
        bGrid.ColumnDefinitions.Add(cDef);
      }
      for(int n = 0; n <= (beamText.Count/beamColumns); n++)
      {
        RowDefinition rDef = new RowDefinition();
        if(n == (beamText.Count/beamColumns)) rDef.Height = new GridLength(36.0);
        else rDef.Height = new GridLength(20.0);
        bGrid.RowDefinitions.Add(rDef);
      }
      for(int n = 0; n < beamText.Count; n++)
      {
        TextBlock txtblk = new TextBlock();
        txtblk.Text = beamText[n];
        txtblk.VerticalAlignment = VerticalAlignment.Center;
        if(n < beamColumns) txtblk.FontWeight = FontWeights.Bold;
        if(txtblk.Text == "OK") txtblk.Foreground = System.Windows.Media.Brushes.Green;
        if(txtblk.Text == "Collision") txtblk.Foreground = System.Windows.Media.Brushes.Red;
        Grid.SetColumn(txtblk, n%beamColumns);
        Grid.SetRow(txtblk, n/beamColumns);
        bGrid.Children.Add(txtblk);
      }
      // Make bottom row for collision zone information
      TextBlock czText = new TextBlock();
      czText.Text = "Collision Zone: ";
      czText.VerticalAlignment = VerticalAlignment.Center;
      czText.HorizontalAlignment = HorizontalAlignment.Center;
      czText.FontWeight = FontWeights.Bold;
      if(c1 == 0.0 && c2 == 0.0) czText.Text += "None";
      else czText.Text += c1.ToString("F1")+"-"+c2.ToString("F1");
      Grid.SetColumn(czText, 0);
      Grid.SetRow(czText, (beamText.Count/beamColumns));
      Grid.SetColumnSpan(czText, beamColumns);
      bGrid.Children.Add(czText);

      return bGrid;
    }

    private void HandleKeyPressEvent(object sender, KeyEventArgs e)
    {
      Key[] viewCheck = new Key[]{Key.W, Key.Q, Key.Z, Key.X, Key.R};
      if(Array.Exists(viewCheck, element => element == e.Key))
      {
        // Adjust camera
        PerspectiveCamera newPCamera = myViewport3D.Camera as PerspectiveCamera;
        if(e.Key == Key.Q) cameraPhi -= 10.0*(Math.PI/180.0);
        if(e.Key == Key.W) cameraPhi += 10.0*(Math.PI/180.0);
        if(e.Key == Key.Z) cameraRadius -= 100.0;
        if(e.Key == Key.X) cameraRadius += 100.0;
        if(e.Key == Key.R){ cameraRadius = 1000.0; cameraPhi = 0.0; }
        double cameraX = cv.x + cameraRadius*Math.Sin(cameraTheta)*Math.Cos(cameraPhi - Math.PI/2.0);
        double cameraY = cv.y + cameraRadius*Math.Sin(cameraTheta)*Math.Sin(cameraPhi - Math.PI/2.0);
        double cameraZ = cv.z + (cameraRadius * Math.Cos(cameraTheta));
        newPCamera.Position = new Point3D(cameraX, cameraY, cameraZ);
        newPCamera.LookDirection = new Vector3D(cv.x-cameraX, cv.y-cameraY, cv.z-cameraZ);
        newPCamera.UpDirection = new Vector3D(0.0, 0.0, 1.0);
        myViewport3D.Camera = newPCamera;
        // Update directional light source (points from camera to origin)
        DirectionalLight newDirLight = ((myViewport3D.Children[0] as ModelVisual3D).Content as Model3DGroup).Children[0] as DirectionalLight;
        newDirLight.Direction = newPCamera.LookDirection;
        ((myViewport3D.Children[0] as ModelVisual3D).Content as Model3DGroup).Children[0] = newDirLight;
      }
    }
  }
}
