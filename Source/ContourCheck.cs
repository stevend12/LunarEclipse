/******************************************************************************/
//
// ContourCheck.cs
// Eclipse Contour Evaluation Script
// Steven Dolly
// Created: March 16, 2020
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

    StructureSet SelectedStructureSet;

    Viewport3D myViewport3D = new Viewport3D();
    double cameraRadius, cameraTheta, cameraPhi;
    VVector cv;

    ComboBox structureList = new ComboBox();
    List<string> structureNames = new List<string>();

    public void Execute(ScriptContext context, Window window)
    {
      ////////////////////////////////////////////////
      // 1. Check for valid plan with structure set //
      ////////////////////////////////////////////////
      PlanSetup plan = context.PlanSetup;
      PlanSum psum = context.PlanSumsInScope.FirstOrDefault();
      if(plan == null && psum == null)
      {
        MessageBox.Show("Error: No plan or plan sum loaded");
        return;
      }
      SelectedStructureSet = plan != null ? plan.StructureSet : psum.StructureSet;
      Structure bodyStructure = SelectedStructureSet.Structures.FirstOrDefault(s => s.Id == "BODY");
      foreach(var s in SelectedStructureSet.Structures)
      {
        structureNames.Add(s.Id);
      }
      MeshGeometry3D bodyContour = bodyStructure.MeshGeometry;
      Color bodyColor = bodyStructure.Color;

      //////////////////////////////////////
      // 3. Display results in new window //
      //////////////////////////////////////

      // 3D-view objects
      Model3DGroup myModel3DGroup = new Model3DGroup();
      GeometryModel3D myGeometryModel = new GeometryModel3D();
      ModelVisual3D myModelVisual3D = new ModelVisual3D();

      // Defines the camera used to view the 3D object.
      PerspectiveCamera myPCamera = new PerspectiveCamera();
      cameraRadius = 800.0;
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

      // Use body contour as mesh
      myGeometryModel.Geometry = bodyContour;

      // Define material (e.g. brush) and apply to the mesh geometries.
      DiffuseMaterial myMaterial = new DiffuseMaterial(new SolidColorBrush(bodyColor));
      myGeometryModel.Material = myMaterial;
      myGeometryModel.BackMaterial = myMaterial;

      // Add the geometry model to the model group.
      myModel3DGroup.Children.Add(myGeometryModel);
      // Add the group of models to the ModelVisual3d.
      myModelVisual3D.Content = myModel3DGroup;
      // Add model visual to viewport
      myViewport3D.Children.Add(myModelVisual3D);

      structureList.ItemsSource = structureNames;
      structureList.SelectedIndex = 1;
      structureList.Height = 32;
      structureList.SelectionChanged += ChangeStructure;

      var mainGrid = new Grid();
      ColumnDefinition colDef1 = new ColumnDefinition();
      ColumnDefinition colDef2 = new ColumnDefinition();
      mainGrid.ColumnDefinitions.Add(colDef1);
      mainGrid.ColumnDefinitions.Add(colDef2);
      RowDefinition rowDef1 = new RowDefinition();
      rowDef1.Height = new System.Windows.GridLength(800);
      mainGrid.RowDefinitions.Add(rowDef1);
      mainGrid.ShowGridLines = true;

      Grid.SetRow(structureList, 0);
      Grid.SetColumn(structureList, 0);
      Grid.SetRow(myViewport3D, 0);
      Grid.SetColumn(myViewport3D, 1);

      mainGrid.Children.Add(structureList);
      mainGrid.Children.Add(myViewport3D);

      // Create window
      window.Title = "ContourCheck";
      //window.Closing += new System.ComponentModel.CancelEventHandler(OnWindowClosing);
      window.Height = 800;
      window.Width = 1000;
      window.Content = mainGrid;
      window.KeyDown += HandleKeyPressEvent;
    }

    private void ChangeStructure(object sender, EventArgs e)
    {
      // Get selected structure
      Structure viewStructure = SelectedStructureSet.Structures.FirstOrDefault(
        s => s.Id == structureNames[structureList.SelectedIndex]);
      MeshGeometry3D viewContour = viewStructure.MeshGeometry;
      Color viewColor = viewStructure.Color;
      // Use body contour as mesh
      GeometryModel3D viewGeometryModel = ((myViewport3D.Children[0] as ModelVisual3D).Content as Model3DGroup).Children[1] as GeometryModel3D;
      viewGeometryModel.Geometry = viewContour;
      // Define material (e.g. brush) and apply to the mesh geometries.
      DiffuseMaterial viewMaterial = new DiffuseMaterial(new SolidColorBrush(viewColor));
      viewGeometryModel.Material = viewMaterial;
      viewGeometryModel.BackMaterial = viewMaterial;
    }

    private void HandleKeyPressEvent(object sender, KeyEventArgs e)
    {
      Key[] rotateCheck = new Key[]{Key.W, Key.S, Key.A, Key.D};
      if(Array.Exists(rotateCheck, element => element == e.Key))
      {
        // Adjust camera
        PerspectiveCamera newPCamera = myViewport3D.Camera as PerspectiveCamera;
        if(e.Key == Key.W) cameraTheta -= 5.0*(Math.PI/180.0);
        if(e.Key == Key.S) cameraTheta += 5.0*(Math.PI/180.0);
        if(e.Key == Key.A) cameraPhi -= 5.0*(Math.PI/180.0);
        if(e.Key == Key.D) cameraPhi += 5.0*(Math.PI/180.0);
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
