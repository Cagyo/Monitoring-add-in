using ESRI.ArcGIS.DefenseSolutions;
using ESRI.ArcGIS.Framework;
using ESRI.ArcGIS.ArcMapUI;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.Display;
using ESRI.ArcGIS.DataSourcesFile;
using ESRI.ArcGIS.Carto;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ESRI.ArcGIS.Geoprocessing;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.CartoUI;
using System.Runtime.InteropServices;

namespace MonitoringSystem
{
    public partial class MainForm : Form
    {
        string gdb;
        public MainForm()
        {
            InitializeComponent();
        }
        

        #region "Add Shapefile Using OpenFileDialog"

        ///<summary>Add a shapefile to the ActiveView using the Windows.Forms.OpenFileDialog control.</summary>
        ///
        ///<param name="activeView">An IActiveView interface</param>
        /// 
        ///<remarks></remarks>
        public void AddShapefileUsingOpenFileDialog(IActiveView activeView)
        {
          //parameter check
          if (activeView == null)
          {
            return;
          }

          // Use the OpenFileDialog Class to choose which shapefile to load.
          System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
          openFileDialog.InitialDirectory = "c:\\";
          openFileDialog.Filter = "Shapefiles (*.shp)|*.shp";
          openFileDialog.FilterIndex = 2;
          openFileDialog.RestoreDirectory = true;
          openFileDialog.Multiselect = false;


          if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
          {
            // The user chose a particular shapefile.

            // The returned string will be the full path, filename and file-extension for the chosen shapefile. Example: "C:\test\cities.shp"
            string shapefileLocation = openFileDialog.FileName;
            textBox2.Text = shapefileLocation;
            if (shapefileLocation != "")
            {
              // Ensure the user chooses a shapefile

              // Create a new ShapefileWorkspaceFactory CoClass to create a new workspace
              IWorkspaceFactory workspaceFactory = new ShapefileWorkspaceFactoryClass();

              // IO.Path.GetDirectoryName(shapefileLocation) returns the directory part of the string. Example: "C:\test\"
              IFeatureWorkspace featureWorkspace = (IFeatureWorkspace)workspaceFactory.OpenFromFile(System.IO.Path.GetDirectoryName(shapefileLocation), 0); // Explicit Cast

              // IO.Path.GetFileNameWithoutExtension(shapefileLocation) returns the base filename (without extension). Example: "cities"
              IFeatureClass featureClass = featureWorkspace.OpenFeatureClass(System.IO.Path.GetFileNameWithoutExtension(shapefileLocation));

              IFeatureLayer featureLayer = new FeatureLayerClass();
              featureLayer.FeatureClass = featureClass;
              featureLayer.Name = featureClass.AliasName;
              featureLayer.Visible = true;
              activeView.FocusMap.AddLayer(featureLayer);

              // Zoom the display to the full extent of all layers in the map
              activeView.Extent = activeView.FullExtent;
              activeView.PartialRefresh(esriViewDrawPhase.esriViewGeography, null, null);
            }
            else
            {
              // The user did not choose a shapefile.
              // Do whatever remedial actions as necessary
              // Windows.Forms.MessageBox.Show("No shapefile chosen", "No Choice #1",
              //                                     Windows.Forms.MessageBoxButtons.OK,
              //                                     Windows.Forms.MessageBoxIcon.Exclamation);
            }
          }
          else
          {
            // The user did not choose a shapefile. They clicked Cancel or closed the dialog by the "X" button.
            // Do whatever remedial actions as necessary.
            // Windows.Forms.MessageBox.Show("No shapefile chosen", "No Choice #2",
            //                                      Windows.Forms.MessageBoxButtons.OK,
            //                                      Windows.Forms.MessageBoxIcon.Exclamation);
          }
        }
        #endregion
        

        private void button1_Click(object sender, EventArgs e)
        {
            //IApplication app;
            //var MxDox = ArcMap.Application.Document;
            var MxDoc = ArcMap.Application.Document as IMxDocument;
            IActiveView activeView = MxDoc.ActiveView;
            
            AddShapefileUsingOpenFileDialog(ArcMap.Document.ActiveView);
            
        }

        #region "Excel import to shape and generating layer (manual)"
        //Импорт данных из экселя в шейп, генерация слоя с населенными пуктами для которых присутствуют данные
        private void button2_Click(object sender, EventArgs e)
        {
            var MxDoc = ArcMap.Application.Document as IMxDocument;
            IActiveView activeView = MxDoc.ActiveView;

            // Get the TOC
            IContentsView IContentsView = MxDoc.CurrentContentsView;
            var layer = activeView.FocusMap.Layer[0] as FeatureLayer;
            //textBox1.Text = layer.FeatureClass.Fields.Field[layer.FeatureClass.Fields.FieldCount-1].Name;
            IClass Towns = layer.FeatureClass as IClass;
            IFeatureClass TownsFeats = layer.FeatureClass;
            //if (layer.FeatureClass.Fields.Field[layer.FeatureClass.Fields.FieldCount - 1].Name == "Emissions")
            //{
                int num = Towns.FindField("Emissions");
                if (num != -1)
                {
                    IField DeletedField = Towns.Fields.Field[num] as IField;
                    Towns.DeleteField(DeletedField);
                }
                num = Towns.FindField("Model");
                if (num != -1)
                {
                    IField DeletedField = Towns.Fields.Field[num] as IField;
                    Towns.DeleteField(DeletedField);
                }
                //Towns.DeleteField
            //}
            

                /*num = Towns.FindField("AllArea");
                if (num != -1)
                {
                    IField DeletedField = Towns.Fields.Field[num] as IField;
                    Towns.DeleteField(DeletedField);
                }*/

                IField Field2 = new FieldClass();
                IFieldEdit fieldEdit2 = (IFieldEdit)Field2;
                fieldEdit2.Name_2 = "Model";
                fieldEdit2.Type_2 = esriFieldType.esriFieldTypeDouble;
                Towns.AddField(Field2);

            IField Field = new FieldClass();
            IFieldEdit fieldEdit = (IFieldEdit)Field;
            fieldEdit.Name_2 = "Emissions";
            fieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
            Towns.AddField(Field);
            
            Dictionary<string, double> data = new Dictionary<string,double>();
            Dictionary<string, double> ParcedData = new Dictionary<string,double>();
            List<string> Types = new List<string>();
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "*.xls|*.xlsx";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox3.Text = openFileDialog1.FileName;
                Loader ldr = new Loader(this, openFileDialog1.FileName);//@"c:\Temp\13_11_2012\data.xlsx"
                try
                {
                    data = ldr.XlsRead();
                    foreach (var pair in data)
                    {
                        string[] result = pair.Key.Split('.');
                        Types.Add(result[0]);
                        //textBox1.Text += result[1].ToLower() + "\r\n";
                        //ParcedData.Add(result[1].ToUpper(), pair.Value);
                    }
                }
                catch (Exception exc)
                {
                    MessageBox.Show(exc.Message);
                }

                num = Towns.FindField("ukrname");
                if (num != -1)
                {
                    IQueryFilter queryFilter = new QueryFilterClass();
                    queryFilter.WhereClause = "FID > 0"; //"U" = domain value for Utility
                    ESRI.ArcGIS.Geodatabase.IFeatureCursor featureCursor = TownsFeats.Search(queryFilter, false);
                    // Get the first feature
                    ESRI.ArcGIS.Geodatabase.IFeature feature = featureCursor.NextFeature();
                    int InData = 0;
                    while (feature != null)
                    {

                        string TownName = feature.get_Value(2).ToString().ToUpper();
                        string TownType = feature.get_Value(3).ToString().ToUpper();
                        //textBox1.Text += TownType + "\r\n";
                        string GenTownName = "";
                        if (TownType == "МІСТО")
                        {
                            GenTownName = "м." + TownName;
                        }
                        else if (TownType == "СЕЛИЩЕ")
                        {
                            GenTownName = "с-ще." + TownName;
                        }
                        else if (TownType == "СЕЛО")
                        {
                            GenTownName = "с." + TownName;
                        }
                        else if (TownType == "СМТ.")
                        {
                            GenTownName = "смт." + TownName;
                        }
                        if (data.ContainsKey(GenTownName))
                        {
                            num = Towns.FindField("Emissions");
                            feature.set_Value(num, data[GenTownName]);
                            feature.Store();
                            //textBox1.Text += "\r\n" + GenTownName;
                            InData++;
                        }

                        feature = featureCursor.NextFeature();
                    }
                    //textBox1.Text = InData.ToString();
                }
                layer = activeView.FocusMap.Layer[0] as FeatureLayer;
                SelectMapFeaturesByAttributeQuery(activeView, layer, "Emissions > 0");
                IFeatureLayer NewLayer = new FeatureLayer();
                IFeatureLayerDefinition2 defLayer = layer as IFeatureLayerDefinition2;
                NewLayer = defLayer.CreateSelectionLayer("Analyzed_localities", true, "", "");
                activeView.FocusMap.AddLayer(NewLayer);
            }
            else MessageBox.Show("Отсутствует выбранный файл");
            groupBox3.Enabled = true;
            groupBox4.Enabled = true;
        }
        #endregion

        public void SetTextbox1(string msg)
        {
            textBox1.Text += msg;
        }
        #region "Построение буферов для городов"
        private void BuildBuffer(string Cities, string result, string distance)
        {

            // Create the geoprocessor.
            IGeoProcessor2 gp = new GeoProcessorClass();
            gp.SetEnvironmentValue("workspace", gdb);
            gp.OverwriteOutput = true;

            // Create an IVariantArray to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(Cities);
            parameters.Add(result);
            parameters.Add(distance);

            // Execute the tool.
            gp.Execute("Buffer_analysis", parameters, null);
        }
        #endregion

        #region "Построение буферов (button)"
        private void CreateBuffers()
        {
            /*var MxDoc = ArcMap.Application.Document as IMxDocument;
            IActiveView activeView = MxDoc.ActiveView;

            // Get the TOC
            IContentsView IContentsView = MxDoc.CurrentContentsView;
            var layer = activeView.FocusMap.Layer[0] as FeatureLayer;
            IFeature Town = layer.FeatureClass.GetFeature(1309);
            var area = Town.Shape as IArea;
            textBox1.Text = area.Area.ToString()+Town.get_Value(4);*/

            string RadiusLess = textBox5.Text;
            string RadiusMore = textBox6.Text;
            BuildBuffer("More20CitiesKm", "More20CitiesKm_buff", RadiusMore);
            BuildBuffer("Less20CitiesKm", "Less20CitiesKm_buff", RadiusLess);
            //BuildBuffer(@"C:\Temp\13_11_2012\Donetsk_region.shp", @"C:\Temp\13_11_2012\Region_Buffer.shp", "10 Kilometers");
            /*foreach (var dataset1 in Towns.GetRows(null,true))
            {

            }*/
        }
        #endregion


        #region "Get Azimuth from Two Points"

        ///<summary>Get the geodetically correct Rhumb Line azimuth between two points.</summary>
/// 
///<param name="fromPoint">An IPoint interface that is the start (or from) location</param>
///<param name="toPoint">An IPoint interface that is the end (or to) location</param>
///<param name="spatialReference">An esriSRGeoCSType enum that is a predefined geographic coordinate system. Example: ESRI.ArcGIS.Geometry.esriSRGeoCSType.esriSRGeoCS_NAD1983</param>
///  
///<returns>A System.Double that represents the true azimuth</returns>
///  
///<remarks></remarks>
public System.Double GetAzimuthFromTwoPoints(IPoint fromPoint, IPoint toPoint, esriSRGeoCSType spatialReference)
{

    // Define the spatial reference of the rhumb line.
    ISpatialReferenceFactory2 spatialReferenceFactory2 = new SpatialReferenceEnvironmentClass();
    ISpatialReference2 spatialReference2 = (ISpatialReference2)spatialReferenceFactory2.CreateSpatialReference((System.Int16)spatialReference);

    // Initialize the MeasurementTool and define the properties of the line.
    // These properties include the line type, which is a rhumb line in this case, and the 
    // spatial reference of the line. 
    IMeasurementTool measurementTool = new MeasurementToolClass();
    measurementTool.SpecialGeolineType = cjmtkSGType.cjmtkSGTRhumbLine;
    measurementTool.SpecialSpatialReference = spatialReference2;

    // Determine the distance and azimuth of the rhumb line based on the start and end point coordinates.
    measurementTool.ConstructByPoints(fromPoint, toPoint);

    // Return the Azimuth.
    return measurementTool.Angle;
}
#endregion

        #region "Select Map Features by Attribute Query"

        ///<summary>Select features in the IActiveView by an attribute query using a SQL syntax in a where clause.</summary>
        /// 
        ///<param name="activeView">An IActiveView interface</param>
        ///<param name="featureLayer">An IFeatureLayer interface to select upon</param>
        ///<param name="whereClause">A System.String that is the SQL where clause syntax to select features. Example: "CityName = 'Redlands'"</param>
        ///  
        ///<remarks>Providing and empty string "" will return all records.</remarks>
        public void SelectMapFeaturesByAttributeQuery(IActiveView activeView, IFeatureLayer featureLayer, System.String whereClause)
        {
          if(activeView == null || featureLayer == null || whereClause == null)
          {
            return;
          }
          IFeatureSelection featureSelection = featureLayer as IFeatureSelection; // Dynamic Cast

          // Set up the query
          IQueryFilter queryFilter = new QueryFilterClass();
          queryFilter.WhereClause = whereClause;

          // Invalidate only the selection cache. Flag the original selection
          activeView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);

          // Perform the selection
          featureSelection.SelectFeatures(queryFilter, esriSelectionResultEnum.esriSelectionResultNew, false);

          // Flag the new selection
          activeView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);
        }
        #endregion

        #region "Save Layer To File"

        ///<summary>Write and Layer to a file on disk.</summary>
        ///  
        ///<param name="layerFilePath">A System.String that is the path and filename for the layer file to be created. Example: "C:\temp\cities.lyr"</param>
        ///<param name="layer">An ILayer interface.</param>
        ///   
        ///<remarks></remarks>
        public void SaveToLayerFile(String layerFilePath, ILayer layer)
        {
          if(layer == null)
          {
            return;
          }
          //create a new LayerFile instance
          ILayerFile layerFile = new LayerFileClass();

          //make sure that the layer file name is valid
          if (System.IO.Path.GetExtension(layerFilePath) != ".lyr")
            return;
          if (layerFile.get_IsPresent(layerFilePath))
            System.IO.File.Delete(layerFilePath);

          //create a new layer file
          layerFile.New(layerFilePath);

          //attach the layer file with the actual layer
          layerFile.ReplaceContents(layer);

          //save the layer file
          layerFile.Save();
        }
        #endregion

        #region "Разбивка городов по площади"
        private void SplitSettlements()
        {
            var MxDoc = ArcMap.Application.Document as IMxDocument;
            IActiveView activeView = MxDoc.ActiveView;
            IContentsView IContentsView = MxDoc.CurrentContentsView;
            var layer = activeView.FocusMap.Layer[0] as FeatureLayer;
            SelectMapFeaturesByAttributeQuery(activeView, layer, "AllArea >= " + textBox4.Text);
            IFeatureLayer NewLayer = new FeatureLayer();
            IFeatureLayerDefinition2 defLayer = layer as IFeatureLayerDefinition2;
            NewLayer = defLayer.CreateSelectionLayer("More20CitiesKm", true, "", "");
            activeView.FocusMap.AddLayer(NewLayer);
            SaveToLayerFile(@"c:\temp\more20CitiesKm.lyr", NewLayer);
            SelectMapFeaturesByAttributeQuery(activeView, layer, "AllArea < " + textBox4.Text);
            NewLayer = defLayer.CreateSelectionLayer("Less20CitiesKm", true, "", "");
            activeView.FocusMap.AddLayer(NewLayer);
            SaveToLayerFile(@"c:\temp\Less20CitiesKm.lyr", NewLayer);


            //IFeatureLayer l1 = activeView.FocusMap.Layer[0] as IFeatureLayer;
            //IFeatureLayer l3 = activeView.FocusMap.Layer[2] as IFeatureLayer;
            
            /*IDataset dataset = layer.FeatureClass.FeatureDataset.Workspace.Datasets[0].Next();
            textBox1.Text = dataset.BrowseName;*/
            //IFeature Town = layer.FeatureClass.GetFeature(1309);
        }
        #endregion
       
        #region "Clear Selected Map Features"

///<summary>Clear the selected features in the IActiveView for a specified IFeatureLayer.</summary>
/// 
///<param name="activeView">An IActiveView interface</param>
///<param name="featureLayer">An IFeatureLayer</param>
/// 
///<remarks></remarks>
public void ClearSelectedMapFeatures(IActiveView activeView, IFeatureLayer featureLayer)
{
  if(activeView == null || featureLayer == null)
  {
    return;
  }
  IFeatureSelection featureSelection = featureLayer as IFeatureSelection; // Dynamic Cast

  // Invalidate only the selection cache. Flag the original selection
  activeView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);

  // Clear the selection
  featureSelection.Clear();

  // Flag the new selection
  activeView.PartialRefresh(esriViewDrawPhase.esriViewGeoSelection, null, null);
}
#endregion
        
        #region "Toggle Visibility Of Composite Layer"

/// <summary>
/// Toggle the visibility on and off for a composite layer in the TOC of a map.
/// </summary>
/// <param name="activeView">An IActiveView interface.</param>
/// <param name="layerIndex">A Int32 that is the index number of the composite layer in the TOC. Example: 0.</param>
/// <remarks>This snippet is useful for toggling the visibility of composite layers. An example of a composite layer is feature linked annotation.</remarks>
public void ToggleVisibleOfCompositeLayer(IActiveView activeView, System.Int32 layerIndex)
{

    IMap map = activeView.FocusMap;
    ILayer layer = map.get_Layer(layerIndex);
    ICompositeLayer2 compositeLayer2 = (ICompositeLayer2)layer;
    System.Int32 compositeLayerIndex = 0;

    if (layer.Visible)
    {

        //Turn the layer visibility off
        layer.Visible = false;

        //Turn each sub-layer (ie. composite layer) visibility off
        for (compositeLayerIndex = 0; compositeLayerIndex < compositeLayer2.Count; compositeLayerIndex++)
        {
            compositeLayer2.get_Layer(compositeLayerIndex).Visible = false;
        }

    }
    else
    {

        //Turn the layer visibility on
        layer.Visible = true;

        //Turn each sub-layer (ie. composite layer) visibility on
        for (compositeLayerIndex = 0; compositeLayerIndex < compositeLayer2.Count; compositeLayerIndex++)
        {
            compositeLayer2.get_Layer(compositeLayerIndex).Visible = true;
        }

    }

    //Refresh the TOC and the map window
    activeView.ContentsChanged();

}
#endregion

        #region "Центроиды"
        private void CreateCentroids()
        {
            var MxDoc = ArcMap.Application.Document as IMxDocument;
            IActiveView activeView = MxDoc.ActiveView;

            // Get the TOC
            IContentsView IContentsView = MxDoc.CurrentContentsView;
            
            var layer = activeView.FocusMap.Layer[4] as FeatureLayer;
            //textBox1.Text = layer.FeatureClass.Fields.Field[layer.FeatureClass.Fields.FieldCount-1].Name;
            ClearSelectedMapFeatures(activeView, layer);

            IGeoProcessor2 gp = new GeoProcessorClass();
            gp.SetEnvironmentValue("workspace", gdb);
            gp.OverwriteOutput = true;

            // Create an IVariantArray to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add("Analyzed_localities");
            parameters.Add("TownsToPoints");
            parameters.Add("CENTROID");

            // Execute the tool.
            gp.Execute("FeatureToPoint_management", parameters, null);
        }
#endregion


        #region "Классификация 5 кл, естественные границы"
        private double[] breaks;
        private double[] breaks2;
        private void ClassifyFeatures(string field, ref double[] breaks2)
        {
            IMxDocument pMxDoc;
            IGeoFeatureLayer pFLayer;
            IFeatureClass pFclass;
            IFeature pFeature;
            IFeatureCursor pFCursor;
            IClassBreaksRenderer pRender;
            ITable pTable;
            IGeoFeatureLayer pGeoLayer;
            IClassifyGEN pClassifyGEN;
            ITableHistogram pTableHistogram;
            IHistogram pHistogram;
            object frqs;
            object xVals;

            pMxDoc = ArcMap.Document;
            pFLayer = pMxDoc.ActiveView.FocusMap.Layer[0] as IGeoFeatureLayer;
            pFclass = pFLayer.FeatureClass;
            pFCursor = pFclass.Search(null, false);
            pFeature = pFCursor.NextFeature();
            pTable = pFclass as ITable;

            pGeoLayer = pFLayer;
            pRender = new ClassBreaksRenderer();
            pClassifyGEN = new NaturalBreaks();

            pTableHistogram = new TableHistogram() as ITableHistogram;
            pHistogram = pTableHistogram as IHistogram;

            pTableHistogram.Field = field;       // matches renderer field
            pTableHistogram.Table = pTable;

            pHistogram.GetHistogram(out xVals, out frqs);

            int classes = Convert.ToInt32(numericUpDown1.Value);
            pClassifyGEN.Classify(xVals, frqs, classes);      // use five classes

            pRender = new ClassBreaksRenderer();
            double[] cb = pClassifyGEN.ClassBreaks as double[];

            breaks2 = cb;
        }
        /*private void ClassifyMainFeatures()
        {
            IMxDocument pMxDoc;
            IGeoFeatureLayer pFLayer;
            IFeatureClass pFclass;
            IFeature pFeature;
            IFeatureCursor pFCursor;
            IClassBreaksRenderer pRender;
            ITable pTable;
            IGeoFeatureLayer pGeoLayer;
            IClassifyGEN pClassifyGEN;
            ITableHistogram pTableHistogram;
            IHistogram pHistogram;
            object frqs;
            object xVals;


            pMxDoc = ArcMap.Document;
            pFLayer = pMxDoc.ActiveView.FocusMap.Layer[0] as IGeoFeatureLayer;
            pFclass = pFLayer.FeatureClass;
            pFCursor = pFclass.Search(null, false);
            pFeature = pFCursor.NextFeature();
            pTable = pFclass as ITable;

            pGeoLayer = pFLayer;
            pRender = new ClassBreaksRenderer();
            pClassifyGEN = new NaturalBreaks();

            pTableHistogram = new TableHistogram() as ITableHistogram;
            pHistogram = pTableHistogram as IHistogram;

            pTableHistogram.Field = "Model";       // matches renderer field
            pTableHistogram.Table = pTable;

            pHistogram.GetHistogram(out xVals, out frqs);

            int classes = Convert.ToInt32(numericUpDown1.Value);
            pClassifyGEN.Classify(xVals, frqs, classes);       // use five classes

            pRender = new ClassBreaksRenderer();
            double[] cb = pClassifyGEN.ClassBreaks as double[];

            breaks = cb;
            
        }*/
        #endregion

        private void SelectLayerByLocation(string layer1, string layer2)
        {
            // Create the geoprocessor.
            IGeoProcessor2 gp = new GeoProcessorClass();
            gp.SetEnvironmentValue("workspace", gdb);
            gp.OverwriteOutput = true;

            // Create an IVariantArray to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(layer1);
            parameters.Add("intersect");
            parameters.Add(layer2);

            // Execute the tool.
            gp.Execute("SelectLayerByLocation_management", parameters, null);
        }

        private void CountRows(string layer, out int count)
        {
            // Create the geoprocessor.
            IGeoProcessor2 gp = new GeoProcessorClass();
            gp.SetEnvironmentValue("workspace", gdb);
            gp.OverwriteOutput = true;

            // Create an IVariantArray to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add(layer);
            
            // Execute the tool.
            count = Convert.ToInt32(gp.Execute("GetCount_management", parameters, null).GetOutput(0).GetAsText());
            //textBox1.Text = str;
            //count = Convert.ToInt32(gp.Execute("GetCount_management", parameters, null).GetOutput(0));
        }

        private void AnalyseCities(int layerID, string Buffers, string Cities, string LinesLocalities)
        {
            double Emission_main = 0;
            double Emission_other = 0;
            string Wind = "";
            double Speed = 0;
            int Fid_main = 0;
            int Fid_other;
            var MxDoc = ArcMap.Application.Document as IMxDocument;
            IActiveView activeView = MxDoc.ActiveView;

            // Get the TOC
            IContentsView IContentsView = MxDoc.CurrentContentsView;
            var layer = activeView.FocusMap.Layer[layerID] as FeatureLayer;
            //textBox1.Text = layer.FeatureClass.Fields.Field[layer.FeatureClass.Fields.FieldCount-1].Name;
            IClass Towns = layer.FeatureClass as IClass;
            IFeatureClass TownsFeats = layer.FeatureClass;
            IFeatureDataset data = TownsFeats.FeatureDataset;
            int count = 0;
            CountRows(Buffers, out count);

            for (int i = 1; i <= count; i++)
            {
                SelectMapFeaturesByAttributeQuery(activeView, layer, "OBJECTID = " + i);

                IFeatureSelection TownsFeatsSelect = layer as IFeatureSelection;
                ISelectionSet TownsFeatsSelectSet = TownsFeatsSelect.SelectionSet;
                if (TownsFeatsSelectSet.Count > 0)
                {
                    IQueryFilter queryFilter = new QueryFilterClass();
                    queryFilter.WhereClause = "OBJECTID > 0"; //"U" = domain value for Utility
                    ICursor Cursor;
                    TownsFeatsSelectSet.Search(queryFilter, false, out Cursor);
                    // Get the first feature
                    IFeatureCursor featureCursor = Cursor as IFeatureCursor;
                    ESRI.ArcGIS.Geodatabase.IFeature feature = featureCursor.NextFeature();
                    while (feature != null)
                    {
                        int num = Towns.FindField("Emissions");
                        Emission_main = Convert.ToDouble(feature.get_Value(num));
                        num = Towns.FindField("SavedID");
                        Fid_main = Convert.ToInt32(feature.get_Value(num));
                        num = Towns.FindField("Wind");
                        Wind = Convert.ToString(feature.get_Value(num));
                        num = Towns.FindField("Speed");
                        Speed = Convert.ToDouble(feature.get_Value(num));
                        feature = featureCursor.NextFeature();
                    }
                }

                SelectLayerByLocation(Cities, Buffers);
                SelectLayerByLocation(LinesLocalities, Buffers);

                var layer2 = activeView.FocusMap.Layer[6] as FeatureLayer;
                //textBox1.Text = layer.FeatureClass.Fields.Field[layer.FeatureClass.Fields.FieldCount-1].Name;
                IClass Towns2 = layer2.FeatureClass as IClass;
                IFeatureClass TownsFeats2 = layer2.FeatureClass;

                //SelectMapFeaturesByAttributeQuery(activeView, layer2, "OBJECTID = 3");
                IFeatureSelection TownsFeatsSelect2 = layer2 as IFeatureSelection;
                ISelectionSet TownsFeatsSelectSet2 = TownsFeatsSelect2.SelectionSet;
                if (TownsFeatsSelectSet2.Count > 0)
                {
                    IQueryFilter queryFilter = new QueryFilterClass();
                    queryFilter.WhereClause = "FID > 0"; //"U" = domain value for Utility
                    ICursor Cursor;
                    TownsFeatsSelectSet2.Search(queryFilter, false, out Cursor);
                    // Get the first feature
                    IFeatureCursor featureCursor = Cursor as IFeatureCursor;
                    ESRI.ArcGIS.Geodatabase.IFeature feature = featureCursor.NextFeature();
                    while (feature != null)
                    {
                        if (Convert.ToInt32(feature.get_Value(0)) != Fid_main)
                        {
                            int num = Towns.FindField("Emissions"); //фильтрация по фиду !!!
                            Emission_other = Convert.ToDouble(feature.get_Value(num));
                            num = Towns.FindField("SavedID");
                            Fid_other = Convert.ToInt32(feature.get_Value(num));
                            //textBox1.Text += "\r\nEmiss: " + Emission_other.ToString() + " " + Fid_other;

                            var layer3 = activeView.FocusMap.Layer[0] as FeatureLayer;
                            IClass Towns3 = layer3.FeatureClass as IClass;
                            IFeatureClass TownsFeats3 = layer3.FeatureClass;

                            SelectMapFeaturesByAttributeQuery(activeView, layer3, "SavedID = " + Fid_main + " or SavedID = " + Fid_other);
                            IFeatureSelection TownsFeatsSelect3 = layer3 as IFeatureSelection;
                            ISelectionSet TownsFeatsSelectSet3 = TownsFeatsSelect3.SelectionSet;
                            if (TownsFeatsSelectSet3.Count > 0)
                            {
                                IQueryFilter queryFilter2 = new QueryFilterClass();
                                queryFilter2.WhereClause = "OBJECTID >= 0"; //"U" = domain value for Utility
                                ICursor Cursor2;
                                TownsFeatsSelectSet3.Search(queryFilter2, false, out Cursor2);
                                // Get the first feature
                                IFeatureCursor featureCursor2 = Cursor2 as IFeatureCursor;
                                ESRI.ArcGIS.Geodatabase.IFeature feature2 = featureCursor2.NextFeature();
                                IPoint point = new PointClass();
                                IPoint point2 = new PointClass();
                                while (feature2 != null)
                                {
                                    num = Towns.FindField("SavedID");
                                    double x = feature2.Shape.Envelope.XMax;
                                    double y = feature2.Shape.Envelope.YMax;
                                    //textBox1.Text += "\r\n" + feature2.get_Value(0) + " " + x.ToString() + " " + y.ToString();
                                    int idf = Convert.ToInt32(feature2.get_Value(num));
                                    if (Convert.ToInt32(feature2.get_Value(num)) == Fid_main)
                                    {
                                        point.PutCoords(x, y);
                                    }
                                    else
                                    {
                                        point2.PutCoords(x, y);
                                    }
                                    feature2 = featureCursor2.NextFeature();
                                }
                                double Distance = GetDistanceFromTwoPoints(point, point2, ESRI.ArcGIS.Geometry.esriSRGeoCSType.esriSRGeoCS_WGS1984);
                                double Azimuth = GetAzimuthFromTwoPoints(point, point2, ESRI.ArcGIS.Geometry.esriSRGeoCSType.esriSRGeoCS_WGS1984);
                                //textBox1.Text += "\r\nDistance: " + Distance.ToString();
                                //textBox1.Text += "\r\nAzimuth: " + Azimuth.ToString();
                                //Distance /= 1000;
                                string[] Rose = Wind.Split(' ');
                                double Rhumb = 0;
                                if (Azimuth > -22 && Azimuth <= 22)
                                    Rhumb = Convert.ToDouble(Rose[0]);
                                else if (Azimuth > 22 && Azimuth <= 67)
                                    Rhumb = Convert.ToDouble(Rose[1]);
                                else if (Azimuth > 67 && Azimuth <= 112)
                                    Rhumb = Convert.ToDouble(Rose[2]);
                                else if (Azimuth > 112 && Azimuth <= 157)
                                    Rhumb = Convert.ToDouble(Rose[3]);
                                else if ((Azimuth > 157 && Azimuth < 180) || Azimuth < -157)
                                    Rhumb = Convert.ToDouble(Rose[4]);
                                else if (Azimuth <= -22 && Azimuth >= -67)
                                    Rhumb = Convert.ToDouble(Rose[7]);
                                else if (Azimuth < -67 && Azimuth >= -112)
                                    Rhumb = Convert.ToDouble(Rose[6]);
                                else if (Azimuth < -112 && Azimuth >= -157)
                                    Rhumb = Convert.ToDouble(Rose[5]);
                                double koef = (Rhumb / 100) * Speed * 1000;
                                double model = (Emission_main * Emission_other) / Math.Pow(Distance + koef, 2);
                                //textBox1.Text += "\r\nRhumb: " + Rhumb;
                                //textBox1.Text += "\r\nModel: " + model;


                                var layer4 = activeView.FocusMap.Layer[1] as FeatureLayer;
                                IClass Towns4 = layer4.FeatureClass as IClass;
                                IFeatureClass TownsFeats4 = layer4.FeatureClass;
                                /*IField Field = new FieldClass();
                                IFieldEdit fieldEdit = (IFieldEdit)Field;
                                fieldEdit.Name_2 = "Model";
                                fieldEdit.Type_2 = esriFieldType.esriFieldTypeDouble;
                                Towns4.AddField(Field);*/
                                //SelectMapFeaturesByAttributeQuery(activeView, layer3, "SavedID = " + Fid_main + " or SavedID = " + Fid_other);
                                IFeatureSelection TownsFeatsSelect4 = layer4 as IFeatureSelection;
                                ISelectionSet TownsFeatsSelectSet4 = TownsFeatsSelect4.SelectionSet;
                                if (TownsFeatsSelectSet4.Count > 0)
                                {
                                    IQueryFilter queryFilter3 = new QueryFilterClass();
                                    queryFilter3.WhereClause = "SavedID = " + Fid_other; //"U" = domain value for Utility
                                    ICursor Cursor3;
                                    TownsFeatsSelectSet4.Search(queryFilter3, false, out Cursor3);
                                    // Get the first feature
                                    IFeatureCursor featureCursor3 = Cursor3 as IFeatureCursor;
                                    ESRI.ArcGIS.Geodatabase.IFeature feature3 = featureCursor3.NextFeature();
                                    while (feature3 != null)
                                    {
                                        num = TownsFeats4.FindField("Model");
                                        double tmp = Convert.ToDouble(feature3.get_Value(num));
                                        tmp += model;
                                        feature3.set_Value(num, tmp);
                                        feature3.Store();
                                        feature3 = featureCursor3.NextFeature();
                                    }
                                }
                            }

                        }
                        feature = featureCursor.NextFeature();

                        //foreach (string str in Rose)
                        //    textBox1.Text += "\r\n" + str;
                    }
                }
            }
        }

        private void MainAnalyze()
        {
            AnalyseCities(2, "Less20CitiesKm_buff", "Analyzed_localities", "Localities_line_split");
            AnalyseCities(3, "More20CitiesKm_buff", "Analyzed_localities", "Localities_line_split");
        }  

        #region "Get Distance from Two Points"

///<summary>Get the geodetically correct Rhumb Line distance between two points.</summary>
/// 
///<param name="fromPoint">An IPoint interface that is the start (or from) location</param>
///<param name="toPoint">An IPoint interface that is the end (or to) location</param>
///<param name="spatialReference">An esriSRGeoCSType enum that is a predefined geographic coordinate system. Example: ESRI.ArcGIS.Geometry.esriSRGeoCSType.esriSRGeoCS_NAD1983</param>
///  
///<returns>A System.Double representing true distance</returns>
///  
///<remarks></remarks>
public System.Double GetDistanceFromTwoPoints(IPoint fromPoint, IPoint toPoint, esriSRGeoCSType spatialReference)
{

  // Define the spatial reference of the rhumb line. 
  ISpatialReferenceFactory2 spatialReferenceFactory2 = new SpatialReferenceEnvironmentClass();
  ISpatialReference2 spatialReference2 = (ISpatialReference2)spatialReferenceFactory2.CreateSpatialReference((System.Int16)spatialReference);

  // Initialize the MeasurementTool and define the properties of the line.
  // These properties include the line type, which is a rhumb line in this case, and the 
  // spatial reference of the line.   
  IMeasurementTool measurementTool = new MeasurementToolClass();
  measurementTool.SpecialGeolineType = cjmtkSGType.cjmtkSGTRhumbLine;
  measurementTool.SpecialSpatialReference = spatialReference2;

  // Determine the distance and azimuth of the rhumb line based on the start and end point coordinates.   
  measurementTool.ConstructByPoints(fromPoint, toPoint);

  // Return the Distance. 
  return measurementTool.Distance;
}
#endregion

       

        private void SplitPolygons()
        {
            // Create the geoprocessor.
            IGeoProcessor2 gp = new GeoProcessorClass();
            gp.SetEnvironmentValue("workspace", gdb);
            gp.OverwriteOutput = true;

            // Create an IVariantArray to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add("Analyzed_localities");
            //parameters.Add("intersect");
            parameters.Add("Localities_line");
            

            // Execute the tool.
            gp.Execute("FeatureToLine_management", parameters, null);

            parameters.RemoveAll();

            parameters.Add("Localities_line");
            parameters.Add("Localities_line_split");
            gp.Execute("SplitLine_management", parameters, null);

            ILayer layer = ArcMap.Document.FocusMap.Layer[2];
            ArcMap.Document.FocusMap.DeleteLayer(layer);
        }

        private void LinesToPoints()
        {
            var MxDoc = ArcMap.Application.Document as IMxDocument;
            IActiveView activeView = MxDoc.ActiveView;

            // Get the TOC
            IContentsView IContentsView = MxDoc.CurrentContentsView;
            var layer = activeView.FocusMap.Layer[1] as FeatureLayer;
            ClearSelectedMapFeatures(activeView, layer);

            IGeoProcessor2 gp = new GeoProcessorClass();
            gp.SetEnvironmentValue("workspace", gdb);
            gp.OverwriteOutput = true;

            // Create an IVariantArray to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add("Localities_line_split");
            //parameters.Add("intersect");
            parameters.Add("Localities_ready");


            // Execute the tool.
            gp.Execute("FeatureVerticesToPoints_management", parameters, null);
            //FeatureVerticesToPoints_management
            parameters.RemoveAll();

            parameters.Add("Localities_ready");
            parameters.Add("Model");
            parameters.Add("Result");
            gp.Execute("NaturalNeighbor_sa", parameters, null);

            parameters.RemoveAll();
            if (checkBox1.Checked == true)
            {
                parameters.Add("Localities_ready");
                parameters.Add("Emissions");
                parameters.Add("Result_Without_Model");
                gp.Execute("NaturalNeighbor_sa", parameters, null);
            }
            //var resRaster = gp.Execute("NaturalNeighbor", parameters, null);
            //resRaster.ReturnValue
            //resRaster.
            /*activeView.FocusMap.Layer[1].Visible = false;
            activeView.FocusMap.Layer[0].Visible = false;
            activeView.FocusMap.Layer[2].Visible = false;
            activeView.FocusMap.Layer[3].Visible = false;
            activeView.FocusMap.Layer[4].Visible = false;
            activeView.FocusMap.Layer[8].Visible = false;
            activeView.FocusMap.Layer[9].Visible = false;*/
            /*ToggleVisibleOfCompositeLayer(activeView, 0);
            ToggleVisibleOfCompositeLayer(activeView, 1);
            ToggleVisibleOfCompositeLayer(activeView, 2);
            ToggleVisibleOfCompositeLayer(activeView, 3);
            ToggleVisibleOfCompositeLayer(activeView, 4);
            ToggleVisibleOfCompositeLayer(activeView, 8);
            ToggleVisibleOfCompositeLayer(activeView, 9);*/
        }

        private void button14_Click(object sender, EventArgs e)
        {
            bool error = false;
            try
            {
                Convert.ToInt32(textBox4.Text);
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);//"Неверное значение порога площади"
                error = true;
            }
            if (error == false)
            {
                SplitSettlements();
                CreateBuffers();
                CreateCentroids();
                SplitPolygons();
                MainAnalyze();
                LinesToPoints();
                ClassifyFeatures("Model",ref breaks);
                ClassifyFeatures("Emissions", ref breaks2);
                groupBox1.Enabled = true;
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = null;
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox7.Text = System.IO.Path.GetDirectoryName(openFileDialog1.FileName);
                gdb = textBox7.Text;
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            var MxDoc = ArcMap.Application.Document as IMxDocument;
            IActiveView activeView = MxDoc.ActiveView;

            // Get the TOC
            IContentsView IContentsView = MxDoc.CurrentContentsView;
            //var layer = activeView.FocusMap.Layer[1] as FeatureLayer;
            //ClearSelectedMapFeatures(activeView, layer);

            IGeoProcessor2 gp = new GeoProcessorClass();
            gp.SetEnvironmentValue("workspace", gdb);
            gp.OverwriteOutput = true;

            // Create an IVariantArray to hold the parameter values.
            IVariantArray parameters = new VarArrayClass();

            // Populate the variant array with parameter values.
            parameters.Add("Result");
            //parameters.Add("intersect");
            parameters.Add("Донецк");
            //parameters.Add("INSIDE");
            parameters.Add("Cutted_Result");

            
            
            gp.Execute("ExtractByMask_sa", parameters, null);
            parameters.RemoveAll();

            string BreaksString = "";
            for (int i = 0; i < breaks.Count()-1; i++)
            {
                BreaksString += breaks[i].ToString() + " " + breaks[i + 1] + " " + (i+1) + ";";//+ breaks[i] + " - " + breaks[i+1]
            }
            textBox1.Text = BreaksString;
            string parcedBreaks = BreaksString.Replace(",", ".");
            parameters.Add("Cutted_Result");
            //parameters.Add("intersect");
            parameters.Add("Value");
            //parameters.Add("INSIDE");
            parameters.Add(BreaksString);
            parameters.Add("ReclassifiedResult");
            //parameters.Add("DATA");
            try
            {
                gp.Execute("Reclassify_sa", parameters, null);
                //textBox1.Text = mesg.Messages.Element[0].ToString();
            }
            catch (COMException ex) { 
                //textBox1.Text = ex.Message; 
            }
            if (checkBox1.Checked == true)
            {
                parameters.RemoveAll();

                // Populate the variant array with parameter values.
                parameters.Add("Result_Without_Model");
                //parameters.Add("intersect");
                parameters.Add("Донецк");
                //parameters.Add("INSIDE");
                parameters.Add("Cutted_Result_Without_Model");



                gp.Execute("ExtractByMask_sa", parameters, null);
                parameters.RemoveAll();

                BreaksString = "";
                for (int i = 0; i < breaks2.Count() - 1; i++)
                {
                    BreaksString += breaks2[i].ToString() + " " + breaks2[i + 1] + " " + (i + 1) + ";";//+ breaks[i] + " - " + breaks[i+1]
                }
                parameters.Add("Cutted_Result_Without_Model");
                //parameters.Add("intersect");
                parameters.Add("Value");
                //parameters.Add("INSIDE");
                parameters.Add(BreaksString);
                parameters.Add("ReclassifiedResult_Without_Model");
                //parameters.Add("DATA");
                try
                {
                    gp.Execute("Reclassify_sa", parameters, null);
                    //textBox1.Text = mesg.Messages.Element[0].ToString();
                }
                catch (COMException ex)
                {
                    //textBox1.Text = ex.Message; 
                }
            }
            
        }

        private void button17_Click(object sender, EventArgs e)
        {
            AddShapefileUsingOpenFileDialog(ArcMap.Document.ActiveView);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            AboutForm f = new AboutForm();
            f.ShowDialog();
        }

    }
}
