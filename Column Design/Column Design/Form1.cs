using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Oasys.AdSec;
using Oasys.AdSec.DesignCode;
using Oasys.AdSec.StandardMaterials;
using Oasys.AdSec.Reinforcement;
using Oasys.AdSec.Reinforcement.Groups;
using Oasys.AdSec.Reinforcement.Layers;
using Oasys.Profiles;
using UnitsNet;
using Oasys.Units;
using System.Drawing;
using System.IO;

namespace Column_Design
{
    public struct ColumnInputData
    {
        public string columnLabel;
        public string story;
        public double story_elevation;
        public string location;
        public double fck;
        public double depth;
        public double width;
        public double diameter;
        public double P;
        public double MMajor;
        public double MMinor;
        public double rebarPtEtabs;
        public double maxVAlongX;
        public double maxVAlongY;
        public double length;
        public string governingCombo;
    }

    public struct ColumnShearData
    {
        public string columnLabel;
        public string story;
        public string location;
        public double maxVAlongX;
        public double maxVAlongY;
    }

    public struct ColumnForceData
    {
        public string columnLabel;
        public string story;
        public double P;
        public double MMajor;
        public double MMinor;
        public double VAlongX;
        public double VAlongY;
        public string outputCase;
        public double station;
    }

    public struct ColumnDesignData
    {
        public double arrangementDepth;
        public double arrangementWidth;
        public int cornerBarDia;
        public int cornerBarCountAlongDepth;
        public int cornerBarCountAlongWidth;
        public int centreBarDia;
        public int centreBarCountAlongDepth;
        public int centreBarCountAlongWidth;
        public double spacingAlongDepth;
        public double spacingAlongWidth;
        public double ptProvided;
        public double tensile_rebar;
        public String rebarDescription;
    }

    public struct DesignDataToExcel
    {
        public ISection designSection;
        public ColumnInputData inputData;
        public ColumnDesignData DesignData;
        public double loadUtilisation;
        public double momentRatio;
        public double factoredShearForceAlongY;
        public double factoredShearForceAlongX;
        public double effectiveDepth;
        public double effectiveWidth;
        public double clearCover;
        public double tauVAlongX;
        public double tauVAlongY;
        public double tauC;
        public double tauCMax;
        public double longitudinalFy;
        public double linkFy;
        public double linkDia;
        public int legsAlongY;
        public int legsAlongX;
        public double VusAlongX;
        public double VusAlongY;
        public double asvProvidedAlongX;
        public double asvProvidedAlongY;
        public double nonConfiningSpacingOne;
        public double nonConfiningSpacingTwo;
        public double nonConfiningSpacingProvided;
        public double minNonConfiningAsvRequired;
        public double confinfingSpacingOne;
        public double confinfingSpacingTwo;
        public double confinfingSpacingThree;
        public double maxConfiningSpacingRequired;
        public double confiningSpacingProvided;
        public double Ag;
        public double Ak;
        public double h;
        public double AshOne;
        public double AshTwo;
        public double minAshRequired;
        public double AshProvided;
        public List<ILoad> checkedMajorEccentricLoads;
        public List<ILoad> checkedMinorEccentricLoads;
        public List<ILoad> checkedOtherLoads;
    }

    public partial class ColumnDesignForm : Form
    {
        public List<ColumnInputData> columnInputData = new List<ColumnInputData>();
        public List<ColumnShearData> columnShearData = new List<ColumnShearData>();
        public List<ColumnForceData> columnForceData = new List<ColumnForceData>();
        public List<DesignDataToExcel> designDataToExcel = new List<DesignDataToExcel>();
        public List<string> uniqueStories = new List<String>();
        public List<string> uniqueColumnLabels = new List<String>();

        public IAdSec adsecApp = IAdSec.Create(IS456.Edition_2000);
        public ISection section = null;
        public ISolution solution = null;
        public List<ILineGroup> lineGroups = new List<ILineGroup>();

        public static class ResultTableColumnIndex
        {
            public static int column_label = 0;
            public static int story = 1;
            public static int story_elevation = 2;
            public static int location = 3;
            public static int fck = 4;
            public static int depth = 5;
            public static int width = 6;
            public static int length = 7;
            public static int load_combination = 8;
            public static int p = 9;
            public static int m33 = 10;
            public static int m22 = 11;
            public static int load_util = 12;
            public static int m_ratio = 13;
            public static int etabs_req = 14;
            public static int adsec_pro = 15;
            public static int tensile_rebar = 16;
            public static int rebar_description = 17;
            public static int vy = 18;
            public static int vx = 19;
            public static int asvy = 20;
            public static int asvx = 21;
            public static int tauc = 22;
            public static int tauvy = 23;
            public static int tauvx = 24;
            public static int max_nc = 25;
            public static int max_c = 26;
            public static int ash_req = 27;
        }

        public ColumnDesignForm()
        {
            InitializeComponent();
            columnsToShow.Items.Add("Column Label");
            columnsToShow.Items.Add("Story");
            columnsToShow.Items.Add("Story Elevation");
            columnsToShow.Items.Add("Location");
            columnsToShow.Items.Add("Fck");
            columnsToShow.Items.Add("Depth");
            columnsToShow.Items.Add("Width");
            columnsToShow.Items.Add("Length");
            columnsToShow.Items.Add("Load Combination");
            columnsToShow.Items.Add("P");
            columnsToShow.Items.Add("M33");
            columnsToShow.Items.Add("M22");
            columnsToShow.Items.Add("Load Utilisation");
            columnsToShow.Items.Add("M/Mu");
            columnsToShow.Items.Add("Rebar Required - ETABS");
            columnsToShow.Items.Add("Rebar Provided - AdSec");
            columnsToShow.Items.Add("Tensile Rebar");
            columnsToShow.Items.Add("Rebar Description");
            columnsToShow.Items.Add("Vy");
            columnsToShow.Items.Add("Vx");
            columnsToShow.Items.Add("Asvy");
            columnsToShow.Items.Add("Asvx");
            columnsToShow.Items.Add("TauC");
            columnsToShow.Items.Add("TauVY");
            columnsToShow.Items.Add("TauVX");
            columnsToShow.Items.Add("Max NC Spacing");
            columnsToShow.Items.Add("Max C Spacing");
            columnsToShow.Items.Add("Ash Required");
            for(int i=0; i< this.columnsToShow.Items.Count; i++)
            {
                if(columnsToShow.Items[i].ToString() != "Story Elevation")
                {
                    columnsToShow.SetItemChecked(i, true);
                }
                else
                {
                    columnsToShow.SetItemChecked(i, false);
                }
            }
            ResultsTable.Columns[ResultTableColumnIndex.story_elevation].Visible = false;
            Include16mmBar.Checked = true;
            Include20mmBar.Checked = true;
            Include25mmBar.Checked = true;
            Include32mmBar.Checked = true;
            UniformRebar.Checked = true;

            IncreaseTauC.Checked = true;
        }

        public void ExtractEtabsForceTables(Excel.Workbook excelWorkBook)
        {
            string @ModelPath = this.modelPath.Text;

            ETABSv1.cOAPI myETABSObject = null;
            int ret = 0;
            ETABSv1.cHelper myHelper;
            try
            {
                myHelper = new ETABSv1.Helper();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot create an instance of the Helper object");
                return;
            }
            try
            {
                //create ETABS object
                myETABSObject = myHelper.CreateObjectProgID("CSI.ETABS.API.ETABSObject");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot start a new instance of the program.");
                return;
            }
            ret = myETABSObject.ApplicationStart();
            ETABSv1.cSapModel mySapModel = myETABSObject.SapModel;
            ret = mySapModel.File.OpenFile(ModelPath);
            mySapModel.SetPresentUnits(ETABSv1.eUnits.kN_m_C);

            Excel.Worksheet readMe = null;
            Excel.Worksheet designInputSheet = null;
            Excel.Worksheet columnShearSheet = null;
            Excel.Worksheet columnForcesSheet = null;
            foreach (Excel.Worksheet sheet in excelWorkBook.Sheets)
            {
                if (sheet.Name == "Concrete Column PMM Envelope")
                {
                    if(sheet.UsedRange.Rows.Count > 3)
                    {
                        sheet.Range[
                            sheet.Cells[4,1], sheet.Cells[sheet.UsedRange.Rows.Count, sheet.UsedRange.Columns.Count]].
                            ClearContents();
                    }

                    //sheet.UsedRange.ClearContents();
                    designInputSheet = sheet;
                }
                else if (sheet.Name == "Concrete Column Shear Envelope")
                {
                    if (sheet.UsedRange.Rows.Count > 3)
                    {
                        sheet.Range[
                            sheet.Cells[4, 1], sheet.Cells[sheet.UsedRange.Rows.Count, sheet.UsedRange.Columns.Count]].
                            ClearContents();
                    }

                    //sheet.UsedRange.ClearContents();
                    columnShearSheet = sheet;
                }
                else if (sheet.Name == "Column Forces")
                {
                    if (sheet.UsedRange.Rows.Count > 3)
                    {
                        sheet.Range[
                            sheet.Cells[4, 1], sheet.Cells[sheet.UsedRange.Rows.Count, sheet.UsedRange.Columns.Count]].
                            ClearContents();
                    }

                    //sheet.UsedRange.ClearContents();
                    columnForcesSheet = sheet;
                }
                else if (sheet.Name == "Read Me")
                {
                    readMe = sheet;
                }
            }

            string[] PMMFieldKeyList = null;
            int PMMTableVersion = 0;
            string[] PMMFieldsKeysIncluded = null;
            int PMMNumberRecords = 0;
            string[] PMMTableData = null;
            ret = mySapModel.DatabaseTables.GetTableForDisplayArray(
                "Concrete Column PMM Envelope - IS 456-2000", 
                ref PMMFieldKeyList, 
                this.ColumnGroupName.Text, 
                ref PMMTableVersion, 
                ref PMMFieldsKeysIncluded, 
                ref PMMNumberRecords, 
                ref PMMTableData);

            string[] SFFieldKeyList = null;
            int SFTableVersion = 0;
            string[] SFFieldsKeysIncluded = null;
            int SFNumberRecords = 0;
            string[] SFTableData = null;
            ret = mySapModel.DatabaseTables.GetTableForDisplayArray(
                "Concrete Column Shear Envelope - IS 456-2000",
                ref SFFieldKeyList,
                this.ColumnGroupName.Text,
                ref SFTableVersion,
                ref SFFieldsKeysIncluded,
                ref SFNumberRecords,
                ref SFTableData);

            string[] CFFieldKeyList = null;
            int CFTableVersion = 0;
            string[] CFFieldsKeysIncluded = null;
            int CFNumberRecords = 0;
            string[] CFTableData = null;
            ret = mySapModel.DatabaseTables.GetTableForDisplayArray(
                "Element Forces - Columns",
                ref CFFieldKeyList,
                this.ColumnGroupName.Text,
                ref CFTableVersion,
                ref CFFieldsKeysIncluded,
                ref CFNumberRecords,
                ref CFTableData);

            int NumberNames = 0;
            string[] MyName = null;
	        string[] PropName = null;
	        string[] StoryName = null;
	        string[] PointName1 = null;
	        string[] PointName2 = null;
	        double[] Point1X = null;
	        double[] Point1Y = null;
	        double[] Point1Z = null;
	        double[] Point2X = null;
	        double[] Point2Y = null;
	        double[] Point2Z = null;
	        double[] Angle = null;
	        double[] Offset1X = null;
	        double[] Offset2X = null;
	        double[] Offset1Y = null;
	        double[] Offset2Y = null;
	        double[] Offset1Z = null;
	        double[] Offset2Z = null;
	        int[] CardinalPoint = null;

            ret = mySapModel.FrameObj.GetAllFrames(
                ref NumberNames,
                ref MyName,
                ref PropName,
                ref StoryName,
                ref PointName1,
                ref PointName2,
                ref Point1X,
                ref Point1Y,
                ref Point1Z,
                ref Point2X,
                ref Point2Y,
                ref Point2Z,
                ref Angle,
                ref Offset1X,
                ref Offset2X,
                ref Offset1Y,
                ref Offset2Y,
                ref Offset1Z,
                ref Offset2Z,
                ref CardinalPoint,
                "Global"
                );
            Dictionary<string, Dictionary<string, double>> frameLengths = new Dictionary<string, Dictionary<string, double>>();
            for(int i =0; i< MyName.Length; i++)
            {
                string givenLabel = null;
                string givenStory = null;
                mySapModel.FrameObj.GetLabelFromName(MyName[i], ref givenLabel, ref givenStory);
                double length = 1000 * Math.Sqrt(
                    Math.Pow(Point2X[i] - Point1X[i], 2) + 
                    Math.Pow(Point2Y[i] - Point1Y[i], 2) + 
                    Math.Pow(Point2Z[i] - Point1Z[i], 2));
                if (!frameLengths.ContainsKey(givenLabel))
                {
                    Dictionary<string, double> storyAndLength = new Dictionary<string, double>();
                    storyAndLength.Add(givenStory, length);
                    frameLengths.Add(givenLabel, storyAndLength);
                }
                else
                {
                    frameLengths[givenLabel].Add(givenStory, length);
                }
            }

            int sectionPropertyIndex = 0;
            int storyIndex = 0;
            int labelIndex = 0;
            for (int i=0; i< PMMFieldsKeysIncluded.Length; i++)
            {
                if (PMMFieldsKeysIncluded[i] == "DesignSect")
                {
                    sectionPropertyIndex = i;
                }
                else if (PMMFieldsKeysIncluded[i] == "Story")
                {
                    storyIndex = i;
                }
                else if (PMMFieldsKeysIncluded[i] == "Label")
                {
                    labelIndex = i;
                }

                //designInputSheet.Cells[2, i+1].Value2 = PMMFieldsKeysIncluded[i];
            }

            int row = 3;
            int column = 1;
            bool dimHeadingsAdded = false;
            for (int i=0; i< PMMTableData.Length; i++)
            {
                if(i % PMMFieldsKeysIncluded.Length == 0)
                {
                    row++;
                    column = 1;
                }
                if (i % PMMFieldsKeysIncluded.Length == sectionPropertyIndex)
                {
                    //if(!dimHeadingsAdded)
                    //{
                    //    designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 1].Value2 = "Depth";
                    //    designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 2].Value2 = "Width";
                    //    designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 3].Value2 = "Diameter";
                    //    designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 4].Value2 = "fck";
                    //    designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 5].Value2 = "Length";
                    //    dimHeadingsAdded = true;
                    //}

                    string FileName = "";
                    string MatProp = "";
                    double depth = 0;
                    double width = 0;
                    double circleDiameter = 0;
                    int Color = 0;
                    string Notes = "";
                    string GUID = "";
                    ret = 0;
                    ret = mySapModel.PropFrame.GetRectangle(
                        PMMTableData[i], ref FileName, ref MatProp, ref depth, ref width, ref Color, ref Notes, ref GUID);
                    if(ret == 1)
                    {
                        depth = 0;
                        width = 0;
                        string circleFileName = "";
                        int circleColor = 0;
                        string circleNotes = "";
                        string circleGUID = "";
                        ret = mySapModel.PropFrame.GetCircle(
                            PMMTableData[i], ref circleFileName, ref MatProp, ref circleDiameter, ref circleColor, ref circleNotes, ref circleGUID);
                    }

                    double fc = 0;
                    bool isLightweight = false;
                    double fcsfactor = 0;
                    int SSType = 0;
                    int SSHysType = 0;
                    double StrainAtfc = 0;
                    double StrainUltimate = 0;
                    double FrictionAngle = 0;
                    double DilatationalAngle = 0;
                    ret = mySapModel.PropMaterial.GetOConcrete(
                        MatProp, ref fc, ref isLightweight, ref fcsfactor, ref SSType, ref SSHysType,
                        ref StrainAtfc, ref StrainUltimate, ref FrictionAngle, ref DilatationalAngle);

                    designInputSheet.Cells[row, PMMFieldsKeysIncluded.Length + 1].Value2 = depth * 1000;
                    designInputSheet.Cells[row, PMMFieldsKeysIncluded.Length + 2].Value2 = width * 1000;
                    designInputSheet.Cells[row, PMMFieldsKeysIncluded.Length + 3].Value2 = circleDiameter * 1000;
                    designInputSheet.Cells[row, PMMFieldsKeysIncluded.Length + 4].Value2 = fc / 1000;
                    designInputSheet.Cells[row, PMMFieldsKeysIncluded.Length + 5].Value2 =
                        frameLengths[PMMTableData[i - (sectionPropertyIndex - labelIndex)]][PMMTableData[i - (sectionPropertyIndex - storyIndex)]];
                }
                designInputSheet.Cells[row, column].Value2 = PMMTableData[i];
                column++;
            }

            //for (int i = 0; i < SFFieldsKeysIncluded.Length; i++)
            //{
            //    columnShearSheet.Cells[2, i + 1].Value2 = SFFieldsKeysIncluded[i];
            //}

            row = 3;
            column = 1;
            for (int i = 0; i < SFTableData.Length; i++)
            {
                if (i % SFFieldsKeysIncluded.Length == 0)
                {
                    row++;
                    column = 1;
                }

                columnShearSheet.Cells[row, column].Value2 = SFTableData[i];
                column++;
            }

            //for (int i = 0; i < CFFieldsKeysIncluded.Length; i++)
            //{
            //    columnForcesSheet.Cells[2, i + 1].Value2 = CFFieldsKeysIncluded[i];
            //}

            row = 3;
            column = 1;
            for (int i = 0; i < CFTableData.Length; i++)
            {
                if (i % CFFieldsKeysIncluded.Length == 0)
                {
                    row++;
                    column = 1;
                }

                columnForcesSheet.Cells[row, column].Value2 = CFTableData[i];
                column++;
            }

            //Close ETABS
            myETABSObject.ApplicationExit(false);
            excelWorkBook.Save();

            //Clean up variables
            mySapModel = null;
            myETABSObject = null;
        }

        private void extractInput_Click(object sender, EventArgs e)
        {
            ResultsTable.Rows.Clear();
            columnsToDesign.Items.Clear();
            DesignStories.Items.Clear();

            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkBook = null;
            Excel.Worksheet designInputSheet = null;
            Excel.Worksheet columnShearSheet = null;
            Excel.Worksheet columnForcesSheet = null;

            excelApplication = new Excel.Application();
            excelApplication.Visible = true;
            string @excelPath = designInput.Text;
            excelWorkBook = excelApplication.Workbooks.Open(excelPath);
            designInputSheet = excelWorkBook.Worksheets["Concrete Column PMM Envelope"];
            columnShearSheet = excelWorkBook.Worksheets["Concrete Column Shear Envelope"];
            columnForcesSheet = excelWorkBook.Worksheets["Column Forces"];

            object[,] columnPMMDataObject = designInputSheet.UsedRange.Value2;
            List<string> columnsLabels = new List<string>();
            int columnIndex = 0;
            int storyIndex = 0;
            int locationIndex = 0;
            int pIndex = 0;
            int mMajorIndex = 0;
            int mMinorIndex = 0;
            int ptRequiredIndex = 0;
            int depthIndex = 0;
            int widthIndex = 0;
            int diameterIndex = 0;
            int fckIndex = 0;
            int lengthIndex = 0;
            int govComboIndex = 0;
            int storyElevationIndex = 0;
            for (int column = 1; column <= designInputSheet.UsedRange.Columns.Count; column++)
            {
                if (column > 15)
                {
                    if (designInputSheet.Cells[2, column].Value2 == null)
                    {
                        break;
                    }
                }

                if (designInputSheet.Cells[2, column].Value2 == "Label")
                {
                    columnIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "Story")
                {
                    storyIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "Location")
                {
                    locationIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "P")
                {
                    pIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "MMajor")
                {
                    mMajorIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "MMinor")
                {
                    mMinorIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "RatioRebar")
                {
                    ptRequiredIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "Depth")
                {
                    depthIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "Width")
                {
                    widthIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "Diameter")
                {
                    diameterIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "fck")
                {
                    fckIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "Length")
                {
                    lengthIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "PMMCombo")
                {
                    govComboIndex = column;
                }
                else if (designInputSheet.Cells[2, column].Value2 == "Story Elevation")
                {
                    storyElevationIndex = column;
                }
            }

            for (int row = 4; row <= designInputSheet.UsedRange.Rows.Count; row++)
            {
                var columnLabel = columnPMMDataObject[row, columnIndex];
                if (columnLabel == null)
                {
                    break;
                }

                columnsLabels.Add((string)columnLabel);

                if (!uniqueStories.Contains((string)columnPMMDataObject[row, storyIndex]))
                {
                    uniqueStories.Add((string)columnPMMDataObject[row, storyIndex]);
                }
                if (!uniqueColumnLabels.Contains((string)columnPMMDataObject[row, columnIndex]))
                {
                    uniqueColumnLabels.Add((string)columnPMMDataObject[row, columnIndex]);
                }

                var thisColumnInputData = new ColumnInputData();
                thisColumnInputData.columnLabel = (string)columnPMMDataObject[row, columnIndex];
                thisColumnInputData.story = (string)columnPMMDataObject[row, storyIndex];
                thisColumnInputData.story_elevation = (double)columnPMMDataObject[row, storyElevationIndex];
                thisColumnInputData.location = (string)columnPMMDataObject[row, locationIndex];
                thisColumnInputData.fck = (double)columnPMMDataObject[row, fckIndex];
                thisColumnInputData.depth = (double)columnPMMDataObject[row, depthIndex];
                thisColumnInputData.width = (double)columnPMMDataObject[row, widthIndex];
                thisColumnInputData.diameter = (double)columnPMMDataObject[row, diameterIndex];
                thisColumnInputData.P = (double)columnPMMDataObject[row, pIndex];
                thisColumnInputData.MMajor = (double)columnPMMDataObject[row, mMajorIndex];
                thisColumnInputData.MMinor = (double)columnPMMDataObject[row, mMinorIndex];
                thisColumnInputData.length = (double)columnPMMDataObject[row, lengthIndex];
                thisColumnInputData.governingCombo = ((string)columnPMMDataObject[row, govComboIndex]).Replace(" ", "");
                double givenPtRequired = 0;
                string numberFormat = designInputSheet.Cells[row, ptRequiredIndex].NumberFormat;
                if (numberFormat.Contains('%'))
                {
                    if(columnPMMDataObject[row, ptRequiredIndex] is double)
                    {
                        givenPtRequired = columnPMMDataObject[row, ptRequiredIndex] != null ?
                            (double)columnPMMDataObject[row, ptRequiredIndex] * 100 :
                            0;
                    }
                    else
                    {
                        givenPtRequired = 0;
                    }
                }
                else
                {
                    givenPtRequired = columnPMMDataObject[row, ptRequiredIndex] != null &&
                        ((string)columnPMMDataObject[row, ptRequiredIndex]).Contains('%') ?
                        double.Parse(((string)columnPMMDataObject[row, ptRequiredIndex]).Replace('%', ' ')) :
                        0;
                }
                thisColumnInputData.rebarPtEtabs = givenPtRequired;

                columnInputData.Add(thisColumnInputData);
            }

            foreach (string columnLabel in columnsLabels.Distinct().ToList())
            {
                columnsToDesign.Items.Add(columnLabel);
            }

            int vMajorIndex = 0;
            int vMinorIndex = 0;
            for (int column = 1; column <= columnShearSheet.UsedRange.Columns.Count; column++)
            {
                if (columnShearSheet.Cells[2, column].Value2 == null)
                {
                    break;
                }

                if (columnShearSheet.Cells[2, column].Value2 == "Label")
                {
                    columnIndex = column;
                }
                else if (columnShearSheet.Cells[2, column].Value2 == "Story")
                {
                    storyIndex = column;
                }
                else if (columnShearSheet.Cells[2, column].Value2 == "Location")
                {
                    locationIndex = column;
                }
                else if (columnShearSheet.Cells[2, column].Value2 == "VMajor")
                {
                    vMajorIndex = column;
                }
                else if (columnShearSheet.Cells[2, column].Value2 == "VMinor")
                {
                    vMinorIndex = column;
                }
            }

            object[,] columnShearDataObject = columnShearSheet.UsedRange.Value2;
            for (int row = 4; row <= columnShearSheet.UsedRange.Rows.Count; row++)
            {
                var columnLabel = columnShearDataObject[row, columnIndex];
                if (columnLabel == null)
                {
                    break;
                }

                var thisColumnShearData = new ColumnShearData();
                thisColumnShearData.columnLabel = (string)columnLabel;
                thisColumnShearData.story = (string)columnShearDataObject[row, storyIndex];
                thisColumnShearData.location = (string)columnShearDataObject[row, locationIndex];
                thisColumnShearData.maxVAlongY =
                    columnShearDataObject[row, vMajorIndex] != null ?
                    Math.Abs((double)columnShearDataObject[row, vMajorIndex]) :
                    0;
                thisColumnShearData.maxVAlongX =
                    columnShearDataObject[row, vMinorIndex] != null ?
                    Math.Abs((double)columnShearDataObject[row, vMinorIndex]) :
                    0;
                columnShearData.Add(thisColumnShearData);
            }

            object[,] columnForceDataObject = columnForcesSheet.UsedRange.Value2;
            int vAlongXIndex = 0;
            int vAlongYIndex = 0;
            int outputCaseIndex = -1;
            int stationIndex = -1;
            for (int column = 1; column <= columnForcesSheet.UsedRange.Columns.Count; column++)
            {
                if (columnForcesSheet.Cells[2, column].Value2 == null)
                {
                    break;
                }

                if (columnForcesSheet.Cells[2, column].Value2 == "Column")
                {
                    columnIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "Story")
                {
                    storyIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "P")
                {
                    pIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "M2")
                {
                    mMinorIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "M3")
                {
                    mMajorIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "V2")
                {
                    vAlongYIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "V3")
                {
                    vAlongXIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "OutputCase")
                {
                    outputCaseIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "Load Case/Combo")
                {
                    outputCaseIndex = column;
                }
                else if (columnForcesSheet.Cells[2, column].Value2 == "Station")
                {
                    stationIndex = column;
                }
            }

            for (int row = 4; row <= columnForcesSheet.UsedRange.Rows.Count; row++)
            {
                var columnLabel = columnForceDataObject[row, columnIndex];
                if (columnLabel == null)
                {
                    break;
                }

                var givenColumnForceData = new ColumnForceData();
                givenColumnForceData.columnLabel = (string)columnForceDataObject[row, columnIndex];
                givenColumnForceData.story = (string)columnForceDataObject[row, storyIndex];
                givenColumnForceData.P = (double)columnForceDataObject[row, pIndex];
                givenColumnForceData.MMajor = (double)columnForceDataObject[row, mMajorIndex];
                givenColumnForceData.MMinor = (double)columnForceDataObject[row, mMinorIndex];
                givenColumnForceData.VAlongX = Math.Abs((double)columnForceDataObject[row, vAlongXIndex]);
                givenColumnForceData.VAlongY = Math.Abs((double)columnForceDataObject[row, vAlongYIndex]);
                var outputCaseOrCombo = (string)columnForceDataObject[row, outputCaseIndex];
                givenColumnForceData.outputCase = outputCaseOrCombo.Replace("Max", "").Replace("Min", "").Replace(" ", "");
                givenColumnForceData.station = (double)columnForceDataObject[row, stationIndex];
                columnForceData.Add(givenColumnForceData);
            }

            columnPMMDataObject = null;
            columnShearDataObject = null;
            columnForceDataObject = null;

            excelWorkBook.Close();
            excelApplication.Quit();

            excelApplication = null;
            excelWorkBook = null;
            designInputSheet = null;
            columnShearSheet = null;
            columnForcesSheet = null;
        }

        private void columnsToDesign_SelectedIndexChanged(object sender, EventArgs e)
        {
            DesignStories.Items.Clear();
            this.maxEtabsRebarPt.Text = "";

            List<string> storyNames = new List<string>();
            foreach (var thisColumnInputData in columnInputData.Where(
                _ => _.columnLabel == this.columnsToDesign.SelectedItem.ToString()).ToList())
            {
                storyNames.Add(thisColumnInputData.story);
            }

            DesignStories.Items.Add("All stories");
            foreach (string storyName in storyNames.Distinct().ToList())
            {
                DesignStories.Items.Add(storyName);
            }
        }

        public double Interpolate(double value, double startX, double endX, double startY, double endY)
        {
            return startY + (endY - startY) * (value - startX) / (endX - startX);
        }

        public double CalculateTauC(double ptProvided, double grade)
        {
            if (14 < grade && grade < 16)
            {
                if (ptProvided <= 0.15)
                {
                    return 0.28;
                }
                else if (0.15 < ptProvided && ptProvided <= 0.25)
                {
                    return Interpolate(ptProvided, 0.15, 0.25, 0.28, 0.35);
                }
                else if (0.25 < ptProvided && ptProvided <= 0.5)
                {
                    return Interpolate(ptProvided, 0.25, 0.5, 0.35, 0.46);
                }
                else if (0.5 < ptProvided && ptProvided <= 0.75)
                {
                    return Interpolate(ptProvided, 0.5, 0.75, 0.46, 0.54);
                }
                else if (0.75 < ptProvided && ptProvided <= 1)
                {
                    return Interpolate(ptProvided, 0.75, 1, 0.54, 0.6);
                }
                else if (1 < ptProvided && ptProvided <= 1.25)
                {
                    return Interpolate(ptProvided, 1, 1.25, 0.6, 0.64);
                }
                else if (1.25 < ptProvided && ptProvided <= 1.5)
                {
                    return Interpolate(ptProvided, 1.25, 1.5, 0.64, 0.68);
                }
                else if (1.5 < ptProvided && ptProvided <= 1.75)
                {
                    return Interpolate(ptProvided, 1.5, 1.75, 0.68, 0.71);
                }
                else if (1.75 < ptProvided && ptProvided <= 2)
                {
                    return Interpolate(ptProvided, 1.75, 2, 0.71, 0.71);
                }
                else if (2 < ptProvided && ptProvided <= 2.25)
                {
                    return Interpolate(ptProvided, 2, 2.25, 0.71, 0.71);
                }
                else if (2.25 < ptProvided && ptProvided <= 2.5)
                {
                    return Interpolate(ptProvided, 2.25, 2.5, 0.71, 0.71);
                }
                else if (2.5 < ptProvided && ptProvided <= 2.75)
                {
                    return Interpolate(ptProvided, 2.5, 2.75, 0.71, 0.71);
                }
                else if (2.75 < ptProvided && ptProvided < 3)
                {
                    return Interpolate(ptProvided, 2.75, 3, 0.71, 0.71);
                }
                else if (3 <= ptProvided)
                {
                    return 0.71;
                }
                else
                    return 0;
            }
            else if (19 < grade && grade < 21)
            {
                if (ptProvided <= 0.15)
                {
                    return 0.28;
                }
                else if (0.15 < ptProvided && ptProvided <= 0.25)
                {
                    return Interpolate(ptProvided, 0.15, 0.25, 0.28, 0.36);
                }
                else if (0.25 < ptProvided && ptProvided <= 0.5)
                {
                    return Interpolate(ptProvided, 0.25, 0.5, 0.36, 0.48);
                }
                else if (0.5 < ptProvided && ptProvided <= 0.75)
                {
                    return Interpolate(ptProvided, 0.5, 0.75, 0.48, 0.56);
                }
                else if (0.75 < ptProvided && ptProvided <= 1)
                {
                    return Interpolate(ptProvided, 0.75, 1, 0.56, 0.62);
                }
                else if (1 < ptProvided && ptProvided <= 1.25)
                {
                    return Interpolate(ptProvided, 1, 1.25, 0.62, 0.67);
                }
                else if (1.25 < ptProvided && ptProvided <= 1.5)
                {
                    return Interpolate(ptProvided, 1.25, 1.5, 0.67, 0.72);
                }
                else if (1.5 < ptProvided && ptProvided <= 1.75)
                {
                    return Interpolate(ptProvided, 1.5, 1.75, 0.72, 0.75);
                }
                else if (1.75 < ptProvided && ptProvided <= 2)
                {
                    return Interpolate(ptProvided, 1.75, 2, 0.75, 0.79);
                }
                else if (2 < ptProvided && ptProvided <= 2.25)
                {
                    return Interpolate(ptProvided, 2, 2.25, 0.79, 0.81);
                }
                else if (2.25 < ptProvided && ptProvided <= 2.5)
                {
                    return Interpolate(ptProvided, 2.25, 2.5, 0.81, 0.82);
                }
                else if (2.5 < ptProvided && ptProvided <= 2.75)
                {
                    return Interpolate(ptProvided, 2.5, 2.75, 0.82, 0.82);
                }
                else if (2.75 < ptProvided && ptProvided < 3)
                {
                    return Interpolate(ptProvided, 2.75, 3, 0.82, 0.82);
                }
                else if (3 <= ptProvided)
                {
                    return 0.82;
                }
                else
                    return 0;
            }
            else if (24 < grade && grade < 26)
            {
                if (ptProvided <= 0.15)
                {
                    return 0.29;
                }
                else if (0.15 < ptProvided && ptProvided <= 0.25)
                {
                    return Interpolate(ptProvided, 0.15, 0.25, 0.29, 0.36);
                }
                else if (0.25 < ptProvided && ptProvided <= 0.5)
                {
                    return Interpolate(ptProvided, 0.25, 0.5, 0.36, 0.49);
                }
                else if (0.5 < ptProvided && ptProvided <= 0.75)
                {
                    return Interpolate(ptProvided, 0.5, 0.75, 0.49, 0.57);
                }
                else if (0.75 < ptProvided && ptProvided <= 1)
                {
                    return Interpolate(ptProvided, 0.75, 1, 0.57, 0.64);
                }
                else if (1 < ptProvided && ptProvided <= 1.25)
                {
                    return Interpolate(ptProvided, 1, 1.25, 0.64, 0.7);
                }
                else if (1.25 < ptProvided && ptProvided <= 1.5)
                {
                    return Interpolate(ptProvided, 1.25, 1.5, 0.7, 0.74);
                }
                else if (1.5 < ptProvided && ptProvided <= 1.75)
                {
                    return Interpolate(ptProvided, 1.5, 1.75, 0.74, 0.78);
                }
                else if (1.75 < ptProvided && ptProvided <= 2)
                {
                    return Interpolate(ptProvided, 1.75, 2, 0.78, 0.82);
                }
                else if (2 < ptProvided && ptProvided <= 2.25)
                {
                    return Interpolate(ptProvided, 2, 2.25, 0.82, 0.85);
                }
                else if (2.25 < ptProvided && ptProvided <= 2.5)
                {
                    return Interpolate(ptProvided, 2.25, 2.5, 0.85, 0.88);
                }
                else if (2.5 < ptProvided && ptProvided <= 2.75)
                {
                    return Interpolate(ptProvided, 2.5, 2.75, 0.88, 0.9);
                }
                else if (2.75 < ptProvided && ptProvided < 3)
                {
                    return Interpolate(ptProvided, 2.75, 3, 0.9, 0.92);
                }
                else if (3 <= ptProvided)
                {
                    return 0.92;
                }
                else
                    return 0;
            }
            else if (29 < grade && grade < 31)
            {
                if (ptProvided <= 0.15)
                {
                    return 0.29;
                }
                else if (0.15 < ptProvided && ptProvided <= 0.25)
                {
                    return Interpolate(ptProvided, 0.15, 0.25, 0.29, 0.37);
                }
                else if (0.25 < ptProvided && ptProvided <= 0.5)
                {
                    return Interpolate(ptProvided, 0.25, 0.5, 0.37, 0.5);
                }
                else if (0.5 < ptProvided && ptProvided <= 0.75)
                {
                    return Interpolate(ptProvided, 0.5, 0.75, 0.5, 0.59);
                }
                else if (0.75 < ptProvided && ptProvided <= 1)
                {
                    return Interpolate(ptProvided, 0.75, 1, 0.59, 0.66);
                }
                else if (1 < ptProvided && ptProvided <= 1.25)
                {
                    return Interpolate(ptProvided, 1, 1.25, 0.66, 0.71);
                }
                else if (1.25 < ptProvided && ptProvided <= 1.5)
                {
                    return Interpolate(ptProvided, 1.25, 1.5, 0.71, 0.76);
                }
                else if (1.5 < ptProvided && ptProvided <= 1.75)
                {
                    return Interpolate(ptProvided, 1.5, 1.75, 0.76, 0.8);
                }
                else if (1.75 < ptProvided && ptProvided <= 2)
                {
                    return Interpolate(ptProvided, 1.75, 2, 0.8, 0.84);
                }
                else if (2 < ptProvided && ptProvided <= 2.25)
                {
                    return Interpolate(ptProvided, 2, 2.25, 0.84, 0.88);
                }
                else if (2.25 < ptProvided && ptProvided <= 2.5)
                {
                    return Interpolate(ptProvided, 2.25, 2.5, 0.88, 0.91);
                }
                else if (2.5 < ptProvided && ptProvided <= 2.75)
                {
                    return Interpolate(ptProvided, 2.5, 2.75, 0.91, 0.94);
                }
                else if (2.75 < ptProvided && ptProvided < 3)
                {
                    return Interpolate(ptProvided, 2.75, 3, 0.94, 0.96);
                }
                else if (3 <= ptProvided)
                {
                    return 0.96;
                }
                else
                    return 0;
            }
            else if (34 < grade && grade < 36)
            {
                if (ptProvided <= 0.15)
                {
                    return 0.29;
                }
                else if (0.15 < ptProvided && ptProvided <= 0.25)
                {
                    return Interpolate(ptProvided, 0.15, 0.25, 0.29, 0.37);
                }
                else if (0.25 < ptProvided && ptProvided <= 0.5)
                {
                    return Interpolate(ptProvided, 0.25, 0.5, 0.37, 0.50);
                }
                else if (0.5 < ptProvided && ptProvided <= 0.75)
                {
                    return Interpolate(ptProvided, 0.5, 0.75, 0.50, 0.59);
                }
                else if (0.75 < ptProvided && ptProvided <= 1)
                {
                    return Interpolate(ptProvided, 0.75, 1, 0.59, 0.67);
                }
                else if (1 < ptProvided && ptProvided <= 1.25)
                {
                    return Interpolate(ptProvided, 1, 1.25, 0.67, 0.73);
                }
                else if (1.25 < ptProvided && ptProvided <= 1.5)
                {
                    return Interpolate(ptProvided, 1.25, 1.5, 0.73, 0.78);
                }
                else if (1.5 < ptProvided && ptProvided <= 1.75)
                {
                    return Interpolate(ptProvided, 1.5, 1.75, 0.78, 0.82);
                }
                else if (1.75 < ptProvided && ptProvided <= 2)
                {
                    return Interpolate(ptProvided, 1.75, 2, 0.82, 0.86);
                }
                else if (2 < ptProvided && ptProvided <= 2.25)
                {
                    return Interpolate(ptProvided, 2, 2.25, 0.86, 0.90);
                }
                else if (2.25 < ptProvided && ptProvided <= 2.5)
                {
                    return Interpolate(ptProvided, 2.25, 2.5, 0.90, 0.93);
                }
                else if (2.5 < ptProvided && ptProvided <= 2.75)
                {
                    return Interpolate(ptProvided, 2.5, 2.75, 0.93, 0.96);
                }
                else if (2.75 < ptProvided && ptProvided < 3)
                {
                    return Interpolate(ptProvided, 2.75, 3, 0.96, 0.99);
                }
                else if (3 <= ptProvided)
                {
                    return 0.99;
                }
                else
                    return 0;
            }
            else if (40 <= grade)
            {
                if (ptProvided <= 0.15)
                {
                    return 0.30;
                }
                else if (0.15 < ptProvided && ptProvided <= 0.25)
                {
                    return Interpolate(ptProvided, 0.15, 0.25, 0.30, 0.38);
                }
                else if (0.25 < ptProvided && ptProvided <= 0.5)
                {
                    return Interpolate(ptProvided, 0.25, 0.5, 0.38, 0.51);
                }
                else if (0.5 < ptProvided && ptProvided <= 0.75)
                {
                    return Interpolate(ptProvided, 0.5, 0.75, 0.51, 0.6);
                }
                else if (0.75 < ptProvided && ptProvided <= 1)
                {
                    return Interpolate(ptProvided, 0.75, 1, 0.6, 0.68);
                }
                else if (1 < ptProvided && ptProvided <= 1.25)
                {
                    return Interpolate(ptProvided, 1, 1.25, 0.68, 0.74);
                }
                else if (1.25 < ptProvided && ptProvided <= 1.5)
                {
                    return Interpolate(ptProvided, 1.25, 1.5, 0.74, 0.79);
                }
                else if (1.5 < ptProvided && ptProvided <= 1.75)
                {
                    return Interpolate(ptProvided, 1.5, 1.75, 0.79, 0.84);
                }
                else if (1.75 < ptProvided && ptProvided <= 2)
                {
                    return Interpolate(ptProvided, 1.75, 2, 0.84, 0.88);
                }
                else if (2 < ptProvided && ptProvided <= 2.25)
                {
                    return Interpolate(ptProvided, 2, 2.25, 0.88, 0.92);
                }
                else if (2.25 < ptProvided && ptProvided <= 2.5)
                {
                    return Interpolate(ptProvided, 2.25, 2.5, 0.92, 0.95);
                }
                else if (2.5 < ptProvided && ptProvided <= 2.75)
                {
                    return Interpolate(ptProvided, 2.5, 2.75, 0.95, 0.98);
                }
                else if (2.75 < ptProvided && ptProvided < 3)
                {
                    return Interpolate(ptProvided, 2.75, 3, 0.98, 1.01);
                }
                else if (3 <= ptProvided)
                {
                    return 1.01;
                }
                else
                    return 0;
            }
            else
                return 0;
        }

        public double CalculateTauCMax(double grade)
        {
            if (14 < grade && grade < 16)
            {
                return 2.5;
            }
            else if (19 < grade && grade < 21) { return 2.8; }
            else if (24 < grade && grade < 26) { return 3.1; }
            else if (29 < grade && grade < 31) { return 3.5; }
            else if (34 < grade && grade < 36) { return 3.7; }
            else if (40 <= grade) { return 4; }
            else
                return 0;
        }

        public Oasys.AdSec.Materials.IMaterial GetSectionMaterial(double materialStrength)
        {
            if(9 < materialStrength && materialStrength < 11)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M10;
            }
            else if (14 < materialStrength && materialStrength < 16)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M15;
            }
            else if (19 < materialStrength && materialStrength < 21)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M20;
            }
            else if (24 < materialStrength && materialStrength < 26)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M25;
            }
            else if (29 < materialStrength && materialStrength < 31)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M30;
            }
            else if (34 < materialStrength && materialStrength < 36)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M35;
            }
            else if (39 < materialStrength && materialStrength < 41)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M40;
            }
            else if (44 < materialStrength && materialStrength < 46)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M45;
            }
            else if (49 < materialStrength && materialStrength < 51)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M50;
            }
            else if (54 < materialStrength && materialStrength < 56)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M55;
            }
            else if (59 < materialStrength && materialStrength < 61)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M60;
            }
            else if (64 < materialStrength && materialStrength < 66)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M65;
            }
            else if (69 < materialStrength && materialStrength < 71)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M70;
            }
            else if (74 < materialStrength && materialStrength < 76)
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M75;
            }
            else
            {
                return Oasys.AdSec.StandardMaterials.Concrete.IS456.Edition_2000.M80;
            }
        }

        public Oasys.AdSec.Materials.IReinforcement GetRebarMaterial(double materialStrength)
        {
            if (249 < materialStrength && materialStrength < 251)
            {
                return Oasys.AdSec.StandardMaterials.Reinforcement.Steel.IS456.Edition_2000.S250;
            }
            else if (414 < materialStrength && materialStrength < 416)
            {
                return Oasys.AdSec.StandardMaterials.Reinforcement.Steel.IS456.Edition_2000.S415;
            }
            else
            {
                return Oasys.AdSec.StandardMaterials.Reinforcement.Steel.IS456.Edition_2000.S500;
            }
        }

        private void analyse_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < ResultsTable.Rows.Count - 1; i++)
            {
                section.Material = GetSectionMaterial((double)ResultsTable.Rows[i].Cells[2].Value);
                section.ReinforcementGroups.Clear();

                double linkMaterial = Double.Parse(this.givenLinkFy.Text);
                double linkDia = Double.Parse(this.linkDia.Text);
                var linkBar = IBarBundle.Create(GetRebarMaterial(linkMaterial), Length.FromMillimeters(linkDia));
                var linkGroup = ILinkGroup.Create(linkBar);
                section.ReinforcementGroups.Add(linkGroup);

                if (this.checkTop.Checked)
                {
                    double mainRebarMaterial = Double.Parse(this.givenTopFy.Text);
                    double topDia = Double.Parse(this.topDia.Text);
                    var topBar = IBarBundle.Create(GetRebarMaterial(mainRebarMaterial), Length.FromMillimeters(topDia));
                    double topSpacing = Double.Parse(this.topSpacing.Text);
                    var topLayer = ILayerByBarPitch.Create(topBar, Length.FromMillimeters(topSpacing));
                    var topGroup = ITemplateGroup.Create(ITemplateGroup.Face.Top);
                    topGroup.Layers.Add(topLayer);
                    section.ReinforcementGroups.Add(topGroup);
                }

                if (this.checkBottom.Checked)
                {
                    double mainRebarMaterial = Double.Parse(this.givenBottomFy.Text);
                    double bottomDia = Double.Parse(this.bottomDia.Text);
                    var bottomBar = IBarBundle.Create(GetRebarMaterial(mainRebarMaterial), Length.FromMillimeters(bottomDia));
                    double bottomSpacing = Double.Parse(this.bottomSpacing.Text);
                    var bottomLayer = ILayerByBarPitch.Create(bottomBar, Length.FromMillimeters(bottomSpacing));
                    var bottomGroup = ITemplateGroup.Create(ITemplateGroup.Face.Bottom);
                    bottomGroup.Layers.Add(bottomLayer);
                    section.ReinforcementGroups.Add(bottomGroup);
                }

                if (this.checkLeft.Checked)
                {
                    double mainRebarMaterial = Double.Parse(this.givenLeftFy.Text);
                    double leftDia = Double.Parse(this.leftDia.Text);
                    var leftBar = IBarBundle.Create(GetRebarMaterial(mainRebarMaterial), Length.FromMillimeters(leftDia));
                    double leftSpacing = Double.Parse(this.leftSpacing.Text);
                    var leftLayer = ILayerByBarPitch.Create(leftBar, Length.FromMillimeters(leftSpacing));
                    var leftGroup = ITemplateGroup.Create(ITemplateGroup.Face.LeftSide);
                    leftGroup.Layers.Add(leftLayer);
                    section.ReinforcementGroups.Add(leftGroup);
                }

                if (this.checkRight.Checked)
                {
                    double mainRebarMaterial = Double.Parse(this.givenRightFy.Text);
                    double rightDia = Double.Parse(this.rightDia.Text);
                    var rightBar = IBarBundle.Create(GetRebarMaterial(mainRebarMaterial), Length.FromMillimeters(rightDia));
                    double rightSpacing = Double.Parse(this.rightSpacing.Text);
                    var rightLayer = ILayerByBarPitch.Create(rightBar, Length.FromMillimeters(rightSpacing));
                    var rightGroup = ITemplateGroup.Create(ITemplateGroup.Face.RightSide);
                    rightGroup.Layers.Add(rightLayer);
                    section.ReinforcementGroups.Add(rightGroup);
                }

                if (this.checkPerimeter.Checked)
                {
                    double mainRebarMaterial = Double.Parse(this.givenPerimFy.Text);
                    double perimDia = Double.Parse(this.perimDia.Text);
                    var perimBar = IBarBundle.Create(GetRebarMaterial(mainRebarMaterial), Length.FromMillimeters(perimDia));
                    double perimSpacing = Double.Parse(this.perimSpacing.Text);
                    var perimLayer = ILayerByBarPitch.Create(perimBar, Length.FromMillimeters(perimSpacing));
                    var perimeterGroup = IPerimeterGroup.Create();
                    perimeterGroup.Layers.Add(perimLayer);
                    section.ReinforcementGroups.Add(perimeterGroup);
                }

                if (this.checkLine.Checked)
                {
                    foreach (var lineGroup in lineGroups)
                    {
                        section.ReinforcementGroups.Add(lineGroup);
                    }
                }

                double cover = Double.Parse(this.rebarCover.Text);
                section.Cover = ICover.Create(Length.FromMillimeters(cover));
                solution = adsecApp.Analyse(section);

                var simplifiedSection = adsecApp.Flatten(section);
                double rebarArea = 0;
                foreach (var barGroup in (simplifiedSection.ReinforcementGroups))
                {
                    if (barGroup is ISingleBars)
                    {
                        var singleBarGroup = (ISingleBars)barGroup;
                        double barDia = singleBarGroup.BarBundle.Diameter.Millimeters;
                        rebarArea += 0.785 * barDia * barDia * singleBarGroup.BarBundle.CountPerBundle * singleBarGroup.Positions.Count;
                    }
                }

                double rebarPtAdSec = 100 * rebarArea /
                    ((Oasys.Profiles.IRectangleProfile)section.Profile).Depth.Millimeters /
                    ((Oasys.Profiles.IRectangleProfile)section.Profile).Width.Millimeters;

                double axialAdSec = -1 * (double)ResultsTable.Rows[i].Cells[2].Value;
                double mMajorAdSec = (double)ResultsTable.Rows[i].Cells[3].Value;
                double mMinorAdSec = (double)ResultsTable.Rows[i].Cells[4].Value;
                var adSecLoad = ILoad.Create(
                    Force.FromKilonewtons(axialAdSec),
                    Moment.FromKilonewtonMeters(mMajorAdSec),
                    Moment.FromKilonewtonMeters(mMinorAdSec));

                var strengthResult = solution.Strength.Check(adSecLoad);
                var momentRanges = strengthResult.MomentRanges;
                double ultimateMoment = 0;
                for (int j = 0; j < momentRanges.Count; j++)
                {
                    ultimateMoment = Math.Max(momentRanges[j].Max.KilonewtonMeters, ultimateMoment);
                }
                double appliedMoment = Math.Sqrt(Math.Pow(mMajorAdSec, 2) + Math.Pow(mMinorAdSec, 2));
                double momentRatio = appliedMoment / ultimateMoment;
                ResultsTable.Rows[i].Cells[5].Value = Math.Round(strengthResult.LoadUtilisation.DecimalFractions, 2);
                ResultsTable.Rows[i].Cells[6].Value = Math.Round(momentRatio, 2);
                ResultsTable.Rows[i].Cells[8].Value = Math.Round(rebarPtAdSec, 3);
                ResultsTable.Rows[i].Cells[7].Value = "";
            }
        }

        private void addLine_Click(object sender, EventArgs e)
        {
            double mainRebarMaterial = Double.Parse(this.givenLineFy.Text);
            double lineDiaValue = double.Parse(this.lineDia.Text);
            double point1XValue = double.Parse(this.point1X.Text);
            double point1YValue = double.Parse(this.point1Y.Text);
            double point2XValue = double.Parse(this.point2X.Text);
            double point2YValue = double.Parse(this.point2Y.Text);
            double lineSpacingValue = double.Parse(this.lineSpacing.Text);
            var lineBar = IBarBundle.Create(GetRebarMaterial(mainRebarMaterial), Length.FromMillimeters(lineDiaValue));
            var lineLayer = ILayerByBarPitch.Create(lineBar, Length.FromMillimeters(lineSpacingValue));
            var firstBarPosition = IPoint.Create(Length.FromMillimeters(point1XValue), Length.FromMillimeters(point1YValue));
            var lastBarPosition = IPoint.Create(Length.FromMillimeters(point2XValue), Length.FromMillimeters(point2YValue));

            var lineGroup = ILineGroup.Create(firstBarPosition, lastBarPosition, lineLayer);
            lineGroups.Add(lineGroup);
            string lineDescription =
                this.givenLineFy.Text + "MPa T" +
                this.lineDia.Text + "-" +
                this.lineSpacing.Text + " (" +
                this.point1X.Text + ", " +
                this.point1Y.Text + ") (" +
                this.point2X.Text + ", " +
                this.point2Y.Text + ")";
            this.lineGroupsAdded.Items.Add(lineDescription);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (this.lineGroupsAdded.Items.Count != 1)
            {
                lineGroups.Remove(lineGroups[this.lineGroupsAdded.SelectedIndex]);
                this.lineGroupsAdded.Items.Remove(this.lineGroupsAdded.SelectedItem);
            }
            else
            {
                this.lineGroupsAdded.Items.Clear();
                lineGroups.Clear();
                this.lineGroupsAdded.Text = "";
            }
        }

        private void addToExcel_Click(object sender, EventArgs e)
        {
            if (designDataToExcel.Count < 1)
                return;

            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkBook = null;
            Excel.Worksheet designOutputSheet = null;

            excelApplication = new Excel.Application();
            excelApplication.Visible = true;
            string @excelPath = this.designInput.Text;
            excelWorkBook = excelApplication.Workbooks.Open(excelPath);
            designOutputSheet = excelWorkBook.Worksheets["Reinforcement Schedule"];
            int startRow = designOutputSheet.UsedRange.Rows.Count + 1;

            for (int i = 0; i < ResultsTable.Rows.Count-1; i++)
            {
                designOutputSheet.Cells[startRow, 1].Value2 = (string)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.column_label].Value;
                designOutputSheet.Cells[startRow, 2].Value2 = (string)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.story].Value;
                designOutputSheet.Cells[startRow, 3].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.story_elevation].Value;
                designOutputSheet.Cells[startRow, 4].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.depth].Value;
                designOutputSheet.Cells[startRow, 5].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.width].Value;
                designOutputSheet.Cells[startRow, 6].Value2 = 0;
                designOutputSheet.Cells[startRow, 7].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.fck].Value;
                designOutputSheet.Cells[startRow, 8].Value2 = double.Parse(rebarCover.Text);
                designOutputSheet.Cells[startRow, 9].Value2 = (string)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.rebar_description].Value;
                designOutputSheet.Cells[startRow, 10].Value2 = "T" + this.linkDia.Text + "-" + this.nonConfiningSpacing.Text + "mm";
                designOutputSheet.Cells[startRow, 11].Value2 = "T" + this.linkDia.Text + "-" + this.confiningSpacing.Text + "mm";
                designOutputSheet.Cells[startRow, 12].Value2 = designDataToExcel[i].legsAlongY;
                designOutputSheet.Cells[startRow, 13].Value2 = designDataToExcel[i].legsAlongX;
                designOutputSheet.Cells[startRow, 14].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.p].Value;
                designOutputSheet.Cells[startRow, 15].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.m33].Value;
                designOutputSheet.Cells[startRow, 16].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.m22].Value;
                designOutputSheet.Cells[startRow, 17].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.etabs_req].Value;
                designOutputSheet.Cells[startRow, 18].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.adsec_pro].Value;
                designOutputSheet.Cells[startRow, 19].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.load_util].Value;
                designOutputSheet.Cells[startRow, 20].Value2 = (double)ResultsTable.Rows[i].Cells[ResultTableColumnIndex.m_ratio].Value;
                startRow += 1;
            }

            for (int i = 0; i < designDataToExcel.Count; i++)
            {
                var givenDesignData = designDataToExcel[i];
                string columnID = givenDesignData.inputData.columnLabel + " " + givenDesignData.inputData.story;

                var appliedLoads = Oasys.Collections.IList<ILoad>.Create();

                appliedLoads.Add(ILoad.Create(
                    Force.FromKilonewtons(-1 * givenDesignData.inputData.P),
                    Moment.FromKilonewtonMeters(givenDesignData.inputData.MMajor),
                    Moment.FromKilonewtonMeters(givenDesignData.inputData.MMinor)));

                foreach (var load in givenDesignData.checkedMajorEccentricLoads)
                {
                    appliedLoads.Add(load);
                }
                foreach (var load in givenDesignData.checkedMinorEccentricLoads)
                {
                    appliedLoads.Add(load);
                }
                foreach (var load in givenDesignData.checkedOtherLoads)
                {
                    appliedLoads.Add(load);
                }

                var jsonConverter = new Oasys.AdSec.IO.Serialization.JsonConverter(IS456.Edition_2000);
                var jsonString = jsonConverter.SectionToJson(givenDesignData.designSection, appliedLoads);
                string @excelFilePath = this.designInput.Text;
                string directoryName = Path.GetDirectoryName(excelFilePath);
                if (!Directory.Exists(Path.Combine(directoryName, "AdSec"))) Directory.CreateDirectory(Path.Combine(directoryName, "AdSec"));
                string adsecFilePath = Path.Combine(directoryName,"AdSec", columnID + ".ads");

                if (File.Exists(adsecFilePath))
                {
                    File.Delete(adsecFilePath);

                    excelApplication.DisplayAlerts = false;
                    excelWorkBook.Worksheets[columnID + " IS456"].Delete();
                    excelWorkBook.Worksheets[columnID + " IS13920"].Delete();
                    excelApplication.DisplayAlerts = true;
                }

                File.WriteAllText(adsecFilePath, jsonString);


                Excel.Worksheet shearCalcIS456Sheet = excelWorkBook.Sheets.Add(
                    Type.Missing, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], Type.Missing, Type.Missing);

                shearCalcIS456Sheet.Name = columnID + " IS456";
                shearCalcIS456Sheet.Cells[2, 3].Value2 = "SHEAR REINFORCEMENT FOR RC COLUMNS";
                shearCalcIS456Sheet.Cells[4, 3].Value2 = "As per IS456:2000";
                shearCalcIS456Sheet.Cells[6, 3].Value2 = "RECTANGULAR COLUMN";
                shearCalcIS456Sheet.Cells[8, 3].Value2 = givenDesignData.inputData.columnLabel;
                shearCalcIS456Sheet.Cells[8, 4].Value2 = givenDesignData.inputData.story;
                shearCalcIS456Sheet.Cells[10, 3].Value2 = "Factored shear force (along X)";
                shearCalcIS456Sheet.Cells[10, 4].Value2 = givenDesignData.factoredShearForceAlongX;
                shearCalcIS456Sheet.Cells[10, 5].Value2 = "kN";
                shearCalcIS456Sheet.Cells[11, 3].Value2 = "Factored shear force (along Y)";
                shearCalcIS456Sheet.Cells[11, 4].Value2 = givenDesignData.factoredShearForceAlongY;
                shearCalcIS456Sheet.Cells[11, 5].Value2 = "kN";
                shearCalcIS456Sheet.Cells[12, 3].Value2 = "B";
                shearCalcIS456Sheet.Cells[12, 4].Value2 = givenDesignData.inputData.width;
                shearCalcIS456Sheet.Cells[12, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[13, 3].Value2 = "D";
                shearCalcIS456Sheet.Cells[13, 4].Value2 = givenDesignData.inputData.depth;
                shearCalcIS456Sheet.Cells[13, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[14, 3].Value2 = "Clear cover";
                shearCalcIS456Sheet.Cells[14, 4].Value2 = givenDesignData.clearCover;
                shearCalcIS456Sheet.Cells[14, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[15, 3].Value2 = "Dia of corner longitudinal bar";
                shearCalcIS456Sheet.Cells[15, 4].Value2 = givenDesignData.DesignData.cornerBarDia;
                shearCalcIS456Sheet.Cells[15, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[16, 3].Value2 = "Dia of centre longitudinal bar";
                shearCalcIS456Sheet.Cells[16, 4].Value2 = givenDesignData.DesignData.centreBarDia;
                shearCalcIS456Sheet.Cells[16, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[17, 3].Value2 = "b";
                shearCalcIS456Sheet.Cells[17, 4].Value2 = givenDesignData.effectiveWidth;
                shearCalcIS456Sheet.Cells[17, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[18, 3].Value2 = "d";
                shearCalcIS456Sheet.Cells[18, 4].Value2 = givenDesignData.effectiveDepth;
                shearCalcIS456Sheet.Cells[18, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[19, 3].Value2 = "tau v (along X)";
                shearCalcIS456Sheet.Cells[19, 4].Value2 = givenDesignData.tauVAlongX;
                shearCalcIS456Sheet.Cells[19, 5].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[19, 6].Value2 = "cl. 40.1";
                shearCalcIS456Sheet.Cells[20, 3].Value2 = "tau v (along Y)";
                shearCalcIS456Sheet.Cells[20, 4].Value2 = givenDesignData.tauVAlongY;
                shearCalcIS456Sheet.Cells[20, 5].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[20, 6].Value2 = "cl. 40.1";
                shearCalcIS456Sheet.Cells[22, 3].Value2 = "tau c max";
                shearCalcIS456Sheet.Cells[22, 4].Value2 = givenDesignData.tauCMax;
                shearCalcIS456Sheet.Cells[22, 5].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[22, 6].Value2 = "Table 20";
                shearCalcIS456Sheet.Cells[23, 3].Value2 = "Pt";
                shearCalcIS456Sheet.Cells[23, 4].Value2 = Math.Round(givenDesignData.DesignData.ptProvided, 2);
                shearCalcIS456Sheet.Cells[23, 5].Value2 = "%";
                shearCalcIS456Sheet.Cells[23, 6].Value2 = "Check 'Reinforcement Schedule' sheet";
                shearCalcIS456Sheet.Cells[24, 3].Value2 = "Tensile Rebar";
                shearCalcIS456Sheet.Cells[24, 4].Value2 = Math.Round(givenDesignData.DesignData.tensile_rebar, 2);
                shearCalcIS456Sheet.Cells[24, 5].Value2 = "%";
                shearCalcIS456Sheet.Cells[24, 6].Value2 = "Check 'Reinforcement Schedule' sheet";
                shearCalcIS456Sheet.Cells[25, 3].Value2 = "fck";
                shearCalcIS456Sheet.Cells[25, 4].Value2 = givenDesignData.inputData.fck;
                shearCalcIS456Sheet.Cells[25, 5].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[27, 3].Value2 = "tau c";
                shearCalcIS456Sheet.Cells[27, 4].Value2 = givenDesignData.tauC;
                shearCalcIS456Sheet.Cells[27, 5].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[27, 6].Value2 = IncreaseTauC.Checked ?
                    "Considering increased shear strength from Table 19 as per cl.40.2.2" :
                    "Table 19";
                shearCalcIS456Sheet.Cells[29, 4].Value2 =
                    Math.Max(givenDesignData.tauVAlongX, givenDesignData.tauVAlongY) > givenDesignData.tauC ?
                    "Shear reinforcement needed" :
                    "Shear reinforcement not needed, provide minimum shear reinforcement";
                shearCalcIS456Sheet.Cells[31, 3].Value2 = "Vus (along X)";
                shearCalcIS456Sheet.Cells[31, 4].Value2 = givenDesignData.VusAlongX;
                shearCalcIS456Sheet.Cells[31, 5].Value2 = "kN";
                shearCalcIS456Sheet.Cells[31, 6].Value2 = "cl. 40.4";
                shearCalcIS456Sheet.Cells[32, 3].Value2 = "Vus (along Y)";
                shearCalcIS456Sheet.Cells[32, 4].Value2 = givenDesignData.VusAlongY;
                shearCalcIS456Sheet.Cells[32, 5].Value2 = "kN";
                shearCalcIS456Sheet.Cells[32, 6].Value2 = "cl. 40.4";
                shearCalcIS456Sheet.Cells[34, 3].Value2 = "fy";
                shearCalcIS456Sheet.Cells[34, 4].Value2 = givenDesignData.linkFy;
                shearCalcIS456Sheet.Cells[34, 5].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[35, 3].Value2 = "Dia of transverse bar";
                shearCalcIS456Sheet.Cells[35, 4].Value2 = givenDesignData.linkDia;
                shearCalcIS456Sheet.Cells[35, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[36, 3].Value2 = "number of legs (along X)";
                shearCalcIS456Sheet.Cells[36, 4].Value2 = givenDesignData.legsAlongX;
                shearCalcIS456Sheet.Cells[37, 3].Value2 = "Asv (along X)";
                shearCalcIS456Sheet.Cells[37, 4].Value2 = givenDesignData.asvProvidedAlongX;
                shearCalcIS456Sheet.Cells[37, 5].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[38, 3].Value2 = "number of legs (along Y)";
                shearCalcIS456Sheet.Cells[38, 4].Value2 = givenDesignData.legsAlongY;
                shearCalcIS456Sheet.Cells[39, 3].Value2 = "Asv (along Y)";
                shearCalcIS456Sheet.Cells[39, 4].Value2 = givenDesignData.asvProvidedAlongY;
                shearCalcIS456Sheet.Cells[39, 5].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[41, 3].Value2 = "Sv required";
                shearCalcIS456Sheet.Cells[41, 4].Value2 = givenDesignData.nonConfiningSpacingOne;
                shearCalcIS456Sheet.Cells[41, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[41, 6].Value2 = "cl. 40.4 (a)";
                shearCalcIS456Sheet.Cells[42, 3].Value2 = "Sv required";
                shearCalcIS456Sheet.Cells[42, 4].Value2 = givenDesignData.nonConfiningSpacingTwo;
                shearCalcIS456Sheet.Cells[42, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[42, 6].Value2 = "cl. 26.5.3.2 (c)";
                shearCalcIS456Sheet.Cells[43, 3].Value2 = "Sv provided";
                shearCalcIS456Sheet.Cells[43, 4].Value2 = givenDesignData.nonConfiningSpacingProvided;
                shearCalcIS456Sheet.Cells[43, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[44, 4].Value2 =
                    givenDesignData.nonConfiningSpacingOne > 0 ?
                    givenDesignData.nonConfiningSpacingProvided <
                    Math.Min(givenDesignData.nonConfiningSpacingOne, givenDesignData.nonConfiningSpacingTwo) ? "SPACING IS OK" : "SPACING IS NOT OK" :
                    givenDesignData.nonConfiningSpacingProvided <
                    givenDesignData.nonConfiningSpacingTwo ? "SPACING IS OK" : "SPACING IS NOT OK";
                if ((string)shearCalcIS456Sheet.Cells[44, 4].Value2 == "SPACING IS NOT OK")
                {
                    shearCalcIS456Sheet.Cells[44, 4].Interior.Color = Color.Red;
                    shearCalcIS456Sheet.Cells[44, 4].Font.Color = Color.White;
                }
                shearCalcIS456Sheet.Cells[46, 3].Value2 = "Minimum shear reinforcement";
                shearCalcIS456Sheet.Cells[46, 4].Value2 = givenDesignData.minNonConfiningAsvRequired;
                shearCalcIS456Sheet.Cells[46, 5].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[46, 6].Value2 = "cl. 26.5.1.6";
                shearCalcIS456Sheet.Cells[47, 4].Value2 =
                    givenDesignData.minNonConfiningAsvRequired <
                    Math.Min(givenDesignData.legsAlongX * Math.PI * 0.25 * givenDesignData.linkDia * givenDesignData.linkDia,
                    givenDesignData.legsAlongY * Math.PI * 0.25 * givenDesignData.linkDia * givenDesignData.linkDia) ? "AREA IS OK" : "AREA IS NOT OK";
                if ((string)shearCalcIS456Sheet.Cells[47, 4].Value2 == "AREA IS NOT OK")
                {
                    shearCalcIS456Sheet.Cells[47, 4].Interior.Color = Color.Red;
                    shearCalcIS456Sheet.Cells[47, 4].Font.Color = Color.White;
                }
                shearCalcIS456Sheet.Cells[49, 3].Value2 = "Hence provide";
                shearCalcIS456Sheet.Cells[49, 4].Value2 = givenDesignData.linkDia;
                shearCalcIS456Sheet.Cells[49, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[49, 6].Value2 = givenDesignData.legsAlongX;
                shearCalcIS456Sheet.Cells[49, 7].Value2 = "legged stirrups along X @";
                shearCalcIS456Sheet.Cells[49, 8].Value2 = givenDesignData.nonConfiningSpacingProvided;
                shearCalcIS456Sheet.Cells[49, 9].Value2 = "mm";
                shearCalcIS456Sheet.Cells[49, 10].Value2 = "c/c";
                shearCalcIS456Sheet.Cells[49, 11].Value2 = "in the non-confining zone";
                shearCalcIS456Sheet.Cells[49, 12].Value2 = "As per IS456:2000";
                shearCalcIS456Sheet.Cells[50, 3].Value2 = "Hence provide";
                shearCalcIS456Sheet.Cells[50, 4].Value2 = givenDesignData.linkDia;
                shearCalcIS456Sheet.Cells[50, 5].Value2 = "mm";
                shearCalcIS456Sheet.Cells[50, 6].Value2 = givenDesignData.legsAlongY;
                shearCalcIS456Sheet.Cells[50, 7].Value2 = "legged stirrups along Y @";
                shearCalcIS456Sheet.Cells[50, 8].Value2 = givenDesignData.nonConfiningSpacingProvided;
                shearCalcIS456Sheet.Cells[50, 9].Value2 = "mm";
                shearCalcIS456Sheet.Cells[50, 10].Value2 = "c/c";
                shearCalcIS456Sheet.Cells[50, 11].Value2 = "in the non-confining zone";
                shearCalcIS456Sheet.Cells[50, 12].Value2 = "As per IS456:2000";

                Excel.Worksheet shearCalcIS13920Sheet = excelWorkBook.Sheets.Add(
                    Type.Missing, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], Type.Missing, Type.Missing);
                shearCalcIS13920Sheet.Name = columnID + " IS13920";
                shearCalcIS13920Sheet.Cells[2, 3].Value2 = "SHEAR REINFORCEMENT FOR RC COLUMNS";
                shearCalcIS13920Sheet.Cells[4, 3].Value2 = "As per IS13920:2016";
                shearCalcIS13920Sheet.Cells[6, 3].Value2 = "RECTANGULAR COLUMN";
                shearCalcIS13920Sheet.Cells[8, 3].Value2 = givenDesignData.inputData.columnLabel;
                shearCalcIS13920Sheet.Cells[8, 4].Value2 = givenDesignData.inputData.story;
                shearCalcIS13920Sheet.Cells[10, 3].Value2 = "Spacing of confining reinforcement (Sv)";
                shearCalcIS13920Sheet.Cells[12, 3].Value2 = "B";
                shearCalcIS13920Sheet.Cells[12, 4].Value2 = givenDesignData.inputData.width;
                shearCalcIS13920Sheet.Cells[12, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[13, 3].Value2 = "D";
                shearCalcIS13920Sheet.Cells[13, 4].Value2 = givenDesignData.inputData.depth;
                shearCalcIS13920Sheet.Cells[13, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[14, 4].Value2 = givenDesignData.confinfingSpacingOne;
                shearCalcIS13920Sheet.Cells[14, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[14, 6].Value2 = "cl. 7.6.1 (b) (1)";
                shearCalcIS13920Sheet.Cells[16, 3].Value2 = "Smallest dia of longitudinal bar";
                shearCalcIS13920Sheet.Cells[16, 4].Value2 = givenDesignData.DesignData.centreBarDia;
                shearCalcIS13920Sheet.Cells[16, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[17, 4].Value2 = givenDesignData.confinfingSpacingTwo;
                shearCalcIS13920Sheet.Cells[17, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[17, 6].Value2 = "cl. 7.6.1 (b) (2)";
                shearCalcIS13920Sheet.Cells[19, 4].Value2 = givenDesignData.confinfingSpacingThree;
                shearCalcIS13920Sheet.Cells[19, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[19, 6].Value2 = "cl. 7.6.1 (b) (3)";
                shearCalcIS13920Sheet.Cells[21, 3].Value2 = "Sv required";
                shearCalcIS13920Sheet.Cells[21, 4].Value2 = givenDesignData.maxConfiningSpacingRequired;
                shearCalcIS13920Sheet.Cells[21, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[22, 3].Value2 = "Sv provided";
                shearCalcIS13920Sheet.Cells[22, 4].Value2 = givenDesignData.confiningSpacingProvided;
                shearCalcIS13920Sheet.Cells[22, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[23, 4].Value2 =
                    givenDesignData.maxConfiningSpacingRequired >= givenDesignData.confiningSpacingProvided ?
                    "SPACING IS OK" : "SPACING IS NOT OK";
                if ((string)shearCalcIS13920Sheet.Cells[23, 4].Value2 == "SPACING IS NOT OK")
                {
                    shearCalcIS13920Sheet.Cells[23, 4].Interior.Color = Color.Red;
                    shearCalcIS13920Sheet.Cells[23, 4].Font.Color = Color.White;
                }
                shearCalcIS13920Sheet.Cells[25, 3].Value2 = "fck";
                shearCalcIS13920Sheet.Cells[25, 4].Value2 = givenDesignData.inputData.fck;
                shearCalcIS13920Sheet.Cells[25, 5].Value2 = "N/mm2";
                shearCalcIS13920Sheet.Cells[26, 3].Value2 = "fy";
                shearCalcIS13920Sheet.Cells[26, 4].Value2 = givenDesignData.linkFy;
                shearCalcIS13920Sheet.Cells[26, 5].Value2 = "N/mm2";
                shearCalcIS13920Sheet.Cells[28, 3].Value2 = "Dia of transverse bar";
                shearCalcIS13920Sheet.Cells[28, 4].Value2 = givenDesignData.linkDia;
                shearCalcIS13920Sheet.Cells[28, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[29, 3].Value2 = "Dia of corner longitudinal bar";
                shearCalcIS13920Sheet.Cells[29, 4].Value2 = givenDesignData.DesignData.cornerBarDia;
                shearCalcIS13920Sheet.Cells[29, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[30, 3].Value2 = "Clear cover";
                shearCalcIS13920Sheet.Cells[30, 4].Value2 = givenDesignData.clearCover;
                shearCalcIS13920Sheet.Cells[30, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[32, 3].Value2 = "Ag";
                shearCalcIS13920Sheet.Cells[32, 4].Value2 = givenDesignData.Ag;
                shearCalcIS13920Sheet.Cells[32, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[32, 6].Value2 = "cl. 7.6.1 (c)";
                shearCalcIS13920Sheet.Cells[33, 3].Value2 = "Ak";
                shearCalcIS13920Sheet.Cells[33, 4].Value2 = givenDesignData.Ak;
                shearCalcIS13920Sheet.Cells[33, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[34, 3].Value2 = "h";
                shearCalcIS13920Sheet.Cells[34, 4].Value2 = givenDesignData.h;
                shearCalcIS13920Sheet.Cells[34, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[35, 3].Value2 = "Ash required one";
                shearCalcIS13920Sheet.Cells[35, 4].Value2 = givenDesignData.AshOne;
                shearCalcIS13920Sheet.Cells[35, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[35, 6].Value2 = "cl. 7.6.1 (c) (2)";
                shearCalcIS13920Sheet.Cells[36, 3].Value2 = "Ash required two";
                shearCalcIS13920Sheet.Cells[36, 4].Value2 = givenDesignData.AshTwo;
                shearCalcIS13920Sheet.Cells[36, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[36, 6].Value2 = "cl. 7.6.1 (c) (2)";
                shearCalcIS13920Sheet.Cells[37, 3].Value2 = "Ash provided";
                shearCalcIS13920Sheet.Cells[37, 4].Value2 = givenDesignData.AshProvided;
                shearCalcIS13920Sheet.Cells[37, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[38, 4].Value2 =
                    givenDesignData.AshProvided > Math.Max(givenDesignData.AshOne, givenDesignData.AshTwo) ?
                    "AREA IS OK" : "AREA IS NOT OK";
                if ((string)shearCalcIS13920Sheet.Cells[38, 4].Value2 == "AREA IS NOT OK")
                {
                    shearCalcIS13920Sheet.Cells[38, 4].Interior.Color = Color.Red;
                    shearCalcIS13920Sheet.Cells[38, 4].Font.Color = Color.White;
                }

                shearCalcIS13920Sheet.Cells[40, 3].Value2 = "Hence provide";
                shearCalcIS13920Sheet.Cells[40, 4].Value2 = givenDesignData.linkDia;
                shearCalcIS13920Sheet.Cells[40, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[40, 6].Value2 = "@";
                shearCalcIS13920Sheet.Cells[40, 7].Value2 = givenDesignData.confiningSpacingProvided;
                shearCalcIS13920Sheet.Cells[40, 8].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[40, 9].Value2 = "as ties in the end zone (confining zone) and lap splices zone";


                //Excel.Worksheet columnDetailSheet = excelWorkBook.Sheets.Add(
                //    Type.Missing, excelWorkBook.Sheets[excelWorkBook.Sheets.Count], Type.Missing, Type.Missing);
                //columnDetailSheet.Name = columnID + " Detail";
                //columnDetailSheet.Cells[1, 1].Value2 = "Depth";
                //columnDetailSheet.Cells[2, 1].Value2 = "Width";
                //columnDetailSheet.Cells[3, 1].Value2 = "fck";
                //columnDetailSheet.Cells[4, 1].Value2 = "Cover";
                //columnDetailSheet.Cells[5, 1].Value2 = "Link dia";
                //columnDetailSheet.Cells[6, 1].Value2 = "Link fy";
                //columnDetailSheet.Cells[1, 2].Value2 = givenDesignData.inputData.depth;
                //columnDetailSheet.Cells[2, 2].Value2 = givenDesignData.inputData.width;
                //columnDetailSheet.Cells[3, 2].Value2 = givenDesignData.inputData.fck;
                //columnDetailSheet.Cells[4, 2].Value2 = givenDesignData.clearCover;
                //columnDetailSheet.Cells[5, 2].Value2 = givenDesignData.linkDia;
                //columnDetailSheet.Cells[6, 2].Value2 = givenDesignData.linkFy;

                //columnDetailSheet.Cells[1, 4].Value2 = "Line Rebar";
                //columnDetailSheet.Cells[1, 5].Value2 = "Bar count";
                //columnDetailSheet.Cells[1, 6].Value2 = "Bar fy";
                //columnDetailSheet.Cells[1, 7].Value2 = "Bar dia";
                //columnDetailSheet.Cells[1, 8].Value2 = "Start point Y";
                //columnDetailSheet.Cells[1, 9].Value2 = "Start point Z";
                //columnDetailSheet.Cells[1, 10].Value2 = "End point Y";
                //columnDetailSheet.Cells[1, 11].Value2 = "End point Z";
                //var givenDesignSection = givenDesignData.designSection;
                //int lineRebarCount = 1;
                //foreach(var rebarGroup in givenDesignSection.ReinforcementGroups)
                //{
                //    if(rebarGroup is ILineGroup)
                //    {
                //        lineRebarCount++;
                //        var group = (ILineGroup)rebarGroup;
                //        var layer = (ILayerByBarCount)group.Layer;
                //        columnDetailSheet.Cells[lineRebarCount, 4].Value2 = lineRebarCount-1;
                //        columnDetailSheet.Cells[lineRebarCount, 5].Value2 = layer.Count;
                //        columnDetailSheet.Cells[lineRebarCount, 6].Value2 = givenDesignData.longitudinalFy;
                //        columnDetailSheet.Cells[lineRebarCount, 7].Value2 = layer.BarBundle.Diameter.Millimeters;
                //        columnDetailSheet.Cells[lineRebarCount, 8].Value2 = Math.Round(group.FirstBarPosition.Y.Millimeters);
                //        columnDetailSheet.Cells[lineRebarCount, 9].Value2 = Math.Round(group.FirstBarPosition.Z.Millimeters);
                //        columnDetailSheet.Cells[lineRebarCount, 10].Value2 = Math.Round(group.LastBarPosition.Y.Millimeters);
                //        columnDetailSheet.Cells[lineRebarCount, 11].Value2 = Math.Round(group.LastBarPosition.Z.Millimeters);
                //    }
                //}
            }

            excelWorkBook.Save();
            excelWorkBook.Close();
            excelApplication.Quit();

            //var firstDesignSection = this.designDataToExcel[0].designSection;
            //var rectangleProfile = (IRectangleProfile)firstDesignSection.Profile;
            //DocumentCollection acDocMgr = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;
            //Document acDoc = acDocMgr.Add("acad.dwt");
            //acDocMgr.MdiActiveDocument = acDoc;

        }

        public double CalculatePtProvided(ISection concreteSection, double cornerBarCount, int cornerBarDia, double centreBarCount, int centreBarDia)
        {
            double ptProvided =
                100 * (centreBarCount * Math.Pow(centreBarDia, 2) * Math.PI * 0.25 +
                cornerBarCount * Math.Pow(cornerBarDia, 2) * Math.PI * 0.25) /
                ((IRectangleProfile)concreteSection.Profile).Depth.Millimeters /
                ((IRectangleProfile)concreteSection.Profile).Width.Millimeters;
            
            return ptProvided;
        }

        public void AddColumnDesignData
            (ISection concreteSection, 
            IList<ColumnDesignData> designDatas, 
            int nAlongDepth, 
            int nCentreBarAlongDepth, 
            int nAlongWidth, 
            int nCentreBarAlongWidth, 
            int cornerBarDia,
            int centreBarDia,
            double arrangementDepth,
            double arrangementWidth,
            double ptRequired)
        {
            int centreBarCount = 2 * (nCentreBarAlongDepth + nCentreBarAlongWidth);
            int cornerBarCount = 2 * (nAlongDepth + nAlongWidth) - centreBarCount - 4;
            double ptProvided = CalculatePtProvided(concreteSection, cornerBarCount, cornerBarDia, centreBarCount, centreBarDia);
            
            //if(designDatas.Count == 0)
            //{
            //    ColumnDesignData columnDesignData = new ColumnDesignData();
            //    columnDesignData.cornerBarCountAlongDepth = 0;
            //    columnDesignData.cornerBarCountAlongWidth = 0;
            //    columnDesignData.cornerBarDia = 0;
            //    columnDesignData.centreBarCountAlongDepth = 0;
            //    columnDesignData.centreBarCountAlongWidth = 0;
            //    columnDesignData.centreBarDia = 0;
            //    columnDesignData.ptProvided = 0;
            //    columnDesignData.spacingAlongDepth = 0;
            //    columnDesignData.spacingAlongWidth = 0;
            //    columnDesignData.arrangementDepth = arrangementDepth;
            //    columnDesignData.arrangementWidth = arrangementWidth;

            //    designDatas.Add(columnDesignData);
            //}

            if (ptRequired < ptProvided)
            {
                double spacingAlongDepth = arrangementDepth / (nAlongDepth - 1);
                double spacingAlongWidth = arrangementWidth / (nAlongWidth - 1);

                ColumnDesignData columnDesignData = new ColumnDesignData();
                columnDesignData.cornerBarCountAlongDepth = nAlongDepth - nCentreBarAlongDepth;
                columnDesignData.cornerBarCountAlongWidth = nAlongWidth - nCentreBarAlongWidth;
                columnDesignData.cornerBarDia = cornerBarDia;
                columnDesignData.centreBarCountAlongDepth = nCentreBarAlongDepth;
                columnDesignData.centreBarCountAlongWidth = nCentreBarAlongWidth;
                columnDesignData.centreBarDia = centreBarDia;
                columnDesignData.ptProvided = ptProvided;
                columnDesignData.spacingAlongDepth = spacingAlongDepth;
                columnDesignData.spacingAlongWidth = spacingAlongWidth;
                columnDesignData.arrangementDepth = arrangementDepth;
                columnDesignData.arrangementWidth = arrangementWidth;

                designDatas.Add(columnDesignData);
            }
        }

        public ColumnDesignData DesignReinforcement(ISection concreteSection, double requiredRebarPt, double maxSpacing)
        {
            IList<ColumnDesignData> designDatas = new List<ColumnDesignData>();
            int linkDia = (int)((ILinkGroup)concreteSection.ReinforcementGroups[0]).BarBundle.Diameter.Millimeters;
            double arrangementDepth =
                ((IRectangleProfile)concreteSection.Profile).Depth.Millimeters -
                2 * concreteSection.Cover.UniformCover.Millimeters -
                2 * linkDia -
                32;// cornerBarDia;
            double arrangementWidth =
                ((IRectangleProfile)concreteSection.Profile).Width.Millimeters -
                2 * concreteSection.Cover.UniformCover.Millimeters -
                2 * linkDia -
                32;// cornerBarDia;
            ColumnDesignData emptyDesignData = new ColumnDesignData();
            emptyDesignData.cornerBarCountAlongDepth = 0;
            emptyDesignData.cornerBarCountAlongWidth = 0;
            emptyDesignData.cornerBarDia = 0;
            emptyDesignData.centreBarCountAlongDepth = 0;
            emptyDesignData.centreBarCountAlongWidth = 0;
            emptyDesignData.centreBarDia = 0;
            emptyDesignData.ptProvided = 0;
            emptyDesignData.spacingAlongDepth = 0;
            emptyDesignData.spacingAlongWidth = 0;
            emptyDesignData.arrangementDepth = arrangementDepth;
            emptyDesignData.arrangementWidth = arrangementWidth;
            designDatas.Add(emptyDesignData);

            IList<int> barDias = new List<int>();
            if (Include16mmBar.Checked)
                barDias.Add(16);
            if (Include20mmBar.Checked)
                barDias.Add(20);
            if (Include25mmBar.Checked)
                barDias.Add(25);
            if (Include32mmBar.Checked)
                barDias.Add(32);
            if (barDias.Count == 0)
                barDias.Add(32);

            IList<double> barSpacings = new List<double>();
            double startSpacing = maxSpacing;
            barSpacings.Add(startSpacing);

            foreach (int cornerBarDia in barDias)
            {                
                foreach (double spacing in barSpacings)
                {
                    int nAlongDepth = (int)Math.Ceiling(arrangementDepth / spacing) + 1;
                    int nAlongWidth = (int)Math.Ceiling(arrangementWidth / spacing) + 1;
                    var maxPossiblePt = 100 * 0.25 * Math.PI * 32 * 32 * ((nAlongDepth + nAlongWidth) * 2 - 4) /
                        ((IRectangleProfile)concreteSection.Profile).Depth.Millimeters /
                        ((IRectangleProfile)concreteSection.Profile).Width.Millimeters;
                    
                    if(requiredRebarPt > maxPossiblePt)
                    {
                        ColumnDesignData columnDesignData = new ColumnDesignData();
                        columnDesignData.cornerBarCountAlongDepth = 0;
                        columnDesignData.cornerBarCountAlongWidth = 0;
                        columnDesignData.cornerBarDia = 0;
                        columnDesignData.centreBarCountAlongDepth = 0;
                        columnDesignData.centreBarCountAlongWidth = 0;
                        columnDesignData.centreBarDia = 0;
                        columnDesignData.ptProvided = 0;
                        columnDesignData.spacingAlongDepth = 0;
                        columnDesignData.spacingAlongWidth = 0;
                        columnDesignData.arrangementDepth = arrangementDepth;
                        columnDesignData.arrangementWidth = arrangementWidth;
                        return columnDesignData;
                    }

                    foreach (int centreBarDia in barDias.Where( _ => _ <= cornerBarDia).ToList())
                    {
                        IList<int> evenCentreBarCount = new List<int>() { 2, 4, 6 };
                        IList<int> oddCentreBarCount = new List<int>() { 1, 3, 5, 7 };

                        if (nAlongDepth % 2 == 0 && nAlongWidth % 2 == 0)
                        {
                            foreach (int centreCountAlongDepth in evenCentreBarCount)
                            {
                                foreach (int centreCountAlongWidth in evenCentreBarCount)
                                {
                                    if (nAlongDepth > centreCountAlongDepth && nAlongWidth > centreCountAlongWidth)
                                    {
                                        AddColumnDesignData(
                                            concreteSection,
                                            designDatas,
                                            nAlongDepth,
                                            centreCountAlongDepth,
                                            nAlongWidth,
                                            centreCountAlongWidth,
                                            cornerBarDia,
                                            centreBarDia,
                                            arrangementDepth,
                                            arrangementWidth,
                                            requiredRebarPt);
                                    }
                                }
                            }
                        }
                        else if (nAlongDepth % 2 == 0 && nAlongWidth % 2 != 0)
                        {
                            foreach (int centreCountAlongDepth in evenCentreBarCount)
                            {
                                foreach (int centreCountAlongWidth in oddCentreBarCount)
                                {
                                    if (nAlongDepth > centreCountAlongDepth && nAlongWidth > centreCountAlongWidth)
                                    {
                                        AddColumnDesignData(
                                            concreteSection,
                                            designDatas,
                                            nAlongDepth,
                                            centreCountAlongDepth,
                                            nAlongWidth,
                                            centreCountAlongWidth,
                                            cornerBarDia,
                                            centreBarDia,
                                            arrangementDepth,
                                            arrangementWidth,
                                            requiredRebarPt);
                                    }
                                }
                            }
                        }
                        else if (nAlongDepth % 2 != 0 && nAlongWidth % 2 != 0)
                        {
                            foreach (int centreCountAlongDepth in oddCentreBarCount)
                            {
                                foreach (int centreCountAlongWidth in oddCentreBarCount)
                                {
                                    if (nAlongDepth > centreCountAlongDepth && nAlongWidth > centreCountAlongWidth)
                                    {
                                        AddColumnDesignData(
                                            concreteSection,
                                            designDatas,
                                            nAlongDepth,
                                            centreCountAlongDepth,
                                            nAlongWidth,
                                            centreCountAlongWidth,
                                            cornerBarDia,
                                            centreBarDia,
                                            arrangementDepth,
                                            arrangementWidth,
                                            requiredRebarPt);
                                    }
                                }
                            }
                        }
                        else if (nAlongDepth % 2 != 0 && nAlongWidth % 2 == 0)
                        {
                            foreach (int centreCountAlongDepth in oddCentreBarCount)
                            {
                                foreach (int centreCountAlongWidth in evenCentreBarCount)
                                {
                                    if (nAlongDepth > centreCountAlongDepth && nAlongWidth > centreCountAlongWidth)
                                    {
                                        AddColumnDesignData(
                                            concreteSection,
                                            designDatas,
                                            nAlongDepth,
                                            centreCountAlongDepth,
                                            nAlongWidth,
                                            centreCountAlongWidth,
                                            cornerBarDia,
                                            centreBarDia,
                                            arrangementDepth,
                                            arrangementWidth,
                                            requiredRebarPt);
                                    }
                                }
                            }
                        }
                    }
                }
            }

            if (designDatas.Count == 1)
                return designDatas[0];

            designDatas.RemoveAt(0);
            double sectionDepth = ((IRectangleProfile)concreteSection.Profile).Depth.Millimeters;
            double sectionWidth = ((IRectangleProfile)concreteSection.Profile).Width.Millimeters;
            if (sectionDepth == sectionWidth)
            {
                foreach (var unequalCornerBarsData in designDatas.Where(
                    _ => _.cornerBarCountAlongDepth != _.cornerBarCountAlongWidth).ToList())
                {
                    designDatas.Remove(unequalCornerBarsData);
                }
            }

            return designDatas.Where(designData => designData.ptProvided == designDatas.Min(_ => _.ptProvided)).ToList()[0];
        }

        public ColumnDesignData DefineSectionReinforcement(ISection concreteSection, double requiredRebarPt, double maxSpacing)
        {
            ColumnDesignData columnDesignData = DesignReinforcement(concreteSection, requiredRebarPt, maxSpacing);
            double arrangementDepth = columnDesignData.arrangementDepth;
            double arrangementWidth = columnDesignData.arrangementWidth;
            int cornerBarDia = columnDesignData.cornerBarDia;
            int centreBarDia = columnDesignData.centreBarDia;
            double spacingAlongDepth = columnDesignData.spacingAlongDepth;
            double spacingAlongWidth = columnDesignData.spacingAlongWidth;
            int centreBarCountAlongDepth = columnDesignData.centreBarCountAlongDepth;
            int centreBarCountAlongWidth = columnDesignData.centreBarCountAlongWidth;
            int cornerBarCountAlongDepth = columnDesignData.cornerBarCountAlongDepth;
            int cornerBarCountAlongWidth = columnDesignData.cornerBarCountAlongWidth;

            if(columnDesignData.ptProvided != 0)
            {
                IBarBundle cornerBar = IBarBundle.Create(
                    GetRebarMaterial(Double.Parse(this.designLongFy.Text)),
                    Length.FromMillimeters(cornerBarDia));
                IBarBundle centreBar = IBarBundle.Create(
                    GetRebarMaterial(Double.Parse(this.designLongFy.Text)),
                    Length.FromMillimeters(centreBarDia));

                ILayerByBarCount layer1 = ILayerByBarCount.Create((int)(0.5 * cornerBarCountAlongWidth), cornerBar);
                ILayerByBarCount layer2 = ILayerByBarCount.Create(centreBarCountAlongWidth, centreBar);
                if ((0.5 * cornerBarCountAlongDepth - 1) > 0)
                {
                    ILayerByBarCount layer3 = ILayerByBarCount.Create((int)(0.5 * cornerBarCountAlongDepth - 1), cornerBar);
                    ILineGroup lineGroup7 = ILineGroup.Create(
                        IPoint.Create(
                            Length.FromMillimeters(-0.5 * arrangementWidth),
                            Length.FromMillimeters(-0.5 * arrangementDepth + spacingAlongDepth)),
                        IPoint.Create(
                            Length.FromMillimeters(-0.5 * arrangementWidth),
                            Length.FromMillimeters(-0.5 * arrangementDepth + (0.5 * cornerBarCountAlongDepth - 1) * spacingAlongDepth)),
                        layer3);
                    ILineGroup lineGroup8 = ILineGroup.Create(
                        IPoint.Create(
                            Length.FromMillimeters(-0.5 * arrangementWidth),
                            Length.FromMillimeters(0.5 * arrangementDepth - spacingAlongDepth)),
                        IPoint.Create(
                            Length.FromMillimeters(-0.5 * arrangementWidth),
                            Length.FromMillimeters(0.5 * arrangementDepth - (0.5 * cornerBarCountAlongDepth - 1) * spacingAlongDepth)),
                        layer3);
                    ILineGroup lineGroup9 = ILineGroup.Create(
                        IPoint.Create(
                            Length.FromMillimeters(0.5 * arrangementWidth),
                            Length.FromMillimeters(-0.5 * arrangementDepth + spacingAlongDepth)),
                        IPoint.Create(
                            Length.FromMillimeters(0.5 * arrangementWidth),
                            Length.FromMillimeters(-0.5 * arrangementDepth + (0.5 * cornerBarCountAlongDepth - 1) * spacingAlongDepth)),
                        layer3);
                    ILineGroup lineGroup10 = ILineGroup.Create(
                        IPoint.Create(
                            Length.FromMillimeters(0.5 * arrangementWidth),
                            Length.FromMillimeters(0.5 * arrangementDepth - spacingAlongDepth)),
                        IPoint.Create(
                            Length.FromMillimeters(0.5 * arrangementWidth),
                            Length.FromMillimeters(0.5 * arrangementDepth - (0.5 * cornerBarCountAlongDepth - 1) * spacingAlongDepth)),
                        layer3);
                    concreteSection.ReinforcementGroups.Add(lineGroup7);
                    concreteSection.ReinforcementGroups.Add(lineGroup8);
                    concreteSection.ReinforcementGroups.Add(lineGroup9);
                    concreteSection.ReinforcementGroups.Add(lineGroup10);
                }
                ILayerByBarCount layer4 = ILayerByBarCount.Create(centreBarCountAlongDepth, centreBar);

                ILineGroup lineGroup1 = ILineGroup.Create(
                    IPoint.Create(
                        Length.FromMillimeters(-0.5 * arrangementWidth),
                        Length.FromMillimeters(-0.5 * arrangementDepth)),
                    IPoint.Create(
                        Length.FromMillimeters(-0.5 * arrangementWidth + (0.5 * cornerBarCountAlongWidth - 1) * spacingAlongWidth),
                        Length.FromMillimeters(-0.5 * arrangementDepth)),
                    layer1);
                ILineGroup lineGroup2 = ILineGroup.Create(
                    IPoint.Create(
                        Length.FromMillimeters(-0.5 * arrangementWidth),
                        Length.FromMillimeters(0.5 * arrangementDepth)),
                    IPoint.Create(
                        Length.FromMillimeters(-0.5 * arrangementWidth + (0.5 * cornerBarCountAlongWidth - 1) * spacingAlongWidth),
                        Length.FromMillimeters(0.5 * arrangementDepth)),
                    layer1);
                ILineGroup lineGroup3 = ILineGroup.Create(
                    IPoint.Create(
                        Length.FromMillimeters(0.5 * arrangementWidth - (0.5 * cornerBarCountAlongWidth - 1) * spacingAlongWidth),
                        Length.FromMillimeters(-0.5 * arrangementDepth)),
                    IPoint.Create(
                        Length.FromMillimeters(0.5 * arrangementWidth),
                        Length.FromMillimeters(-0.5 * arrangementDepth)),
                    layer1);
                ILineGroup lineGroup4 = ILineGroup.Create(
                    IPoint.Create(
                        Length.FromMillimeters(0.5 * arrangementWidth - (0.5 * cornerBarCountAlongWidth - 1) * spacingAlongWidth),
                        Length.FromMillimeters(0.5 * arrangementDepth)),
                    IPoint.Create(
                        Length.FromMillimeters(0.5 * arrangementWidth),
                        Length.FromMillimeters(0.5 * arrangementDepth)),
                    layer1);
                ILineGroup lineGroup5 = ILineGroup.Create(
                    IPoint.Create(
                        Length.FromMillimeters(-0.5 * arrangementWidth + 0.5 * cornerBarCountAlongWidth * spacingAlongWidth),
                        Length.FromMillimeters(-0.5 * arrangementDepth)),
                    IPoint.Create(
                        Length.FromMillimeters(0.5 * arrangementWidth - 0.5 * cornerBarCountAlongWidth * spacingAlongWidth),
                        Length.FromMillimeters(-0.5 * arrangementDepth)),
                    layer2);
                ILineGroup lineGroup6 = ILineGroup.Create(
                    IPoint.Create(
                        Length.FromMillimeters(-0.5 * arrangementWidth + 0.5 * cornerBarCountAlongWidth * spacingAlongWidth),
                        Length.FromMillimeters(0.5 * arrangementDepth)),
                    IPoint.Create(
                        Length.FromMillimeters(0.5 * arrangementWidth - 0.5 * cornerBarCountAlongWidth * spacingAlongWidth),
                        Length.FromMillimeters(0.5 * arrangementDepth)),
                    layer2);
                ILineGroup lineGroup11 = ILineGroup.Create(
                    IPoint.Create(
                        Length.FromMillimeters(-0.5 * arrangementWidth),
                        Length.FromMillimeters(-0.5 * arrangementDepth + 0.5 * cornerBarCountAlongDepth * spacingAlongDepth)),
                    IPoint.Create(
                        Length.FromMillimeters(-0.5 * arrangementWidth),
                        Length.FromMillimeters(0.5 * arrangementDepth - 0.5 * cornerBarCountAlongDepth * spacingAlongDepth)),
                    layer4);
                ILineGroup lineGroup12 = ILineGroup.Create(
                    IPoint.Create(
                        Length.FromMillimeters(0.5 * arrangementWidth),
                        Length.FromMillimeters(-0.5 * arrangementDepth + 0.5 * cornerBarCountAlongDepth * spacingAlongDepth)),
                    IPoint.Create(
                        Length.FromMillimeters(0.5 * arrangementWidth),
                        Length.FromMillimeters(0.5 * arrangementDepth - 0.5 * cornerBarCountAlongDepth * spacingAlongDepth)),
                    layer4);

                concreteSection.ReinforcementGroups.Add(lineGroup1);
                concreteSection.ReinforcementGroups.Add(lineGroup2);
                concreteSection.ReinforcementGroups.Add(lineGroup3);
                concreteSection.ReinforcementGroups.Add(lineGroup4);
                concreteSection.ReinforcementGroups.Add(lineGroup5);
                concreteSection.ReinforcementGroups.Add(lineGroup6);
                concreteSection.ReinforcementGroups.Add(lineGroup11);
                concreteSection.ReinforcementGroups.Add(lineGroup12);

                String rebarDescription;
                if (cornerBarDia == centreBarDia)
                    rebarDescription =
                        (2 * cornerBarCountAlongDepth +
                        2 * cornerBarCountAlongWidth +
                        2 * centreBarCountAlongDepth +
                        2 * centreBarCountAlongWidth - 4).ToString() + "-T" + cornerBarDia.ToString();
                else
                    rebarDescription =
                        (2 * cornerBarCountAlongDepth +
                        2 * cornerBarCountAlongWidth - 4).ToString() + "-T" + cornerBarDia.ToString() + " + " +
                        (2 * centreBarCountAlongDepth +
                        2 * centreBarCountAlongWidth).ToString() + "-T" + centreBarDia.ToString();
                columnDesignData.rebarDescription = rebarDescription;
            }
            else
            {
                columnDesignData.rebarDescription = "";
            }

            return columnDesignData;
        }

        private void Design_Click(object sender, EventArgs e)
        {
            designDataToExcel.Clear();
            var givenDesignDataToExcelList = new List<DesignDataToExcel>();

            for (int design_row = 0; design_row < ResultsTable.Rows.Count - 1; design_row++)
            {
                double depth = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.depth].Value;
                double width = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.width].Value;

                if (depth == 0 || width == 0) continue;

                string givenColumnLabel = ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.column_label].Value.ToString();
                string givenStory = ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.story].Value.ToString();
                double fck = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.fck].Value;
                double cover = double.Parse(rebarCover.Text);
                double ptRequired = ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.etabs_req].Value is string ?
                    double.Parse((string)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.etabs_req].Value) :
                    (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.etabs_req].Value;

                IProfile profile = IRectangleProfile.Create(Length.FromMillimeters(depth), Length.FromMillimeters(width));
                section = ISection.Create(profile, GetSectionMaterial(fck));
                section.Cover = ICover.Create(Length.FromMillimeters(cover));

                List<ColumnForceData> matchingForceData = columnForceData.Where(_ =>
                _.columnLabel == givenColumnLabel && _.story == givenStory).ToList();
                List<ColumnForceData> matchingForceDataGLC = matchingForceData.Where(_ =>
                _.outputCase == ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.load_combination].Value.ToString()).ToList();

                // Loads to check
                double axialAdSec = -1 * (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.p].Value;
                double mMajorAdSec = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m33].Value;
                double mMinorAdSec = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m22].Value;
                var adSecLoad = ILoad.Create(
                    Force.FromKilonewtons(axialAdSec),
                    Moment.FromKilonewtonMeters(mMajorAdSec),
                    Moment.FromKilonewtonMeters(mMinorAdSec));
                List<ILoad> loadsToBeChecked = new List<ILoad>();
                var givenDesignDataToExcel = new DesignDataToExcel();
                givenDesignDataToExcel.inputData.depth = depth;
                givenDesignDataToExcel.inputData.width = width;
                givenDesignDataToExcel.inputData.P = Math.Round(axialAdSec * -1, 2);
                givenDesignDataToExcel.inputData.MMajor = Math.Round(mMajorAdSec, 2);
                givenDesignDataToExcel.inputData.MMinor = Math.Round(mMinorAdSec, 2);
                givenDesignDataToExcel.checkedMajorEccentricLoads = new List<ILoad>();
                givenDesignDataToExcel.checkedMinorEccentricLoads = new List<ILoad>();
                givenDesignDataToExcel.checkedOtherLoads = new List<ILoad>();

                double lengthOfColumn = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.length].Value;
                double previousStation = 0;
                List<double> appliedAxialForGLC = new List<double>();
                int indexOfP = -1;
                foreach (var givenForceData in matchingForceDataGLC)
                {
                    if (givenForceData.station == 0)
                    {
                        indexOfP++;
                        previousStation = 0;
                    }

                    if (Math.Abs(lengthOfColumn * 0.5 - givenForceData.station * 1000) <
                        Math.Abs(lengthOfColumn * 0.5 - previousStation))
                    {
                        if (appliedAxialForGLC.Count - 1 != indexOfP)
                        {
                            appliedAxialForGLC.Add(givenForceData.P);
                        }
                        else
                        {
                            appliedAxialForGLC[indexOfP] = givenForceData.P;
                        }
                    }
                    previousStation = givenForceData.station * 1000;
                }
                indexOfP = -1;
                foreach (var givenForceData in matchingForceDataGLC)
                {
                    if (givenForceData.station == 0)
                        indexOfP++;

                    double ex = OverrideMajorEcc.Checked ? double.Parse(minEccMajor.Text) : 
                        Math.Max((0.65 * lengthOfColumn / 500) + (depth / 30), 20);
                    double minMajorMoment = appliedAxialForGLC[indexOfP] * ex / 1000;

                    ILoad loadOne = ILoad.Create(
                        Force.FromKilonewtons(appliedAxialForGLC[indexOfP]),
                        Moment.FromKilonewtonMeters(minMajorMoment),
                        Moment.FromKilonewtonMeters(givenForceData.MMinor));
                    loadsToBeChecked.Add(loadOne);
                    givenDesignDataToExcel.checkedMajorEccentricLoads.Add(loadOne);
                }
                indexOfP = -1;
                foreach (var givenForceData in matchingForceDataGLC)
                {
                    if (givenForceData.station == 0)
                        indexOfP++;

                    double ey = OverrideMinorEcc.Checked ? double.Parse(minEccMinor.Text) :
                        Math.Max((0.65 * lengthOfColumn / 500) + (width / 30), 20);
                    double minMinorMoment = appliedAxialForGLC[indexOfP] * ey / 1000;

                    ILoad loadTwo = ILoad.Create(
                        Force.FromKilonewtons(appliedAxialForGLC[indexOfP]),
                        Moment.FromKilonewtonMeters(givenForceData.MMajor),
                        Moment.FromKilonewtonMeters(minMinorMoment));
                    loadsToBeChecked.Add(loadTwo);
                    givenDesignDataToExcel.checkedMinorEccentricLoads.Add(loadTwo);
                }
                foreach (var givenForceData in matchingForceData)
                {
                    ILoad load = ILoad.Create(
                        Force.FromKilonewtons(givenForceData.P),
                        Moment.FromKilonewtonMeters(givenForceData.MMajor),
                        Moment.FromKilonewtonMeters(givenForceData.MMinor));
                    loadsToBeChecked.Add(load);
                    givenDesignDataToExcel.checkedOtherLoads.Add(load);
                }

                ColumnDesignData designData = new ColumnDesignData();
                designData.rebarDescription = "";

                IStrengthResult strengthResult = null;
                IStrengthResult otherStrengthResult = null;
                double momentRatio = 0;
                double loadUtilisation = 0;
                double otherMomentRatio = 0;
                double otherLoadUtilisation = 0;
                double linkMaterial = double.Parse(givenLinkFy.Text);
                double linkDia = double.Parse(this.linkDia.Text);
                var linkBar = IBarBundle.Create(GetRebarMaterial(linkMaterial), Length.FromMillimeters(linkDia));
                var linkGroup = ILinkGroup.Create(linkBar);
                givenDesignDataToExcel.inputData.rebarPtEtabs = Math.Round(ptRequired, 2);

                double arrangementDepth = depth - 2 * section.Cover.UniformCover.Millimeters - 2 * linkDia - 32;// cornerBarDia;
                double arrangementWidth = width - 2 * section.Cover.UniformCover.Millimeters - 2 * linkDia - 32;// cornerBarDia;

                int nAlongDepth = (int)Math.Ceiling(arrangementDepth / double.Parse(maximumSpacing.Text)) + 1;
                int nAlongWidth = (int)Math.Ceiling(arrangementWidth / double.Parse(maximumSpacing.Text)) + 1;
                var minPossiblePt = 100 * 0.25 * Math.PI * 16 * 16 * ((nAlongDepth + nAlongWidth) * 2 - 4) / depth / width;

                while (ptRequired != 0)
                {
                    ISection previousDesignSection = ISection.Create(section.Profile, section.Material);
                    previousDesignSection.Cover = section.Cover;
                    foreach (var rebarGroup in section.ReinforcementGroups)
                    {
                        previousDesignSection.ReinforcementGroups.Add(rebarGroup);
                    }

                    var previousDesignData = designData;
                    double previousMomentRatio = momentRatio;
                    double previousLoadUtilisation = loadUtilisation;

                    section.ReinforcementGroups.Clear();
                    section.ReinforcementGroups.Add(linkGroup);

                    designData = DefineSectionReinforcement(section, ptRequired, double.Parse(maximumSpacing.Text));

                    if (designData.ptProvided != 0)
                    {
                        solution = adsecApp.Analyse(section);
                        strengthResult = solution.Strength.Check(adSecLoad);

                        double tensile_rebar_area = 0;
                        var simplified_section = adsecApp.Flatten(section);
                        foreach (var rebar_group in simplified_section.ReinforcementGroups)
                        {
                            if (!(rebar_group is ISingleBars))
                            {
                                continue;
                            }
                            foreach(var rebar_position in ((ISingleBars)rebar_group).Positions)
                            {
                                if (strengthResult.Deformation.StrainAt(rebar_position).MicroStrain > 0)
                                {
                                    tensile_rebar_area += 0.25 * Math.PI * Math.Pow(((ISingleBars)rebar_group).BarBundle.Diameter.Millimeters, 2);
                                }
                            }
                        }
                        double tensile_rebar = 100 * tensile_rebar_area / section.Profile.Area().SquareMillimeters;
                        designData.tensile_rebar = tensile_rebar;

                        var momentRanges = strengthResult.MomentRanges;
                        double ultimateMoment = 0;
                        for (int j = 0; j < momentRanges.Count; j++)
                        {
                            ultimateMoment = Math.Max(momentRanges[j].Max.KilonewtonMeters, ultimateMoment);
                        }
                        double appliedMoment = Math.Sqrt(Math.Pow(mMajorAdSec, 2) + Math.Pow(mMinorAdSec, 2));
                        momentRatio = appliedMoment / ultimateMoment;
                        loadUtilisation = strengthResult.LoadUtilisation.DecimalFractions;

                        foreach (var loadToCheck in loadsToBeChecked)
                        {
                            otherStrengthResult = solution.Strength.Check(loadToCheck);

                            var otherMomentRanges = otherStrengthResult.MomentRanges;
                            double otherUltimateMoment = 0;
                            for (int j = 0; j < otherMomentRanges.Count; j++)
                            {
                                otherUltimateMoment = Math.Max(otherMomentRanges[j].Max.KilonewtonMeters, otherUltimateMoment);
                            }
                            double otherAppliedMoment = Math.Sqrt(
                                Math.Pow(loadToCheck.YY.KilonewtonMeters, 2) +
                                Math.Pow(loadToCheck.ZZ.KilonewtonMeters, 2));

                            otherMomentRatio = Math.Max(otherMomentRatio, otherAppliedMoment / otherUltimateMoment);
                            otherLoadUtilisation = Math.Max(otherLoadUtilisation, otherStrengthResult.LoadUtilisation.DecimalFractions);
                        }

                        double limit = double.Parse(this.UtilisationLimit.Text) / 100;
                        if (limit > 1 || limit <= 0)
                            limit = 1;

                        if (momentRatio > limit || loadUtilisation > limit || otherMomentRatio > limit || otherLoadUtilisation > limit)
                        {
                            if (momentRatio > 1 || loadUtilisation > 1 || otherMomentRatio > 1 || otherLoadUtilisation > 1)
                            {
                                section = previousDesignSection;
                                designData = previousDesignData;
                                momentRatio = previousMomentRatio;
                                loadUtilisation = previousLoadUtilisation;
                            }

                            givenDesignDataToExcel.loadUtilisation = loadUtilisation;
                            givenDesignDataToExcel.momentRatio = momentRatio;
                            givenDesignDataToExcel.DesignData = designData;
                            givenDesignDataToExcel.designSection = section;
                            givenDesignDataToExcelList.Add(givenDesignDataToExcel);
                            
                            ptRequired = 0;
                        }
                        else
                        {
                            if ((minPossiblePt < ptRequired + 0.05) && !(ptRequired < 0.805))
                            {
                                ptRequired = ptRequired - 0.05;
                            }
                            else
                            {
                                givenDesignDataToExcel.loadUtilisation = loadUtilisation;
                                givenDesignDataToExcel.momentRatio = momentRatio;
                                givenDesignDataToExcel.DesignData = designData;
                                givenDesignDataToExcel.designSection = section;
                                givenDesignDataToExcelList.Add(givenDesignDataToExcel);

                                ptRequired = 0;
                            }
                        }
                    }
                    else
                    {
                        givenDesignDataToExcel.loadUtilisation = loadUtilisation;
                        givenDesignDataToExcel.momentRatio = momentRatio;
                        givenDesignDataToExcel.DesignData = designData;
                        givenDesignDataToExcel.designSection = section;
                        givenDesignDataToExcelList.Add(givenDesignDataToExcel);

                        ptRequired = 0;
                    }
                }
            }

            if(ResultsTable.Rows.Count > 2 && UniformRebar.Checked)
            {
                double previousDepth = 0;
                double previousWidth = 0;
                double maxPT = 0;
                int startIndex = 0;
                ColumnDesignData maxPTData = new ColumnDesignData();
                Oasys.Collections.IList<IGroup> maxPTRebarGroup = null;
                for (int i = 0; i < givenDesignDataToExcelList.Count; i++)
                {
                    if (i == 0)
                    {
                        previousDepth = givenDesignDataToExcelList[i].inputData.depth;
                        previousWidth = givenDesignDataToExcelList[i].inputData.width;
                        maxPT = givenDesignDataToExcelList[i].DesignData.ptProvided;
                        maxPTData = givenDesignDataToExcelList[i].DesignData;
                        maxPTRebarGroup = givenDesignDataToExcelList[i].designSection.ReinforcementGroups;
                        continue;
                    }

                    if (previousDepth == givenDesignDataToExcelList[i].inputData.depth &&
                        previousWidth == givenDesignDataToExcelList[i].inputData.width)
                    {
                        if (maxPT < givenDesignDataToExcelList[i].DesignData.ptProvided)
                        {
                            maxPT = givenDesignDataToExcelList[i].DesignData.ptProvided;
                            maxPTData = givenDesignDataToExcelList[i].DesignData;
                            maxPTRebarGroup = 
                                givenDesignDataToExcelList[i].designSection.ReinforcementGroups;
                        }

                        if (i == (givenDesignDataToExcelList.Count - 1))
                        {
                            for (int j = startIndex; j < givenDesignDataToExcelList.Count; j++)
                            {
                                var dataToUpdate = givenDesignDataToExcelList[j];
                                dataToUpdate.DesignData = maxPTData;
                                dataToUpdate.designSection.ReinforcementGroups = maxPTRebarGroup;
                                if(maxPTData.ptProvided !=0)
                                {
                                    var solution = adsecApp.Analyse(dataToUpdate.designSection);
                                    var strengthResult = solution.Strength.Check(
                                        ILoad.Create(
                                            Force.FromKilonewtons(-1 * dataToUpdate.inputData.P),
                                            Moment.FromKilonewtonMeters(dataToUpdate.inputData.MMajor),
                                            Moment.FromKilonewtonMeters(dataToUpdate.inputData.MMinor)));

                                    double tensile_rebar_area = 0;
                                    var simplified_section = adsecApp.Flatten(dataToUpdate.designSection);
                                    foreach (var rebar_group in simplified_section.ReinforcementGroups)
                                    {
                                        if (!(rebar_group is ISingleBars))
                                        {
                                            continue;
                                        }
                                        foreach (var rebar_position in ((ISingleBars)rebar_group).Positions)
                                        {
                                            if (strengthResult.Deformation.StrainAt(rebar_position).MicroStrain > 0)
                                            {
                                                tensile_rebar_area +=
                                                    0.25 * Math.PI * Math.Pow(((ISingleBars)rebar_group).BarBundle.Diameter.Millimeters, 2);
                                            }
                                        }
                                    }
                                    double tensile_rebar = 100 * tensile_rebar_area / section.Profile.Area().SquareMillimeters;
                                    dataToUpdate.DesignData.tensile_rebar = tensile_rebar;

                                    var momentRanges = strengthResult.MomentRanges;
                                    double ultimateMoment = 0;
                                    for (int k = 0; k < momentRanges.Count; k++)
                                    {
                                        ultimateMoment = Math.Max(momentRanges[k].Max.KilonewtonMeters, ultimateMoment);
                                    }
                                    double appliedMoment =
                                        Math.Sqrt(Math.Pow(dataToUpdate.inputData.MMajor, 2) + Math.Pow(dataToUpdate.inputData.MMinor, 2));
                                    var momentRatio = appliedMoment / ultimateMoment;
                                    var loadUtilisation = strengthResult.LoadUtilisation.DecimalFractions;
                                    dataToUpdate.loadUtilisation = loadUtilisation;
                                    dataToUpdate.momentRatio = momentRatio;
                                }
                                else
                                {
                                    dataToUpdate.loadUtilisation = 0;
                                    dataToUpdate.momentRatio = 0;
                                }

                                givenDesignDataToExcelList[j] = dataToUpdate;
                            }
                        }
                    }
                    else
                    {
                        for (int j = startIndex; j < i; j++)
                        {
                            var dataToUpdate = givenDesignDataToExcelList[j];
                            dataToUpdate.DesignData = maxPTData;
                            dataToUpdate.designSection.ReinforcementGroups = maxPTRebarGroup;
                            if(maxPTData.ptProvided !=0)
                            {
                                var solution = adsecApp.Analyse(dataToUpdate.designSection);
                                var strengthResult = solution.Strength.Check(
                                    ILoad.Create(
                                        Force.FromKilonewtons(-1 * dataToUpdate.inputData.P),
                                        Moment.FromKilonewtonMeters(dataToUpdate.inputData.MMajor),
                                        Moment.FromKilonewtonMeters(dataToUpdate.inputData.MMinor)));

                                double tensile_rebar_area = 0;
                                var simplified_section = adsecApp.Flatten(dataToUpdate.designSection);
                                foreach (var rebar_group in simplified_section.ReinforcementGroups)
                                {
                                    if (!(rebar_group is ISingleBars))
                                    {
                                        continue;
                                    }
                                    foreach (var rebar_position in ((ISingleBars)rebar_group).Positions)
                                    {
                                        if (strengthResult.Deformation.StrainAt(rebar_position).MicroStrain > 0)
                                        {
                                            tensile_rebar_area += 
                                                0.25 * Math.PI * Math.Pow(((ISingleBars)rebar_group).BarBundle.Diameter.Millimeters, 2);
                                        }
                                    }
                                }
                                double tensile_rebar = 100 * tensile_rebar_area / section.Profile.Area().SquareMillimeters;
                                dataToUpdate.DesignData.tensile_rebar = tensile_rebar;

                                var momentRanges = strengthResult.MomentRanges;
                                double ultimateMoment = 0;
                                for (int k = 0; k < momentRanges.Count; k++)
                                {
                                    ultimateMoment = Math.Max(momentRanges[k].Max.KilonewtonMeters, ultimateMoment);
                                }
                                double appliedMoment =
                                    Math.Sqrt(Math.Pow(dataToUpdate.inputData.MMajor, 2) + Math.Pow(dataToUpdate.inputData.MMinor, 2));
                                var momentRatio = appliedMoment / ultimateMoment;
                                var loadUtilisation = strengthResult.LoadUtilisation.DecimalFractions;
                                dataToUpdate.loadUtilisation = loadUtilisation;
                                dataToUpdate.momentRatio = momentRatio;
                            }
                            else
                            {
                                dataToUpdate.loadUtilisation = 0;
                                dataToUpdate.momentRatio = 0;
                            }
                            givenDesignDataToExcelList[j] = dataToUpdate;
                        }

                        startIndex = i;
                        maxPT = givenDesignDataToExcelList[i].DesignData.ptProvided;
                        maxPTData = givenDesignDataToExcelList[i].DesignData;
                        maxPTRebarGroup = givenDesignDataToExcelList[i].designSection.ReinforcementGroups;
                        previousDepth = givenDesignDataToExcelList[i].inputData.depth;
                        previousWidth = givenDesignDataToExcelList[i].inputData.width;
                    }
                }
            }

            for (int design_row = 0; design_row < ResultsTable.Rows.Count - 1; design_row++)
            {
                double depth = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.depth].Value;
                double width = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.width].Value;
                
                if (depth == 0 || width == 0) continue;

                var givenDesignDataToExcel = givenDesignDataToExcelList[0];
                if (givenDesignDataToExcelList.Count > 0)
                {
                    givenDesignDataToExcelList.RemoveAt(0);
                }
                var designData = givenDesignDataToExcel.DesignData;
                var loadUtilisation = givenDesignDataToExcel.loadUtilisation;
                var momentRatio = givenDesignDataToExcel.momentRatio;
                var section = givenDesignDataToExcel.designSection;
                double cover = double.Parse(rebarCover.Text);
                double linkMaterial = double.Parse(givenLinkFy.Text);
                double linkDia = double.Parse(this.linkDia.Text);

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.load_util].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.load_util].Style.ForeColor = Color.Black;

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m_ratio].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m_ratio].Style.ForeColor = Color.Black;

                if (loadUtilisation > 1)
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.load_util].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.load_util].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.load_util].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.load_util].Style.ForeColor = Color.White;
                }

                if (momentRatio > 1)
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m_ratio].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m_ratio].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m_ratio].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m_ratio].Style.ForeColor = Color.White;
                }

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.load_util].Value = Math.Round(loadUtilisation, 2);
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.m_ratio].Value = Math.Round(momentRatio, 2);
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.rebar_description].Value = designData.rebarDescription;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.adsec_pro].Value = Math.Round(designData.ptProvided, 2);
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tensile_rebar].Value = Math.Round(designData.tensile_rebar, 2);

                double factoredShearForceAlongY = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.vy].Value;
                double factoredShearForceAlongX = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.vx].Value;
                double effectiveDepth = 0;
                double effectiveWidth = 0;

                double tauVAlongY = 0;
                double tauVAlongX = 0;
                double tauC = 0;
                double tauCMax = CalculateTauCMax((double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.fck].Value);
                int legsAlongY = 0;
                int legsAlongX = 0;
                double vusAlongY = 0;
                double vusAlongX = 0;
                double asvAlongY = 0;
                double asvAlongX = 0;
                double spacingRequiredAlongY = 0;
                double spacingRequiredAlongX = 0;
                double spacingRequired = 0;
                double minAsvRequired = 0;
                double nonConfiningSpacingOne = 0;
                double nonConfiningSpacingTwo = 0;
                double spacingTwo = 0;
                double spacingOne = 0;
                double spacingThree = 0;
                double maxConfiningSpacing = 0;
                double Ag = 0;
                double Ak = 0;
                double AshOne = 0;
                double AshTwo = 0;
                double AshRequired = 0;

                if (designData.ptProvided != 0)
                {
                    // Shear reinf. calculation as per IS 456
                    effectiveDepth = depth - cover - linkDia - designData.cornerBarDia * 0.5;
                    effectiveWidth = width - cover - linkDia - designData.cornerBarDia * 0.5;
                    tauVAlongY = 1000 * factoredShearForceAlongY / effectiveDepth / width;
                    tauVAlongX = 1000 * factoredShearForceAlongX / effectiveWidth / depth;
                    double fck = (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.fck].Value;
                    tauC = CalculateTauC(designData.tensile_rebar, fck);

                    if(IncreaseTauC.Checked)
                    {
                        double axial_force = 1000 * (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.p].Value;
                        if(axial_force > 0) tauC = tauC * Math.Min(1.5, 1 + 3 * axial_force / depth / width / fck);
                    }

                    legsAlongY = designData.cornerBarCountAlongWidth + designData.centreBarCountAlongWidth;
                    legsAlongX = designData.cornerBarCountAlongDepth + designData.centreBarCountAlongDepth;
                    vusAlongY = (1000 * factoredShearForceAlongY - tauC * effectiveDepth * width);
                    vusAlongX = (1000 * factoredShearForceAlongX - tauC * effectiveWidth * depth);
                    asvAlongY = (Math.PI / 4) * linkDia * linkDia * legsAlongY;
                    asvAlongX = (Math.PI / 4) * linkDia * linkDia * legsAlongX;
                    nonConfiningSpacingTwo = Math.Min(Math.Min(depth, width), Math.Min(16 * designData.centreBarDia, 300));
                    spacingRequiredAlongY = 0.87 * double.Parse(givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY > 0 ?
                        Math.Min(0.87 * double.Parse(givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY, nonConfiningSpacingTwo):
                        nonConfiningSpacingTwo;
                    spacingRequiredAlongX = 0.87 * double.Parse(givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX > 0 ?
                        Math.Min(0.87 * double.Parse(givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX, nonConfiningSpacingTwo) :
                        nonConfiningSpacingTwo; // IS456 pg.49
                    spacingRequired = Math.Min(spacingRequiredAlongY, spacingRequiredAlongX);
                    minAsvRequired = 
                        0.4 * width * double.Parse(nonConfiningSpacing.Text) / 0.87 / double.Parse(givenLinkFy.Text); // IS456 pg.48

                    nonConfiningSpacingOne =
                        0.87 * double.Parse(givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY > 0 &&
                        0.87 * double.Parse(givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX > 0 ?
                        Math.Min(0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY,
                        0.87 * double.Parse(givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX) : (
                        0.87 * double.Parse(givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY > 0 ?
                        0.87 * double.Parse(givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY : (
                        0.87 * double.Parse(givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX > 0 ?
                        0.87 * double.Parse(givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX : 0));

                    // Shear reinf. calculation as per IS 13920
                    spacingTwo = designData.centreBarDia == 16 ? 6 * 20 : 6 * designData.centreBarDia;
                    spacingOne = spacingTwo; // Math.Min(depth, width) / 4;
                    spacingThree = spacingTwo; // 100;
                    maxConfiningSpacing = Math.Min(spacingOne, Math.Min(spacingTwo, spacingThree));

                    Ag = depth * width;
                    Ak = (depth - 2 * section.Cover.UniformCover.Millimeters) *
                        (width - 2 * section.Cover.UniformCover.Millimeters);
                    AshOne = 0.18 * double.Parse(this.confiningSpacing.Text) *
                        Math.Max(designData.spacingAlongDepth, designData.spacingAlongWidth) *
                        (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.fck].Value *
                        (Ag / Ak - 1) /
                        double.Parse(this.givenLinkFy.Text);
                    AshTwo = 0.05 * double.Parse(this.confiningSpacing.Text) *
                        Math.Max(designData.spacingAlongDepth, designData.spacingAlongWidth) *
                        (double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.fck].Value /
                        double.Parse(this.givenLinkFy.Text);
                    if(IsGravityColumn.Checked)
                    {
                        AshOne = 0.5 * AshOne;
                        AshTwo = 0.5 * AshTwo;
                    }
                    AshRequired = Math.Max(AshOne, AshTwo);
                }

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvx].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvx].Style.ForeColor = Color.Black;

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvx].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvx].Style.ForeColor = Color.Black;

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvy].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvy].Style.ForeColor = Color.Black;

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvy].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvy].Style.ForeColor = Color.Black;

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_nc].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_nc].Style.ForeColor = Color.Black;

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_c].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_c].Style.ForeColor = Color.Black;

                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.ash_req].Style.BackColor = Color.White;
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.ash_req].Style.ForeColor = Color.Black;

                designTauCMax.Text = Math.Round(tauCMax, 1).ToString();
                designMinAsv.Text = Math.Ceiling(minAsvRequired).ToString();
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauc].Value = Math.Round(tauC, 2);
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvx].Value = Math.Round(tauVAlongX, 2);
                if(tauCMax < tauVAlongX)
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvx].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvx].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvx].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvx].Style.ForeColor = Color.White;
                }
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvx].Value = Math.Floor(asvAlongX);
                if (asvAlongX < minAsvRequired)
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvx].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvx].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvx].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvx].Style.ForeColor = Color.White;
                }
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvy].Value = Math.Round(tauVAlongY, 2);
                if (tauCMax < tauVAlongY)
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvy].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvy].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvy].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.tauvy].Style.ForeColor = Color.White;
                }
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvy].Value = Math.Floor(asvAlongY);
                if (asvAlongY < minAsvRequired)
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvy].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvy].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvy].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.asvy].Style.ForeColor = Color.White;
                }
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_nc].Value = Math.Floor(spacingRequired);
                if (spacingRequired < double.Parse(this.nonConfiningSpacing.Text))
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_nc].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_nc].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_nc].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_nc].Style.ForeColor = Color.White;
                }
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_c].Value = Math.Floor(maxConfiningSpacing);
                if (maxConfiningSpacing < double.Parse(this.confiningSpacing.Text))
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_c].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_c].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_c].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.max_c].Style.ForeColor = Color.White;
                }
                ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.ash_req].Value = Math.Ceiling(AshRequired);
                this.AshProvided.Text = Math.Floor(0.25 * Math.PI * linkDia * linkDia).ToString();
                if (Math.Floor(0.25 * Math.PI * linkDia * linkDia) < AshRequired)
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.ash_req].Style.BackColor = Color.Red;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.ash_req].Style.ForeColor = Color.White;
                }
                else
                {
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.ash_req].Style.BackColor = Color.Green;
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.ash_req].Style.ForeColor = Color.White;
                }

                givenDesignDataToExcel.inputData.columnLabel = 
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.column_label].Value.ToString();
                givenDesignDataToExcel.inputData.story = 
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.story].Value.ToString();
                givenDesignDataToExcel.inputData.location = 
                    ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.location].Value.ToString();
                givenDesignDataToExcel.inputData.diameter = 0;
                givenDesignDataToExcel.inputData.fck = 
                    Math.Round((double)ResultsTable.Rows[design_row].Cells[ResultTableColumnIndex.fck].Value, 0);
                givenDesignDataToExcel.DesignData = designData;
                givenDesignDataToExcel.factoredShearForceAlongY = Math.Round(factoredShearForceAlongY, 2);
                givenDesignDataToExcel.factoredShearForceAlongX = Math.Round(factoredShearForceAlongX, 2);
                givenDesignDataToExcel.effectiveDepth = Math.Round(effectiveDepth, 0);
                givenDesignDataToExcel.effectiveWidth = Math.Round(effectiveWidth, 0);
                givenDesignDataToExcel.tauVAlongY = Math.Round(tauVAlongY, 2);
                givenDesignDataToExcel.tauVAlongX = Math.Round(tauVAlongX, 2);
                givenDesignDataToExcel.tauC = Math.Round(tauC, 2);
                givenDesignDataToExcel.tauCMax = Math.Round(tauCMax, 2);
                givenDesignDataToExcel.longitudinalFy = Math.Round(double.Parse(designLongFy.Text), 0);
                givenDesignDataToExcel.linkFy = Math.Round(linkMaterial, 0);
                givenDesignDataToExcel.linkDia = Math.Round(linkDia, 0);
                givenDesignDataToExcel.legsAlongY = legsAlongY;
                givenDesignDataToExcel.legsAlongX = legsAlongX;
                givenDesignDataToExcel.VusAlongY = Math.Round(vusAlongY / 1000, 2);
                givenDesignDataToExcel.VusAlongX = Math.Round(vusAlongX / 1000, 2);
                givenDesignDataToExcel.asvProvidedAlongX = Math.Floor(asvAlongX);
                givenDesignDataToExcel.asvProvidedAlongY = Math.Floor(asvAlongY);
                givenDesignDataToExcel.nonConfiningSpacingOne = Math.Floor(nonConfiningSpacingOne);
                givenDesignDataToExcel.nonConfiningSpacingTwo = Math.Floor(nonConfiningSpacingTwo);
                givenDesignDataToExcel.minNonConfiningAsvRequired = Math.Ceiling(minAsvRequired);
                givenDesignDataToExcel.nonConfiningSpacingProvided = Math.Floor(double.Parse(nonConfiningSpacing.Text));
                givenDesignDataToExcel.clearCover = Math.Round(cover, 0);
                givenDesignDataToExcel.confinfingSpacingOne = Math.Floor(spacingOne);
                givenDesignDataToExcel.confinfingSpacingTwo = Math.Floor(spacingTwo);
                givenDesignDataToExcel.confinfingSpacingThree = Math.Floor(spacingThree);
                givenDesignDataToExcel.maxConfiningSpacingRequired = Math.Floor(maxConfiningSpacing);
                givenDesignDataToExcel.Ag = Math.Round(Ag, 0);
                givenDesignDataToExcel.Ak = Math.Round(Ak, 0);
                givenDesignDataToExcel.h = Math.Ceiling(Math.Max(designData.spacingAlongDepth, designData.spacingAlongWidth));
                givenDesignDataToExcel.AshOne = Math.Ceiling(AshOne);
                givenDesignDataToExcel.AshTwo = Math.Ceiling(AshTwo);
                givenDesignDataToExcel.minAshRequired = Math.Ceiling(AshRequired);
                givenDesignDataToExcel.AshProvided = Math.Floor(0.25 * Math.PI * linkDia * linkDia);
                givenDesignDataToExcel.confiningSpacingProvided = Math.Ceiling(double.Parse(this.confiningSpacing.Text));
                designDataToExcel.Add(givenDesignDataToExcel);
            }
        }

        private void columnsToShow_SelectedIndexChanged(object sender, EventArgs e)
        {
            var checkedColumns = this.columnsToShow.CheckedIndices;
            foreach(var checkedColumn in checkedColumns)
            {
                ResultsTable.Columns[(int)checkedColumn].Visible = true;
            }
            for(int i=0; i< this.columnsToShow.Items.Count; i++)
            {
                if(!checkedColumns.Contains(i))
                {
                    ResultsTable.Columns[i].Visible = false;
                }
            }
        }

        private void ExtractETABSTables_Click(object sender, EventArgs e)
        {
            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkBook = null;

            excelApplication = new Excel.Application();
            excelApplication.Visible = true;
            string @excelPath = this.designInput.Text;
            excelWorkBook = excelApplication.Workbooks.Open(excelPath);

            ExtractEtabsForceTables(excelWorkBook);

            excelWorkBook.Close();
            excelApplication.Quit();
            excelApplication = null;
            excelWorkBook = null;
        }

        private void BrowseInputFIle_Click(object sender, EventArgs e)
        {
            var file_browse_dialog = new OpenFileDialog();
            file_browse_dialog.Title = "Browse Input File";
            var dialog_result = file_browse_dialog.ShowDialog();
            if (dialog_result == DialogResult.OK)
            {
                designInput.ForeColor = Color.Black;
                designInput.Text = file_browse_dialog.FileName;
            }
        }

        private void designInput_Enter(object sender, EventArgs e)
        {
            if (designInput.Text == "Enter Input File Path...")
            {
                designInput.ForeColor = Color.Black;
                designInput.Text = "";
            }
        }

        private void designInput_Leave(object sender, EventArgs e)
        {
            if (designInput.Text.Length == 0)
            {
                designInput.ForeColor = Color.Gray;
                designInput.Text = "Enter Input File Path...";
            }
        }

        private void UpdateDesignRows_Click(object sender, EventArgs e)
        {
            if (DesignStories.CheckedItems.Contains("All stories"))
            {
                for (int i = 0; i < DesignStories.Items.Count; i++)
                {
                    DesignStories.SetItemChecked(i, true);
                }
            }
            for (int design_row_index = 0; design_row_index < ResultsTable.Rows.Count - 1; design_row_index++)
            {
                if (ResultsTable.Rows[design_row_index].Cells[ResultTableColumnIndex.column_label].Value.ToString() == columnsToDesign.Text)
                {
                    ResultsTable.Rows.Remove(ResultsTable.Rows[design_row_index]);
                    design_row_index--;
                }
            }

            double maxRebarPtEtabs = 0;
            double fckValue = 0;
            double depthValue = 0;
            double widthValue = 0;
            double diameterValue = 0;
            double pTopValue = 0;
            double mMajorTopValue = 0;
            double mMinorTopValue = 0;
            double pBottomValue = 0;
            double mMajorBottomValue = 0;
            double mMinorBottomValue = 0;
            double rebarPtEtabsValue = 0;
            double rebarPtEtabsBottomValue = 0;
            double vAlongXTop = 0;
            double vAlongYTop = 0;
            double vAlongXBottom = 0;
            double vAlongYBottom = 0;
            string governing_combo = "";

            var thisColumnInputDatas = columnInputData.Where(
                _ => _.columnLabel == columnsToDesign.Text).ToList();
            thisColumnInputDatas.Reverse();

            foreach (var thisColumnInputData in thisColumnInputDatas)
            {
                if (maxRebarPtEtabs < thisColumnInputData.rebarPtEtabs)
                {
                    maxRebarPtEtabs = thisColumnInputData.rebarPtEtabs;
                }

                if (DesignStories.CheckedItems.Contains(thisColumnInputData.story))
                {
                    fckValue = thisColumnInputData.fck;
                    depthValue = thisColumnInputData.depth;
                    widthValue = thisColumnInputData.width;
                    diameterValue = thisColumnInputData.diameter;

                    List<ColumnForceData> matchingForceData = columnForceData.Where(_ =>
                    (_.columnLabel == thisColumnInputData.columnLabel &&
                    _.story == thisColumnInputData.story)).ToList();
                    double maxVAlongXForce = matchingForceData.Count > 0 ? matchingForceData.Max(_ => _.VAlongX) : 0;
                    double maxVAlongYForce = matchingForceData.Count > 0 ? matchingForceData.Max(_ => _.VAlongY) : 0;
                    List<ColumnShearData> matchingShearData = columnShearData.Where(_ =>
                    (_.columnLabel == thisColumnInputData.columnLabel &&
                    _.story == thisColumnInputData.story &&
                    _.location == thisColumnInputData.location)).ToList();

                    if (thisColumnInputData.location == "Bottom")
                    {
                        pBottomValue = thisColumnInputData.P;
                        mMajorBottomValue = thisColumnInputData.MMajor;
                        mMinorBottomValue = thisColumnInputData.MMinor;
                        rebarPtEtabsBottomValue = thisColumnInputData.rebarPtEtabs;
                        vAlongXBottom = matchingShearData.Count > 0 ?
                            Math.Max(matchingShearData[0].maxVAlongX, maxVAlongXForce) : maxVAlongXForce;
                        vAlongYBottom = matchingShearData.Count > 0 ?
                            Math.Max(matchingShearData[0].maxVAlongY, maxVAlongYForce) : maxVAlongYForce;
                        governing_combo = thisColumnInputData.governingCombo;
                    }
                    else if (thisColumnInputData.location == "Top")
                    {
                        pTopValue = thisColumnInputData.P;
                        mMajorTopValue = thisColumnInputData.MMajor;
                        mMinorTopValue = thisColumnInputData.MMinor;
                        rebarPtEtabsValue = thisColumnInputData.rebarPtEtabs;
                        vAlongXTop = matchingShearData.Count > 0 ?
                            Math.Max(matchingShearData[0].maxVAlongX, maxVAlongXForce) : maxVAlongXForce;
                        vAlongYTop = matchingShearData.Count > 0 ?
                            Math.Max(matchingShearData[0].maxVAlongY, maxVAlongYForce) : maxVAlongYForce;

                        if (rebarPtEtabsBottomValue < rebarPtEtabsValue)
                        {
                            ResultsTable.Rows.Add(
                                columnsToDesign.Text,
                                thisColumnInputData.story,
                                thisColumnInputData.story_elevation,
                                "Top",
                                Math.Round(thisColumnInputData.fck, 0),
                                Math.Round(thisColumnInputData.depth, 0),
                                Math.Round(thisColumnInputData.width, 0),
                                Math.Round(thisColumnInputData.length, 0),
                                thisColumnInputData.governingCombo,
                                Math.Round(pTopValue, 0),
                                Math.Round(mMajorTopValue, 0),
                                Math.Round(mMinorTopValue, 0),
                                0.0,
                                0.0,
                                Math.Round(rebarPtEtabsValue, 2),
                                0.0,
                                0.0,
                                "",
                                Math.Max(Math.Round(vAlongYTop, 0), Math.Round(vAlongYBottom, 0)),
                                Math.Max(Math.Round(vAlongXTop, 0), Math.Round(vAlongXBottom, 0)),
                                0.0,
                                0.0,
                                0.0,
                                0.0,
                                0.0,
                                0.0,
                                0.0,
                                0.0);
                        }
                        else
                        {
                            ResultsTable.Rows.Add(
                                columnsToDesign.Text,
                                thisColumnInputData.story,
                                thisColumnInputData.story_elevation,
                                "Bottom",
                                Math.Round(thisColumnInputData.fck, 0),
                                Math.Round(thisColumnInputData.depth, 0),
                                Math.Round(thisColumnInputData.width, 0),
                                Math.Round(thisColumnInputData.length, 0),
                                governing_combo,
                                Math.Round(pBottomValue, 0),
                                Math.Round(mMajorBottomValue, 0),
                                Math.Round(mMinorBottomValue, 0),
                                0.0,
                                0.0,
                                Math.Round(rebarPtEtabsBottomValue, 2),
                                0.0,
                                0.0,
                                "",
                                Math.Max(Math.Round(vAlongYTop, 0), Math.Round(vAlongYBottom, 0)),
                                Math.Max(Math.Round(vAlongXTop, 0), Math.Round(vAlongXBottom, 0)),
                                0.0,
                                0.0,
                                0.0,
                                0.0,
                                0.0,
                                0.0,
                                0.0,
                                0.0);
                        }
                    }
                }
            }

            //maxRebarPtEtabs *= 100;
            this.maxEtabsRebarPt.Text = maxRebarPtEtabs.ToString();
        }

        private void rebarCover_Leave(object sender, EventArgs e)
        {
            if (!double.TryParse(rebarCover.Text, out _) || double.Parse(rebarCover.Text) <= 0)
            {
                errorProvider1.SetError(rebarCover, "Cover must be a positive number");
                rebarCover.Text = (40).ToString();
                return;
            }
            else
            {
                errorProvider1.SetError(rebarCover, "");
                errorProvider1.Clear();
            }
        }

        private void ClearDesignTable_Click(object sender, EventArgs e)
        {
            ResultsTable.Rows.Clear();
        }
    }
}