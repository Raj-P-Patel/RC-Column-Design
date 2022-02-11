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

namespace Column_Design
{
    public struct ColumnInputData
    {
        public string columnLabel;
        public string story;
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
        public List<String> uniqueStories = new List<String>();
        public List<String> uniqueColumnLabels = new List<String>();

        public IAdSec adsecApp = IAdSec.Create(IS456.Edition_2000);
        public ISection section = null;
        public ISolution solution = null;
        public List<ILineGroup> lineGroups = new List<ILineGroup>();

        public ColumnDesignForm()
        {
            InitializeComponent();
            this.columnsToShow.Items.Add("Story");
            this.columnsToShow.Items.Add("Location");
            this.columnsToShow.Items.Add("fck");
            this.columnsToShow.Items.Add("Depth");
            this.columnsToShow.Items.Add("Width");
            this.columnsToShow.Items.Add("Length");
            this.columnsToShow.Items.Add("Load Combination");
            this.columnsToShow.Items.Add("P");
            this.columnsToShow.Items.Add("MMajor");
            this.columnsToShow.Items.Add("MMinor");
            this.columnsToShow.Items.Add("Lutil");
            this.columnsToShow.Items.Add("MUtil");
            this.columnsToShow.Items.Add("RebarDesc");
            this.columnsToShow.Items.Add("AdSecPt");
            this.columnsToShow.Items.Add("ETABSPt");
            this.columnsToShow.Items.Add("Vy");
            this.columnsToShow.Items.Add("Vx");
            this.columnsToShow.Items.Add("AsvY");
            this.columnsToShow.Items.Add("AsvX");
            this.columnsToShow.Items.Add("TauC");
            this.columnsToShow.Items.Add("TauVY");
            this.columnsToShow.Items.Add("TauVX");
            this.columnsToShow.Items.Add("maxNC");
            this.columnsToShow.Items.Add("maxC");
            this.columnsToShow.Items.Add("minAsh");
            for(int i=0; i< this.columnsToShow.Items.Count; i++)
            {
                this.columnsToShow.SetItemChecked(i, true);
            }

            this.minEccOverride.Items.Add("ecc. major (mm)");
            this.minEccOverride.Items.Add("ecc. minor (mm)");
            for (int i = 0; i < this.minEccOverride.Items.Count; i++)
            {
                this.minEccOverride.SetItemChecked(i, false);
            }
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
                    sheet.UsedRange.ClearContents();
                    designInputSheet = sheet;
                }
                else if (sheet.Name == "Concrete Column Shear Envelope")
                {
                    sheet.UsedRange.ClearContents();
                    columnShearSheet = sheet;
                }
                else if (sheet.Name == "Column Forces")
                {
                    sheet.UsedRange.ClearContents();
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

                designInputSheet.Cells[2, i+1].Value2 = PMMFieldsKeysIncluded[i];
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
                    if(!dimHeadingsAdded)
                    {
                        designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 1].Value2 = "Depth";
                        designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 2].Value2 = "Width";
                        designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 3].Value2 = "Diameter";
                        designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 4].Value2 = "fck";
                        designInputSheet.Cells[2, PMMFieldsKeysIncluded.Length + 5].Value2 = "Length";
                        dimHeadingsAdded = true;
                    }

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

            for (int i = 0; i < SFFieldsKeysIncluded.Length; i++)
            {
                columnShearSheet.Cells[2, i + 1].Value2 = SFFieldsKeysIncluded[i];
            }

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

            for (int i = 0; i < CFFieldsKeysIncluded.Length; i++)
            {
                columnForcesSheet.Cells[2, i + 1].Value2 = CFFieldsKeysIncluded[i];
            }

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
            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkBook = null;
            Excel.Worksheet designInputSheet = null;
            Excel.Worksheet columnShearSheet = null;
            Excel.Worksheet columnForcesSheet = null;

            excelApplication = new Excel.Application();
            excelApplication.Visible = true;
            string @excelPath = this.designInput.Text;
            excelWorkBook = excelApplication.Workbooks.Open(excelPath);

            foreach (Excel.Worksheet sheet in excelWorkBook.Sheets)
            {
                if (sheet.Name == "Concrete Column PMM Envelope")
                {
                    designInputSheet = sheet;
                }
                else if (sheet.Name == "Concrete Column Shear Envelope")
                {
                    columnShearSheet = sheet;
                }
                else if (sheet.Name == "Column Forces")
                {
                    columnForcesSheet = sheet;
                }
            }

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
                String numberFormat = designInputSheet.Cells[row, ptRequiredIndex].NumberFormat;
                if (numberFormat.Contains('%'))
                {
                    givenPtRequired = columnPMMDataObject[row, ptRequiredIndex] != null ?
                        (double)columnPMMDataObject[row, ptRequiredIndex] * 100 :
                        0;
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
                this.columnsToDesign.Items.Add(columnLabel);
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
            this.storyName.Items.Clear();
            this.storiesToDesign.Items.Clear();
            this.maxEtabsRebarPt.Text = "";
            this.ResultsTable.Rows.Clear();

            List<string> storyNames = new List<string>();
            foreach (var thisColumnInputData in columnInputData.Where(
                _ => _.columnLabel == this.columnsToDesign.SelectedItem.ToString()).ToList())
            {
                storyNames.Add(thisColumnInputData.story);
            }
            
            this.storyName.Items.Add("All stories");
            foreach (string storyName in storyNames.Distinct().ToList())
            {
                this.storyName.Items.Add(storyName);
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
            for (int i=0; i< this.ResultsTable.Rows.Count-1; i++)
            {
                section.Material = GetSectionMaterial((double)this.ResultsTable.Rows[i].Cells[2].Value);
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

                double axialAdSec = -1 * (double)this.ResultsTable.Rows[i].Cells[2].Value;
                double mMajorAdSec = (double)this.ResultsTable.Rows[i].Cells[3].Value;
                double mMinorAdSec = (double)this.ResultsTable.Rows[i].Cells[4].Value;
                var adSecLoad = ILoad.Create(
                    Force.FromKilonewtons(axialAdSec),
                    Moment.FromKilonewtonMeters(mMajorAdSec),
                    Moment.FromKilonewtonMeters(mMinorAdSec));

                var strengthResult = solution.Strength.Check(adSecLoad);
                var momentRanges = strengthResult.MomentRanges;
                double ultimateMoment = 0;
                for(int j=0; j< momentRanges.Count; j++)
                {
                    ultimateMoment = Math.Max(momentRanges[j].Max.KilonewtonMeters, ultimateMoment);
                }
                double appliedMoment = Math.Sqrt(Math.Pow(mMajorAdSec, 2) + Math.Pow(mMinorAdSec, 2));
                double momentRatio = appliedMoment / ultimateMoment;
                this.ResultsTable.Rows[i].Cells[5].Value = Math.Round(strengthResult.LoadUtilisation.DecimalFractions,2);
                this.ResultsTable.Rows[i].Cells[6].Value = Math.Round(momentRatio,2);
                this.ResultsTable.Rows[i].Cells[8].Value = Math.Round(rebarPtAdSec, 3);
                this.ResultsTable.Rows[i].Cells[7].Value = "";
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
            if(this.lineGroupsAdded.Items.Count != 1)
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
            Excel.Application excelApplication = null;
            Excel.Workbook excelWorkBook = null;
            Excel.Worksheet designOutputSheet = null;

            excelApplication = new Excel.Application();
            excelApplication.Visible = true;
            string @excelPath = this.designInput.Text;
            excelWorkBook = excelApplication.Workbooks.Open(excelPath);

            foreach (Excel.Worksheet sheet in excelWorkBook.Sheets)
            {
                if (sheet.Name == "AdSec Results")
                {
                    designOutputSheet = sheet;
                    break;
                }
            }
            int startRow = designOutputSheet.UsedRange.Rows.Count + 1;

            for (int i = 0; i < this.ResultsTable.Rows.Count-1; i++)
            {
                designOutputSheet.Cells[startRow, 1].Value2 = this.columnsToDesign.Text;
                designOutputSheet.Cells[startRow, 2].Value2 = (string)this.ResultsTable.Rows[i].Cells[0].Value;
                designOutputSheet.Cells[startRow, 3].Value2 = (string)this.ResultsTable.Rows[i].Cells[1].Value;
                designOutputSheet.Cells[startRow, 4].Value2 = (double)this.ResultsTable.Rows[i].Cells[3].Value;
                designOutputSheet.Cells[startRow, 5].Value2 = (double)this.ResultsTable.Rows[i].Cells[4].Value;
                designOutputSheet.Cells[startRow, 6].Value2 = 0;
                designOutputSheet.Cells[startRow, 7].Value2 = (double)this.ResultsTable.Rows[i].Cells[2].Value;
                designOutputSheet.Cells[startRow, 8].Value2 = (double)this.ResultsTable.Rows[i].Cells[7].Value;
                designOutputSheet.Cells[startRow, 9].Value2 = (double)this.ResultsTable.Rows[i].Cells[8].Value;
                designOutputSheet.Cells[startRow, 10].Value2 = (double)this.ResultsTable.Rows[i].Cells[9].Value;
                designOutputSheet.Cells[startRow, 11].Value2 = (double)ResultsTable.Rows[i].Cells[10].Value;
                designOutputSheet.Cells[startRow, 12].Value2 = (double)ResultsTable.Rows[i].Cells[11].Value;
                designOutputSheet.Cells[startRow, 13].Value2 = (double)this.ResultsTable.Rows[i].Cells[14].Value;
                designOutputSheet.Cells[startRow, 14].Value2 = this.maxEtabsRebarPt.Text;
                designOutputSheet.Cells[startRow, 15].Value2 = (double)this.ResultsTable.Rows[i].Cells[13].Value;
                designOutputSheet.Cells[startRow, 16].Value2 = (string)this.ResultsTable.Rows[i].Cells[12].Value;
                startRow += 1;
            }

            for (int i = 0; i < this.designDataToExcel.Count; i++)
            {
                var givenDesignData = designDataToExcel[i];
                string columnID = givenDesignData.inputData.columnLabel + " "
                    + givenDesignData.inputData.story;

                Excel.Worksheet shearCalcIS456Sheet = excelWorkBook.Sheets.Add(
                    excelWorkBook.Sheets[excelWorkBook.Sheets.Count], Type.Missing, Type.Missing, Type.Missing);
                shearCalcIS456Sheet.Name = columnID + " IS456";
                shearCalcIS456Sheet.Cells[2, 3].Value2 = "SHEAR REINFORCEMENT FOR RC COLUMNS";
                shearCalcIS456Sheet.Cells[4, 3].Value2 = "As per IS456:2000";
                shearCalcIS456Sheet.Cells[6, 3].Value2 = "RECTANGULAR COLUMN";
                shearCalcIS456Sheet.Cells[8, 3].Value2 = givenDesignData.inputData.columnLabel;
                shearCalcIS456Sheet.Cells[8, 4].Value2 = givenDesignData.inputData.story;
                shearCalcIS456Sheet.Cells[10, 3].Value2 = "Factored shear force (along X)";
                shearCalcIS456Sheet.Cells[10, 5].Value2 = givenDesignData.factoredShearForceAlongX;
                shearCalcIS456Sheet.Cells[10, 6].Value2 = "kN";
                shearCalcIS456Sheet.Cells[11, 3].Value2 = "Factored shear force (along Y)";
                shearCalcIS456Sheet.Cells[11, 5].Value2 = givenDesignData.factoredShearForceAlongY;
                shearCalcIS456Sheet.Cells[11, 6].Value2 = "kN";
                shearCalcIS456Sheet.Cells[12, 4].Value2 = "B";
                shearCalcIS456Sheet.Cells[12, 5].Value2 = givenDesignData.inputData.width;
                shearCalcIS456Sheet.Cells[12, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[13, 4].Value2 = "D";
                shearCalcIS456Sheet.Cells[13, 5].Value2 = givenDesignData.inputData.depth;
                shearCalcIS456Sheet.Cells[13, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[14, 4].Value2 = "Clear cover";
                shearCalcIS456Sheet.Cells[14, 5].Value2 = givenDesignData.clearCover;
                shearCalcIS456Sheet.Cells[14, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[15, 4].Value2 = "Dia of corner longitudinal bar";
                shearCalcIS456Sheet.Cells[15, 5].Value2 = givenDesignData.DesignData.cornerBarDia;
                shearCalcIS456Sheet.Cells[15, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[16, 4].Value2 = "Dia of centre longitudinal bar";
                shearCalcIS456Sheet.Cells[16, 5].Value2 = givenDesignData.DesignData.centreBarDia;
                shearCalcIS456Sheet.Cells[16, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[17, 4].Value2 = "b";
                shearCalcIS456Sheet.Cells[17, 5].Value2 = givenDesignData.effectiveWidth;
                shearCalcIS456Sheet.Cells[17, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[18, 4].Value2 = "d";
                shearCalcIS456Sheet.Cells[18, 5].Value2 = givenDesignData.effectiveDepth;
                shearCalcIS456Sheet.Cells[18, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[19, 4].Value2 = "tau v (along X)";
                shearCalcIS456Sheet.Cells[19, 5].Value2 = givenDesignData.tauVAlongX;
                shearCalcIS456Sheet.Cells[19, 6].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[20, 4].Value2 = "tau v (along Y)";
                shearCalcIS456Sheet.Cells[20, 5].Value2 = givenDesignData.tauVAlongY;
                shearCalcIS456Sheet.Cells[20, 6].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[22, 4].Value2 = "tau c max";
                shearCalcIS456Sheet.Cells[22, 5].Value2 = givenDesignData.tauCMax;
                shearCalcIS456Sheet.Cells[22, 6].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[22, 7].Value2 = "Check IS456:2000";
                shearCalcIS456Sheet.Cells[24, 4].Value2 = "Pst";
                shearCalcIS456Sheet.Cells[24, 5].Value2 = givenDesignData.DesignData.ptProvided;
                shearCalcIS456Sheet.Cells[24, 6].Value2 = "%";
                shearCalcIS456Sheet.Cells[24, 7].Value2 = "Check your design";
                shearCalcIS456Sheet.Cells[25, 4].Value2 = "fck";
                shearCalcIS456Sheet.Cells[25, 5].Value2 = givenDesignData.inputData.fck;
                shearCalcIS456Sheet.Cells[25, 6].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[25, 7].Value2 = "Check your design";
                shearCalcIS456Sheet.Cells[27, 4].Value2 = "tau c";
                shearCalcIS456Sheet.Cells[27, 5].Value2 = givenDesignData.tauC;
                shearCalcIS456Sheet.Cells[27, 6].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[27, 7].Value2 = "Check IS456:2000";
                shearCalcIS456Sheet.Cells[29, 5].Value2 =
                    Math.Max(givenDesignData.tauVAlongX, givenDesignData.tauVAlongY) > givenDesignData.tauC ?
                    "Shear reinforcement needed" :
                    "Shear reinforcement not needed, provide minimum shear reinforcement";
                shearCalcIS456Sheet.Cells[31, 4].Value2 = "Vus (along X)";
                shearCalcIS456Sheet.Cells[31, 5].Value2 = givenDesignData.VusAlongX;
                shearCalcIS456Sheet.Cells[31, 6].Value2 = "kN";
                shearCalcIS456Sheet.Cells[32, 4].Value2 = "Vus (along Y)";
                shearCalcIS456Sheet.Cells[32, 5].Value2 = givenDesignData.VusAlongY;
                shearCalcIS456Sheet.Cells[32, 6].Value2 = "kN";
                shearCalcIS456Sheet.Cells[34, 4].Value2 = "fy";
                shearCalcIS456Sheet.Cells[34, 5].Value2 = givenDesignData.linkFy;
                shearCalcIS456Sheet.Cells[34, 6].Value2 = "N/mm2";
                shearCalcIS456Sheet.Cells[35, 4].Value2 = "Dia of bar";
                shearCalcIS456Sheet.Cells[35, 5].Value2 = givenDesignData.linkDia;
                shearCalcIS456Sheet.Cells[35, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[36, 4].Value2 = "number of legs (along X)";
                shearCalcIS456Sheet.Cells[36, 5].Value2 = givenDesignData.legsAlongX;
                shearCalcIS456Sheet.Cells[37, 4].Value2 = "Asv (along X)";
                shearCalcIS456Sheet.Cells[37, 5].Value2 = givenDesignData.asvProvidedAlongX;
                shearCalcIS456Sheet.Cells[37, 6].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[38, 4].Value2 = "number of legs (along Y)";
                shearCalcIS456Sheet.Cells[38, 5].Value2 = givenDesignData.legsAlongY;
                shearCalcIS456Sheet.Cells[39, 4].Value2 = "Asv (along Y)";
                shearCalcIS456Sheet.Cells[39, 5].Value2 = givenDesignData.asvProvidedAlongY;
                shearCalcIS456Sheet.Cells[39, 6].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[41, 4].Value2 = "Sv required (cl. 40.4)";
                shearCalcIS456Sheet.Cells[41, 5].Value2 = givenDesignData.nonConfiningSpacingOne;
                shearCalcIS456Sheet.Cells[41, 6].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[42, 4].Value2 = "Sv required (cl. 26.5.1.5)";
                shearCalcIS456Sheet.Cells[42, 5].Value2 = givenDesignData.nonConfiningSpacingTwo;
                shearCalcIS456Sheet.Cells[42, 6].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[43, 4].Value2 = "Sv adopted";
                shearCalcIS456Sheet.Cells[43, 5].Value2 = givenDesignData.nonConfiningSpacingProvided;
                shearCalcIS456Sheet.Cells[43, 6].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[44, 5].Value2 = 
                    givenDesignData.nonConfiningSpacingProvided < 
                    Math.Min(givenDesignData.nonConfiningSpacingOne, givenDesignData.nonConfiningSpacingTwo) ? "SPACING IS OK" : "SPACING IS NOT OK";
                shearCalcIS456Sheet.Cells[46, 4].Value2 = "Minimum shear reinforcement (cl. 26.5.1.6)";
                shearCalcIS456Sheet.Cells[46, 5].Value2 = givenDesignData.minNonConfiningAsvRequired;
                shearCalcIS456Sheet.Cells[46, 6].Value2 = "mm2";
                shearCalcIS456Sheet.Cells[47, 5].Value2 = 
                    givenDesignData.minNonConfiningAsvRequired < 
                    Math.Min(givenDesignData.legsAlongX * Math.PI * 0.25 * givenDesignData.linkDia * givenDesignData.linkDia, 
                    givenDesignData.legsAlongY * Math.PI * 0.25 * givenDesignData.linkDia * givenDesignData.linkDia) ? "OK" : "NOT OK";
                shearCalcIS456Sheet.Cells[49, 4].Value2 = "Hence provide";
                shearCalcIS456Sheet.Cells[49, 5].Value2 = givenDesignData.linkDia;
                shearCalcIS456Sheet.Cells[49, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[49, 7].Value2 = givenDesignData.legsAlongX;
                shearCalcIS456Sheet.Cells[49, 8].Value2 = "legged stirrups along X @";
                shearCalcIS456Sheet.Cells[49, 9].Value2 = givenDesignData.nonConfiningSpacingProvided;
                shearCalcIS456Sheet.Cells[49, 10].Value2 = "mm";
                shearCalcIS456Sheet.Cells[49, 11].Value2 = "c/c";
                shearCalcIS456Sheet.Cells[49, 12].Value2 = "in the non-confining zone";
                shearCalcIS456Sheet.Cells[49, 13].Value2 = "As per IS456:2000";
                shearCalcIS456Sheet.Cells[50, 4].Value2 = "Hence provide";
                shearCalcIS456Sheet.Cells[50, 5].Value2 = givenDesignData.linkDia;
                shearCalcIS456Sheet.Cells[50, 6].Value2 = "mm";
                shearCalcIS456Sheet.Cells[50, 7].Value2 = givenDesignData.legsAlongY;
                shearCalcIS456Sheet.Cells[50, 8].Value2 = "legged stirrups along Y @";
                shearCalcIS456Sheet.Cells[50, 9].Value2 = givenDesignData.nonConfiningSpacingProvided;
                shearCalcIS456Sheet.Cells[50, 10].Value2 = "mm";
                shearCalcIS456Sheet.Cells[50, 11].Value2 = "c/c";
                shearCalcIS456Sheet.Cells[50, 12].Value2 = "in the non-confining zone";
                shearCalcIS456Sheet.Cells[50, 13].Value2 = "As per IS456:2000";

                Excel.Worksheet shearCalcIS13920Sheet = excelWorkBook.Sheets.Add(
                    excelWorkBook.Sheets[excelWorkBook.Sheets.Count], Type.Missing, Type.Missing, Type.Missing);
                shearCalcIS13920Sheet.Name = columnID + " IS13920";
                shearCalcIS13920Sheet.Cells[2, 3].Value2 = "SHEAR REINFORCEMENT FOR RC COLUMNS";
                shearCalcIS13920Sheet.Cells[4, 3].Value2 = "As per IS13920:2016";
                shearCalcIS13920Sheet.Cells[6, 3].Value2 = "RECTANGULAR COLUMN";
                shearCalcIS13920Sheet.Cells[8, 3].Value2 = givenDesignData.inputData.columnLabel;
                shearCalcIS13920Sheet.Cells[8, 4].Value2 = givenDesignData.inputData.story;
                shearCalcIS13920Sheet.Cells[10, 3].Value2 = "Spacing of confining reinforcement";
                shearCalcIS13920Sheet.Cells[12, 3].Value2 = "B";
                shearCalcIS13920Sheet.Cells[12, 4].Value2 = givenDesignData.inputData.width;
                shearCalcIS13920Sheet.Cells[12, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[13, 3].Value2 = "D";
                shearCalcIS13920Sheet.Cells[13, 4].Value2 = givenDesignData.inputData.depth;
                shearCalcIS13920Sheet.Cells[13, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[14, 4].Value2 = givenDesignData.confinfingSpacingOne;
                shearCalcIS13920Sheet.Cells[14, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[14, 6].Value2 = "As per IS13920:2016";
                shearCalcIS13920Sheet.Cells[16, 3].Value2 = "Smallest dia of longitudinal bar";
                shearCalcIS13920Sheet.Cells[16, 4].Value2 = givenDesignData.DesignData.centreBarDia;
                shearCalcIS13920Sheet.Cells[16, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[17, 4].Value2 = givenDesignData.confinfingSpacingTwo;
                shearCalcIS13920Sheet.Cells[17, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[17, 6].Value2 = "As per IS13920:2016";
                shearCalcIS13920Sheet.Cells[19, 4].Value2 = givenDesignData.confinfingSpacingThree;
                shearCalcIS13920Sheet.Cells[19, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[19, 6].Value2 = "As per IS13920:2016";
                shearCalcIS13920Sheet.Cells[21, 3].Value2 = "Maximum spacing of confining reinforcement";
                shearCalcIS13920Sheet.Cells[21, 4].Value2 = givenDesignData.maxConfiningSpacingRequired;
                shearCalcIS13920Sheet.Cells[21, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[23, 3].Value2 = "Assumed spacing of links (Sv)";
                shearCalcIS13920Sheet.Cells[23, 4].Value2 = givenDesignData.confiningSpacingProvided;
                shearCalcIS13920Sheet.Cells[23, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[24, 3].Value2 = "fck";
                shearCalcIS13920Sheet.Cells[24, 4].Value2 = givenDesignData.inputData.fck;
                shearCalcIS13920Sheet.Cells[24, 5].Value2 = "N/mm2";
                shearCalcIS13920Sheet.Cells[25, 3].Value2 = "fy";
                shearCalcIS13920Sheet.Cells[25, 4].Value2 = givenDesignData.linkFy;
                shearCalcIS13920Sheet.Cells[25, 5].Value2 = "N/mm2";
                shearCalcIS13920Sheet.Cells[27, 3].Value2 = "Dia of transverse reinforcement";
                shearCalcIS13920Sheet.Cells[27, 4].Value2 = givenDesignData.linkDia;
                shearCalcIS13920Sheet.Cells[27, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[28, 3].Value2 = "Dia of corner longitudinal reinforcement";
                shearCalcIS13920Sheet.Cells[28, 4].Value2 = givenDesignData.DesignData.cornerBarDia;
                shearCalcIS13920Sheet.Cells[28, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[29, 3].Value2 = "Clear cover";
                shearCalcIS13920Sheet.Cells[29, 4].Value2 = givenDesignData.clearCover;
                shearCalcIS13920Sheet.Cells[29, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[31, 3].Value2 = "Ag";
                shearCalcIS13920Sheet.Cells[31, 4].Value2 = givenDesignData.Ag;
                shearCalcIS13920Sheet.Cells[31, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[32, 3].Value2 = "Ak";
                shearCalcIS13920Sheet.Cells[32, 4].Value2 = givenDesignData.Ak;
                shearCalcIS13920Sheet.Cells[32, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[33, 3].Value2 = "h";
                shearCalcIS13920Sheet.Cells[33, 4].Value2 = givenDesignData.h;
                shearCalcIS13920Sheet.Cells[33, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[34, 3].Value2 = "Ash for the first case";
                shearCalcIS13920Sheet.Cells[34, 4].Value2 = givenDesignData.AshOne;
                shearCalcIS13920Sheet.Cells[34, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[35, 3].Value2 = "Ash for the second case";
                shearCalcIS13920Sheet.Cells[35, 4].Value2 = givenDesignData.AshTwo;
                shearCalcIS13920Sheet.Cells[35, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[36, 3].Value2 = "Ash provided";
                shearCalcIS13920Sheet.Cells[36, 4].Value2 = givenDesignData.AshProvided;
                shearCalcIS13920Sheet.Cells[36, 5].Value2 = "mm2";
                shearCalcIS13920Sheet.Cells[37, 4].Value2 = 
                    givenDesignData.AshProvided > Math.Max(givenDesignData.AshOne, givenDesignData.AshTwo) ?
                    "OK" : "NOT OK";
                shearCalcIS13920Sheet.Cells[39, 3].Value2 = "Max spacing of confining reinforcement";
                shearCalcIS13920Sheet.Cells[39, 4].Value2 = givenDesignData.maxConfiningSpacingRequired;
                shearCalcIS13920Sheet.Cells[39, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[39, 6].Value2 = "As per IS13920:2016";
                shearCalcIS13920Sheet.Cells[40, 3].Value2 = "Hence provide";
                shearCalcIS13920Sheet.Cells[40, 4].Value2 = givenDesignData.linkDia;
                shearCalcIS13920Sheet.Cells[40, 5].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[40, 6].Value2 = "@";
                shearCalcIS13920Sheet.Cells[40, 7].Value2 = givenDesignData.confiningSpacingProvided;
                shearCalcIS13920Sheet.Cells[40, 8].Value2 = "mm";
                shearCalcIS13920Sheet.Cells[40, 9].Value2 = "as ties in the end zone (confining zone) and lap splices zone";


                Excel.Worksheet columnForcesSheet = excelWorkBook.Sheets.Add(
                    excelWorkBook.Sheets[excelWorkBook.Sheets.Count], Type.Missing, Type.Missing, Type.Missing);
                columnForcesSheet.Name = columnID + " Forces";
                columnForcesSheet.Cells[1, 1].Value2 = "Column Label";
                columnForcesSheet.Cells[2, 1].Value2 = "Story";
                columnForcesSheet.Cells[1, 4].Value2 = "P (kN)";
                columnForcesSheet.Cells[1, 5].Value2 = "MMajor (kNm)";
                columnForcesSheet.Cells[1, 6].Value2 = "MMinor (kNm)";
                columnForcesSheet.Cells[1, 2].Value2 = givenDesignData.inputData.columnLabel;
                columnForcesSheet.Cells[2, 2].Value2 = givenDesignData.inputData.story;
                columnForcesSheet.Cells[2, 4].Value2 = givenDesignData.inputData.P;
                columnForcesSheet.Cells[2, 5].Value2 = givenDesignData.inputData.MMajor;
                columnForcesSheet.Cells[2, 6].Value2 = givenDesignData.inputData.MMinor;
                columnForcesSheet.Cells[2, 7].Value2 = "PMM envelope";
                int startingRow = 3;
                foreach (var load in givenDesignData.checkedMajorEccentricLoads)
                {
                    columnForcesSheet.Cells[startingRow, 4].Value2 = Math.Round(-1 * load.X.Kilonewtons, 2);
                    columnForcesSheet.Cells[startingRow, 5].Value2 = Math.Round(load.YY.KilonewtonMeters, 2);
                    columnForcesSheet.Cells[startingRow, 6].Value2 = Math.Round(load.ZZ.KilonewtonMeters, 2);
                    columnForcesSheet.Cells[startingRow, 7].Value2 = "Minimum eccentric moment for major bending";
                    startingRow++;
                }
                foreach (var load in givenDesignData.checkedMinorEccentricLoads)
                {
                    columnForcesSheet.Cells[startingRow, 4].Value2 = Math.Round(-1 * load.X.Kilonewtons, 2);
                    columnForcesSheet.Cells[startingRow, 5].Value2 = Math.Round(load.YY.KilonewtonMeters, 2);
                    columnForcesSheet.Cells[startingRow, 6].Value2 = Math.Round(load.ZZ.KilonewtonMeters, 2);
                    columnForcesSheet.Cells[startingRow, 7].Value2 = "Minimum eccentric moment for minor bending";
                    startingRow++;
                }
                foreach (var load in givenDesignData.checkedOtherLoads)
                {
                    columnForcesSheet.Cells[startingRow, 4].Value2 = Math.Round(-1 * load.X.Kilonewtons, 2);
                    columnForcesSheet.Cells[startingRow, 5].Value2 = Math.Round(load.YY.KilonewtonMeters, 2);
                    columnForcesSheet.Cells[startingRow, 6].Value2 = Math.Round(load.ZZ.KilonewtonMeters, 2);
                    startingRow++;
                }
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

        public ColumnDesignData designReinforcement(ISection concreteSection, double requiredRebarPt, double maxSpacing)
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

            IList<int> barDias = new List<int>() { 32, 25, 20, 16 };
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
                        IList<int> oddCentreBarCount = new List<int>() { 3, 5, 7 };

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
            ColumnDesignData columnDesignData = designReinforcement(concreteSection, requiredRebarPt, maxSpacing);
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
            this.designDataToExcel.Clear();

            for (int i = 0; i < this.ResultsTable.Rows.Count - 1; i++)
            {
                if((double)this.ResultsTable.Rows[i].Cells[3].Value == 0 || (double)this.ResultsTable.Rows[i].Cells[4].Value == 0)
                {
                    continue;
                }

                String givenColumnLabel = this.columnsToDesign.Text;
                String givenStory = this.ResultsTable.Rows[i].Cells[0].Value.ToString();
                List<ColumnForceData> matchingForceData = columnForceData.Where(_ =>
                (_.columnLabel == givenColumnLabel &&
                _.story == givenStory)).ToList();
                List<ColumnForceData> matchingForceDataGLC = matchingForceData.Where(_ =>
                _.outputCase == this.ResultsTable.Rows[i].Cells[6].Value.ToString()).ToList();

                IProfile profile = null; 
                profile = IRectangleProfile.Create(
                    Length.FromMillimeters((double)this.ResultsTable.Rows[i].Cells[3].Value), 
                    Length.FromMillimeters((double)this.ResultsTable.Rows[i].Cells[4].Value));
                section = ISection.Create(profile, GetSectionMaterial((double)this.ResultsTable.Rows[i].Cells[2].Value));

                double cover = Double.Parse(this.rebarCover.Text);
                section.Cover = ICover.Create(Length.FromMillimeters(cover));
                
                double ptRequired = this.ResultsTable.Rows[i].Cells[14].Value is string?
                    double.Parse((string)this.ResultsTable.Rows[i].Cells[14].Value) :
                    (double)this.ResultsTable.Rows[i].Cells[14].Value;
                
                // Loads to check
                double axialAdSec = -1 * (double)this.ResultsTable.Rows[i].Cells[7].Value;
                double mMajorAdSec = (double)this.ResultsTable.Rows[i].Cells[8].Value;
                double mMinorAdSec = (double)this.ResultsTable.Rows[i].Cells[9].Value;
                var adSecLoad = ILoad.Create(
                    Force.FromKilonewtons(axialAdSec),
                    Moment.FromKilonewtonMeters(mMajorAdSec),
                    Moment.FromKilonewtonMeters(mMinorAdSec));
                List<ILoad> loadsToBeChecked = new List<ILoad>();
                var givenDesignDataToExcel = new DesignDataToExcel();
                givenDesignDataToExcel.checkedMajorEccentricLoads = new List<ILoad>();
                givenDesignDataToExcel.checkedMinorEccentricLoads = new List<ILoad>();
                givenDesignDataToExcel.checkedOtherLoads = new List<ILoad>();

                double lengthOfColumn = (double)this.ResultsTable.Rows[i].Cells[5].Value;
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

                    if (Math.Abs(lengthOfColumn * 0.5 - givenForceData.station*1000) <
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
                    previousStation = givenForceData.station*1000;
                }
                indexOfP = -1;
                foreach (var givenForceData in matchingForceDataGLC)
                {
                    if (givenForceData.station == 0)
                        indexOfP++;

                    double ex = this.minEccOverride.CheckedIndices.Contains(0) ?
                        Double.Parse(this.minEccMajor.Text) : 
                        (0.65 * lengthOfColumn / 500) + ((double)this.ResultsTable.Rows[i].Cells[3].Value / 30);
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
                    
                    double ey = this.minEccOverride.CheckedIndices.Contains(1) ?
                        Double.Parse(this.minEccMinor.Text) : 
                        (0.65 * lengthOfColumn / 500) + ((double)this.ResultsTable.Rows[i].Cells[4].Value / 30);
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
                double linkMaterial = Double.Parse(this.givenLinkFy.Text);
                double linkDia = Double.Parse(this.linkDia.Text);
                var linkBar = IBarBundle.Create(GetRebarMaterial(linkMaterial), Length.FromMillimeters(linkDia));
                var linkGroup = ILinkGroup.Create(linkBar);
                givenDesignDataToExcel.inputData.rebarPtEtabs = Math.Round(ptRequired, 2);
                double arrangementDepth =
                    ((IRectangleProfile)section.Profile).Depth.Millimeters -
                    2 * section.Cover.UniformCover.Millimeters -
                    2 * linkDia -
                    32;// cornerBarDia;
                double arrangementWidth =
                    ((IRectangleProfile)section.Profile).Width.Millimeters -
                    2 * section.Cover.UniformCover.Millimeters -
                    2 * linkDia -
                    32;// cornerBarDia;

                int nAlongDepth = (int)Math.Ceiling(arrangementDepth / Double.Parse(this.maximumSpacing.Text)) + 1;
                int nAlongWidth = (int)Math.Ceiling(arrangementWidth / Double.Parse(this.maximumSpacing.Text)) + 1;
                var minPossiblePt = 100 * 0.25 * Math.PI * 16 * 16 * ((nAlongDepth + nAlongWidth) * 2 - 4) /
                    ((IRectangleProfile)section.Profile).Depth.Millimeters /
                    ((IRectangleProfile)section.Profile).Width.Millimeters;

                while (ptRequired != 0)
                {
                    ISection previousDesignSection = section;
                    var previousDesignData = designData;
                    double previousMomentRatio = momentRatio;
                    double previousLoadUtilisation = loadUtilisation;

                    section.ReinforcementGroups.Clear();
                    section.ReinforcementGroups.Add(linkGroup);

                    designData = DefineSectionReinforcement(
                        section,
                        ptRequired,
                        Double.Parse(this.maximumSpacing.Text));

                    if (designData.ptProvided != 0)
                    {
                        solution = adsecApp.Analyse(section);

                        strengthResult = solution.Strength.Check(adSecLoad);
                        var momentRanges = strengthResult.MomentRanges;
                        double ultimateMoment = 0;
                        for (int j = 0; j < momentRanges.Count; j++)
                        {
                            ultimateMoment = Math.Max(momentRanges[j].Max.KilonewtonMeters, ultimateMoment);
                        }
                        double appliedMoment = Math.Sqrt(Math.Pow(mMajorAdSec, 2) + Math.Pow(mMinorAdSec, 2));
                        momentRatio = appliedMoment / ultimateMoment;
                        loadUtilisation = strengthResult.LoadUtilisation.DecimalFractions;
                        
                        foreach(var loadToCheck in loadsToBeChecked)
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

                        if (momentRatio > 1 || loadUtilisation > 1 || otherMomentRatio > 1 || otherLoadUtilisation > 1)
                        {
                            section = previousDesignSection;
                            designData = previousDesignData;
                            momentRatio = previousMomentRatio;
                            loadUtilisation = previousLoadUtilisation;
                            ptRequired = 0;
                        }
                        else
                        {
                            if (minPossiblePt < ptRequired + 0.05)
                            {
                                ptRequired = ptRequired - 0.05;
                            }
                            else
                            {
                                ptRequired = 0;
                            }
                        }

                        //if (momentRatio < 0.9 && loadUtilisation < 0.9)
                        //{
                        //    if (designData.ptProvided != previousDesignData.ptProvided)
                        //    {
                        //        ptRequired = ptRequired - 0.1;
                        //    }
                        //    else
                        //    {
                        //        ptRequired = 0;
                        //    }
                        //}
                        //else if (momentRatio > 0.95 || loadUtilisation > 0.95)
                        //{
                        //    designData = previousDesignData;
                        //    momentRatio = previousMomentRatio;
                        //    loadUtilisation = previousLoadUtilisation;
                        //    ptRequired = 0;
                        //}
                        //else
                        //{
                        //    ptRequired = 0;
                        //}
                    }
                    else
                    {
                        //this.maximumSpacing.Text = (Double.Parse(this.maximumSpacing.Text) - 5).ToString();
                        ptRequired = 0;
                    }
                }

                this.ResultsTable.Rows[i].Cells[10].Value = Math.Round(loadUtilisation, 2);
                this.ResultsTable.Rows[i].Cells[11].Value = Math.Round(momentRatio, 2);
                this.ResultsTable.Rows[i].Cells[12].Value = designData.rebarDescription;
                this.ResultsTable.Rows[i].Cells[13].Value = Math.Round(designData.ptProvided, 2);

                double factoredShearForceAlongY = (double)this.ResultsTable.Rows[i].Cells[15].Value;
                double factoredShearForceAlongX = (double)this.ResultsTable.Rows[i].Cells[16].Value;
                double depth = ((IRectangleProfile)section.Profile).Depth.Millimeters;
                double width = ((IRectangleProfile)section.Profile).Width.Millimeters;
                double effectiveDepth = 0;
                double effectiveWidth = 0;

                double tauVAlongY = 0;
                double tauVAlongX = 0;
                double tauC = 0;
                double tauCMax = CalculateTauCMax((double)this.ResultsTable.Rows[i].Cells[2].Value);
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
                double spacingOne = Math.Min(depth, width) / 4;
                double spacingTwo = 6 * designData.centreBarDia;
                double spacingThree = 100;
                double maxConfiningSpacing = Math.Min(spacingOne, Math.Min(spacingTwo, spacingThree));
                double Ag = 0;
                double Ak = 0;
                double AshOne = 0;
                double AshTwo = 0;
                double AshRequired = 0;
                
                if (designData.ptProvided != 0)
                {
                    // Shear reinf. calculation as per IS 456
                    effectiveDepth =
                        depth - section.Cover.UniformCover.Millimeters - linkDia - designData.cornerBarDia * 0.5;
                    effectiveWidth =
                        width - section.Cover.UniformCover.Millimeters - linkDia - designData.cornerBarDia * 0.5;
                    tauVAlongY = 1000 * factoredShearForceAlongY / effectiveDepth / width;
                    tauVAlongX = 1000 * factoredShearForceAlongX / effectiveWidth / depth;
                    tauC = CalculateTauC(designData.ptProvided, (double)this.ResultsTable.Rows[i].Cells[2].Value);
                    legsAlongY = designData.cornerBarCountAlongWidth + designData.centreBarCountAlongWidth;
                    legsAlongX = designData.cornerBarCountAlongDepth + designData.centreBarCountAlongDepth;
                    vusAlongY = (1000 * factoredShearForceAlongY - tauC * effectiveDepth * width);
                    vusAlongX = (1000 * factoredShearForceAlongX - tauC * effectiveWidth * depth);
                    asvAlongY = (Math.PI / 4) * linkDia * linkDia * legsAlongY;
                    asvAlongX = (Math.PI / 4) * linkDia * linkDia * legsAlongX;
                    spacingRequiredAlongY = 0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY > 0 ?
                        Math.Min(0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY, 0.75 * effectiveDepth) :
                        0.75 * effectiveDepth;
                    spacingRequiredAlongX = 0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX > 0 ?
                        Math.Min(0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX, 0.75 * effectiveWidth) :
                        0.75 * effectiveWidth; // IS456 pg.47
                    spacingRequired = Math.Min(spacingRequiredAlongY, spacingRequiredAlongX);
                    minAsvRequired = 0.4 * width * double.Parse(this.nonConfiningSpacing.Text) 
                        / 0.87 / double.Parse(this.givenLinkFy.Text); // IS456 pg.48

                    nonConfiningSpacingOne = Math.Min(
                        0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY > 0 ?
                        0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongY * effectiveDepth / vusAlongY :
                        0,
                        0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX > 0 ?
                        0.87 * double.Parse(this.givenLinkFy.Text) * asvAlongX * effectiveWidth / vusAlongX :
                        0);
                    nonConfiningSpacingTwo = Math.Min(0.75 * effectiveDepth, 0.75 * effectiveWidth);

                    // Shear reinf. calculation as per IS 13920
                    spacingOne = Math.Min(depth, width) / 4;
                    //spacingTwo = 6 * designData.centreBarDia;
                    spacingTwo = designData.centreBarDia == 16 ? 6 * 20 : 6 * designData.centreBarDia;
                    spacingThree = 100;
                    maxConfiningSpacing = Math.Min(spacingOne, Math.Min(spacingTwo, spacingThree));

                    Ag = depth * width;
                    Ak = (depth - 2 * section.Cover.UniformCover.Millimeters) * 
                        (width - 2 * section.Cover.UniformCover.Millimeters);
                    AshOne = 0.18 * double.Parse(this.confiningSpacing.Text) *
                        Math.Max(designData.spacingAlongDepth, designData.spacingAlongWidth) *
                        (double)this.ResultsTable.Rows[i].Cells[2].Value *
                        (Ag / Ak - 1) /
                        double.Parse(this.givenLinkFy.Text);
                    AshTwo = 0.05 * double.Parse(this.confiningSpacing.Text) *
                        Math.Max(designData.spacingAlongDepth, designData.spacingAlongWidth) *
                        (double)this.ResultsTable.Rows[i].Cells[2].Value /
                        double.Parse(this.givenLinkFy.Text);
                    AshRequired = Math.Max(AshOne, AshTwo);
                }

                this.designTauCMax.Text = Math.Round(tauCMax, 1).ToString();
                this.designMinAsv.Text = Math.Ceiling(minAsvRequired).ToString();
                this.ResultsTable.Rows[i].Cells[19].Value = Math.Round(tauC, 2);
                this.ResultsTable.Rows[i].Cells[21].Value = Math.Round(tauVAlongX, 2);
                this.ResultsTable.Rows[i].Cells[18].Value = Math.Floor(asvAlongX);
                this.ResultsTable.Rows[i].Cells[20].Value = Math.Round(tauVAlongY, 2);
                this.ResultsTable.Rows[i].Cells[17].Value = Math.Floor(asvAlongY);
                this.ResultsTable.Rows[i].Cells[22].Value = Math.Floor(spacingRequired);
                this.ResultsTable.Rows[i].Cells[23].Value = Math.Floor(maxConfiningSpacing);
                this.ResultsTable.Rows[i].Cells[24].Value = Math.Ceiling(AshRequired);
                this.AshProvided.Text = Math.Floor(0.25 * Math.PI * linkDia * linkDia).ToString();

                givenDesignDataToExcel.designSection = section;
                givenDesignDataToExcel.inputData = new ColumnInputData();
                givenDesignDataToExcel.inputData.columnLabel = this.columnsToDesign.Text;
                givenDesignDataToExcel.inputData.story = this.ResultsTable.Rows[i].Cells[0].Value.ToString();
                givenDesignDataToExcel.inputData.P = Math.Round(axialAdSec * -1, 2);
                givenDesignDataToExcel.inputData.MMajor = Math.Round(mMajorAdSec, 2);
                givenDesignDataToExcel.inputData.MMinor = Math.Round(mMinorAdSec, 2);
                givenDesignDataToExcel.inputData.location = this.ResultsTable.Rows[i].Cells[1].Value.ToString();
                givenDesignDataToExcel.inputData.width = Math.Round(width, 0);
                givenDesignDataToExcel.inputData.depth = Math.Round(depth, 0);
                givenDesignDataToExcel.inputData.diameter = 0;
                givenDesignDataToExcel.inputData.fck = Math.Round((double)this.ResultsTable.Rows[i].Cells[2].Value, 0);

                givenDesignDataToExcel.DesignData = designData;

                givenDesignDataToExcel.loadUtilisation = Math.Round(loadUtilisation, 2);
                givenDesignDataToExcel.momentRatio = Math.Round(momentRatio, 2);
                givenDesignDataToExcel.factoredShearForceAlongY = Math.Round(factoredShearForceAlongY, 2);
                givenDesignDataToExcel.factoredShearForceAlongX = Math.Round(factoredShearForceAlongX, 2);
                givenDesignDataToExcel.effectiveDepth = Math.Round(effectiveDepth, 0);
                givenDesignDataToExcel.effectiveWidth = Math.Round(effectiveWidth, 0);
                givenDesignDataToExcel.tauVAlongY = Math.Round(tauVAlongY, 2);
                givenDesignDataToExcel.tauVAlongX = Math.Round(tauVAlongX, 2);
                givenDesignDataToExcel.tauC = Math.Round(tauC, 2);
                givenDesignDataToExcel.tauCMax = Math.Round(tauCMax, 2);
                givenDesignDataToExcel.longitudinalFy = Math.Round(double.Parse(this.designLongFy.Text), 0);
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
                givenDesignDataToExcel.nonConfiningSpacingProvided = Math.Floor(double.Parse(this.nonConfiningSpacing.Text));
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
                this.designDataToExcel.Add(givenDesignDataToExcel);
                section.ReinforcementGroups.Clear();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if(!this.storiesToDesign.Items.Contains(this.storyName.Text))
            {
                this.storiesToDesign.Items.Add(this.storyName.Text);
            }

            this.ResultsTable.Rows.Clear();
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

            foreach (var thisColumnInputData in columnInputData.Where(
                _ => _.columnLabel == this.columnsToDesign.SelectedItem.ToString()).ToList())
            {
                if (maxRebarPtEtabs < thisColumnInputData.rebarPtEtabs)
                {
                    maxRebarPtEtabs = thisColumnInputData.rebarPtEtabs;
                }

                if (this.storiesToDesign.Items.Contains("All stories") || this.storiesToDesign.Items.Contains(thisColumnInputData.story))
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

                    if (thisColumnInputData.location == "Top")
                    {
                        List<ColumnShearData> matchingShearData = columnShearData.Where(_ =>
                        (_.columnLabel == thisColumnInputData.columnLabel &&
                        _.story == thisColumnInputData.story &&
                        _.location == thisColumnInputData.location)).ToList();

                        pTopValue = thisColumnInputData.P;
                        mMajorTopValue = thisColumnInputData.MMajor;
                        mMinorTopValue = thisColumnInputData.MMinor;
                        rebarPtEtabsValue = thisColumnInputData.rebarPtEtabs;
                        vAlongXTop = matchingShearData.Count > 0 ?
                            Math.Max(matchingShearData[0].maxVAlongX, maxVAlongXForce) : maxVAlongXForce;
                        vAlongYTop = matchingShearData.Count > 0 ?
                            Math.Max(matchingShearData[0].maxVAlongY, maxVAlongYForce): maxVAlongYForce;
                    }
                    else if (thisColumnInputData.location == "Bottom")
                    {
                        List<ColumnShearData> matchingShearData = columnShearData.Where(_ =>
                        (_.columnLabel == thisColumnInputData.columnLabel &&
                        _.story == thisColumnInputData.story &&
                        _.location == thisColumnInputData.location)).ToList();

                        pBottomValue = thisColumnInputData.P;
                        mMajorBottomValue = thisColumnInputData.MMajor;
                        mMinorBottomValue = thisColumnInputData.MMinor;
                        rebarPtEtabsBottomValue = thisColumnInputData.rebarPtEtabs;
                        vAlongXBottom = matchingShearData.Count > 0 ?
                            Math.Max(matchingShearData[0].maxVAlongX, maxVAlongXForce) : maxVAlongXForce;
                        vAlongYBottom = matchingShearData.Count > 0 ?
                            Math.Max(matchingShearData[0].maxVAlongY, maxVAlongYForce) : maxVAlongYForce;

                        if (rebarPtEtabsBottomValue < rebarPtEtabsValue)
                        {
                            this.ResultsTable.Rows.Add(
                                thisColumnInputData.story,
                                "Top",
                                Math.Round(thisColumnInputData.fck, 0),
                                Math.Round(thisColumnInputData.depth, 0),
                                Math.Round(thisColumnInputData.width, 0),
                                Math.Round(thisColumnInputData.length, 0),
                                thisColumnInputData.governingCombo,
                                Math.Round(pTopValue, 2),
                                 Math.Round(mMajorTopValue, 2),
                                 Math.Round(mMinorTopValue, 2),
                                0.0,
                                0.0,
                                "",
                                0.0,
                                Math.Round(rebarPtEtabsValue, 2),
                                Math.Max(Math.Round(vAlongYTop,2), Math.Round(vAlongYBottom,2)),
                                Math.Max(Math.Round(vAlongXTop, 2), Math.Round(vAlongXBottom, 2)),
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
                            this.ResultsTable.Rows.Add(
                                thisColumnInputData.story,
                                "Bottom",
                                Math.Round(thisColumnInputData.fck, 0),
                                Math.Round(thisColumnInputData.depth, 0),
                                Math.Round(thisColumnInputData.width, 0),
                                Math.Round(thisColumnInputData.length, 0),
                                thisColumnInputData.governingCombo,
                                Math.Round(pBottomValue, 2),
                                Math.Round(mMajorBottomValue, 2),
                                Math.Round(mMinorBottomValue, 2),
                                0.0,
                                0.0,
                                "",
                                0.0,
                                Math.Round(rebarPtEtabsBottomValue, 2),
                                Math.Max(Math.Round(vAlongYTop, 2), Math.Round(vAlongYBottom, 2)),
                                Math.Max(Math.Round(vAlongXTop, 2), Math.Round(vAlongXBottom, 2)),
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

        private void columnsToShow_SelectedIndexChanged(object sender, EventArgs e)
        {
            var checkedColumns = this.columnsToShow.CheckedIndices;
            foreach(var checkedColumn in checkedColumns)
            {
                this.ResultsTable.Columns[(int)checkedColumn].Visible = true;
            }
            for(int i=0; i< this.columnsToShow.Items.Count; i++)
            {
                if(!checkedColumns.Contains(i))
                {
                    this.ResultsTable.Columns[i].Visible = false;
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if(this.storiesToDesign.SelectedItem != null)
            {
                if ((string)this.storiesToDesign.SelectedItem == "All stories")
                {
                    ResultsTable.Rows.Clear();
                    this.storiesToDesign.Items.Clear();
                    this.storiesToDesign.Text = "";
                    return;
                }

                for (int i = 0; i < ResultsTable.Rows.Count; i++)
                {
                    if ((string)ResultsTable.Rows[i].Cells[0].Value == (string)this.storiesToDesign.SelectedItem)
                    {
                        ResultsTable.Rows.RemoveAt(i);
                        break;
                    }
                }

                this.storiesToDesign.Items.Remove(this.storiesToDesign.Items[this.storiesToDesign.SelectedIndex]);
                this.storiesToDesign.Text = "";
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
        }
    }
}
