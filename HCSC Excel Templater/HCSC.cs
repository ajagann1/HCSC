using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using CommandLine;
using CommandLine.Text;


namespace HCSC_Excel_Templater
{
    //This has to be the hackiest, laziest, piece of crap I've put together in a while.
    class HCSC
    {
        //Acts the same way yargs does in javascript
        class Options
        {
            [Option('t', "cTemplate", Required = true,
                HelpText = "The filepath of the CLABSI/CAUTI template to be converted to.")]
            public string cTemplatePath { get; set; }

            [Option('d', "cdiTemplate", Required = false,
                HelpText = "The filepath of the CDI template to be converted to.")]
            public string cdiTemplatePath { get; set; }

            [Option('c', "crosswalk", Required = true,
                HelpText = "The filepath of the CLABSI/CAUTI/CDI template to be converted to.")]
            public string crosswalkPath { get; set; }

            [Option('s', "stateCodes", Required = true,
                HelpText = "The filepath to the list of states and their codes.")]
            public string stateCodesPath { get; set; }

            [Option('h', "hospitals", Required = true,
                HelpText = "The filepath to the list of hospitals to be used.")]
            public string HospitalList { get; set; }

            [Option('u', "cautiInputFolder", Required = true,
                HelpText = "The filepath to the folder the CAUTI input documents are stored in.")]
            public string CautiInputPath { get; set; }

            [Option('b', "clabsiInputFolder", Required = true,
                HelpText = "The filepath to the folder the CLABSI input documents are stored in.")]
            public string ClabsiInputPath { get; set; }

            [Option('d', "cdiInputFolder", Required = false,
                HelpText = "The filepath to the folder the CDI input documents are stored in.")]
            public string CdiInputPath { get; set; }

            [Option('o', "outputFolder", Required = false,
                HelpText = "The filepath to the folder the output documents will be stored in.")]
            public string OutputPath { get; set; }

            [ParserState]
            public IParserState LastParserState { get; set; }

            [HelpOption]
            public string GetUsage()
            {
                return HelpText.AutoBuild(this,
                  (HelpText current) => HelpText.DefaultParsingErrorsHandler(this, current));
            }
        }
        static void Main(string[] args)
        {
            var options = new Options();
            if (Parser.Default.ParseArguments(args, options))
            {
                XLWorkbook readOnly = new XLWorkbook("C:\\Users\\Arvind\\Documents\\HCSC Ad Hoc Data\\XXXXXX_CLABSI_CAUTI_Preview - Copy.xlsx");
                //Set required inputs (if the path is incorrect or doesn't exist, should throw an error)
                var cTemplate = new XLWorkbook(options.cTemplatePath);
                var crosswalk = new XLWorkbook(options.crosswalkPath);
                var stateCodes = new XLWorkbook(options.stateCodesPath);
                var hospitalList = new XLWorkbook(options.HospitalList);

                //Check optional inputs
                var cdiTemplate = options.cdiTemplatePath != null && options.cdiTemplatePath != "" ? new XLWorkbook(options.cdiTemplatePath) : null;
                String outputPath = options.OutputPath != null && options.OutputPath != "" ? options.OutputPath : ".";
                String cautiInputPath = options.CautiInputPath != null && options.CautiInputPath != "" ? options.CautiInputPath : "";
                String clabsiInputPath = options.ClabsiInputPath != null && options.ClabsiInputPath != "" ? options.ClabsiInputPath : "";
                String cdiInputPath = options.CdiInputPath != null && options.CdiInputPath != "" ? options.CdiInputPath : "";

                //Catch some edge case errors
                if (cautiInputPath == "" && clabsiInputPath == "" && cdiInputPath == "")
                {
                    throw new Exception("No input data given");
                }
                if (cdiInputPath == "" && cdiTemplate != null)
                {
                    throw new Exception("If a CDI template is inputted, CDI data is required");
                }
                if (cdiInputPath != "" && cdiTemplate == null)
                {
                    throw new Exception("If CDI data is inputted, a CDI template is required");
                }

                //This block parses the template for all data types
                //Honestly unnecessary but I appreciate the cleaner logic that comes from doing this step
                #region ParseCrosswalk
                var dataTemplates = new XLWorkbook();
                var cauti = dataTemplates.AddWorksheet("CAUTI");
                var clabsi = dataTemplates.AddWorksheet("CLABSI");
                var cdi = dataTemplates.AddWorksheet("CDI");

                var tws = crosswalk.Worksheet(1);
                var row = tws.FirstRowUsed();
                var lastRow = tws.LastRowUsed();
                String cellVal = row.Cell(1).Value.ToString();
                IXLAddress topLeftAddress, botRightAddress;

                while (cellVal.ToLower().IndexOf("clabsi") == -1 && cellVal.ToLower().IndexOf("cauti") == -1 && cellVal.ToLower().IndexOf("cdi") == -1)
                {
                    row = row.RowBelow();
                    if (row.IsEmpty())
                    {
                        throw new Exception("Template not formatted correctly");
                    }
                    cellVal = row.Cell(1).Value.ToString();
                }

                //Initialize individual templates for each data type
                for (int x = 0; x < 3; x++)
                {
                    cellVal = row.Cell(1 + x * 2).Value.ToString();
                    //Select relevant region from the crosswalk template
                    topLeftAddress = row.Cell(1 + 2 * x).Address;
                    botRightAddress = lastRow.Cell(2 + 2 * x).Address;
                    var range = tws.Range(topLeftAddress, botRightAddress);

                    if (cellVal.ToLower().IndexOf("clabsi") != -1)
                    {
                        clabsi.Cell(1, 1).Value = "Placeholder for Quarter";
                        clabsi.Cell(2, 1).Value = "Placeholder for Hospital";
                        clabsi.Cell(3, 1).Value = range;
                        clabsi.Cell(3, 3).Value = "Facility Data";
                        clabsi.Cell(3, 4).Value = "State Data";
                        clabsi.Cell(3, 5).Value = "National Data";
                    }
                    else if (cellVal.ToLower().IndexOf("cauti") != -1)
                    {
                        cauti.Cell(1, 1).Value = "Placeholder for Quarter";
                        cauti.Cell(2, 1).Value = "Placeholder for Hospital";
                        cauti.Cell(3, 1).Value = range;
                        cauti.Cell(3, 3).Value = "Facility Data";
                        cauti.Cell(3, 4).Value = "State Data";
                        cauti.Cell(3, 5).Value = "National Data";
                    }
                    else if (cellVal.ToLower().IndexOf("cdi") != -1)
                    {
                        cdi.Cell(1, 1).Value = "Placeholder for Quarter";
                        cdi.Cell(2, 1).Value = "Placeholder for Hospital";
                        cdi.Cell(3, 1).Value = range;
                        cdi.Cell(3, 3).Value = "Facility Data";
                        cdi.Cell(3, 4).Value = "State Data";
                        cdi.Cell(3, 5).Value = "National Data";
                    }
                }
                #endregion
                //Test
                //dataTemplates.SaveAs(outputPath + "/testTemplates.xlsx");
                Dictionary<String, String> hospitalDict = createHospitalDictionary(hospitalList);
                Dictionary<String, String> stateCodeDict = createStateCodeDictionary(stateCodes);

                //Empty out the default values for the templatte
                for(int x = 1; x <= cTemplate.Worksheets.Count; x++)
                {
                    for(int y = 5; y <= 6; y++)
                    {
                        for(int z = 2; z <= 12; z++)
                        {
                            cTemplate.Worksheet(x).Cell(y, z).Value = "";
                        }
                    }
                }
                
                
                Dictionary<String, XLWorkbook> clabsiReports = parseData(dataTemplates, hospitalDict, stateCodeDict, "CLABSI", clabsiInputPath);
                Dictionary<String, XLWorkbook> cautiReports = parseData(dataTemplates, hospitalDict, stateCodeDict, "CAUTI", cautiInputPath);

                foreach (KeyValuePair<String, XLWorkbook> report in clabsiReports)
                {
                    XLWorkbook outputTemplate = new XLWorkbook();
                    String hospitalId = "";
                    XLWorkbook clabsiWorkbook = report.Value;
                    XLWorkbook cautiWorkbook = cautiReports[report.Key];
                    for (int x = 1; x <= report.Value.Worksheets.Count; x++)
                    {
                        var clabsiSheet = report.Value.Worksheet(x);
                        var cautiSheet = cautiWorkbook.Worksheet(clabsiSheet.Name);
                        String clabcautSheetName = clabsiSheet.Name.Replace('_', ' ');
                        String relevantQ = clabcautSheetName.Substring(clabcautSheetName.IndexOf('Q'), 2);
                        int sheetNumber = -1;
                        for (int y = 1; y <= cTemplate.Worksheets.Count; y++)
                        {
                            String tempSheetName = cTemplate.Worksheet(y).Name.Substring(0, cTemplate.Worksheet(y).Name.IndexOf('-'));

                            if (tempSheetName.Contains(relevantQ))
                            {
                                sheetNumber = y;
                                break;
                            }
                        }
                        
                        cTemplate.Worksheet(sheetNumber).CopyTo(outputTemplate, cTemplate.Worksheet(sheetNumber).Name);
                        var outputSheet = outputTemplate.Worksheet(x);

                        //Table header
                        //outputSheet.Cell(1, 1).Value = "Table " + sheetNumber.ToString() + ": " + clabsiSheet.Cell(1, 1).Value;
                        outputSheet.Cell(2, 1).Value = clabsiSheet.Cell(2, 1).Value;

                        String temp = clabsiSheet.Cell(2, 1).Value.ToString();
                        hospitalId = temp.Substring(0, temp.IndexOf(':') - 1);

                        var headerRow = outputSheet.Row(4);
                        var clabsiOutRow = outputSheet.Row(5);
                        var cautiOutRow = outputSheet.Row(6);
                        int headerCounter = 1;
                        while (!headerRow.Cell(headerCounter).IsEmpty())
                        {
                            var sheetRow = clabsiSheet.Row(4);
                            while (!sheetRow.Cell(2).IsEmpty())
                            {
                                if (sheetRow.Cell(2).Value.ToString().Equals(headerRow.Cell(headerCounter).Value.ToString(), StringComparison.Ordinal))
                                {
                                    break;
                                }
                                sheetRow = sheetRow.RowBelow();
                            }
                            //If the header piece is mapped, put in a value, else, leave it as is
                            clabsiOutRow.Cell(headerCounter).Value = !sheetRow.Cell(2).IsEmpty() ? sheetRow.Cell(3).Value : clabsiOutRow.Cell(headerCounter).Value;
                            headerCounter++;
                        }
                        headerCounter = 1;
                        while (!headerRow.Cell(headerCounter).IsEmpty())
                        {
                            var sheetRow = cautiSheet.Row(4);
                            while (!sheetRow.Cell(2).IsEmpty())
                            {
                                if (sheetRow.Cell(2).Value.ToString().Equals(headerRow.Cell(headerCounter).Value.ToString(), StringComparison.Ordinal))
                                {
                                    break;
                                }
                                sheetRow = sheetRow.RowBelow();
                            }
                            //If the header piece is mapped, put in a value, else, leave it as is
                            cautiOutRow.Cell(headerCounter).Value = !sheetRow.Cell(2).IsEmpty() ? sheetRow.Cell(3).Value : cautiOutRow.Cell(headerCounter).Value;
                            headerCounter++;
                        }

                        outputSheet.Range(outputSheet.Cell(4, 1), outputSheet.Cell(6, 12)).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;

                        //outputSheet.Protect();
                        //outputSheet.Range(outputSheet.FirstCellUsed(), outputSheet.LastCellUsed()).Style.Protection.SetLocked(true);


                    }
                    outputPath = outputPath[outputPath.Length - 1] == '\\' ? outputPath.Substring(0, outputPath.Length - 1) : outputPath;
                    String output = outputPath + "\\output";
                    if (!Directory.Exists(outputPath))
                    {
                        Directory.CreateDirectory(outputPath);
                    }

                    for(int x = 1; x <= outputTemplate.Worksheets.Count; x++)
                    {
                        String name = outputTemplate.Worksheet(x).Name;
                        String[] nameParts = name.Split(' ');
                        if(Convert.ToInt32(nameParts[1]) != x)
                        {
                            outputTemplate.Worksheet(x).Position = Convert.ToInt32(nameParts[1]);
                            x = 0;
                        }
                    }
                    outputTemplate.SaveAs(outputPath + "\\" + hospitalId + "_CAUTI_CLABSI_Preview.xlsx");
                    System.IO.FileInfo objFileInfo = new System.IO.FileInfo(outputPath + "\\" + hospitalId + "_CAUTI_CLABSI_Preview.xlsx");
                    objFileInfo.IsReadOnly = true;
                }

                if (cdiInputPath != "")
                {
                    Dictionary<String, XLWorkbook> cdiReports = parseData(dataTemplates, hospitalDict, stateCodeDict, "CDI", cdiInputPath);
                    //TODO: Put together a suitable output setup for CDI (whenever that's a thing)
                }

            }
            else
            {
                throw new Exception("Required arguments not inputted");
            }
        }

        public static Dictionary<String, String> createHospitalDictionary(XLWorkbook hospitalList)
        {
            Dictionary<String, String> hospitals = new Dictionary<string, string>();
            var row = hospitalList.Worksheet(1).FirstRowUsed();
            row = row.Cell(1).Value.ToString() == "CCN" ? row.RowBelow() : row;
            while (!row.IsEmpty())
            {
                String key = row.Cell(1).Value.ToString().Length == 5 ? "0" + row.Cell(1).Value.ToString() : row.Cell(1).Value.ToString();
                String value = row.Cell(2).Value.ToString().Length == 5 ? "0" + row.Cell(2).Value.ToString() : row.Cell(2).Value.ToString();
                hospitals.Add(key, value);
                row = row.RowBelow();
            }
            return hospitals;
        }

        public static Dictionary<String, String> createStateCodeDictionary(XLWorkbook stateCodes)
        {
            Dictionary<String, String> stateCodesDict = new Dictionary<String, String>();
            var row = stateCodes.Worksheet(1).FirstRowUsed();
            row = row.Cell(1).Value.ToString().Contains("State Code") ? row.RowBelow() : row;
            while (!row.IsEmpty())
            {
                String[] values;
                String rowVal = row.Cell(1).Value.ToString();
                if (rowVal.Contains(","))
                {
                    values = rowVal.Split(',');
                    foreach (String value in values)
                    {
                        stateCodesDict.Add(value.Trim(), row.Cell(2).Value.ToString());
                    }
                }
                else
                {
                    stateCodesDict.Add(rowVal, row.Cell(2).Value.ToString());
                }
                row = row.RowBelow();
            }
            return stateCodesDict;
        }

        public static Dictionary<String, XLWorkbook> parseData(XLWorkbook t, Dictionary<String, String> hospitals, Dictionary<String, String> stateCodes, String type, String filepath)
        {
            Dictionary<String, XLWorkbook> quarters = new Dictionary<string, XLWorkbook>();
            DirectoryInfo dir = new DirectoryInfo(@filepath);
            FileInfo[] data = dir.GetFiles("*.txt");

            //Return all individual dataType file indices in the data array
            int[] facLocations = data.Select((v, k) => v.Name.Contains("fac") ? k : -1).Where(k => k != -1).ToArray();
            int[] stateLocations = data.Select((v, k) => v.Name.Contains("state") ? k : -1).Where(k => k != -1).ToArray();
            int[] natLocations = data.Select((v, k) => v.Name.Contains("natl") ? k : -1).Where(k => k != -1).ToArray();



            foreach (int facLocation in facLocations)
            {
                //Get the current quarter
                String[] fileInfo = data[facLocation].Name.Split('_');
                String quarter = "";
                for(int x = 0; x <= fileInfo.Length; x++)
                {
                        if (fileInfo[x].ToLower().Contains("q"))
                        {
                            quarter = fileInfo[x - 1] + "_" + fileInfo[x];
                            break;
                        }
                }

                //Open the quarter-associated state and national files and get headers
                String line = "";
                String[] stateHeaders = null;
                foreach (int stateLoc in stateLocations)
                {
                    if (data[stateLoc].Name.Contains(quarter))
                    {
                        using (StreamReader stateFile = new StreamReader(data[stateLoc].FullName))
                        {
                            line = stateFile.ReadLine();
                            stateHeaders = line.Split('\t');
                        }
                        break;
                    }
                }

                String[] natHeaders = null;
                String[] natInfo = null;
                foreach (int natLoc in natLocations)
                {
                    if (data[natLoc].Name.Contains(quarter))
                    {
                        using (StreamReader natFile = new StreamReader(data[natLoc].FullName))
                        {
                            line = natFile.ReadLine();
                            natHeaders = line.Split('\t');
                            natInfo = natFile.ReadLine().Split('\t');
                        }
                        break;
                    }
                }

                XLWorkbook hospitalData = new XLWorkbook();
                List<String> usedHospitals = new List<String>();

                //Get fac header and iterate through data, searching for matching hospital ids to the list of hospitals
                using (StreamReader facFile = new StreamReader(data[facLocation].FullName))
                {
                    line = facFile.ReadLine();
                    String[] facHeaders = line.Split('\t');

                    while ((line = facFile.ReadLine()) != null)
                    {
                        String[] linePieces = line.Split('\t');
                        String hospitalId = linePieces[0];
                        if(hospitalId.Length == 5)
                        {
                            hospitalId = "0" + hospitalId;
                        }
                        //Check Id
                        if (hospitals.ContainsKey(hospitalId))
                        {
                            usedHospitals.Add(hospitalId);
                            //Get the state info
                            String state = stateCodes[hospitalId.Substring(0, 2)];
                            
                            String[] stateLineData = null;
                            foreach (int stateLoc in stateLocations)
                            {
                                if (data[stateLoc].Name.Contains(quarter))
                                {
                                    using (StreamReader stateFile = new StreamReader(data[stateLoc].FullName))
                                    {
                                        stateFile.ReadLine();
                                        String stateLine = "";
                                        
                                        while ((stateLine = stateFile.ReadLine()) != null)
                                        {
                                            stateLineData = stateLine.Split('\t');
                                            if (stateLineData[0] == state)
                                            {
                                                break;
                                            }
                                        }
                                    }
                                    break;
                                }
                            }


                            //Copy the appropriate template to a sheet associated with the hospital
                            String sheetName = hospitalId;
                            //hospitals[linePieces[0]].Length > 31 ? hospitals[linePieces[0]].Substring(0, 31).Replace('/', ' ') : hospitals[linePieces[0]];
                            
                            t.Worksheet(type).CopyTo(hospitalData, sheetName);
                            var hospitalSheet = hospitalData.Worksheet(sheetName);
                            var row = hospitalSheet.FirstRowUsed();

                            while (!row.Cell(1).Value.ToString().Contains("_"))
                            {
                                row = row.RowBelow();
                            }
                            while (!row.IsEmpty())
                            {
                                String header = row.Cell(1).Value.ToString();
                                //Find data in files
                                int loc = Array.IndexOf(facHeaders, header);
                                int stateLoc = Array.IndexOf(stateHeaders, header);
                                int natLoc = Array.IndexOf(natHeaders, header);

                                //Populate row
                                row.Cell(3).Value = loc != -1 ? linePieces[loc] : (stateLoc != -1 ? stateLineData[stateLoc] : (natLoc != -1 ? natInfo[natLoc] : ""));

                                //Go to next row
                                row = row.RowBelow();
                            }
                        }
                    }
                }

                //Populate the values for hospitals that don't have info for this quarter
                List<String> notUsedHospitals = new List<String>();

                foreach(KeyValuePair<String, String> hospital in hospitals)
                {
                    if (!usedHospitals.Contains(hospital.Key))
                    {
                        notUsedHospitals.Add(hospital.Key);
                    }
                }

                foreach (String hospital in notUsedHospitals)
                {
                    String state = stateCodes[hospital.Substring(0, 2)];
                    String[] stateData = null;
                    foreach (int stateLoc in stateLocations)
                    {
                        if (data[stateLoc].Name.Contains(quarter))
                        {
                            using (StreamReader stateFile = new StreamReader(data[stateLoc].FullName))
                            {
                                stateFile.ReadLine();
                                String stateLine = "";

                                while ((stateLine = stateFile.ReadLine()) != null)
                                {
                                    stateData = stateLine.Split('\t');
                                    if (stateData[0] == state)
                                    {
                                        break;
                                    }
                                }
                            }
                            break;
                        }
                    }

                    //Copy the appropriate template to a sheet associated with the hospital
                    String sheetName = hospital;
                    //hospitals[linePieces[0]].Length > 31 ? hospitals[linePieces[0]].Substring(0, 31).Replace('/', ' ') : hospitals[linePieces[0]];

                    t.Worksheet(type).CopyTo(hospitalData, sheetName);
                    var hospitalSheet = hospitalData.Worksheet(sheetName);
                    var row = hospitalSheet.FirstRowUsed();

                    while (!row.Cell(1).Value.ToString().Contains("_"))
                    {
                        row = row.RowBelow();
                    }
                    while (!row.IsEmpty())
                    {
                        String header = row.Cell(1).Value.ToString();
                        //Find data in files
                        int stateLoc = Array.IndexOf(stateHeaders, header);
                        int natLoc = Array.IndexOf(natHeaders, header);

                        //Populate row
                        row.Cell(3).Value = stateLoc != -1 ? stateData[stateLoc] : (natLoc != -1 ? natInfo[natLoc] : "");

                        //Go to next row
                        row = row.RowBelow();
                    }
                }

                quarters[quarter] = hospitalData;
            }

            Dictionary<String, XLWorkbook> reports = new Dictionary<string, XLWorkbook>();
            foreach(KeyValuePair<String, String> hospital in hospitals)
            {
                XLWorkbook report = new XLWorkbook();
                String sheetName = hospital.Key;

                foreach (KeyValuePair<String, XLWorkbook> quarterData in quarters)
                {
                    try
                    {
                        quarterData.Value.Worksheet(sheetName).CopyTo(report, quarterData.Key);
                    }
                    catch
                    {
                        report.AddWorksheet(quarterData.Key);
                    }
                    report.Worksheet(quarterData.Key).Cell(1, 1).Value = quarterData.Key;
                    report.Worksheet(quarterData.Key).Cell(2, 1).Value = hospital.Key + ": " + hospital.Value;
                }

                reports.Add(hospital.Key, report);
                /*output = output[output.Length - 1] == '\\' ? output.Substring(0, output.Length - 1) : output;
                String outputPath = output + "\\output\\" + type;
                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                }
                report.SaveAs(outputPath + "\\" + sheetName + ".xlsx");*/
            }

            /*foreach (XLWorkbook report in reports)
            {
                XLWorkbook outputTemplate = new XLWorkbook();
                String hospitalId = "";
                for(int x = 1; x < report.Worksheets.Count; x++)
                {
                    var sheet = report.Worksheet(x);
                    template.Worksheet(1).CopyTo(outputTemplate, sheet.Name);
                    var outputSheet = outputTemplate.Worksheet(x);

                    //Table header
                    outputSheet.Cell(1, 1).Value = "Table " +  x.ToString() + ": " + sheet.Cell(1, 1).Value;
                    outputSheet.Cell(2, 1).Value = sheet.Cell(2, 1).Value;

                    String temp = sheet.Cell(2, 1).Value.ToString();
                    hospitalId = temp.Substring(0, temp.IndexOf(':'));

                    var headerRow = outputSheet.Row(4);
                    var outputRow = outputSheet.LastRowUsed().RowBelow();
                    int headerCounter = 1;
                    while (!headerRow.Cell(headerCounter).IsEmpty())
                    {
                        var sheetRow = sheet.Row(4);
                        while (!sheetRow.Cell(2).IsEmpty())
                        {
                            if(sheetRow.Cell(2).Value.ToString().Equals(headerRow.Cell(headerCounter).Value.ToString(), StringComparison.Ordinal))
                            {
                                break;
                            }
                            sheetRow = sheetRow.RowBelow();
                        }
                        outputRow.Cell(headerCounter).Value = sheetRow.Cell(3).Value;
                        headerCounter++;
                    }
                }
                output = output[output.Length - 1] == '\\' ? output.Substring(0, output.Length - 1) : output;
                String outputPath = output + "\\output\\" + type;
                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                }
                outputTemplate.SaveAs(outputPath + "\\" + hospitalId + "_" + type + ".xlsx");
            }*/
            return reports;
        }

    }
}
