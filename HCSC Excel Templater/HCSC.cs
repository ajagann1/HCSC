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
    class HCSC
    {
        //Acts the same way yargs does in javascript
        class Options
        {
            [Option('t', "template", Required = true,
                HelpText = "The filepath of the CLABSI/CAUTI/CDI template to be converted to.")]
            public string TemplatePath { get; set; }

            [Option('s', "stateCodes", Required = true,
                HelpText = "The filepath to the list of states and their codes.")]
            public string stateCodesPath { get; set; }

            [Option('h', "hospitals", Required = true,
                HelpText = "The filepath to the list of hospitals to be used.")]
            public string HospitalList { get; set; }

            [Option('u', "cautiInputFolder", Required = false,
                HelpText = "The filepath to the folder the CAUTI input documents are stored in.")]
            public string CautiInputPath { get; set; }

            [Option('b', "clabsiInputFolder", Required = false,
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
                //Check inputs
                var template = new XLWorkbook(options.TemplatePath);
                var stateCodes = new XLWorkbook(options.stateCodesPath);
                var hospitalList = new XLWorkbook(options.HospitalList);
                var dataTemplates = new XLWorkbook();
                String outputPath = options.OutputPath != null && options.OutputPath != "" ? options.OutputPath : ".";
                String cautiInputPath = options.CautiInputPath != null && options.CautiInputPath != "" ? options.CautiInputPath : "";
                String clabsiInputPath = options.ClabsiInputPath != null && options.ClabsiInputPath != "" ? options.ClabsiInputPath : "";
                String cdiInputPath = options.CdiInputPath != null && options.CdiInputPath != "" ? options.CdiInputPath : "";
                if (cautiInputPath == "" && clabsiInputPath == "" && cdiInputPath == "")
                {
                    throw new Exception("No input data given");
                }


                //This block parses the template for all data types
                //Honestly unnecessary but I appreciate the cleaner logic that comes from doing this step
                #region ParseTemplate
                var cauti = dataTemplates.AddWorksheet("CAUTI");
                var clabsi = dataTemplates.AddWorksheet("CLABSI");
                var cdi = dataTemplates.AddWorksheet("CDI");

                var tws = template.Worksheet(1);
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

                if (clabsiInputPath != "")
                {
                    parseData(dataTemplates, hospitalDict, stateCodeDict, "CLABSI", clabsiInputPath, outputPath);
                }
                if(cautiInputPath != "")
                {
                    parseData(dataTemplates, hospitalDict, stateCodeDict, "CAUTI", cautiInputPath, outputPath);
                }
                if(cdiInputPath != "")
                {
                    parseData(dataTemplates, hospitalDict, stateCodeDict, "CDI", cdiInputPath, outputPath);
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
                hospitals.Add(row.Cell(1).Value.ToString(), row.Cell(2).Value.ToString());
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

        public static void parseData(XLWorkbook t, Dictionary<String, String> hospitals, Dictionary<String, String> stateCodes, String type, String filepath, String output)
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
                foreach (string file in fileInfo)
                {
                    if (file.ToLower().Contains("q"))
                    {
                        quarter = file;
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

                //Get fac header and iterate through data, searching for matching hospital ids to the list of hospitals
                using (StreamReader facFile = new StreamReader(data[facLocation].FullName))
                {
                    line = facFile.ReadLine();
                    String[] facHeaders = line.Split('\t');

                    while ((line = facFile.ReadLine()) != null)
                    {
                        String[] linePieces = line.Split('\t');
                        //Check Id
                        if (hospitals.ContainsKey(linePieces[0]))
                        {
                            //Get the state info
                            String state = stateCodes[linePieces[0].Substring(0, 2)];
                            
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
                            String sheetName = hospitals[linePieces[0]].Length > 31 ? hospitals[linePieces[0]].Substring(0, 31).Replace('/', ' ') : hospitals[linePieces[0]];
                            
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
                                row.Cell(3).Value = loc != -1 ? linePieces[loc] : "";
                                row.Cell(4).Value = stateLoc != -1 ? stateLineData[stateLoc] : "";
                                row.Cell(5).Value = natLoc != -1 ? natInfo[natLoc] : "";

                                //Go to next row
                                row = row.RowBelow();
                            }
                        }
                    }
                }
                quarters[quarter] = hospitalData;
            }

            foreach(KeyValuePair<String, String> hospital in hospitals)
            {
                XLWorkbook report = new XLWorkbook();
                String sheetName = hospital.Value.Length > 31 ? hospital.Value.Substring(0, 31).Replace('/', ' ') : hospital.Value;

                foreach (KeyValuePair<String, XLWorkbook> quarterData in quarters)
                {
                    quarterData.Value.Worksheet(sheetName).CopyTo(report, quarterData.Key);
                    report.Worksheet(quarterData.Key).Cell(1, 1).Value = quarterData.Key;
                    report.Worksheet(quarterData.Key).Cell(2, 1).Value = hospital.Key + ' ' + hospital.Value;
                }

                output = output[output.Length - 1] == '\\' ? output.Substring(0, output.Length - 1) : output;
                String outputPath = output + "\\output\\" + type;
                if (!Directory.Exists(outputPath))
                {
                    Directory.CreateDirectory(outputPath);
                }
               report.SaveAs(outputPath + "\\" + sheetName + ".xlsx");
            }

        }

    }
}
