using System.Data;
using System.Xml.Linq;
using ClosedXML.Excel;


    IXLWorkbook workbook = new XLWorkbook();
    var worksheet = workbook.Worksheets.Add("XmlToExcel");

    XDocument xmlDoc = XDocument.Load("/Users/abderahman/RiderProjects/LeoniTest/LeoniTest/XmlFile/CC24E-MAIN_HARNESS_XX24 NV-HX2ML.xml");

    // Getting all wires element inside a List
    //List<XElement> wires = xmlDoc.Descendants().Where(des => des.Name.LocalName == "wire").ToList();
    
    DataSet dataset = new DataSet();
    using (var reader = xmlDoc.CreateReader())
    {
        dataset.ReadXml(reader);
    }

    int current = 1;
    foreach (DataTable table in dataset.Tables)
    {
        if (table.TableName == "wire")
        {
            Console.WriteLine(table.TableName);

                worksheet.Cell(current, current).Value= table.TableName;
                worksheet.Cell(current,current).Style.Font.Bold = true;

                for (int col = 0; col < table.Columns.Count; col++)
                {
                    worksheet.Cell(current+1, col+1).Value = table.Columns[col].ColumnName;
                    worksheet.Cell(current+1, col+1).Style.Font.FontSize = 20;
                    worksheet.Cell(current+1, col+1).Style.Font.Bold = true;
                    worksheet.Cell(current+1, col+1).Style.Fill.BackgroundColor = XLColor.Yellow;
                    worksheet.Columns().AdjustToContents();
                }

                foreach (DataRow row in table.Rows)
                {
                    current++;
                    for (int col = 0; col < table.Columns.Count; col++)
                    {
                        worksheet.Cell(current+2, col+1).Value = row[col].ToString();
                    }
                }
        }
    }

    
    //
    // Console.WriteLine(wires.Count + " ===> Wires foud in file");
    //
    // foreach (var wire in wires )
    // {
    //     Console.WriteLine("Wire display name: "+wire.Attribute("displayName").Value);
    //     
    //     worksheet.Cell("A"+).InsertData(wire.Attribute("displayName").Value);
    //     //worksheet.Cells = wire.Attribute("displayName").Value;
    //     
    //     var connectionWire = wire.Descendants().Where(des => des.Name.LocalName == "connection").ToList();
    //     foreach (var connection in connectionWire)
    //     {
    //         Console.WriteLine("Connection pinRef: " + connection.Attribute("pinref").Value);
    //     }
    //     Console.WriteLine("========End of wire===============");
    // }
    //
     workbook.SaveAs("/Users/abderahman/RiderProjects/LeoniTest/LeoniTest/ExcelG1.xlsx");
