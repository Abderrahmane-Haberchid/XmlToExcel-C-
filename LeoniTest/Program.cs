


using System.Xml.Linq;
using ClosedXML.Excel;


IXLWorkbook workbook = new XLWorkbook();
var worksheet = workbook.Worksheets.Add("XmlToExcel");

XDocument xmlDoc = XDocument.Load("/Users/abderahman/RiderProjects/LeoniTest/LeoniTest/XmlFile/CC24E-MAIN_HARNESS_XX24 NV-HX2ML.xml");

var wires = xmlDoc.Descendants().Where(des => des.Name.LocalName == "wire").ToList();

Console.WriteLine(wires.Count + " ===> Wires foud in file");

foreach (var wire in wires )
{
    Console.WriteLine("Wire display name: "+wire.Attribute("displayName").Value);
    
    worksheet.Cell("A1").InsertData(wire.Attribute("displayName").Value);
    
    var connectionWire = wire.Descendants().Where(des => des.Name.LocalName == "connection").ToList();
    foreach (var connection in connectionWire)
    {
        Console.WriteLine("Connection pinRef: " + connection.Attribute("pinref").Value);
    }
    Console.WriteLine("========End of wire===============");
}

workbook.SaveAs("/Users/abderahman/RiderProjects/LeoniTest/LeoniTest/ExcelG1.xlsx");
