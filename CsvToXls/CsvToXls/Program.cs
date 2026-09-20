using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security;
using System.Text;
using System.Xml.Linq;

namespace CsvToXls
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args[0] == "toxls")
            {
                ConvertToXls(args[1], args[2]);
            }
            else if (args[0] == "tocsv")
            {
                ConvertToCsv(args[1], args[2]);
            }
        }

        public static void ConvertToXls(string csvPath, string xlsxPath)
        {
            var rows = File.ReadAllLines(csvPath);

            using (var fileStream = new FileStream(xlsxPath, FileMode.Create))
            {
                using (var archive = new ZipArchive(fileStream, ZipArchiveMode.Create))
                {
                    AddEntry(archive, "[Content_Types].xml", GetContentTypesXml());
                    AddEntry(archive, "_rels/.rels", GetRelsXml());
                    AddEntry(archive, "xl/workbook.xml", GetWorkbookXml());
                    AddEntry(archive, "xl/_rels/workbook.xml.rels", GetWorkbookRelsXml());
                    AddEntry(archive, "xl/worksheets/sheet1.xml", GetWorksheetXml(rows));
                }
            }
        }

        private static void AddEntry(ZipArchive archive, string path, string content)
        {
            var entry = archive.CreateEntry(path);

            using (var writer = new StreamWriter(entry.Open(), Encoding.UTF8))
            {
                writer.Write(content);
            }  
        }

        private static string GetWorksheetXml(string[] csvLines)
        {
            var sb = new StringBuilder();

            sb.AppendLine(@"<?xml version=""1.0"" encoding=""UTF-8""?>");
            sb.AppendLine(@"<worksheet xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main"">");
            sb.AppendLine("<sheetData>");

            int rowIndex = 1;

            foreach (var line in csvLines)
            {
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                int separatorIndex = line.IndexOf(';');

                if (separatorIndex < 0)
                    continue;

                string key = EscapeXml(line.Substring(0, separatorIndex));
                string value = EscapeXml(line.Substring(separatorIndex + 1));

                sb.AppendLine($@"<row r=""{rowIndex}"">");
                sb.AppendLine($@"<c r=""A{rowIndex}"" t=""inlineStr""><is><t>{key}</t></is></c>");
                sb.AppendLine($@"<c r=""B{rowIndex}"" t=""inlineStr""><is><t>{value}</t></is></c>");
                sb.AppendLine("</row>");

                rowIndex++;
            }

            sb.AppendLine("</sheetData>");
            sb.AppendLine("</worksheet>");

            return sb.ToString();
        }

        private static string EscapeXml(string value)
        {
            return SecurityElement.Escape(value) ?? string.Empty;
        }

        private static string GetContentTypesXml() =>
    @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Types xmlns=""http://schemas.openxmlformats.org/package/2006/content-types"">
    <Default Extension=""rels"" ContentType=""application/vnd.openxmlformats-package.relationships+xml""/>
    <Default Extension=""xml"" ContentType=""application/xml""/>
    <Override PartName=""/xl/workbook.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml""/>
    <Override PartName=""/xl/worksheets/sheet1.xml"" ContentType=""application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml""/>
</Types>";

        private static string GetRelsXml() =>
    @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1""
                  Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument""
                  Target=""xl/workbook.xml""/>
</Relationships>";

        private static string GetWorkbookXml() =>
    @"<?xml version=""1.0"" encoding=""UTF-8""?>
<workbook xmlns=""http://schemas.openxmlformats.org/spreadsheetml/2006/main""
          xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships"">
    <sheets>
        <sheet name=""Translations"" sheetId=""1"" r:id=""rId1""/>
    </sheets>
</workbook>";

        private static string GetWorkbookRelsXml() =>
    @"<?xml version=""1.0"" encoding=""UTF-8""?>
<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">
    <Relationship Id=""rId1""
                  Type=""http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""
                  Target=""worksheets/sheet1.xml""/>
</Relationships>";

        public static void ConvertToCsv(string pathXls, string pathCsv)
        {
            Console.WriteLine("ConvertToCsv STARTED");

            using (var archive = ZipFile.OpenRead(pathXls))
            {
                var entry = archive.GetEntry("xl/worksheets/sheet1.xml");
                if (entry == null)
                    throw new FileNotFoundException("sheet1.xml not found");

                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

                var sharedStrings = LoadSharedStrings(archive, ns);

                using (var stream = entry.Open())
                {
                    var doc = XDocument.Load(stream);

                    using (var writer = new StreamWriter(pathCsv, false)) 
                    {
                        foreach (var row in doc.Descendants(ns + "row"))
                        {
                            string key = null;
                            string value = null;

                            foreach (var cell in row.Elements(ns + "c"))
                            {
                                var refAttr = (string)cell.Attribute("r");
                                var cellValue = GetCellValue(cell, ns, sharedStrings);

                                if (refAttr != null && refAttr.StartsWith("A"))
                                    key = cellValue;
                                else if (refAttr != null && refAttr.StartsWith("B"))
                                    value = cellValue;
                            }

                            if (!string.IsNullOrWhiteSpace(key))
                            {
                                writer.WriteLine($"{EscapeCsv(key)};{EscapeCsv(value)}");
                            }
                        }

                        writer.Flush();
                        Console.WriteLine("ConvertToCsv FINISHED");
                    } // overwrite file
                }
            }

        }

        private static string GetCellValue(
    XElement cell,
    XNamespace ns,
    List<string> sharedStrings)
        {
            string type = (string)cell.Attribute("t");

            if (type == "inlineStr")
            {
                return string.Concat(
                    cell.Descendants(ns + "t")
                        .Select(t => t.Value));
            }

            string rawValue = cell.Element(ns + "v")?.Value ?? "";

            if (type == "s" &&
                int.TryParse(rawValue, out int index) &&
                index >= 0 &&
                index < sharedStrings.Count)
            {
                return sharedStrings[index];
            }

            return rawValue;
        }

        private static string EscapeCsv(string value)
        {
            if (value == null)
                return string.Empty;

            if (value.Contains(';') ||
                value.Contains('"') ||
                value.Contains('\n') ||
                value.Contains('\r'))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }

            return value;
        }

        private static List<string> LoadSharedStrings(
    ZipArchive archive,
    XNamespace ns)
        {
            var entry = archive.GetEntry("xl/sharedStrings.xml");

            if (entry == null)
                return new List<string>();

            using (var stream = entry.Open())
            {
                var doc = XDocument.Load(stream);

                return doc.Descendants(ns + "si")
                    .Select(si => string.Concat(
                        si.Descendants(ns + "t")
                            .Select(t => t.Value)))
                    .ToList();
            }
        }
    } 
}
