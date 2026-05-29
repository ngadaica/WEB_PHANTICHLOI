using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Web.Hosting;

namespace WEB_PHANTICHLOI.Models
{
    public class PartMasterLookupResult
    {
        public string Material { get; set; }
        public string Description { get; set; }
        public List<string> Vendors { get; set; }
        public bool IsAvailable { get; set; }
        public string SourcePath { get; set; }
        public string ErrorMessage { get; set; }
    }

    public static class PartMasterLookupService
    {
        private const string DefaultExcelPath = @"\\apbivnsh15\COMMON\90_Temporary\R&D-DE\Phan tich tieng keu_異音解析\TOTAL LIST PART 08.04.2026.xlsx";
        private const string DefaultLocalRelativePath = @"~/App_Data/PartMaster/TOTAL LIST PART 08.04.2026.xlsx";

        private static readonly object SyncRoot = new object();
        private static List<PartMasterRow> _cachedRows;
        private static DateTime _cacheExpiresAt = DateTime.MinValue;
        private static bool _licenseConfigured;
        private static string _cachedSourcePath;

        public static PartMasterLookupResult GetPartInfo(string material)
        {
            var code = (material ?? string.Empty).Trim();
            var result = new PartMasterLookupResult
            {
                Material = code,
                Description = string.Empty,
                Vendors = new List<string>(),
                IsAvailable = false,
                SourcePath = string.Empty,
                ErrorMessage = string.Empty
            };

            if (string.IsNullOrWhiteSpace(code))
            {
                return result;
            }

            string sourcePath;
            string errorMessage;
            if (!TryGetAvailableExcelPath(out sourcePath, out errorMessage))
            {
                result.ErrorMessage = errorMessage;
                return result;
            }

            try
            {
                var materialKey = NormalizeMaterial(code);
                var rows = GetCachedPartRows(sourcePath)
                    .Where(x => string.Equals(x.MaterialKey, materialKey, StringComparison.OrdinalIgnoreCase))
                    .ToList();

                var description = rows
                    .Select(x => x.Description)
                    .FirstOrDefault(x => !string.IsNullOrWhiteSpace(x)) ?? string.Empty;

                result.Description = description;
                result.Vendors = rows
                    .Where(x => string.IsNullOrWhiteSpace(description)
                        || string.Equals(x.Description, description, StringComparison.OrdinalIgnoreCase)
                        || string.IsNullOrWhiteSpace(x.Description))
                    .Select(x => x.Vendor)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(x => x)
                    .ToList();
                result.IsAvailable = true;
                result.SourcePath = sourcePath;

                return result;
            }
            catch (Exception ex)
            {
                LogError(ex);
                result.ErrorMessage = ex.Message;
                result.SourcePath = sourcePath;
                return result;
            }
        }

        private static List<PartMasterRow> GetCachedPartRows(string sourcePath)
        {
            if (_cachedRows != null
                && DateTime.Now <= _cacheExpiresAt
                && string.Equals(_cachedSourcePath, sourcePath, StringComparison.OrdinalIgnoreCase))
            {
                return _cachedRows;
            }

            lock (SyncRoot)
            {
                if (_cachedRows != null
                    && DateTime.Now <= _cacheExpiresAt
                    && string.Equals(_cachedSourcePath, sourcePath, StringComparison.OrdinalIgnoreCase))
                {
                    return _cachedRows;
                }

                _cachedRows = ReadPartRows(sourcePath);
                _cachedSourcePath = sourcePath;
                _cacheExpiresAt = DateTime.Now.AddMinutes(10);
                return _cachedRows;
            }
        }

        private static List<PartMasterRow> ReadPartRows(string sourcePath)
        {
            var rows = new List<PartMasterRow>();

            if (string.IsNullOrWhiteSpace(sourcePath) || !SafeFileExists(sourcePath))
            {
                return rows;
            }

            EnsureLicenseConfigured();

            using (var stream = File.Open(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var package = new ExcelPackage(stream))
            {
                var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                if (worksheet == null || worksheet.Dimension == null)
                {
                    return rows;
                }

                var headerRow = FindHeaderRow(worksheet);
                if (headerRow < 1)
                {
                    return rows;
                }

                var materialIndex = FindColumnIndex(worksheet, headerRow, "MATERIAL");
                var descriptionIndex = FindColumnIndex(worksheet, headerRow, "DESCRIPTION (EN)");
                var vendorIndex = FindColumnIndex(worksheet, headerRow, "VENDOR NAME");

                if (materialIndex < 1 || descriptionIndex < 1 || vendorIndex < 1)
                {
                    return rows;
                }

                for (var row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    var material = GetCellText(worksheet.Cells[row, materialIndex]);
                    if (string.IsNullOrWhiteSpace(material))
                    {
                        continue;
                    }

                    rows.Add(new PartMasterRow
                    {
                        Material = material,
                        MaterialKey = NormalizeMaterial(material),
                        Description = GetCellText(worksheet.Cells[row, descriptionIndex]),
                        Vendor = GetCellText(worksheet.Cells[row, vendorIndex])
                    });
                }
            }

            return rows;
        }

        private static bool TryGetAvailableExcelPath(out string sourcePath, out string errorMessage)
        {
            var candidates = GetCandidatePaths().ToList();

            foreach (var candidate in candidates)
            {
                if (SafeFileExists(candidate))
                {
                    sourcePath = candidate;
                    errorMessage = string.Empty;
                    return true;
                }
            }

            sourcePath = string.Empty;
            errorMessage = "Không tìm thấy file Part Master hoặc IIS không có quyền đọc file.";
            return false;
        }

        private static IEnumerable<string> GetCandidatePaths()
        {
            var configuredPrimary = ConfigurationManager.AppSettings["PartMasterExcelPath"];
            var configuredFallback = ConfigurationManager.AppSettings["PartMasterExcelFallbackPath"];
            var localFallback = HostingEnvironment.MapPath(DefaultLocalRelativePath);

            return new[]
            {
                configuredPrimary,
                DefaultExcelPath,
                configuredFallback,
                localFallback
            }
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Select(x => x.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase);
        }

        private static bool SafeFileExists(string path)
        {
            try
            {
                return !string.IsNullOrWhiteSpace(path) && File.Exists(path);
            }
            catch (Exception ex)
            {
                LogError(ex);
                return false;
            }
        }

        private static void EnsureLicenseConfigured()
        {
            if (_licenseConfigured)
            {
                return;
            }

            lock (SyncRoot)
            {
                if (_licenseConfigured)
                {
                    return;
                }

                ExcelPackage.License.SetNonCommercialOrganization("WEB_PHANTICHLOI");
                _licenseConfigured = true;
            }
        }

        private static int FindHeaderRow(ExcelWorksheet worksheet)
        {
            var maxRow = Math.Min(worksheet.Dimension.End.Row, 10);
            for (var row = worksheet.Dimension.Start.Row; row <= maxRow; row++)
            {
                var materialIndex = FindColumnIndex(worksheet, row, "MATERIAL");
                var descriptionIndex = FindColumnIndex(worksheet, row, "DESCRIPTION (EN)");
                var vendorIndex = FindColumnIndex(worksheet, row, "VENDOR NAME");

                if (materialIndex > 0 && descriptionIndex > 0 && vendorIndex > 0)
                {
                    return row;
                }
            }

            return -1;
        }

        private static int FindColumnIndex(ExcelWorksheet worksheet, int headerRow, string headerName)
        {
            for (var column = worksheet.Dimension.Start.Column; column <= worksheet.Dimension.End.Column; column++)
            {
                if (string.Equals(GetCellText(worksheet.Cells[headerRow, column]), headerName, StringComparison.OrdinalIgnoreCase))
                {
                    return column;
                }
            }

            return -1;
        }

        private static string GetCellText(ExcelRange cell)
        {
            var text = cell == null ? string.Empty : (cell.Text ?? string.Empty);

            if (string.IsNullOrWhiteSpace(text) && cell != null && cell.Value != null)
            {
                text = Convert.ToString(cell.Value);
            }

            return (text ?? string.Empty).Trim();
        }

        private static string NormalizeMaterial(string material)
        {
            var normalized = (material ?? string.Empty).Trim();
            if (normalized.EndsWith(".0", StringComparison.Ordinal))
            {
                normalized = normalized.Substring(0, normalized.Length - 2);
            }

            return normalized.Replace(" ", string.Empty);
        }

        private static void LogError(Exception ex)
        {
            try
            {
                var logPath = HostingEnvironment.MapPath("~/App_Data/part-master-lookup.log");
                if (string.IsNullOrWhiteSpace(logPath))
                {
                    return;
                }

                var directory = Path.GetDirectoryName(logPath);
                if (!string.IsNullOrWhiteSpace(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                File.AppendAllText(
                    logPath,
                    string.Format(
                        "[{0:yyyy-MM-dd HH:mm:ss}] {1}{2}{2}",
                        DateTime.Now,
                        ex,
                        Environment.NewLine));
            }
            catch
            {
            }
        }

        private class PartMasterRow
        {
            public string Material { get; set; }
            public string MaterialKey { get; set; }
            public string Description { get; set; }
            public string Vendor { get; set; }
        }
    }
}