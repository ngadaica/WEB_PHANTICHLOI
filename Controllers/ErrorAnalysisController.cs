using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using WEB_PHANTICHLOI.Models;

namespace WEB_PHANTICHLOI.Controllers
{
    public class ErrorAnalysisController : Controller
    {
        private readonly ErrorAnalysisSqlRepository _repository = new ErrorAnalysisSqlRepository();
        private static readonly string[] PersonInChargeEditors = { "giangng", "manhny", "dapa", "nghiatr" };
        private static readonly string[] DeleteAuthorizedUsers = { "nghiatr", "dapa", "dungny", "luando", "tudo", "chinhlx", "giangng" };

        private class ChartPoint
        {
            public string Label { get; set; }
            public double Value { get; set; }
        }

        public ActionResult Index(ErrorAnalysisIndexViewModel model)
        {
            model = model ?? new ErrorAnalysisIndexViewModel();

            var allErrors = _repository.GetAll() ?? new List<ErrorAnalysis>();
            PopulateManagementNumbers(allErrors);

            var filteredErrors = ApplyIndexFilters(allErrors, model).ToList();
            var pageSize = model.PageSize > 0 ? model.PageSize : 50;
            var totalCount = filteredErrors.Count;
            var totalPages = totalCount == 0 ? 1 : (int)Math.Ceiling(totalCount / (double)pageSize);
            var currentPage = model.CurrentPage > 0 ? model.CurrentPage : 1;

            if (currentPage > totalPages)
            {
                currentPage = totalPages;
            }

            model.Items = filteredErrors
                .Skip((currentPage - 1) * pageSize)
                .Take(pageSize)
                .ToList();
            model.PageSize = pageSize;
            model.CurrentPage = currentPage;
            model.TotalCount = totalCount;
            model.TotalPages = totalPages;
            model.InvestigatingCount = filteredErrors.Count(IsInvestigatingStatus);
            model.CountermeasureDiscussionCount = filteredErrors.Count(IsCountermeasureDiscussionStatus);
            model.CloseCount = filteredErrors.Count(IsCloseStatus);

            ViewBag.CanDeleteErrorAnalysis = CanDeleteErrorAnalysis(GetCurrentUserDisplayName());
            PopulateIndexFilterData(model, allErrors);
            return View(model);
        }

        [HttpGet]
        public ActionResult ExportExcel(ErrorAnalysisIndexViewModel model)
        {
            model = model ?? new ErrorAnalysisIndexViewModel();

            var allErrors = _repository.GetAll() ?? new List<ErrorAnalysis>();
            PopulateManagementNumbers(allErrors);

            var filteredErrors = ApplyIndexFilters(allErrors, model).ToList();

            ExcelPackage.License.SetNonCommercialPersonal("WEB_PHANTICHLOI");

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("DanhSachLoi");

                var headers = new[]
                {
                    "Số quản lý",
                    "Nội dung vấn đề",
                    "Phân loại nguyên nhân",
                    "Thời gian phát sinh",
                    "Công đoạn phát sinh",
                    "Model",
                    "Chủng loại",
                    "Đảm nhiệm",
                    "Nguyên nhân, nội dung điều tra",
                    "Hành động tiếp theo",
                    "Đối sách tạm thời",
                    "Đối sách cố hữu",
                    "Xác nhận hiệu quả",
                    "Thời gian điều tra",
                    "B-act",
                    "Trạng thái B-action",
                    "Dừng xuất hàng",
                    "レポート",
                    "Tài liệu",
                    "Lỗi mãn tính"
                };

                for (var i = 0; i < headers.Length; i++)
                {
                    worksheet.Cells[1, i + 1].Value = headers[i];
                }

                using (var headerRange = worksheet.Cells[1, 1, 1, headers.Length])
                {
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    headerRange.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(47, 103, 189));
                    headerRange.Style.Font.Color.SetColor(System.Drawing.Color.White);
                    headerRange.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    headerRange.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                }

                for (var rowIndex = 0; rowIndex < filteredErrors.Count; rowIndex++)
                {
                    var item = filteredErrors[rowIndex];
                    var row = rowIndex + 2;

                    worksheet.Cells[row, 1].Value = string.IsNullOrWhiteSpace(item.ManagementNumber) ? item.Id.ToString() : item.ManagementNumber;
                    worksheet.Cells[row, 2].Value = item.ProblemContent;
                    worksheet.Cells[row, 3].Value = item.CauseClassification;
                    worksheet.Cells[row, 4].Value = item.OccurrenceDate.HasValue
                        ? item.OccurrenceDate.Value.ToString("dd/MM/yyyy HH:mm")
                        : item.OccurrencePeriod;
                    worksheet.Cells[row, 5].Value = item.OccurrenceProcess;
                    worksheet.Cells[row, 6].Value = item.Model;
                    worksheet.Cells[row, 7].Value = string.IsNullOrWhiteSpace(item.DefectClassification) ? GetCategoryDisplayValue(item.Category) : item.DefectClassification;
                    worksheet.Cells[row, 8].Value = item.PersonInCharge;
                    worksheet.Cells[row, 9].Value = item.DetailedCause;
                    worksheet.Cells[row, 10].Value = item.NextActionSummary;
                    worksheet.Cells[row, 11].Value = item.TemporaryMeasureSummary;
                    worksheet.Cells[row, 12].Value = item.PermanentMeasureSummary;
                    worksheet.Cells[row, 13].Value = string.IsNullOrWhiteSpace(item.MeasureCompletionTime) ? item.Sharing : item.MeasureCompletionTime;
                    worksheet.Cells[row, 14].Value = item.InvestigationHours;
                    worksheet.Cells[row, 15].Value = item.BAction;
                    worksheet.Cells[row, 16].Value = item.BActionStatus;
                    worksheet.Cells[row, 17].Value = item.ShipmentStop;
                    worksheet.Cells[row, 18].Value = item.Report;
                    worksheet.Cells[row, 19].Value = string.Join(Environment.NewLine,
                        ErrorAnalysisDisplayHelper.GetAttachmentPaths(item.AttachmentPath)
                            .Select(ErrorAnalysisDisplayHelper.GetAttachmentDisplayName));
                    worksheet.Cells[row, 20].Value = item.ChronicDefect;
                }

                if (filteredErrors.Any())
                {
                    using (var dataRange = worksheet.Cells[2, 1, filteredErrors.Count + 1, headers.Length])
                    {
                        dataRange.Style.WrapText = true;
                        dataRange.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
                        dataRange.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                        dataRange.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    }
                }

                worksheet.View.FreezePanes(2, 1);
                worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

                worksheet.Column(2).Width = 32;
                worksheet.Column(9).Width = 36;
                worksheet.Column(10).Width = 30;
                worksheet.Column(11).Width = 28;
                worksheet.Column(12).Width = 28;
                worksheet.Column(19).Width = 28;

                var fileName = "DanhSachLoi_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xlsx";
                var bytes = package.GetAsByteArray();
                return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        public ActionResult Create()
        {
            PopulateLookupData();

            var currentUser = GetCurrentUserDisplayName();
            ViewBag.CurrentUser = currentUser;
            ViewBag.CanEditPersonInCharge = CanEditPersonInCharge(currentUser);

            return View(new ErrorAnalysis
            {
                Investigator = currentUser,
                PersonInCharge = currentUser,
                UpdatedBy = currentUser
            });
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create(ErrorAnalysis errorAnalysis, IEnumerable<HttpPostedFileBase> attachments, IEnumerable<HttpPostedFileBase> images)
        {
            PopulateLookupData();
            var currentUser = GetCurrentUserDisplayName();
            ViewBag.CurrentUser = currentUser;
            ViewBag.CanEditPersonInCharge = CanEditPersonInCharge(currentUser);

            if (errorAnalysis == null)
            {
                ModelState.AddModelError(string.Empty, "Không đọc được dữ liệu từ form gửi lên.");
                return View(new ErrorAnalysis());
            }

            if (!CanEditPersonInCharge(currentUser))
            {
                errorAnalysis.PersonInCharge = currentUser;
            }

            errorAnalysis.Category = GetCategoryDisplayValue(errorAnalysis.Category);
            errorAnalysis.AttachmentPath = SaveAttachments(attachments);
            errorAnalysis.ImagePath = SaveAttachments(images);
            PrepareDerivedFields(errorAnalysis, currentUser);

            errorAnalysis.UpdatedBy = currentUser;
            errorAnalysis.UpdatedDate = DateTime.Now;

            _repository.Add(errorAnalysis);
            return RedirectToAction("Index");
        }

        public ActionResult Edit(int id)
        {
            var errorAnalysis = _repository.GetById(id);
            if (errorAnalysis == null)
            {
                return HttpNotFound();
            }

            PopulateLookupData();
            var currentUser = GetCurrentUserDisplayName();
            ViewBag.CurrentUser = currentUser;
            ViewBag.CanEditPersonInCharge = CanEditPersonInCharge(currentUser);
            errorAnalysis.Investigator = currentUser;
            errorAnalysis.Category = GetCategoryDisplayValue(errorAnalysis.Category);
            PopulateManagementNumbers(new List<ErrorAnalysis> { errorAnalysis });
            return View(errorAnalysis);
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit(ErrorAnalysis errorAnalysis, IEnumerable<HttpPostedFileBase> attachments, IEnumerable<HttpPostedFileBase> images)
        {
            PopulateLookupData();
            var currentUser = GetCurrentUserDisplayName();
            ViewBag.CurrentUser = currentUser;
            ViewBag.CanEditPersonInCharge = CanEditPersonInCharge(currentUser);

            if (errorAnalysis == null)
            {
                return HttpNotFound();
            }

            var existing = _repository.GetById(errorAnalysis.Id);
            if (existing == null)
            {
                return HttpNotFound();
            }

            var uploadedAttachments = attachments == null
                ? new List<HttpPostedFileBase>()
                : attachments.Where(x => x != null && x.ContentLength > 0).ToList();
            var uploadedImages = images == null
                ? new List<HttpPostedFileBase>()
                : images.Where(x => x != null && x.ContentLength > 0).ToList();

            if (uploadedAttachments.Any())
            {
                var newAttachmentPath = SaveAttachments(uploadedAttachments);
                DeleteExistingAttachments(existing.AttachmentPath);
                errorAnalysis.AttachmentPath = newAttachmentPath;
            }
            else
            {
                errorAnalysis.AttachmentPath = existing.AttachmentPath;
            }

            if (uploadedImages.Any())
            {
                var newImagePath = SaveAttachments(uploadedImages);
                DeleteExistingAttachments(existing.ImagePath);
                errorAnalysis.ImagePath = newImagePath;
            }
            else
            {
                errorAnalysis.ImagePath = existing.ImagePath;
            }

            errorAnalysis.Category = GetCategoryDisplayValue(errorAnalysis.Category);
            PrepareDerivedFields(errorAnalysis, currentUser);
            if (!CanEditPersonInCharge(currentUser))
            {
                errorAnalysis.PersonInCharge = existing.PersonInCharge;
            }
            else if (string.IsNullOrWhiteSpace(errorAnalysis.PersonInCharge))
            {
                errorAnalysis.PersonInCharge = existing.PersonInCharge;
            }

            errorAnalysis.UpdatedBy = currentUser;
            errorAnalysis.UpdatedDate = DateTime.Now;

            _repository.Update(errorAnalysis);
            return RedirectToAction("Index");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Delete(int id)
        {
            if (!CanDeleteErrorAnalysis(GetCurrentUserDisplayName()))
            {
                return new HttpUnauthorizedResult("Bạn không có quyền xóa bản ghi này.");
            }

            var existing = _repository.GetById(id);
            if (existing != null)
            {
                DeleteExistingAttachments(existing.AttachmentPath);
                DeleteExistingAttachments(existing.ImagePath);
            }

            _repository.Delete(id);
            return RedirectToAction("Index");
        }

        public ActionResult Details(int id)
        {
            var errorAnalysis = _repository.GetById(id);
            if (errorAnalysis == null)
            {
                return HttpNotFound();
            }

            PopulateManagementNumbers(new List<ErrorAnalysis> { errorAnalysis });
            errorAnalysis.Category = GetCategoryDisplayValue(errorAnalysis.Category);
            return View(errorAnalysis);
        }

        public ActionResult Search(string keyword, string factory, string status, string model)
        {
            var errors = _repository.Search(keyword, factory, status, model);
            PopulateManagementNumbers(errors);
            ViewBag.Keyword = keyword;
            ViewBag.FactoryFilter = factory;
            ViewBag.StatusFilter = status;
            ViewBag.ModelFilter = model;
            ViewBag.FactoryOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.Factory, "-- Tất cả --");
            ViewBag.ProgressOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.InvestigationProgress, "-- Tất cả --");

            return View("Search", errors);
        }

        private static IEnumerable<ErrorAnalysis> ApplyIndexFilters(IEnumerable<ErrorAnalysis> errors, ErrorAnalysisIndexViewModel model)
        {
            var query = errors ?? Enumerable.Empty<ErrorAnalysis>();

            query = query.Where(x => MatchesStatusGroup(x, model.FilterStatusGroup));
            query = query.Where(x => ContainsText(x.ManagementNumber, model.FilterManagementNumber));
            query = query.Where(x => ContainsText(x.ProblemContent, model.FilterProblemContent));
            query = query.Where(x => EqualsText(x.CauseClassification, model.FilterCauseClassification));
            query = query.Where(x => ContainsText(x.OccurrencePeriod, model.FilterOccurrencePeriod));
            query = query.Where(x => MatchesOccurrenceMonth(x, model.FilterOccurrenceMonth));
            query = query.Where(x => EqualsText(x.OccurrenceProcess, model.FilterOccurrenceProcess));
            query = query.Where(x => EqualsText(x.Model, model.FilterModel));
            query = query.Where(x => EqualsText(GetDefectDisplayValue(x), model.FilterDefectClassification));
            query = query.Where(x => EqualsText(x.PersonInCharge, model.FilterPersonInCharge));
            query = query.Where(x => ContainsText(x.DetailedCause, model.FilterDetailedCause));
            query = query.Where(x => ContainsText(x.NextActionSummary, model.FilterNextAction));
            query = query.Where(x => ContainsText(x.TemporaryMeasureSummary, model.FilterTemporaryMeasure));
            query = query.Where(x => ContainsText(x.PermanentMeasureSummary, model.FilterPermanentMeasure));
            query = query.Where(x => ContainsText(GetEffectConfirmationValue(x), model.FilterEffectConfirmation));
            query = query.Where(x => ContainsText(GetInvestigationHoursValue(x), model.FilterInvestigationHours));
            query = query.Where(x => EqualsText(x.BAction, model.FilterBAction));
            query = query.Where(x => EqualsText(x.BActionStatus, model.FilterBActionStatus));
            query = query.Where(x => EqualsText(x.ShipmentStop, model.FilterShipmentStop));
            query = query.Where(x => EqualsText(x.Report, model.FilterReport));
            query = query.Where(x => EqualsText(x.ChronicDefect, model.FilterChronicDefect));
            query = query.Where(x => ContainsText(GetAttachmentFilterValue(x), model.FilterAttachments));

            query = query.OrderBy(x => GetOccurrenceDateDistance(x))
                         .ThenByDescending(x => x.OccurrenceDate ?? DateTime.MinValue)
                         .ThenByDescending(x => x.Id);

            return query;
        }

        private static bool MatchesStatusGroup(ErrorAnalysis error, string filterStatusGroup)
        {
            if (string.IsNullOrWhiteSpace(filterStatusGroup))
            {
                return true;
            }

            var normalized = filterStatusGroup.Trim().ToLowerInvariant();

            if (normalized == "close")
            {
                return IsCloseStatus(error);
            }

            if (normalized == "countermeasure")
            {
                return IsCountermeasureDiscussionStatus(error);
            }

            if (normalized == "investigating")
            {
                return IsInvestigatingStatus(error);
            }

            return true;
        }

        private void PopulateIndexFilterData(ErrorAnalysisIndexViewModel model, IEnumerable<ErrorAnalysis> errors)
        {
            var list = errors ?? Enumerable.Empty<ErrorAnalysis>();

            model.CauseOptions = BuildFilterOptions(ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.CauseClassification, "-- tất cả --"));
            model.ProcessOptions = BuildFilterOptions(ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.OccurrenceProcess, "-- tất cả --"));
            model.ModelOptions = BuildFilterOptions(ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.Model, "-- tất cả --"));
            model.PersonInChargeOptions = BuildFilterOptions(list.Select(x => x.PersonInCharge));
            model.DefectOptions = BuildFilterOptions(list.Select(GetDefectDisplayValue));
            model.BActionOptions = BuildFilterOptions(list.Select(x => x.BAction));
            model.BActionStatusOptions = BuildFilterOptions(ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.BActionStatus, "-- tất cả --"));
            model.ShipmentStopOptions = BuildFilterOptions(ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.ShipmentStop, "-- tất cả --"));
            model.ReportOptions = BuildFilterOptions(list.Select(x => x.Report));
            model.ChronicDefectOptions = BuildFilterOptions(ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.ChronicDefect, "-- tất cả --"));
        }

        private static string GetDefectDisplayValue(ErrorAnalysis error)
        {
            return string.IsNullOrWhiteSpace(error.DefectClassification) ? GetCategoryDisplayValue(error.Category) : error.DefectClassification;
        }

        public static string GetCategoryDisplayValue(string category)
        {
            return ErrorAnalysisReferenceData.NormalizeCategory(category);
        }

        private static string GetEffectConfirmationValue(ErrorAnalysis error)
        {
            return string.IsNullOrWhiteSpace(error.MeasureCompletionTime) ? error.Sharing : error.MeasureCompletionTime;
        }

        private static string GetInvestigationHoursValue(ErrorAnalysis error)
        {
            return error.InvestigationHours.HasValue ? error.InvestigationHours.Value.ToString("0.0") : string.Empty;
        }

        private static string GetAttachmentFilterValue(ErrorAnalysis error)
        {
            return string.Join(" ", ErrorAnalysisDisplayHelper.GetDocumentAttachmentPaths(error.AttachmentPath)
                .Select(x => ErrorAnalysisDisplayHelper.GetAttachmentDisplayName(x, 200)));
        }

        private static bool IsCloseStatus(ErrorAnalysis error)
        {
            var status = error != null ? error.Status : null;
            return !string.IsNullOrWhiteSpace(status)
                && status.IndexOf("Close", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool IsCountermeasureDiscussionStatus(ErrorAnalysis error)
        {
            var status = error != null ? error.Status : null;
            return !string.IsNullOrWhiteSpace(status)
                && (status.IndexOf("Đang thảo luận", StringComparison.OrdinalIgnoreCase) >= 0
                    || status.IndexOf("対策", StringComparison.OrdinalIgnoreCase) >= 0
                    || status.IndexOf("cố hữu", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static bool IsInvestigatingStatus(ErrorAnalysis error)
        {
            var status = error != null ? error.Status : null;
            return !string.IsNullOrWhiteSpace(status)
                && !IsCloseStatus(error)
                && !IsCountermeasureDiscussionStatus(error)
                && (status.IndexOf("調査中", StringComparison.OrdinalIgnoreCase) >= 0
                    || status.IndexOf("Đang", StringComparison.OrdinalIgnoreCase) >= 0
                    || status.IndexOf("điều tra", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static bool ContainsText(string source, string filter)
        {
            if (string.IsNullOrWhiteSpace(filter))
            {
                return true;
            }

            return (source ?? string.Empty).IndexOf(filter.Trim(), StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static bool MatchesOccurrenceMonth(ErrorAnalysis error, string filterOccurrenceMonth)
        {
            if (string.IsNullOrWhiteSpace(filterOccurrenceMonth))
            {
                return true;
            }

            DateTime selectedMonth;
            if (!DateTime.TryParseExact(filterOccurrenceMonth + "-01", "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out selectedMonth))
            {
                return true;
            }

            return error != null
                && error.OccurrenceDate.HasValue
                && error.OccurrenceDate.Value.Year == selectedMonth.Year
                && error.OccurrenceDate.Value.Month == selectedMonth.Month;
        }

        private static bool EqualsText(string source, string filter)
        {
            if (string.IsNullOrWhiteSpace(filter))
            {
                return true;
            }

            return string.Equals((source ?? string.Empty).Trim(), filter.Trim(), StringComparison.OrdinalIgnoreCase);
        }

        private static double GetOccurrenceDateDistance(ErrorAnalysis error)
        {
            if (error == null || !error.OccurrenceDate.HasValue)
            {
                return double.MaxValue;
            }

            return Math.Abs((error.OccurrenceDate.Value - DateTime.Now).TotalSeconds);
        }

        private void DeleteExistingAttachments(string attachmentPath)
        {
            foreach (var path in ErrorAnalysisDisplayHelper.GetAttachmentPaths(attachmentPath))
            {
                var physicalPath = Server.MapPath(path);
                if (!string.IsNullOrWhiteSpace(physicalPath) && System.IO.File.Exists(physicalPath))
                {
                    System.IO.File.Delete(physicalPath);
                }
            }
        }

        private static List<SelectListItem> BuildFilterOptions(IEnumerable<SelectListItem> items)
        {
            return items
                .Where(x => x != null)
                .GroupBy(x => x.Value ?? string.Empty)
                .Select(g => g.First())
                .ToList();
        }

        private static List<SelectListItem> BuildFilterOptions(IEnumerable<string> values)
        {
            var items = new List<SelectListItem>
            {
                new SelectListItem { Value = string.Empty, Text = "-- tất cả --" }
            };

            items.AddRange(values
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .Select(x => new SelectListItem { Value = x, Text = x }));

            return items;
        }

        public ActionResult Chart(ErrorAnalysisChartViewModel model)
        {
            model = model ?? new ErrorAnalysisChartViewModel();
            model.SelectedProgresses = NormalizeChartSelections(model.SelectedProgresses);
            model.SelectedCategories = NormalizeChartSelections(model.SelectedCategories);
            model.SelectedModels = NormalizeChartSelections(model.SelectedModels);
            model.SelectedCauseClassifications = NormalizeChartSelections(model.SelectedCauseClassifications);
            model.SelectedPersonInCharges = NormalizeChartSelections(model.SelectedPersonInCharges);
            model.SelectedDefectClassifications = NormalizeChartSelections(model.SelectedDefectClassifications);
            model.SelectedChronicDefects = NormalizeChartSelections(model.SelectedChronicDefects);
            model.SelectedOccurrencePeriods = NormalizeChartSelections(model.SelectedOccurrencePeriods);

            var allErrors = _repository.GetAll() ?? new List<ErrorAnalysis>();
            var errors = ApplyChartFilters(allErrors, model).ToList();

            var progressItems = BuildChartItems(errors, x => x.InvestigationProgress);
            var causeItems = BuildChartItems(errors, x => x.CauseClassification);
            var investigatorItems = BuildChartItems(errors, x => x.Investigator, 10);
            var factoryItems = BuildChartItems(errors, x => x.Factory);
            var modelItems = BuildChartItems(errors, x => x.Model, 10);
            var monthlyItems = BuildMonthlyItems(errors);

            model.TotalCount = errors.Count;
            model.ProgressLabels = progressItems.Select(x => x.Label).ToList();
            model.ProgressData = progressItems.Select(x => x.Value).ToList();
            model.CauseLabels = causeItems.Select(x => x.Label).ToList();
            model.CauseData = causeItems.Select(x => x.Value).ToList();
            model.InvestigatorLabels = investigatorItems.Select(x => x.Label).ToList();
            model.InvestigatorData = investigatorItems.Select(x => x.Value).ToList();
            model.FactoryLabels = factoryItems.Select(x => x.Label).ToList();
            model.FactoryData = factoryItems.Select(x => x.Value).ToList();
            model.ModelLabels = modelItems.Select(x => x.Label).ToList();
            model.ModelData = modelItems.Select(x => x.Value).ToList();
            model.MonthlyLabels = monthlyItems.Select(x => x.Label).ToList();
            model.MonthlyData = monthlyItems.Select(x => x.Value).ToList();
            model.Charts = BuildDashboardCharts(errors);

            PopulateChartFilterOptions(model, allErrors);
            return View(model);
        }

        private static List<string> NormalizeChartSelections(IEnumerable<string> values)
        {
            return (values ?? Enumerable.Empty<string>())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }

        private static IEnumerable<ErrorAnalysis> ApplyChartFilters(IEnumerable<ErrorAnalysis> errors, ErrorAnalysisChartViewModel model)
        {
            var query = errors ?? Enumerable.Empty<ErrorAnalysis>();

            query = query.Where(x => MatchesChartFilter(NormalizeLabel(x.InvestigationProgress), model.SelectedProgresses));
            query = query.Where(x => MatchesChartFilter(NormalizeLabel(GetCategoryDisplayValue(x.Category)), model.SelectedCategories));
            query = query.Where(x => MatchesChartFilter(NormalizeLabel(x.Model), model.SelectedModels));
            query = query.Where(x => MatchesChartFilter(NormalizeLabel(x.CauseClassification), model.SelectedCauseClassifications));
            query = query.Where(x => MatchesChartFilter(NormalizeLabel(x.PersonInCharge), model.SelectedPersonInCharges));
            query = query.Where(x => MatchesChartFilter(NormalizeLabel(GetDefectDisplayValue(x)), model.SelectedDefectClassifications));
            query = query.Where(x => MatchesChartFilter(NormalizeLabel(x.ChronicDefect), model.SelectedChronicDefects));
            query = query.Where(x => MatchesOccurrencePeriodFilter(x, model.SelectedOccurrencePeriods));

            return query;
        }

        private static bool MatchesChartFilter(string value, IEnumerable<string> selectedValues)
        {
            var selections = selectedValues ?? Enumerable.Empty<string>();
            return !selections.Any() || selections.Any(x => string.Equals(x, value, StringComparison.OrdinalIgnoreCase));
        }

        private static bool MatchesOccurrencePeriodFilter(ErrorAnalysis error, IEnumerable<string> selectedValues)
        {
            var selections = selectedValues ?? Enumerable.Empty<string>();
            if (!selections.Any())
            {
                return true;
            }

            var periodKey = GetOccurrencePeriodKey(error);
            return !string.IsNullOrWhiteSpace(periodKey)
                && selections.Any(x => string.Equals(x, periodKey, StringComparison.OrdinalIgnoreCase));
        }

        private static void PopulateChartFilterOptions(ErrorAnalysisChartViewModel model, IEnumerable<ErrorAnalysis> allErrors)
        {
            var list = allErrors ?? Enumerable.Empty<ErrorAnalysis>();

            model.ProgressOptions = BuildChartFilterOptions(list.Select(x => x.InvestigationProgress));
            model.CategoryOptions = BuildChartFilterOptions(list.Select(x => GetCategoryDisplayValue(x.Category)));
            model.ModelOptions = BuildChartFilterOptions(list.Select(x => x.Model));
            model.CauseClassificationOptions = BuildChartFilterOptions(list.Select(x => x.CauseClassification));
            model.PersonInChargeOptions = BuildChartFilterOptions(list.Select(x => x.PersonInCharge));
            model.DefectClassificationOptions = BuildChartFilterOptions(list.Select(GetDefectDisplayValue));
            model.ChronicDefectOptions = BuildChartFilterOptions(list.Select(x => x.ChronicDefect));
            model.OccurrencePeriodOptions = BuildChartOccurrencePeriodOptions(list);
        }

        private static List<string> BuildChartOccurrencePeriodOptions(IEnumerable<ErrorAnalysis> errors)
        {
            return (errors ?? Enumerable.Empty<ErrorAnalysis>())
                .Select(GetOccurrencePeriodKey)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(x => x)
                .ToList();
        }

        private static string GetOccurrencePeriodKey(ErrorAnalysis error)
        {
            return error != null && error.OccurrenceDate.HasValue
                ? error.OccurrenceDate.Value.ToString("yyyy-MM")
                : null;
        }

        private static List<string> BuildChartFilterOptions(IEnumerable<string> values)
        {
            return (values ?? Enumerable.Empty<string>())
                .Select(NormalizeLabel)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();
        }

        [HttpGet]
        public JsonResult GetPartMasterInfo(string material)
        {
            var result = PartMasterLookupService.GetPartInfo(material);
            return Json(new
            {
                success = result != null && result.IsAvailable,
                material = result != null ? result.Material : string.Empty,
                description = result != null ? result.Description : string.Empty,
                vendors = result != null ? result.Vendors : new List<string>(),
                message = result != null ? result.ErrorMessage : "Lookup failed.",
                sourcePath = result != null ? result.SourcePath : string.Empty
            }, JsonRequestBehavior.AllowGet);
        }

        private void PopulateLookupData()
        {
            ViewBag.StageOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.StageClassification, "-- Chọn --");
            ViewBag.ProcessOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.OccurrenceProcess, "-- Chọn --");
            ViewBag.ProgressOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.InvestigationProgress, "-- Chọn --");
            ViewBag.CauseOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.CauseClassification, "-- Chọn --");
            ViewBag.TempMeasureOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.TemporaryMeasureClassification, "-- Chọn --");
            ViewBag.PermMeasureOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.PermanentMeasureClassification, "-- Chọn --");
            ViewBag.FactoryOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.Factory, "-- Chọn --");
            ViewBag.ShipmentStopOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.ShipmentStop, "-- Chọn --");
            ViewBag.InvestigatorOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.Investigator, "-- Chọn --");
            ViewBag.CategoryOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.Category, "-- Chọn --");
            ViewBag.ModelOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.Model, "-- Chọn --");
            ViewBag.SupplierOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.Supplier, "-- Chọn --");
            ViewBag.PhenomenonOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.Phenomenon, "-- Chọn --");
            ViewBag.ChronicDefectOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.ChronicDefect, "-- Chọn --");
            ViewBag.BActionStatusOptions = ErrorAnalysisReferenceData.GetSelectList(ErrorAnalysisLookupGroups.BActionStatus, "-- Chọn --");
        }

        private string SaveAttachments(IEnumerable<HttpPostedFileBase> attachments)
        {
            if (attachments == null)
            {
                return null;
            }

            var uploadFolder = Server.MapPath("~/Uploads");

            if (!System.IO.Directory.Exists(uploadFolder))
            {
                System.IO.Directory.CreateDirectory(uploadFolder);
            }

            var savedFiles = new List<string>();

            foreach (var attachment in attachments.Where(x => x != null && x.ContentLength > 0))
            {
                var originalFileName = System.IO.Path.GetFileName(attachment.FileName);
                if (string.IsNullOrWhiteSpace(originalFileName))
                {
                    continue;
                }

                var uniqueFileName = Guid.NewGuid().ToString("N") + "_" + originalFileName;
                var path = System.IO.Path.Combine(uploadFolder, uniqueFileName);
                attachment.SaveAs(path);
                savedFiles.Add("~/Uploads/" + uniqueFileName);
            }

            return savedFiles.Any() ? string.Join("|", savedFiles) : null;
        }

        private void PrepareDerivedFields(ErrorAnalysis errorAnalysis, string currentUser)
        {
            errorAnalysis.Status = errorAnalysis.InvestigationProgress;
            errorAnalysis.OccurrencePeriod = errorAnalysis.OccurrenceDate.HasValue
                ? errorAnalysis.OccurrenceDate.Value.ToString("dd/MM/yyyy HH:mm")
                : null;
            errorAnalysis.NightShiftDays = CalculateNightShiftDays(errorAnalysis);
            errorAnalysis.InvestigationHours = CalculateInvestigationHours(errorAnalysis);

            errorAnalysis.Investigator = currentUser;
            if (string.IsNullOrWhiteSpace(errorAnalysis.PersonInCharge))
            {
                errorAnalysis.PersonInCharge = currentUser;
            }
            errorAnalysis.ImageStatus = errorAnalysis.Phenomenon;
            errorAnalysis.NextActionSummary = errorAnalysis.NextAction;
            errorAnalysis.TemporaryMeasureSummary = errorAnalysis.TemporaryMeasureDetail;
            errorAnalysis.PermanentMeasureSummary = errorAnalysis.PermanentMeasureDetail;
        }

        private static double? CalculateInvestigationHours(ErrorAnalysis errorAnalysis)
        {
            if (errorAnalysis == null
                || !errorAnalysis.InvestigationStartTime.HasValue
                || !errorAnalysis.InvestigationEndTime.HasValue)
            {
                return null;
            }

            var diffDays = (errorAnalysis.InvestigationEndTime.Value - errorAnalysis.InvestigationStartTime.Value).TotalDays;
            var daysOff = errorAnalysis.DaysOff ?? 0;
            var nightShiftDays = errorAnalysis.NightShiftDays ?? 0;
            var measurementWaitHours = errorAnalysis.MeasurementWaitHours ?? 0;
            var nightBreakHours = CalculateNightBreakHours(errorAnalysis.InvestigationStartTime.Value);
            var dayBreakHours = CalculateDayBreakHours(errorAnalysis.InvestigationStartTime.Value, errorAnalysis.InvestigationEndTime.Value);
            var startTime = errorAnalysis.InvestigationStartTime.Value.TimeOfDay;
            var sameDate = errorAnalysis.InvestigationStartTime.Value.Date == errorAnalysis.InvestigationEndTime.Value.Date;
            var baseHours = (diffDays - daysOff) * 24;
            var hours = startTime > new TimeSpan(20, 0, 0) && sameDate
                ? baseHours
                : baseHours - (nightShiftDays - daysOff) * nightBreakHours;

            hours = hours - dayBreakHours - measurementWaitHours;
            return Math.Round(hours, 1, MidpointRounding.AwayFromZero);
        }

        private static int? CalculateNightShiftDays(ErrorAnalysis errorAnalysis)
        {
            if (errorAnalysis == null
                || !errorAnalysis.InvestigationStartTime.HasValue
                || !errorAnalysis.InvestigationEndTime.HasValue)
            {
                return errorAnalysis?.NightShiftDays;
            }

            var start = errorAnalysis.InvestigationStartTime.Value;
            var end = errorAnalysis.InvestigationEndTime.Value;
            if (end < start)
            {
                return 0;
            }

            var dayDiff = (end.Date - start.Date).Days;
            var value = start.TimeOfDay > new TimeSpan(19, 30, 0)
                ? dayDiff - 1
                : dayDiff - (end < start.Date.AddDays(1).Add(new TimeSpan(7, 30, 0)) ? 1 : 0);

            return Math.Max(0, value);
        }

        private static double CalculateNightBreakHours(DateTime investigationStartTime)
        {
            var startTime = investigationStartTime.TimeOfDay;
            var nightStart = new TimeSpan(17, 30, 0);
            var nightEnd = new TimeSpan(7, 50, 0);
            var breakTime = startTime > nightStart
                ? (TimeSpan.FromDays(1) - startTime) + nightEnd
                : (TimeSpan.FromDays(1) - nightStart) + nightEnd;

            return breakTime.TotalHours;
        }

        private static double CalculateDayBreakHours(DateTime investigationStartTime, DateTime investigationEndTime)
        {
            if (investigationEndTime < investigationStartTime)
            {
                return 0;
            }

            var breaks = new[]
            {
                new { Start = new TimeSpan(10, 0, 0), End = new TimeSpan(10, 10, 0) },
                new { Start = new TimeSpan(12, 0, 0), End = new TimeSpan(12, 40, 0) },
                new { Start = new TimeSpan(15, 0, 0), End = new TimeSpan(15, 10, 0) }
            };
            var inclusiveDays = (investigationEndTime.Date - investigationStartTime.Date).Days + 1;
            var totalBreakTime = TimeSpan.FromHours(inclusiveDays);
            var startTime = investigationStartTime.TimeOfDay;
            var endTime = investigationEndTime.TimeOfDay;

            foreach (var breakTime in breaks)
            {
                totalBreakTime -= TimeSpan.FromTicks(Math.Max(0, Min(startTime, breakTime.End).Ticks - breakTime.Start.Ticks));
                totalBreakTime -= TimeSpan.FromTicks(Math.Max(0, breakTime.End.Ticks - Max(endTime, breakTime.Start).Ticks));
            }

            return Math.Max(0, totalBreakTime.TotalHours);
        }

        private static TimeSpan Min(TimeSpan left, TimeSpan right)
        {
            return left < right ? left : right;
        }

        private static TimeSpan Max(TimeSpan left, TimeSpan right)
        {
            return left > right ? left : right;
        }

        private void PopulateManagementNumbers(IList<ErrorAnalysis> errors)
        {
            if (errors == null || errors.Count == 0)
            {
                return;
            }

            var allErrors = _repository.GetAll() ?? new List<ErrorAnalysis>();
            var sequenceMap = allErrors
                .Where(x => x != null)
                .OrderBy(x => x.Id)
                .GroupBy(x => NormalizeModelCode(x.Model))
                .ToDictionary(
                    g => g.Key,
                    g => g.Select((item, index) => new { item.Id, Sequence = index + 1 })
                          .ToDictionary(x => x.Id, x => x.Sequence));

            foreach (var error in errors)
            {
                if (error == null)
                {
                    continue;
                }

                var modelCode = NormalizeModelCode(error.Model);
                int sequence;
                if (sequenceMap.ContainsKey(modelCode) && sequenceMap[modelCode].TryGetValue(error.Id, out sequence))
                {
                    error.ManagementNumber = string.Format("{0}_{1:000}", modelCode, sequence);
                }
                else
                {
                    error.ManagementNumber = string.Format("{0}_{1:000}", modelCode, error.Id);
                }
            }
        }

        private static string NormalizeModelCode(string model)
        {
            return string.IsNullOrWhiteSpace(model) ? "NA" : model.Trim();
        }

        private static bool CanEditPersonInCharge(string currentUser)
        {
            return !string.IsNullOrWhiteSpace(currentUser)
                && PersonInChargeEditors.Contains(currentUser.Trim(), StringComparer.OrdinalIgnoreCase);
        }

        private static bool CanDeleteErrorAnalysis(string currentUser)
        {
            return !string.IsNullOrWhiteSpace(currentUser)
                && DeleteAuthorizedUsers.Contains(currentUser.Trim(), StringComparer.OrdinalIgnoreCase);
        }

        private static string GetCurrentUserDisplayName()
        {
            var identity = System.Web.HttpContext.Current?.User?.Identity;
            if (identity == null)
            {
                var logonIdentity = System.Web.HttpContext.Current?.Request?.LogonUserIdentity;
                if (logonIdentity == null || string.IsNullOrWhiteSpace(logonIdentity.Name))
                {
                    return "Unknown";
                }

                return ExtractUserName(logonIdentity.Name);
            }

            var name = identity.Name;
            if (string.IsNullOrWhiteSpace(name))
            {
                var logonIdentity = System.Web.HttpContext.Current?.Request?.LogonUserIdentity;
                return logonIdentity == null || string.IsNullOrWhiteSpace(logonIdentity.Name)
                    ? "Unknown"
                    : ExtractUserName(logonIdentity.Name);
            }

            return ExtractUserName(name);
        }

        private static string ExtractUserName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return "Unknown";
            }

            var separatorIndex = name.LastIndexOf('\\');
            return separatorIndex >= 0 && separatorIndex < name.Length - 1
                ? name.Substring(separatorIndex + 1)
                : name;
        }

        private static List<ChartItem> BuildChartItems(IEnumerable<ErrorAnalysis> errors, Func<ErrorAnalysis, string> selector, int take = 0)
        {
            IEnumerable<ChartItem> query = errors
                .GroupBy(x => NormalizeLabel(selector(x)))
                .Select(x => new ChartItem
                {
                    Label = x.Key,
                    Value = x.Count()
                })
                .OrderByDescending(x => x.Value)
                .ThenBy(x => x.Label);

            if (take > 0)
            {
                query = query.Take(take);
            }

            return query.ToList();
        }

        private static List<ChartItem> BuildMonthlyItems(IEnumerable<ErrorAnalysis> errors)
        {
            return errors
                .Where(x => x.OccurrenceDate.HasValue)
                .GroupBy(x => new
                {
                    x.OccurrenceDate.Value.Year,
                    x.OccurrenceDate.Value.Month
                })
                .OrderBy(x => x.Key.Year)
                .ThenBy(x => x.Key.Month)
                .Select(x => new ChartItem
                {
                    Label = x.Key.Month.ToString("00") + "/" + x.Key.Year,
                    Value = x.Count()
                })
                .ToList();
        }

        private static List<ErrorAnalysisChartCardViewModel> BuildDashboardCharts(IEnumerable<ErrorAnalysis> errors)
        {
            var list = (errors ?? Enumerable.Empty<ErrorAnalysis>()).ToList();

            return new List<ErrorAnalysisChartCardViewModel>
            {
                BuildSimpleBarChart(
                    "modelOccurrenceChart",
                    "SỐ VỤ PHÁT SINH THEO MODEL",
                    list.GroupBy(x => NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)))
                        .Select(g => new ChartPoint { Label = g.Key, Value = g.Count() })
                        .OrderByDescending(x => x.Value)
                        .ThenBy(x => x.Label)
                        .ToList(),
                    "Số vụ",
                    "#f59e0b"),

                BuildPieChart(
                    "causeRatioChart",
                    "PHÂN LOẠI TỶ LỆ THEO NGUYÊN NHÂN LỖI",
                    list.GroupBy(x => NormalizeLabel(x.CauseClassification))
                        .Select(g => new ChartPoint { Label = g.Key, Value = g.Count() })
                        .OrderByDescending(x => x.Value)
                        .ThenBy(x => x.Label)
                        .ToList()),

                BuildCountAndMetricByModelChart(
                    "avgHoursByCategoryChart",
                    "THỜI GIAN ĐIỀU TRA TRUNG BÌNH THEO MODEL",
                    list,
                    x => NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)),
                    x => x.InvestigationHours,
                    values => values.Any() ? values.Average() : 0,
                    "Time (h)"),

                BuildStackedMonthlyChart(
                    "avgHoursByMonthChart",
                    "TỔNG THỜI GIAN ĐIỀU TRA TRUNG BÌNH THEO THÁNG",
                    list.Where(x => x.OccurrenceDate.HasValue && x.InvestigationHours.HasValue).ToList(),
                    x => NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)),
                    group => group.Any() ? group.Average(y => y.InvestigationHours ?? 0) : 0,
                    "Time (h)"),

                BuildCountAndMetricByModelChart(
                    "avgHoursByModelChart",
                    "THỜI GIAN ĐIỀU TRA TRUNG BÌNH THEO TỪNG MODEL",
                    list,
                    x => NormalizeLabel(x.Model),
                    x => x.InvestigationHours,
                    values => values.Any() ? values.Average() : 0,
                    "Time (h)"),

                BuildCountAndMetricByModelChart(
                    "totalHoursByModelChart",
                    "TỔNG THỜI GIAN ĐIỀU TRA THEO MODEL",
                    list,
                    x => NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)),
                    x => x.InvestigationHours,
                    values => values.Sum(),
                    "Time (h)"),

                BuildStackedByLabelChart(
    "investigatorDefectChart",
    "SỐ VỤ ĐIỀU TRA THEO ĐẢM NHIỆM",
    list,
    x => NormalizeLabel(x.PersonInCharge),
    x => NormalizeChartCategoryLabel(GetDefectDisplayValue(x)),
    0,
    6,
    "Số vụ",
    false),

                BuildStackedMonthlyChart(
                    "totalDefectsByMonthChart",
                    "TỔNG SỐ LỖI ĐIỀU TRA THEO THÁNG",
                    list.Where(x => x.OccurrenceDate.HasValue).ToList(),
                    x => NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)),
                    group => group.Count(),
                    "Số vụ"),

                BuildCauseRatioByCategoryChart(list),

                BuildSimpleBarChart(
                    "unknownCauseModelCountChart",
                    "SỐ VỤ KHÔNG RÕ NN THEO MODEL",
                    list.Where(IsUnknownCauseChartItem)
                        .GroupBy(x => NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)))
                        .Select(g => new ChartPoint { Label = g.Key, Value = g.Count() })
                        .OrderByDescending(x => x.Value)
                        .ThenBy(x => x.Label)
                        .ToList(),
                    "Số vụ",
                    "#3b82f6"),

                BuildMonthlyAggregateChart(
                    "unknownCauseMonthCountChart",
                    "SỐ VỤ KHÔNG RÕ NN THEO THÁNG",
                    list.Where(x => IsUnknownCauseChartItem(x) && x.OccurrenceDate.HasValue).ToList(),
                    group => group.Count(),
                    "Số vụ",
                    "#0ea5e9"),

                BuildSimpleBarChart(
                    "unknownCauseAvgHoursByModelChart",
                    "THỜI GIAN ĐIỀU TRA KHÔNG RÕ NN TRUNG BÌNH THEO MODEL",
                    list.Where(x => IsUnknownCauseChartItem(x) && x.InvestigationHours.HasValue)
                        .GroupBy(x => NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)))
                        .Select(g => new ChartPoint { Label = g.Key, Value = g.Average(y => y.InvestigationHours ?? 0) })
                        .OrderByDescending(x => x.Value)
                        .ThenBy(x => x.Label)
                        .ToList(),
                    "Time (h)",
                    "#f97316"),

                BuildMonthlyAggregateChart(
                    "unknownCauseAvgHoursByMonthChart",
                    "THỜI GIAN ĐIỀU TRA LỖI KHÔNG RÕ NGUYÊN NHÂN TRUNG BÌNH THEO THÁNG",
                    list.Where(x => IsUnknownCauseChartItem(x) && x.OccurrenceDate.HasValue && x.InvestigationHours.HasValue).ToList(),
                    group => group.Average(y => y.InvestigationHours ?? 0),
                    "Time (h)",
                    "#8b5cf6")
            };
        }

        private static ErrorAnalysisChartCardViewModel BuildSimpleBarChart(string id, string title, IEnumerable<ChartPoint> data, string yAxisTitle, string color)
        {
            List<ChartPoint> items = data != null ? data.ToList() : new List<ChartPoint>();
            if (!items.Any())
            {
                items.Add(new ChartPoint { Label = "Chưa có dữ liệu", Value = 0d });
            }

            return new ErrorAnalysisChartCardViewModel
            {
                Id = id,
                Title = title,
                ChartType = "bar",
                ShowLegend = false,
                ShowDataLabels = true,
                YAxisTitle = yAxisTitle,
                Labels = items.Select(x => x.Label).ToList<string>(),
                Datasets = new List<ErrorAnalysisChartDatasetViewModel>
                {
                    new ErrorAnalysisChartDatasetViewModel
                    {
                        Label = yAxisTitle,
                        Data = items.Select(x => x.Value).ToList<double>(),
                        BackgroundColor = color,
                        BorderColor = color,
                        BorderWidth = 1,
                        Order = 1
                    }
                }
            };
        }

        private static ErrorAnalysisChartCardViewModel BuildPieChart(string id, string title, IEnumerable<ChartPoint> data)
        {
            List<ChartPoint> items = data != null ? data.ToList() : new List<ChartPoint>();
            if (!items.Any())
            {
                items.Add(new ChartPoint { Label = "Chưa có dữ liệu", Value = 1d });
            }

            var colors = GetColorPalette(items.Count);
            return new ErrorAnalysisChartCardViewModel
            {
                Id = id,
                Title = title,
                ChartType = "pie",
                ShowLegend = true,
                ShowDataLabels = true,
                Labels = items.Select(x => x.Label).ToList<string>(),
                Datasets = new List<ErrorAnalysisChartDatasetViewModel>
                {
                    new ErrorAnalysisChartDatasetViewModel
                    {
                        Label = title,
                        Data = items.Select(x => x.Value).ToList<double>(),
                        BackgroundColor = colors,
                        BorderColor = "#ffffff",
                        BorderWidth = 2,
                        Order = 1
                    }
                }
            };
        }

        private static ErrorAnalysisChartCardViewModel BuildCountAndMetricByModelChart(
            string id,
            string title,
            IEnumerable<ErrorAnalysis> errors,
            Func<ErrorAnalysis, string> labelSelector,
            Func<ErrorAnalysis, double?> metricSelector,
            Func<IEnumerable<double>, double> aggregator,
            string metricAxisTitle)
        {
            var grouped = (errors ?? Enumerable.Empty<ErrorAnalysis>())
                .GroupBy(labelSelector)
                .Select(g => new
                {
                    Label = g.Key,
                    Count = g.Count(),
                    Metric = aggregator(g.Select(x => metricSelector(x) ?? 0))
                })
                .OrderByDescending(x => x.Count)
                .ThenBy(x => x.Label)
                .ToList();

            if (!grouped.Any())
            {
                grouped.Add(new { Label = "Chưa có dữ liệu", Count = 0, Metric = 0d });
            }

            return new ErrorAnalysisChartCardViewModel
            {
                Id = id,
                Title = title,
                ChartType = "bar",
                ShowLegend = true,
                ShowDataLabels = true,
                YAxisTitle = "Số vụ",
                SecondaryYAxisTitle = metricAxisTitle,
                CanvasHeight = 400,
                CanvasMinWidth = grouped.Count > 8 ? grouped.Count * 88 : (int?)null,
                Labels = grouped.Select(x => x.Label).ToList(),
                Datasets = new List<ErrorAnalysisChartDatasetViewModel>
                {
                    new ErrorAnalysisChartDatasetViewModel
                    {
                        Label = "Số vụ",
                        Type = "bar",
                        Data = grouped.Select(x => (double)x.Count).ToList(),
                        BackgroundColor = "#facc15",
                        BorderColor = "#eab308",
                        BorderWidth = 1,
                        YAxisId = "y",
                        Order = 2
                    },
                    new ErrorAnalysisChartDatasetViewModel
                    {
                        Label = metricAxisTitle,
                        Type = "line",
                        Data = grouped.Select(x => Math.Round(x.Metric, 2)).ToList(),
                        BackgroundColor = "rgba(59,130,246,0.15)",
                        BorderColor = "#3b82f6",
                        BorderWidth = 2,
                        Fill = false,
                        Tension = 0.25,
                        YAxisId = "y1",
                        Order = 1
                    }
                }
            };
        }

        private static ErrorAnalysisChartCardViewModel BuildStackedMonthlyChart(
            string id,
            string title,
            List<ErrorAnalysis> errors,
            Func<ErrorAnalysis, string> seriesSelector,
            Func<List<ErrorAnalysis>, double> aggregateSelector,
            string yAxisTitle)
        {
            var list = errors ?? new List<ErrorAnalysis>();
            var monthLabels = list.Where(x => x.OccurrenceDate.HasValue)
                .Select(x => new DateTime(x.OccurrenceDate.Value.Year, x.OccurrenceDate.Value.Month, 1))
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            if (!monthLabels.Any())
            {
                monthLabels.Add(DateTime.Today);
            }

            var series = list.Select(seriesSelector)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();

            if (!series.Any())
            {
                series.Add("Chưa có dữ liệu");
            }

            var colors = GetColorPalette(series.Count);
            var datasets = new List<ErrorAnalysisChartDatasetViewModel>();
            for (var i = 0; i < series.Count; i++)
            {
                var currentSeries = series[i];
                datasets.Add(new ErrorAnalysisChartDatasetViewModel
                {
                    Label = currentSeries,
                    Type = "bar",
                    Stack = "stack-0",
                    BackgroundColor = colors[i],
                    BorderColor = colors[i],
                    BorderWidth = 1,
                    Data = monthLabels.Select(month =>
                        aggregateSelector(list.Where(x => x.OccurrenceDate.HasValue
                            && x.OccurrenceDate.Value.Year == month.Year
                            && x.OccurrenceDate.Value.Month == month.Month
                            && string.Equals(seriesSelector(x), currentSeries, StringComparison.OrdinalIgnoreCase)).ToList()))
                        .Select(x => Math.Round(x, 2))
                        .ToList()
                });
            }

            return new ErrorAnalysisChartCardViewModel
            {
                Id = id,
                Title = title,
                ChartType = "bar",
                Stacked = true,
                ShowLegend = true,
                ShowDataLabels = false,
                YAxisTitle = yAxisTitle,
                Labels = monthLabels.Select(GetMonthLabel).ToList(),
                Datasets = datasets
            };
        }

        private static ErrorAnalysisChartCardViewModel BuildStackedByLabelChart(
            string id,
            string title,
            IEnumerable<ErrorAnalysis> errors,
            Func<ErrorAnalysis, string> labelSelector,
            Func<ErrorAnalysis, string> seriesSelector,
            int maxLabels,
            int maxSeries,
            string yAxisTitle,
            bool horizontal)
        {
            var list = (errors ?? Enumerable.Empty<ErrorAnalysis>()).ToList();
            IEnumerable<ChartPoint> labelQuery = list.GroupBy(labelSelector)
                .Select(g => new ChartPoint { Label = g.Key, Value = g.Count() })
                .OrderByDescending(x => x.Value)
                .ThenBy(x => x.Label);

            if (maxLabels > 0)
            {
                labelQuery = labelQuery.Take(maxLabels);
            }

            var labels = labelQuery
                .Select(x => x.Label)
                .ToList();

            if (!labels.Any())
            {
                labels.Add("Chưa có dữ liệu");
            }

            var series = list.GroupBy(seriesSelector)
                .Select(g => new { Label = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .ThenBy(x => x.Label)
                .Take(maxSeries)
                .Select(x => x.Label)
                .ToList();

            if (!series.Any())
            {
                series.Add("Chưa có dữ liệu");
            }

            var colors = GetColorPalette(series.Count);
            var datasets = new List<ErrorAnalysisChartDatasetViewModel>();
            for (var i = 0; i < series.Count; i++)
            {
                var currentSeries = series[i];
                datasets.Add(new ErrorAnalysisChartDatasetViewModel
                {
                    Label = currentSeries,
                    Type = "bar",
                    Stack = "stack-0",
                    BackgroundColor = colors[i],
                    BorderColor = colors[i],
                    BorderWidth = 1,
                    Data = labels.Select(label => (double)list.Count(x => string.Equals(labelSelector(x), label, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(seriesSelector(x), currentSeries, StringComparison.OrdinalIgnoreCase))).ToList()
                });
            }

            return new ErrorAnalysisChartCardViewModel
            {
                Id = id,
                Title = title,
                ChartType = "bar",
                CanvasHeight = horizontal ? Math.Max(400, labels.Count * 44) : 400,
                CanvasMinWidth = horizontal ? (int?)null : Math.Max(960, labels.Count * 88),
                Stacked = true,
                Horizontal = horizontal,
                ShowLegend = true,
                ShowDataLabels = false,
                YAxisTitle = yAxisTitle,
                Labels = labels,
                Datasets = datasets
            };
        }

        private static ErrorAnalysisChartCardViewModel BuildCauseRatioByCategoryChart(IEnumerable<ErrorAnalysis> errors)
        {
            var list = (errors ?? Enumerable.Empty<ErrorAnalysis>()).ToList();
            var totalCount = list.Count;
            var causes = list.GroupBy(x => NormalizeLabel(x.CauseClassification))
                .Select(g => new { Label = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .ThenBy(x => x.Label)
                .Select(x => x.Label)
                .ToList();

            if (!causes.Any())
            {
                causes.Add("Chưa có dữ liệu");
            }

            var categories = list.GroupBy(x => NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)))
                .Select(g => new { Label = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .ThenBy(x => x.Label)
                .Take(6)
                .Select(x => x.Label)
                .ToList();

            if (!categories.Any())
            {
                categories.Add("Chưa có dữ liệu");
            }

            var colors = GetColorPalette(categories.Count);
            var datasets = new List<ErrorAnalysisChartDatasetViewModel>();
            for (var i = 0; i < categories.Count; i++)
            {
                var category = categories[i];
                datasets.Add(new ErrorAnalysisChartDatasetViewModel
                {
                    Label = category,
                    Type = "bar",
                    Stack = "stack-0",
                    BackgroundColor = colors[i],
                    BorderColor = colors[i],
                    BorderWidth = 1,
                    Data = causes.Select(cause =>
                    {
                        if (totalCount == 0)
                        {
                            return 0d;
                        }

                        var count = list.Count(x => string.Equals(NormalizeLabel(x.CauseClassification), cause, StringComparison.OrdinalIgnoreCase)
                            && string.Equals(NormalizeChartCategoryLabel(GetCategoryDisplayValue(x.Category)), category, StringComparison.OrdinalIgnoreCase));
                        return Math.Round((count * 100d) / totalCount, 2);
                    }).ToList()
                });
            }

            return new ErrorAnalysisChartCardViewModel
            {
                Id = "causeRatioByCategoryChart",
                Title = "PHÂN LOẠI TỶ LỆ LỖI THEO NGUYÊN NHÂN / CHỦNG LOẠI",
                ChartType = "bar",
                Stacked = true,
                ShowLegend = true,
                ShowDataLabels = true,
                YAxisTitle = "Tỷ lệ lỗi %",
                Labels = causes,
                Datasets = datasets
            };
        }

        private static ErrorAnalysisChartCardViewModel BuildMonthlyAggregateChart(
            string id,
            string title,
            List<ErrorAnalysis> errors,
            Func<List<ErrorAnalysis>, double> aggregateSelector,
            string yAxisTitle,
            string color)
        {
            var list = errors ?? new List<ErrorAnalysis>();
            var items = list.Where(x => x.OccurrenceDate.HasValue)
                .GroupBy(x => new { x.OccurrenceDate.Value.Year, x.OccurrenceDate.Value.Month })
                .OrderBy(g => g.Key.Year)
                .ThenBy(g => g.Key.Month)
                .Select(g => new ChartPoint
                {
                    Label = GetMonthLabel(new DateTime(g.Key.Year, g.Key.Month, 1)),
                    Value = Math.Round(aggregateSelector(g.ToList()), 2)
                })
                .ToList();

            return BuildSimpleBarChart(id, title, items, yAxisTitle, color);
        }

        private static bool IsUnknownCauseChartItem(ErrorAnalysis error)
        {
            var value = error != null ? error.CauseClassification : null;
            return !string.IsNullOrWhiteSpace(value)
                && (value.IndexOf("Không rõ", StringComparison.OrdinalIgnoreCase) >= 0
                    || value.IndexOf("原因不明", StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static List<string> GetColorPalette(int count)
        {
            var palette = new[]
            {
                "#4f46e5", "#0ea5e9", "#10b981", "#f59e0b", "#ef4444", "#8b5cf6",
                "#14b8a6", "#f97316", "#64748b", "#ec4899", "#84cc16", "#3b82f6"
            };

            var colors = new List<string>();
            for (var i = 0; i < Math.Max(count, 1); i++)
            {
                colors.Add(palette[i % palette.Length]);
            }

            return colors;
        }

        private static string GetMonthLabel(DateTime value)
        {
            return value.ToString("MMM", CultureInfo.InvariantCulture);
        }

        private static string NormalizeLabel(string value)
        {
            return string.IsNullOrWhiteSpace(value) ? "Chưa cập nhật" : value;
        }

        private static string NormalizeChartCategoryLabel(string value)
        {
            var normalized = NormalizeLabel(value);
            return string.Equals(normalized, "LM (TD2a, PD7,PD7 DASH, PJ8)/ Domino", StringComparison.OrdinalIgnoreCase)
                ? "LM/Domino"
                : normalized;
        }
        public ActionResult LookupManagement()
        {
            return View(BuildLookupManagementViewModel());
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult CreateLookup(string groupCode, string value, string displayText, int? sortOrder)
        {
            if (!IsSupportedLookupGroup(groupCode))
            {
                return HttpNotFound();
            }

            value = (value ?? string.Empty).Trim();
            displayText = string.IsNullOrWhiteSpace(displayText) ? value : displayText.Trim();

            if (string.IsNullOrWhiteSpace(value))
            {
                TempData["LookupError"] = "Giá trị không được để trống.";
                return RedirectToAction("LookupManagement");
            }

            using (var context = new ErrorAnalysisDbContext())
            {
                var exists = context.ErrorAnalysisLookups.Any(x =>
                    x.GroupCode == groupCode &&
                    x.Value == value);

                if (exists)
                {
                    TempData["LookupError"] = "Dữ liệu đã tồn tại.";
                    return RedirectToAction("LookupManagement");
                }

                var nextSortOrder = sortOrder ?? context.ErrorAnalysisLookups
                    .Where(x => x.GroupCode == groupCode)
                    .Select(x => (int?)x.SortOrder)
                    .Max().GetValueOrDefault() + 1;

                context.ErrorAnalysisLookups.Add(new ErrorAnalysisLookup
                {
                    GroupCode = groupCode,
                    Value = value,
                    DisplayText = displayText,
                    SortOrder = nextSortOrder,
                    IsActive = true
                });

                context.SaveChanges();
            }

            TempData["LookupMessage"] = "Đã thêm danh mục.";
            return RedirectToAction("LookupManagement");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult UpdateLookup(int id, string value, string displayText, int sortOrder, bool isActive)
        {
            value = (value ?? string.Empty).Trim();
            displayText = string.IsNullOrWhiteSpace(displayText) ? value : displayText.Trim();

            if (string.IsNullOrWhiteSpace(value))
            {
                TempData["LookupError"] = "Giá trị không được để trống.";
                return RedirectToAction("LookupManagement");
            }

            using (var context = new ErrorAnalysisDbContext())
            {
                var item = context.ErrorAnalysisLookups.FirstOrDefault(x => x.Id == id);
                if (item == null || !IsSupportedLookupGroup(item.GroupCode))
                {
                    return HttpNotFound();
                }

                var exists = context.ErrorAnalysisLookups.Any(x =>
                    x.Id != id &&
                    x.GroupCode == item.GroupCode &&
                    x.Value == value);

                if (exists)
                {
                    TempData["LookupError"] = "Dữ liệu đã tồn tại.";
                    return RedirectToAction("LookupManagement");
                }

                item.Value = value;
                item.DisplayText = displayText;
                item.SortOrder = sortOrder;
                item.IsActive = isActive;

                context.SaveChanges();
            }

            TempData["LookupMessage"] = "Đã cập nhật danh mục.";
            return RedirectToAction("LookupManagement");
        }

        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteLookup(int id)
        {
            using (var context = new ErrorAnalysisDbContext())
            {
                var item = context.ErrorAnalysisLookups.FirstOrDefault(x => x.Id == id);
                if (item == null || !IsSupportedLookupGroup(item.GroupCode))
                {
                    return HttpNotFound();
                }

                context.ErrorAnalysisLookups.Remove(item);
                context.SaveChanges();
            }

            TempData["LookupMessage"] = "Đã xóa danh mục.";
            return RedirectToAction("LookupManagement");
        }

        private static ErrorAnalysisLookupManagementViewModel BuildLookupManagementViewModel()
        {
            using (var context = new ErrorAnalysisDbContext())
            {
                return new ErrorAnalysisLookupManagementViewModel
                {
                    Models = context.ErrorAnalysisLookups
                        .Where(x => x.GroupCode == ErrorAnalysisLookupGroups.Model)
                        .OrderBy(x => x.SortOrder)
                        .ThenBy(x => x.DisplayText)
                        .ToList(),

                    Suppliers = context.ErrorAnalysisLookups
                        .Where(x => x.GroupCode == ErrorAnalysisLookupGroups.Supplier)
                        .OrderBy(x => x.SortOrder)
                        .ThenBy(x => x.DisplayText)
                        .ToList()
                };
            }
        }

        private static bool IsSupportedLookupGroup(string groupCode)
        {
            return groupCode == ErrorAnalysisLookupGroups.Model
                || groupCode == ErrorAnalysisLookupGroups.Supplier;
        }

        private class ChartItem
        {
            public string Label { get; set; }
            public int Value { get; set; }
        }
    }
}
