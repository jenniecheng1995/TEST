using GISFCU.EIP4.HR.Controllers;
using GISFCU.EIP4.HR.Filters;
using GISFCU.EIP4.HR.Helpers;
using GISFCU.EIP4.HRLibs.Models;
using GISFCU.EIP4.HRLibs.Sources;
using GISFCU.EIP4.HRLibs.ViewModels;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace GISFCU.EIP4.HR.Areas.Assess.Controllers
{
    [Authorize]
    public class ManagerAssessController : BaseController
    {

        public static string _year;
        public static string _stage;
        public static DateTime? _periodSDate;
        public static DateTime? _periodEDate;
        public static AssessPeriod _period;

        public ManagerAssessController()
        {
            _period = AssessSource.GetAssessPeriod();
            _year = _period.PYear;
            _stage = _period.Stage;
            _periodSDate = _period.PeriodRangeFrom;
            _periodEDate = _period.PeriodRangeTo;
        }

        /// <summary>
        /// 主管評核頁面
        /// </summary>
        /// <returns></returns>
        [UserAuthorize]
        public ActionResult Index(bool isReload = false)
        {
            var res = AssessSource.GetAssessResultList(_year, _stage, _user.DepNo);

            ViewBag.DepName = _user.DepName;
            ViewBag.AssessYear = _year;
            ViewBag.AMAssessPeriod = string.Format("{0} ~ {1}",
                                     _period.AMPeriodStart.Value.ToShortDateString(),
                                     _period.AMPeriodEnd.Value.ToShortDateString());
            ViewBag.IsInAssessPeriod = AssessSource.IsInAssessPeriod();
            ViewBag.IsDepNeedAssess = AssessSource.IsDepNeedAssess(_user.DepNo);

            if (isReload)
                return PartialView("_AssessResult", res.data);
            else
                return View(res.data);
        }

        /// <summary>
        /// 編輯評核結果
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Edit(List<AssessResultViewModel> model)
        {
            if (ModelState.IsValid)
            {
                int successNum = 0;
                int errorNum = 0;

                foreach (var item in model)
                {
                    var comments = new AssessCommentsViewModel
                    {
                        EmpNo = item.EmpNo,
                        SelfComment = item.SelfComment,
                        MemberComment = item.MemberComment,
                        ManagerComment = item.ManagerComment
                    };

                    var ranking = new AssessRankingViewModel
                    {
                        EmpNo = item.EmpNo,
                        SelfScore = item.SelfScore,
                        MemberScore = item.MemberScore,
                        ManagerScore = item.ManagerScore
                    };

                    var res = AssessSource.SaveAssessResultList(_year, _stage, _periodSDate, _periodEDate, comments, ranking);

                    if (res.isSuccess)
                        successNum++;
                    else
                        errorNum++;
                }

                if (errorNum > 0)
                    return Json(AjaxResponse.Error(string.Format("失敗{0}筆，成功更新{1}筆", errorNum, successNum)));
                else
                    return Json(AjaxResponse.Success(string.Format("成功更新{0}筆", successNum)));
            }

            return Json(AjaxResponse.ModelStateError());
        }

        /// <summary>
        /// 上傳評核結果
        /// </summary>
        /// <param name="file">檔案</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Upload(HttpPostedFileBase file)
        {
            using (var package = new ExcelPackage(file.InputStream))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                var header = new string[] { "員工名稱", "職稱", "職等", "年資",
                                            "自我評核分數", "自我評核評語",
                                            "互評評核分數", "互評評核評語",
                                            "主管評核分數(A)", "主管評核評語",
                                            "獎懲分數(B)", "總計(A+B)", "等第"};

                var startRow = 1;
                var startCol = 1;
                var endRow = worksheet.Dimension.End.Row;

                for (int col = startCol; col < header.Length; col++)
                {
                    var name = worksheet.Cells[startRow, col].Value ?? string.Empty;

                    if (name.ToString() != header[col - 1]) return Json(AjaxResponse.Error("格式不符"));
                }

                int successNum = 0;
                int errorNum = 0;

                for (int row = 2; row <= endRow; row++)
                {
                    int? selfScore = 0,
                         memberScore = 0,
                         managerScore = 0;

                    string selfComment = "",
                           memberComment = "",
                           managerComment = "";

                    var errMsg = new List<string>();
                    bool isNotInt = false;

                    var resCol = worksheet.Cells[row, header.Length + 1];

                    var empName = ConvObjectToString(worksheet.Cells[row, 1].Value);
                    var empNo = EmployeeSource.GetEmpNo(_user.DepName, empName);

                    if (empNo == null) goto ErrorMsg;

                    selfComment = ConvObjectToString(worksheet.Cells[row, 6].Value);
                    memberComment = ConvObjectToString(worksheet.Cells[row, 8].Value);
                    managerComment = ConvObjectToString(worksheet.Cells[row, 10].Value);

                    if (CheckInt(worksheet.Cells[row, 5].Value))
                        selfScore = ConvObjectToInt(worksheet.Cells[row, 5].Value);
                    else
                        isNotInt = true;

                    if (CheckInt(worksheet.Cells[row, 7].Value))
                        memberScore = ConvObjectToInt(worksheet.Cells[row, 7].Value);
                    else
                        isNotInt = true;

                    if (CheckInt(worksheet.Cells[row, 9].Value))
                        managerScore = ConvObjectToInt(worksheet.Cells[row, 9].Value);
                    else
                        isNotInt = true;

                    if (empNo == null)
                        errMsg.Add("查無此人員");
                    if (isNotInt)
                        errMsg.Add("分數只允許輸入數字");
                    if (selfScore > 100 || memberScore > 100 || managerScore > 100)
                        errMsg.Add("分數不可高於100");
                    if (selfScore < 0 || memberScore < 0 || managerScore < 0)
                        errMsg.Add("分數不可為負值");

                    ErrorMsg:
                    if (errMsg.Count == 0)
                    {
                        var comments = new AssessCommentsViewModel
                        {
                            EmpNo = empNo,
                            SelfComment = selfComment,
                            MemberComment = memberComment,
                            ManagerComment = managerComment
                        };

                        var ranking = new AssessRankingViewModel
                        {
                            EmpNo = empNo,
                            SelfScore = selfScore,
                            MemberScore = memberScore,
                            ManagerScore = managerScore
                        };

                        var res = AssessSource.SaveAssessResultList(_year, _stage, _periodSDate, _periodEDate, comments, ranking);

                        if (res.isSuccess)
                        {
                            resCol.Value = "成功";
                            successNum++;
                        }
                        else
                        {
                            resCol.Value = res.error;
                            errorNum++;
                        }
                    }
                    else
                    {
                        resCol.Value = string.Join("、", errMsg);
                        errorNum++;
                    }
                }

                MemoryStream stream = new MemoryStream();
                package.SaveAs(stream);
                stream.Position = 0;

                TempData["UploadResult"] = stream;
                TempData["FileName"] = file.FileName;

                if (errorNum > 0)
                    return Json(AjaxResponse.Error(string.Format("失敗{0}筆，成功更新{1}筆，請查看上傳結果", errorNum, successNum)));
                else
                    return Json(AjaxResponse.Success(string.Format("成功更新{0}筆", successNum)));
            }
        }

        /// <summary>
        /// 下載匯入結果
        /// </summary>
        /// <returns></returns>
        public ActionResult DownloadUploadResult()
        {
            if (TempData["UploadResult"] == null) return Json(AjaxResponse.Error("無上傳結果"), JsonRequestBehavior.AllowGet);

            var stream = (MemoryStream)TempData["UploadResult"];
            var fileName = (string)TempData["FileName"] ?? "評核";

            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            return File(stream, contentType, fileName.Replace(".xlsx", "") + "_匯入結果.xlsx");
        }

        /// <summary>
        /// 取得處別出勤紀錄
        /// </summary>
        public ActionResult GetDuty()
        {
            var res = AssessSource.GetDutyLog(_periodSDate, _periodEDate, _user.DepNo);
            return Json(res, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 取得處別請假紀錄
        /// </summary>
        public ActionResult GetLeave()
        {
            var res = AssessSource.GetLeaveLog(_periodSDate, _periodEDate, _user.DepNo);
            return Json(res, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 取得處別訓練發展統計
        /// </summary>
        public ActionResult GetTraining()
        {
            var res = AssessSource.GetTrainingStat(_periodSDate, _periodEDate, _user.DepNo);
            return Json(res, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 取得處別獎懲紀錄
        /// </summary>
        public ActionResult GetRewardPunish()
        {
            var res = AssessSource.GetRewardPunish(_periodSDate, _periodEDate, _user.DepNo);
            return Json(res, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 取得處別獎懲紀錄
        /// </summary>
        public ActionResult GetOther()
        {
            var res = AssessSource.GetOther(_periodSDate, _periodEDate, _user.DepNo);
            return Json(res, JsonRequestBehavior.AllowGet);
        }

        /// <summary>
        /// 下載評核名單
        /// </summary>
        /// <returns></returns>
        public ActionResult Download()
        {
            string sheetName = "評核名單";

            string fileName = string.Format("{0}{1}評核名單{2}.xlsx", _year, _user.DepName, DateTime.Now.ToString("yyyyMMddHHmmss"));

            var source = AssessSource.GetAssessList(_year, _stage, _user.DepNo);

            if (!source.isSuccess) return Content(string.Format("<script>alert.error('','{0}')</script>", source.error));

            var partColor = Color.FromArgb(31, 78, 120);

            var ranking = AssessSource.GetRankingList();

            var rankingFormula = new List<string>();

            foreach (var item in ranking)
            {
                rankingFormula.Add("L{row}<" + (item.EndScore + 1) + ",\"" + item.Ranking + "\"");
            }

            string rankingRes = "IFS(" + string.Join(",", rankingFormula) + ")";

            var selector = new List<EpplusSelector>
            {
                new EpplusSelector { Field = "EmpName", Name = "員工名稱" },
                new EpplusSelector { Field = "JobCode", Name = "職稱" },
                new EpplusSelector { Field = "JobLevel", Name = "職等" },
                new EpplusSelector { Field = "WorkingYears", Name = "年資" },
                new EpplusSelector { Field = "SelfScore", Name = "自我評核分數" },
                new EpplusSelector { Field = "SelfComment", Name = "自我評核評語" },
                new EpplusSelector { Field = "MemberScore", Name = "互評評核分數" },
                new EpplusSelector { Field = "MemberComment", Name = "互評評核評語" },
                new EpplusSelector { Field = "ManagerScore", Name = "主管評核分數(A)" },
                new EpplusSelector { Field = "ManagerComment", Name = "主管評核評語" },
                new EpplusSelector { Field = "RewardPunishScore", Name = "獎懲分數(B)", HeaderColor = partColor},
                new EpplusSelector { Field = "", Name = "總計(A+B)", Formula = "SUM(I{row}+K{row})", HeaderColor = partColor},
                new EpplusSelector { Field = "", Name = "等第", Formula = rankingRes, HeaderColor = partColor},
                new EpplusSelector { Field = "", Name = "匯入結果" },
            };

            var stream = EpplusHelper.CreateSelExcel(sheetName, source.data, selector);

            string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

            return File(stream, contentType, fileName);
        }

        /// <summary>
        /// 轉換為字串
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private string ConvObjectToString(object value)
        {
            return value == null ? string.Empty : value.ToString();
        }

        /// <summary>
        /// 轉換為字串
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private bool CheckInt(object value)
        {
            int outValue;
            if (value == null || (value != null && int.TryParse(value.ToString(), out outValue)))
                return true;
            else
                return false;
        }

        /// <summary>
        /// 轉換為字串
        /// </summary>
        /// <param name="value"></param>
        /// <returns></returns>
        private int? ConvObjectToInt(object value)
        {
            return value == null ? null : (int?)int.Parse(value.ToString());
        }
    }
}