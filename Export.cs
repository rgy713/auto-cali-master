using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Drawing;

namespace comTest
{
    class Export
    {
        public float[][][] m_Measurements;
        public float[][][] m_Errors;
        public string[] m_BodyStr;
        public string[] m_BodyAngleStr;
        public string[] m_RateStr;
        public float[] m_TestWeightLevels;
        public float[][] m_TestWeightResults;
        public string[] m_WeightPositionLevels;
        public string m_MachineType;
        public string m_SerialNumber;
        public string m_ProductSerialNumber;
        public float[][][] m_PhaseAngles;
        public float[][] m_OriginValues;
        public void export()
        {
            //Create a new ExcelPackage
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                try
                {
                    DateTime nowDate = DateTime.Now;

                    //Set some properties of the Excel document
                    excelPackage.Workbook.Properties.Author = "AINST";
                    excelPackage.Workbook.Properties.Title = "成品检验报告-CNDS-0318";
                    excelPackage.Workbook.Properties.Subject = "AINST";
                    excelPackage.Workbook.Properties.Created = nowDate;

                    //Create the WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("报告单");

                    double ra = 1.2;
                    worksheet.Column(1).Width = 10.5 * ra; //A
                    worksheet.Column(2).Width = 5.63 * ra; //B
                    worksheet.Column(3).Width = 4.75 * ra; //C
                    worksheet.Column(4).Width = 5.88 * ra; //D
                    worksheet.Column(5).Width = 10.13 * ra; //E
                    worksheet.Column(6).Width = 6.5 * ra; //F
                    worksheet.Column(7).Width = 6.5 * ra; //G
                    worksheet.Column(8).Width = 6.5 * ra; //H
                    worksheet.Column(9).Width = 6.5 * ra; //I
                    worksheet.Column(10).Width = 6.5 * ra; //J
                    worksheet.Column(11).Width = 6.5 * ra; //K
                    worksheet.Column(12).Width = 6.5 * ra; //L
                    worksheet.Column(13).Width = 13.5 * ra; //M

                    worksheet.Cells["A1:M73"].Style.Font.Size = 10;
                    worksheet.Cells["A1:M73"].Style.Font.Name = "宋体";
                    worksheet.Cells["A1:M73"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A18:J18"].Style.Border.Top.Color.SetColor(Color.Black);
                    worksheet.Cells["A1:M73"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A18:J18"].Style.Border.Right.Color.SetColor(Color.Black);
                    worksheet.Cells["A1:M73"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A18:J18"].Style.Border.Bottom.Color.SetColor(Color.Black);
                    worksheet.Cells["A1:M73"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A18:J18"].Style.Border.Left.Color.SetColor(Color.Black);

                    worksheet.Row(1).Height = 15;
                    worksheet.Cells["A1:M1"].Merge = true;
                    worksheet.Cells["A1"].Style.Font.Size = 10.5f;
                    worksheet.Cells["A1"].Value = "编号：CX0901-R01                                                                      版本：V1.0";

                    worksheet.Row(2).Height = 21;
                    worksheet.Cells["A2:M2"].Merge = true;
                    worksheet.Cells["A2"].Style.Font.Size = 16;
                    worksheet.Cells["A2"].Style.Font.Bold = true;
                    worksheet.Cells["A2"].Value = "江苏康爱营养科技有限责任公司";
                    worksheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Row(3).Height = 21;
                    worksheet.Cells["A3:M3"].Merge = true;
                    worksheet.Cells["A3"].Style.Font.Size = 16;
                    worksheet.Cells["A3"].Style.Font.Bold = true;
                    worksheet.Cells["A3"].Value = "成品检验报告单";
                    worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Row(4).Height = 17.25;
                    worksheet.Cells["A4:M4"].Merge = true;
                    worksheet.Cells["A4"].Style.Font.Size = 10.5f;
                    worksheet.Cells["A4"].Value = string.Format("编号：{0}", m_SerialNumber);
                    worksheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Row(5).Height = 29.5;
                    worksheet.Cells["A5"].Value = "产品名称";
                    worksheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B5:D5"].Merge = true;
                    worksheet.Cells["B5"].Value = "临床营养检测分析仪";
                    worksheet.Cells["B5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E5"].Value = "规格型号";
                    worksheet.Cells["E5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["F5:H5"].Merge = true;
                    worksheet.Cells["F5"].Value = m_MachineType;
                    worksheet.Cells["F5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["I5:K5"].Merge = true;
                    worksheet.Cells["I5"].Value = "产品编号";
                    worksheet.Cells["I5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["L5:M5"].Merge = true;
                    worksheet.Cells["L5"].Value = m_ProductSerialNumber;
                    worksheet.Cells["L5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(6).Height = 29.5;
                    worksheet.Cells["A6"].Value = "生产日期";
                    worksheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B6:D6"].Merge = true;
                    worksheet.Cells["B6"].Value = nowDate;
                    worksheet.Cells["B6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B6"].Style.Numberformat.Format = "yyyy.MM.dd";

                    worksheet.Cells["E6"].Value = "使用期限";
                    worksheet.Cells["E6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["F6:H6"].Merge = true;
                    worksheet.Cells["F6"].Value = "十年";
                    worksheet.Cells["F6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["I6:K6"].Merge = true;
                    worksheet.Cells["I6"].Value = "检验数量";
                    worksheet.Cells["I6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["L6:M6"].Merge = true;
                    worksheet.Cells["L6"].Value = "1台";
                    worksheet.Cells["L6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(7).Height = 47.75;

                    worksheet.Cells["A7"].Value = "送检日期";
                    worksheet.Cells["A7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B7:D7"].Merge = true;
                    worksheet.Cells["B7"].Value = nowDate;
                    worksheet.Cells["B7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B7"].Style.Numberformat.Format = "yyyy.MM.dd";

                    worksheet.Cells["E7"].Value = "检验依据";
                    worksheet.Cells["E7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["F7:M7"].Merge = true;
                    worksheet.Cells["F7"].Value = "1.《临床营养检测分析仪AiNST-CNDS成品质量标准及检验规程》ZL-JY(CNDS）-02\n2.《老化测试作业指导书》ZL - JY（CNDS & KNDS）-01";
                    worksheet.Cells["F7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F7"].Style.WrapText = true;

                    worksheet.Row(8).Height = 39.5;
                    worksheet.Row(9).Height = 39;

                    worksheet.Cells["A8:A9"].Merge = true;
                    worksheet.Cells["A8"].Value = "检验设备及编号";
                    worksheet.Cells["A8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A8"].Style.WrapText = true;

                    worksheet.Cells["B8:D8"].Merge = true;
                    worksheet.Cells["B8"].Value = "示波器（ZL005）";
                    worksheet.Cells["B8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E8:G8"].Merge = true;
                    worksheet.Cells["E8"].Value = "阻抗采集工装1（ZLGZ001）\n阻抗采集工装2（ZLGZ002）";
                    worksheet.Cells["E8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E8"].Style.WrapText = true;

                    worksheet.Cells["H8:J8"].Merge = true;
                    worksheet.Cells["H8"].Value = "数字万用表（ZL001）";
                    worksheet.Cells["H8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["K8:M8"].Merge = true;
                    worksheet.Cells["K8"].Value = "砝码（ZL006）";
                    worksheet.Cells["K8"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K8"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B9:D9"].Merge = true;
                    worksheet.Cells["B9"].Value = "医用漏电流测试仪（ZL003）";
                    worksheet.Cells["B9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B9"].Style.WrapText = true;

                    worksheet.Cells["E9:G9"].Merge = true;
                    worksheet.Cells["E9"].Value = "医用耐压测试仪（ZL004）";
                    worksheet.Cells["E9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E9"].Style.WrapText = true;

                    worksheet.Cells["H9:J9"].Merge = true;
                    worksheet.Cells["H9"].Value = "医用接地电阻测试仪（ZL002）";
                    worksheet.Cells["H9"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H9"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H9"].Style.WrapText = true;

                    worksheet.Cells["K9:M9"].Merge = true;

                    worksheet.Row(10).Height = 39;

                    worksheet.Cells["A10"].Value = "检验项目";
                    worksheet.Cells["A10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A10"].Style.Font.Bold = true;

                    worksheet.Cells["B10:F10"].Merge = true;
                    worksheet.Cells["B10"].Value = "具体内容";
                    worksheet.Cells["B10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B10"].Style.Font.Bold = true;

                    worksheet.Cells["G10:L10"].Merge = true;
                    worksheet.Cells["G10"].Value = "具体内容";
                    worksheet.Cells["G10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G10"].Style.Font.Bold = true;

                    worksheet.Cells["M10"].Value = "单项结论";
                    worksheet.Cells["M10"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M10"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["M10"].Style.Font.Bold = true;

                    worksheet.Row(11).Height = 121;

                    worksheet.Cells["A11"].Value = "2.1外观和一般性能";
                    worksheet.Cells["A11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A11"].Style.WrapText = true;

                    worksheet.Cells["B11:F11"].Merge = true;
                    worksheet.Cells["B11"].Value = "a)仪器外观应整齐、色泽均匀、无伤痕、划痕等缺陷；\nb)仪器上文字和图示标志应清晰可见；\nc)仪器控制机构应灵活可靠、紧固部位无松动；\nd)仪器的塑料件应无起泡、开裂、变形、飞边现象；\ne)电极表面应平整、无毛刺；\nf)液晶显示器显示应清晰。";
                    worksheet.Cells["B11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B11"].Style.WrapText = true;

                    worksheet.Cells["G11:L11"].Merge = true;
                    worksheet.Cells["G11"].Value = "仪器外观整齐、无伤痕划痕；\n仪器上文字、图标清晰；\n仪器控制机构灵活、紧固部位无松动；\n仪器的塑料件无起泡、开裂、变形、飞边；\n电极表面平整、无毛刺；\n液晶显示器晰";
                    worksheet.Cells["G11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G11"].Style.WrapText = true;

                    worksheet.Cells["M11"].Value = "√合格 □不合格";
                    worksheet.Cells["M11"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M11"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(12).Height = 26.25;
                    worksheet.Row(13).Height = 26.25;

                    for (int i = 14; i <= 31; i++)
                    {
                        worksheet.Row(i).Height = 30;
                    }

                    worksheet.Cells["A12:A29"].Merge = true;
                    worksheet.Cells["A12"].Value = "2.3阻抗测量";
                    worksheet.Cells["A12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A12"].Style.WrapText = true;

                    worksheet.Cells["B12:C29"].Merge = true;
                    worksheet.Cells["B12"].Value = "阻抗测量功能：上肢阻抗测量范围：200Ω～690Ω，测量允差±5%；躯干阻抗测量范围：10Ω～35Ω，测量允差±10%；下肢阻抗测量范围：160Ω～500Ω，测量允差±5%。";
                    worksheet.Cells["B12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B12"].Style.WrapText = true;

                    worksheet.Cells["D12:F13"].Merge = true;
                    worksheet.Cells["D12"].Value = "电阻值";
                    worksheet.Cells["D12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["G12:H12"].Merge = true;
                    worksheet.Cells["G12"].Value = string.Format("上肢\n（{0}、{1}）", m_BodyStr[1], m_BodyStr[4]);
                    worksheet.Cells["G12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G12"].Style.WrapText = true;

                    worksheet.Cells["G13"].Value = "最大值";
                    worksheet.Cells["G13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["H13"].Value = "最小值";
                    worksheet.Cells["H13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["I12:J12"].Merge = true;
                    worksheet.Cells["I12"].Value = string.Format("下肢\n（{0}、{1}）", m_BodyStr[0], m_BodyStr[3]);
                    worksheet.Cells["I12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I12"].Style.WrapText = true;

                    worksheet.Cells["I13"].Value = "最大值";
                    worksheet.Cells["I13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["J13"].Value = "最小值";
                    worksheet.Cells["J13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["K12:L12"].Merge = true;
                    worksheet.Cells["K12"].Value = string.Format("躯干\n（{0}）", m_BodyStr[2]);
                    worksheet.Cells["K12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["K12"].Style.WrapText = true;

                    worksheet.Cells["K13"].Value = "最大值";
                    worksheet.Cells["K13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["L13"].Value = "最小值";
                    worksheet.Cells["L13"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L13"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M12:M13"].Merge = true;
                    worksheet.Cells["M12"].Value = "/";
                    worksheet.Cells["M12"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M12"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    string[] unitDescs = {
                        "阻抗校准工装测量（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）",
                        "阻抗校准工装测量（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）",
                        "阻抗校准工装测量（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）",
                        "阻抗校准工装测量（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）",
                        "阻抗校准工装测量（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）",
                        "阻抗校准工装测量（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）",
                        "阻抗采集工装（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）",
                        "阻抗采集工装（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）",
                        "阻抗采集工装（{0}={1}Ω、{2}={3}Ω、{4}={5}Ω、{6}={7}Ω、{8}={9}Ω）"
                    };
                    string[] valueFields = { "G", "H", "I", "J", "K", "L" };

                    int unitRow = 14;
                    for (int unitIndex = 0; unitIndex < 8; unitIndex++)
                    {
                        int rowIndex = unitRow + 2 * unitIndex;
                        int rowIndex1 = rowIndex + 1;

                        worksheet.Cells[string.Format("D{0}:E{1}", rowIndex, rowIndex1)].Merge = true;
                        string d0 = string.Format("D{0}", rowIndex);
                        float[] org = m_OriginValues[unitIndex];
                        worksheet.Cells[d0].Value = string.Format(unitDescs[unitIndex], m_BodyStr[1], org[1], m_BodyStr[4], org[4], m_BodyStr[0], org[0], m_BodyStr[3], org[3], m_BodyStr[2], org[2]);
                        worksheet.Cells[d0].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[d0].Style.WrapText = true;

                        string f0 = string.Format("F{0}", rowIndex);
                        worksheet.Cells[f0].Value = "测量值";
                        worksheet.Cells[f0].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[f0].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        string f1 = string.Format("F{0}", rowIndex1);
                        worksheet.Cells[f1].Value = "误差值";
                        worksheet.Cells[f1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[f1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        if (m_Measurements != null && unitIndex < m_Measurements.Length)
                        {
                            (double max, double min)[] measureValue = new (double max, double min)[3];

                            float[][] measureValues = m_Measurements[unitIndex];

                            float[] ra_values = measureValues[1];
                            float[] la_values = measureValues[4];
                            float[] tt_values = measureValues[2];
                            float[] rl_values = measureValues[0];
                            float[] ll_values = measureValues[3];

                            float[] ra_la = new float[ra_values.Length + la_values.Length];
                            ra_values.CopyTo(ra_la, 0);
                            la_values.CopyTo(ra_la, ra_values.Length);

                            float[] rl_ll = new float[rl_values.Length + ll_values.Length];
                            rl_values.CopyTo(rl_ll, 0);
                            ll_values.CopyTo(rl_ll, rl_values.Length);

                            measureValue[0] = (ra_la.Max(), ra_la.Min());
                            measureValue[1] = (rl_ll.Max(), rl_ll.Min());
                            measureValue[2] = (tt_values.Max(), tt_values.Min());

                            for (int i = 0; i < measureValue.Length; i++)
                            {
                                (double max, double min) maxmin = measureValue[i];

                                if (maxmin.max >= 0)
                                {
                                    string maxField = string.Format("{0}{1}", valueFields[i * 2], rowIndex);
                                    worksheet.Cells[maxField].Value = maxmin.max;
                                    worksheet.Cells[maxField].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[maxField].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    worksheet.Cells[maxField].Style.Numberformat.Format = "0.00";
                                }
                                
                                if (maxmin.min >= 0)
                                {
                                    string minField = string.Format("{0}{1}", valueFields[i * 2 + 1], rowIndex);
                                    worksheet.Cells[minField].Value = maxmin.min;
                                    worksheet.Cells[minField].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                    worksheet.Cells[minField].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    worksheet.Cells[minField].Style.Numberformat.Format = "0.00";
                                }
                            }
                        }

                        if (m_Errors != null && unitIndex < m_Errors.Length)
                        {
                            (double max, double min)[] errorValue = new (double max, double min)[3];

                            float[][] errorValues = m_Errors[unitIndex];

                            float[] ra_values = errorValues[1];
                            float[] la_values = errorValues[4];
                            float[] tt_values = errorValues[2];
                            float[] rl_values = errorValues[0];
                            float[] ll_values = errorValues[3];

                            float[] ra_la = new float[ra_values.Length + la_values.Length];
                            ra_values.CopyTo(ra_la, 0);
                            la_values.CopyTo(ra_la, ra_values.Length);

                            float[] rl_ll = new float[rl_values.Length + ll_values.Length];
                            rl_values.CopyTo(rl_ll, 0);
                            ll_values.CopyTo(rl_ll, rl_values.Length);

                            errorValue[0] = (ra_la.Max(), ra_la.Min());
                            errorValue[1] = (rl_ll.Max(), rl_ll.Min());
                            errorValue[2] = (tt_values.Max(), tt_values.Min());

                            for (int i = 0; i < errorValue.Length; i++)
                            {
                                (double max, double min) maxmin = errorValue[i];

                                string maxField = string.Format("{0}{1}", valueFields[i * 2], rowIndex1);
                                worksheet.Cells[maxField].Value = maxmin.max == -1 ? "" : (maxmin.max == 0 ? "0" : string.Format("{0}{1}%", maxmin.max > 0 ? "+" : "", Math.Round(maxmin.max, 2)));
                                worksheet.Cells[maxField].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[maxField].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                string minField = string.Format("{0}{1}", valueFields[i * 2 + 1], rowIndex1);
                                worksheet.Cells[minField].Value = maxmin.min == -1 ? "" : (maxmin.min == 0 ? "0" : string.Format("{0}{1}%", maxmin.min > 0 ? "+" : "", Math.Round(maxmin.min, 2)));
                                worksheet.Cells[minField].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[minField].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }

                            worksheet.Cells[string.Format("M{0}:M{1}", rowIndex, rowIndex1)].Merge = true;
                            string m0 = string.Format("M{0}", rowIndex);
                            worksheet.Cells[m0].Value = "□合格 □不合格";
                            worksheet.Cells[m0].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[m0].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                    }

                    worksheet.Row(30).Height = 30;
                    worksheet.Row(31).Height = 30;
                    worksheet.Row(32).Height = 30;
                    worksheet.Row(33).Height = 30;
                    worksheet.Row(34).Height = 30;
                    worksheet.Row(35).Height = 30;
                    worksheet.Row(36).Height = 30;

                    worksheet.Cells["A30:A36"].Merge = true;
                    worksheet.Cells["A30"].Value = "2.4相位角测量";
                    worksheet.Cells["A30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A30"].Style.WrapText = true;

                    worksheet.Cells["B30:C36"].Merge = true;
                    worksheet.Cells["B30"].Value = "±0.2°";
                    worksheet.Cells["B30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B30"].Style.WrapText = true;

                    worksheet.Cells["D30:E30"].Merge = true;
                    worksheet.Cells["D30"].Value = "相位角值";
                    worksheet.Cells["D30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["D31:D32"].Merge = true;
                    worksheet.Cells["D31"].Value = "50Hz";
                    worksheet.Cells["D31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D31"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D31"].Style.WrapText = true;

                    worksheet.Cells["E31"].Value = "测量值";
                    worksheet.Cells["E31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E31"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E32"].Value = "误差值";
                    worksheet.Cells["E32"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E32"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["D33:D34"].Merge = true;
                    worksheet.Cells["D33"].Value = "250Hz";
                    worksheet.Cells["D33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D33"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D33"].Style.WrapText = true;

                    worksheet.Cells["E33"].Value = "测量值";
                    worksheet.Cells["E33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E33"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E34"].Value = "误差值";
                    worksheet.Cells["E34"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E34"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["D35:D36"].Merge = true;
                    worksheet.Cells["D35"].Value = "500Hz";
                    worksheet.Cells["D35"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D35"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D35"].Style.WrapText = true;

                    worksheet.Cells["E35"].Value = "测量值";
                    worksheet.Cells["E35"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E35"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E36"].Value = "误差值";
                    worksheet.Cells["E36"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E36"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    string[] valueFields1 = { "F", "G", "H", "I", "J", "K", "L" };

                    for (int i = 0; i < 7; i++)
                    {
                        string field = string.Format("{0}30", valueFields1[i]);
                        worksheet.Cells[field].Value = m_BodyAngleStr[i];
                        worksheet.Cells[field].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[field].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    }

                    worksheet.Cells["M30"].Value = "/";
                    worksheet.Cells["M30"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M31:M32"].Merge = true;
                    worksheet.Cells["M31"].Value = "□合格 □不合格";
                    worksheet.Cells["M31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M31"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M33:M34"].Merge = true;
                    worksheet.Cells["M33"].Value = "□合格 □不合格";
                    worksheet.Cells["M33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M33"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M35:M36"].Merge = true;
                    worksheet.Cells["M35"].Value = "□合格 □不合格";
                    worksheet.Cells["M35"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M35"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    for (int i = 0; i < 3; i++)
                    {
                        for (int j = 0; j < 7; j++)
                        {
                            if(m_PhaseAngles[0][j][i] > -1)
                            {
                                string field = string.Format("{0}{1}", valueFields1[j], 31 + i * 2);
                                worksheet.Cells[field].Value = m_PhaseAngles[3][j][i];
                                worksheet.Cells[field].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                                worksheet.Cells[field].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                worksheet.Cells[field].Style.Numberformat.Format = "0.00";
                            }                            
                        }
                    }

                    worksheet.Row(37).Height = 29.25;

                    worksheet.Cells["A37"].Value = "2.5输出电流";
                    worksheet.Cells["A37"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A37"].Style.WrapText = true;

                    worksheet.Cells["B37:F37"].Merge = true;
                    worksheet.Cells["B37"].Value = "最大输出电流≤450μA。";
                    worksheet.Cells["B37"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["G37:L37"].Merge = true;
                    worksheet.Cells["G37"].Value = "最大测量值为320μA";
                    worksheet.Cells["M37"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G37"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M37"].Value = "□合格 □不合格";
                    worksheet.Cells["M37"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M37"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(38).Height = 30;
                    worksheet.Row(39).Height = 45;
                    worksheet.Row(40).Height = 45;
                    worksheet.Row(41).Height = 30;
                    worksheet.Row(42).Height = 30;

                    worksheet.Cells["A38:A42"].Merge = true;
                    worksheet.Cells["A38"].Value = "2.6体重测量功能";
                    worksheet.Cells["A38"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A38"].Style.WrapText = true;

                    worksheet.Cells["B38:E40"].Merge = true;
                    worksheet.Cells["B38"].Value = "体重测量功能\na)在体重测量范围10kg～120kg内，体重测量允差±0.5kg；\nb)同一载荷在不同位置的示值，其示值允差应不大于±0.5kg；\nc)载荷超过额定量程时应显示“NaN”，并有文字提示“体重超出检测范围”；";
                    worksheet.Cells["B38"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B38"].Style.WrapText = true;

                    worksheet.Cells["F38"].Value = "测量结果";
                    worksheet.Cells["F38"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F38"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F38"].Style.WrapText = true;

                    Boolean isMerge = false;
                    for (int i = 0; i < m_TestWeightLevels.Length; i++)
                    {
                        string fieldName = valueFields[isMerge ? i + 1 : i];

                        int weightLevel = (int)m_TestWeightLevels[i];

                        if(weightLevel == 15)
                        {
                            worksheet.Cells[string.Format("{0}38:{1}38", valueFields[i], valueFields[i + 1])].Merge = true;
                            worksheet.Cells[string.Format("{0}39:{1}39", valueFields[i], valueFields[i + 1])].Merge = true;
                            worksheet.Cells[string.Format("{0}40:{1}40", valueFields[i], valueFields[i + 1])].Merge = true;
                            isMerge = true;
                        }

                        string field = string.Format("{0}38", fieldName);
                        worksheet.Cells[field].Value = string.Format("{0}kg", weightLevel);
                        worksheet.Cells[field].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[field].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                        string vfield = string.Format("{0}39", fieldName);
                        
                        worksheet.Cells[vfield].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[vfield].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[vfield].Style.WrapText = true;

                        string efield = string.Format("{0}40", fieldName);
                        worksheet.Cells[efield].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[efield].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[efield].Style.WrapText = true;

                        float[] testWeights = m_TestWeightResults[i];

                        List<string> weightValues = new List<string>();
                        List<string> weightErrorValues = new List<string>();
                        int testedIndex = 0;

                        for (int j = 0; j < testWeights.Length; j++)
                        {
                            float testWeight = testWeights[j];
                            string weightPos = m_WeightPositionLevels[j];

                            if (testWeight >= 0)
                            {
                                testedIndex = j;
                                weightValues.Add(string.Format("{0} {1}", weightPos, Math.Round(testWeight, 2)));
                                weightErrorValues.Add(string.Format("{0} {1}", weightPos, Math.Round(testWeight - weightLevel, 2)));
                            }
                        }

                        
                        {
                            
                        }
                        if (weightLevel == 15)
                        {
                            worksheet.Cells[vfield].Value = string.Join(",", weightValues);
                            worksheet.Cells[efield].Value = string.Join(",", weightErrorValues);
                        }
                        else if (weightValues.Count == 1 && testedIndex == 0)
                        {
                            float testWeight = testWeights[0];

                            if (testWeight >= 0)
                            {
                                worksheet.Cells[vfield].Value = Math.Round(testWeight, 2);
                                worksheet.Cells[efield].Value = Math.Round(testWeight - weightLevel, 2);
                            }
                        }
                    }
                    worksheet.Cells["L38"].Value = "125kg";
                    worksheet.Cells["L38"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L38"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M38"].Value = "/";
                    worksheet.Cells["M38"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M38"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["F39"].Value = "测量值";
                    worksheet.Cells["F39"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F39"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["F40"].Value = "误差值";
                    worksheet.Cells["F40"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F40"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["L39:L40"].Merge = true;
                    worksheet.Cells["L39"].Value = "显示“NaN”,并有文字提示“体重超出检测范围”";
                    worksheet.Cells["L39"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["L39"].Style.WrapText = true;

                    worksheet.Cells["M39:M40"].Merge = true;
                    worksheet.Cells["M39"].Value = "□合格 □不合格";
                    worksheet.Cells["M39"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M39"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B41:E42"].Merge = true;
                    worksheet.Cells["B41"].Value = "d)载荷超过额定量程的125%，1分钟后卸载，秤体表面无损坏，性能符合a)的要求。";
                    worksheet.Cells["B41"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B41"].Style.WrapText = true;

                    worksheet.Cells["F41"].Value = "测量值";
                    worksheet.Cells["F41"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F41"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["F42"].Value = "误差值";
                    worksheet.Cells["F42"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F42"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["J41:K42"].Merge = true;
                    worksheet.Cells["J41"].Value = "/";
                    worksheet.Cells["J41"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J41"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["J41"].Style.WrapText = true;

                    worksheet.Cells["L41:L42"].Merge = true;
                    worksheet.Cells["L41"].Value = "1分钟后卸载，秤体表面无损坏。";
                    worksheet.Cells["L41"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["L41"].Style.WrapText = true;

                    worksheet.Cells["M41:M42"].Merge = true;
                    worksheet.Cells["M41"].Value = "□合格 □不合格";
                    worksheet.Cells["M41"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M41"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(43).Height = 168;

                    worksheet.Cells["A43"].Value = "2.7.1 软件主要功能—数据采集功能";
                    worksheet.Cells["A43"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A43"].Style.WrapText = true;

                    worksheet.Cells["B43:F43"].Merge = true;
                    worksheet.Cells["B43"].Value = "a)具有新建患者信息，已建患者信息查询功能；\nb)人体成分分析：具有理想体重范围、体重、细胞内液、细胞外液、蛋白质、无机盐、肌肉量、体脂肪、总体水、去脂体重、体脂百分比、基础代谢率、体质指数(BMI)、骨骼肌、肌肉分析指标；\nc)实验室数据：可录入实验室检查结果；\nd)NRS - 2002：可录入营养风险筛查检查结果；\ne)PG - SGA：可录入肿瘤患者主观整体营养评估检查结果；\nf)膳食调查：可录入简明膳食调查法评分结果；\ng)体能消耗调查：可录入体能消耗调查结果；\nh)干预疗法：可录入营养处方用量和次数。";
                    worksheet.Cells["B43"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B43"].Style.WrapText = true;

                    worksheet.Cells["G43:L43"].Merge = true;
                    worksheet.Cells["G43"].Value = "a)患者新建功能正常、可以正常查询\nb)人体成分分析可以正常检测\nc)实验室数据可正常录入检查结果\nd)NRS - 2002可正常录入营养风险筛查结果\ne)PG - SGA可正常录入肿瘤患者主观整体营养评估检查结果\nf)膳食调查录入功能正常\ng)体能消耗调查录入功能正常\nh)干预疗法录入功能正常";
                    worksheet.Cells["G43"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G43"].Style.WrapText = true;

                    worksheet.Cells["M43"].Value = "□合格 □不合格";
                    worksheet.Cells["M43"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M43"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(44).Height = 39.5;

                    worksheet.Cells["A44"].Value = "2.7.2 档案管理功能";
                    worksheet.Cells["A44"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A44"].Style.WrapText = true;

                    worksheet.Cells["B44:F44"].Merge = true;
                    worksheet.Cells["B44"].Value = "记录所有测试过的用户的数据。提供列表显示，并且可以查看任意一次历史测试的结果。";
                    worksheet.Cells["B44"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B44"].Style.WrapText = true;

                    worksheet.Cells["G44:L44"].Merge = true;
                    worksheet.Cells["G44"].Value = "列表显示正常，点击任意一条可查看历史信息，无报错";
                    worksheet.Cells["G44"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G44"].Style.WrapText = true;

                    worksheet.Cells["M44"].Value = "□合格 □不合格";
                    worksheet.Cells["M44"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M44"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(45).Height = 35.25;

                    worksheet.Cells["A45"].Value = "2.7.3 信息统计功能";
                    worksheet.Cells["A45"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A45"].Style.WrapText = true;

                    worksheet.Cells["B45:F45"].Merge = true;
                    worksheet.Cells["B45"].Value = "可以根据不同条件做出筛选，并以图形形式展示统计结果。";
                    worksheet.Cells["B45"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B45"].Style.WrapText = true;

                    worksheet.Cells["G45:L45"].Merge = true;
                    worksheet.Cells["G45"].Value = "信息统计模块，信息、图表正常显示，无报错";
                    worksheet.Cells["G45"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G45"].Style.WrapText = true;

                    worksheet.Cells["M45"].Value = "□合格 □不合格";
                    worksheet.Cells["M45"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M45"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(46).Height = 222.5;

                    worksheet.Cells["A46"].Value = "2.7.4软件主要功能（系统设置功能）";
                    worksheet.Cells["A46"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A46"].Style.WrapText = true;

                    worksheet.Cells["B46:F46"].Merge = true;
                    worksheet.Cells["B46"].Value = "a)基本信息：显示系统名称、系统名称（英文）、系统标题、版本、产品型号、版权；\nb)参数设置：可以对医院名称、NRS - 2002报表标题、PG - SGA报表标题、体重校准、时间校准、实验室数据、科室床号进行设置；\nc)权限管理：具有新建权限信息、编辑权限信息、删除权限信息，其中管理员权限只有编辑查看功能；\nd)账号管理：具有新建账号信息、编辑账号信息、删除账号信息，其中分析仪管理员账号只有编辑查看功能；\ne)营养补充剂(肠内)：具有全营养、特定全营养、非全营养新建信息，编辑信息，删除信息功能，已建信息查询功能；\nf)营养补充剂(肠外)：具有添加肠外药品信息、删除肠外药品信息功能，已增肠外药品查询功能。";
                    worksheet.Cells["B46"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B46"].Style.WrapText = true;

                    worksheet.Cells["G46:L46"].Merge = true;
                    worksheet.Cells["G46"].Value = "a)基本信息页面功能正常\nb)参数设置页面功能正常\nc)权限管理页面功能正常\nd)账号管理页面功能正常\nd)营养补充剂(肠内)页面功能正常\nf)营养补充剂(肠外)页面功能正常";
                    worksheet.Cells["G46"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G46"].Style.WrapText = true;

                    worksheet.Cells["M46"].Value = "□合格 □不合格";
                    worksheet.Cells["M46"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M46"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(47).Height = 81.75;

                    worksheet.Cells["A47"].Value = "2.8.1网络安全功能—数据接口";
                    worksheet.Cells["A47"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A47"].Style.WrapText = true;

                    worksheet.Cells["B47:F47"].Merge = true;
                    worksheet.Cells["B47"].Value = "a)USB接口：可用于系统信息统计数据保存至存储媒介；\nb)网络接口：可连接局域网，数据使用http协议和tcp / ip协议；\nc)存储格式：用户数据以.frm和.ibd 两种格式储存在该设备中。";
                    worksheet.Cells["B47"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B47"].Style.WrapText = true;

                    worksheet.Cells["G47:L47"].Merge = true;
                    worksheet.Cells["G47"].Value = "a)USB接口正常，插入U盘可以正常拷贝文件\nb)网络接口正常，输入项目地址 127.0.0.1 / cnds100 可以正常打开系统\nc)存储方式正常，mysql安装目录查看文件格式一致";
                    worksheet.Cells["G47"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G47"].Style.WrapText = true;

                    worksheet.Cells["M47"].Value = "□合格 □不合格";
                    worksheet.Cells["M47"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M47"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(48).Height = 66.5;

                    worksheet.Cells["A48"].Value = "2.8.2用户访问限制";
                    worksheet.Cells["A48"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A48"].Style.WrapText = true;

                    worksheet.Cells["B48:F48"].Merge = true;
                    worksheet.Cells["B48"].Value = "a)用户应通过登录窗口，经ID号、密码验证正确后，方可进入系统；\nb)不同类型的用户（分析仪管理员、分析仪操作者），登录后的界面、管理的内容及权限不同。";
                    worksheet.Cells["B48"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B48"].Style.WrapText = true;

                    worksheet.Cells["G48:L48"].Merge = true;
                    worksheet.Cells["G48"].Value = "a)正确用户名密码，可正常登录系统，错误不能登录\nb)不同用户权限不一致，登录后看到的页面不同\n111用户登录，所有权限\n001用户登录，只有数据采集权限";
                    worksheet.Cells["G48"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G48"].Style.WrapText = true;

                    worksheet.Cells["M48"].Value = "□合格 □不合格";
                    worksheet.Cells["M48"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M48"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(49).Height = 33;

                    worksheet.Cells["A49"].Value = "2.8.3数据保密性";
                    worksheet.Cells["A49"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A49"].Style.WrapText = true;

                    worksheet.Cells["B49:F49"].Merge = true;
                    worksheet.Cells["B49"].Value = "只有在ID号、密码验证通过后，其测试数据，才具有可得性";
                    worksheet.Cells["B49"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B49"].Style.WrapText = true;

                    worksheet.Cells["G49:L49"].Merge = true;
                    worksheet.Cells["G49"].Value = "在未登录、登录状态下，数据保密性功能符合要求";
                    worksheet.Cells["G49"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G49"].Style.WrapText = true;

                    worksheet.Cells["M49"].Value = "□合格 □不合格";
                    worksheet.Cells["M49"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M49"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(50).Height = 45;

                    worksheet.Cells["A50"].Value = "2.9电气安全保护接地";
                    worksheet.Cells["A50"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A50"].Style.WrapText = true;

                    worksheet.Cells["B50:F50"].Merge = true;
                    worksheet.Cells["B50"].Value = "GB9706.1-2007标准中18）要求：具有电源输入插口的保护接地点与已保护接地的所有可触及金属部分之间的阻抗，不应超过0.1Ω； ";
                    worksheet.Cells["B50"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B50"].Style.WrapText = true;

                    worksheet.Cells["G50:L50"].Merge = true;
                    worksheet.Cells["G50"].Value = "测量数值为______Ω";
                    worksheet.Cells["G50"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G50"].Style.WrapText = true;

                    worksheet.Cells["M50"].Value = "□合格 □不合格";
                    worksheet.Cells["M50"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M50"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(51).Height = 47.75;

                    worksheet.Cells["A51"].Value = "2.9电气安全对地漏电流";
                    worksheet.Cells["A51"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A51"].Style.WrapText = true;

                    worksheet.Cells["B51:F51"].Merge = true;
                    worksheet.Cells["B51"].Value = "GB9706.1-2007标准中19）要求：正常状态≤0.5mA；单一故障状态≤1mA ";
                    worksheet.Cells["B51"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B51"].Style.WrapText = true;

                    worksheet.Cells["G51:L51"].Merge = true;
                    worksheet.Cells["G51"].Value = "正常状态：    ______mA\n单一故障状态：______mA";
                    worksheet.Cells["G51"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G51"].Style.WrapText = true;

                    worksheet.Cells["M51"].Value = "□合格 □不合格";
                    worksheet.Cells["M51"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M51"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(52).Height = 36;

                    worksheet.Cells["A52"].Value = "2.9电气安全外壳漏电流";
                    worksheet.Cells["A52"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A52"].Style.WrapText = true;

                    worksheet.Cells["B52:F52"].Merge = true;
                    worksheet.Cells["B52"].Value = "GB9706.1-2007标准中19）要求：正常状态≤0.1mA；单一故障状态≤0.5mA";
                    worksheet.Cells["B52"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B52"].Style.WrapText = true;

                    worksheet.Cells["G52:L52"].Merge = true;
                    worksheet.Cells["G52"].Value = "正常状态：    ______mA\n单一故障状态：______mA";
                    worksheet.Cells["G52"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G52"].Style.WrapText = true;

                    worksheet.Cells["M52"].Value = "□合格 □不合格";
                    worksheet.Cells["M52"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M52"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(53).Height = 48;

                    worksheet.Cells["A53"].Value = "2.9电气安全患者漏电流（ac）";
                    worksheet.Cells["A53"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A53"].Style.WrapText = true;

                    worksheet.Cells["B53:F53"].Merge = true;
                    worksheet.Cells["B53"].Value = "GB9706.1-2007标准中19）要求：正常状态≤0.1mA；单一故障状态≤0.5mA";
                    worksheet.Cells["B53"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B53"].Style.WrapText = true;

                    worksheet.Cells["G53:L53"].Merge = true;
                    worksheet.Cells["G53"].Value = "正常状态：    ______mA\n单一故障状态：______mA";
                    worksheet.Cells["G53"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G53"].Style.WrapText = true;

                    worksheet.Cells["M53"].Value = "□合格 □不合格";
                    worksheet.Cells["M53"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M53"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(54).Height = 48;

                    worksheet.Cells["A54"].Value = "2.9电气安全患者漏电流（dc）";
                    worksheet.Cells["A54"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A54"].Style.WrapText = true;

                    worksheet.Cells["B54:F54"].Merge = true;
                    worksheet.Cells["B54"].Value = "GB9706.1-2007标准中19）要求：正常状态≤0.01mA；单一故障状态≤0.05mA";
                    worksheet.Cells["B54"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B54"].Style.WrapText = true;

                    worksheet.Cells["G54:L54"].Merge = true;
                    worksheet.Cells["G54"].Value = "正常状态：    ______mA\n单一故障状态：______mA";
                    worksheet.Cells["G54"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G54"].Style.WrapText = true;

                    worksheet.Cells["M54"].Value = "□合格 □不合格";
                    worksheet.Cells["M54"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M54"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(55).Height = 69.75;

                    worksheet.Cells["A55"].Value = "2.9电气安全患者辅助漏电流（应用部分加网电压）";
                    worksheet.Cells["A55"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A55"].Style.WrapText = true;

                    worksheet.Cells["B55:F55"].Merge = true;
                    worksheet.Cells["B55"].Value = "B9706.1-2007标准中19）：单一故障状态≤5mA";
                    worksheet.Cells["B55"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B55"].Style.WrapText = true;

                    worksheet.Cells["G55:L55"].Merge = true;
                    worksheet.Cells["G55"].Value = "单一故障状态：______mA";
                    worksheet.Cells["G55"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G55"].Style.WrapText = true;

                    worksheet.Cells["M55"].Value = "□合格 □不合格";
                    worksheet.Cells["M55"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M55"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(56).Height = 48;

                    worksheet.Cells["A56"].Value = "2.9电气安全患者辅助漏电流（dc）";
                    worksheet.Cells["A56"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A56"].Style.WrapText = true;

                    worksheet.Cells["B56:F56"].Merge = true;
                    worksheet.Cells["B56"].Value = "GB9706.1-2007标准中19）：正常状态≤0.01mA；单一故障状态≤0.05mA";
                    worksheet.Cells["B56"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B56"].Style.WrapText = true;

                    worksheet.Cells["G56:L56"].Merge = true;
                    worksheet.Cells["G56"].Value = "正常状态：    ______mA\n单一故障状态：______mA";
                    worksheet.Cells["G56"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G56"].Style.WrapText = true;

                    worksheet.Cells["M56"].Value = "□合格 □不合格";
                    worksheet.Cells["M56"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M56"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(57).Height = 48;

                    worksheet.Cells["A57"].Value = "2.9电气安全患者辅助漏电流（ac）";
                    worksheet.Cells["A57"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A57"].Style.WrapText = true;

                    worksheet.Cells["B57:F57"].Merge = true;
                    worksheet.Cells["B57"].Value = "GB9706.1-2007标准中19）要求：正常状态≤0.1mA；单一故障状态≤0.5mA";
                    worksheet.Cells["B57"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B57"].Style.WrapText = true;

                    worksheet.Cells["G57:L57"].Merge = true;
                    worksheet.Cells["G57"].Value = "正常状态：    ______mA\n单一故障状态：______mA";
                    worksheet.Cells["G57"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G57"].Style.WrapText = true;

                    worksheet.Cells["M57"].Value = "□合格 □不合格";
                    worksheet.Cells["M57"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M57"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(58).Height = 48;

                    worksheet.Cells["A58"].Value = "2.9电气安全电介质强度A - al";
                    worksheet.Cells["A58"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A58"].Style.WrapText = true;

                    worksheet.Cells["B58:F58"].Merge = true;
                    worksheet.Cells["B58"].Value = "GB9706.1-2007标准中20）要求：A.带电部分与已保护接地的可触及金属部分之间1500V";
                    worksheet.Cells["B58"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B58"].Style.WrapText = true;

                    worksheet.Cells["G58:L58"].Merge = true;
                    worksheet.Cells["G58"].Value = "测量通过，无击穿，无闪络";
                    worksheet.Cells["G58"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G58"].Style.WrapText = true;

                    worksheet.Cells["M58"].Value = "□合格 □不合格";
                    worksheet.Cells["M58"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M58"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(59).Height = 48;

                    worksheet.Cells["A59"].Value = "2.9电气安全电介质强度A - a2";
                    worksheet.Cells["A59"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A59"].Style.WrapText = true;

                    worksheet.Cells["B59:F59"].Merge = true;
                    worksheet.Cells["B59"].Value = "GB9706.1-2007标准中20）要求：B.带电部分与未保护接地外壳之间4000V";
                    worksheet.Cells["B59"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B59"].Style.WrapText = true;

                    worksheet.Cells["G59:L59"].Merge = true;
                    worksheet.Cells["G59"].Value = "测量通过，无击穿，无闪络";
                    worksheet.Cells["G59"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G59"].Style.WrapText = true;

                    worksheet.Cells["M59"].Value = "□合格 □不合格";
                    worksheet.Cells["M59"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M59"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(60).Height = 48;

                    worksheet.Cells["A60"].Value = "2.9电气安全电介质强度B - d";
                    worksheet.Cells["A60"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A60"].Style.WrapText = true;

                    worksheet.Cells["B60:F60"].Merge = true;
                    worksheet.Cells["B60"].Value = "GB9706.1-2007标准中20）要求：C. 信号输入及信号输出部分在内的外壳之间，1500V";
                    worksheet.Cells["B60"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B60"].Style.WrapText = true;

                    worksheet.Cells["G60:L60"].Merge = true;
                    worksheet.Cells["G60"].Value = "测量通过，无击穿，无闪络";
                    worksheet.Cells["G60"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G60"].Style.WrapText = true;

                    worksheet.Cells["M60"].Value = "□合格 □不合格";
                    worksheet.Cells["M60"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M60"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(61).Height = 48;

                    worksheet.Cells["A61"].Value = "2.9电气安全电介质强度B - a";
                    worksheet.Cells["A61"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A61"].Style.WrapText = true;

                    worksheet.Cells["B61:F61"].Merge = true;
                    worksheet.Cells["B61"].Value = "GB9706.1-2007标准中20）要求：D.应用部分与带电部分之间，4000V";
                    worksheet.Cells["B61"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B61"].Style.WrapText = true;

                    worksheet.Cells["G61:L61"].Merge = true;
                    worksheet.Cells["G61"].Value = "测量通过，无击穿，无闪络";
                    worksheet.Cells["G61"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G61"].Style.WrapText = true;

                    worksheet.Cells["M61"].Value = "□合格 □不合格";
                    worksheet.Cells["M61"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M61"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    for (int i = 62; i <= 68; i++)
                    {
                        worksheet.Row(i).Height = 35;
                    }

                    worksheet.Cells["A62:A68"].Merge = true;
                    worksheet.Cells["A62"].Value = "老化测试";
                    worksheet.Cells["A62"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A62"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B62:D64"].Merge = true;
                    worksheet.Cells["B62"].Value = "（1）上肢阻抗测量范围：200Ω～690Ω，测量允差±5%；躯干阻抗测量范围：10Ω～35Ω，测量允差±10%；下肢阻抗测量范围：160Ω～500Ω，测量允差±5%。";
                    worksheet.Cells["B62"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B62"].Style.WrapText = true;

                    worksheet.Cells["E62:F62"].Merge = true;
                    worksheet.Cells["E62"].Value = "测量结果";
                    worksheet.Cells["E62"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E62"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E63:F63"].Merge = true;
                    worksheet.Cells["E63"].Value = "阻抗采集工装2";
                    worksheet.Cells["E63"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E63"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E64:F64"].Merge = true;
                    worksheet.Cells["E64"].Value = "误差值";
                    worksheet.Cells["E64"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E64"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["G62:H62"].Merge = true;
                    worksheet.Cells["G62"].Value = string.Format("上肢\n（{0}、{1}）", m_BodyStr[0], m_BodyStr[1]);
                    worksheet.Cells["G62"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G62"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G62"].Style.WrapText = true;

                    worksheet.Cells["G63"].Value = "最大值";
                    worksheet.Cells["G63"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G63"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["H63"].Value = "最小值";
                    worksheet.Cells["H63"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H63"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["I62:J62"].Merge = true;
                    worksheet.Cells["I62"].Value = string.Format("下肢\n（{0}、{1}）", m_BodyStr[3], m_BodyStr[4]);
                    worksheet.Cells["I62"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I62"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["I62"].Style.WrapText = true;

                    worksheet.Cells["I63"].Value = "最大值";
                    worksheet.Cells["I63"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["I63"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["J63"].Value = "最小值";
                    worksheet.Cells["J63"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["J63"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["K62:L62"].Merge = true;
                    worksheet.Cells["K62"].Value = string.Format("躯干\n（{0}）", m_BodyStr[2]);
                    worksheet.Cells["K62"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K62"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["K62"].Style.WrapText = true;

                    worksheet.Cells["K63"].Value = "最大值";
                    worksheet.Cells["K63"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["K63"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["L63"].Value = "最小值";
                    worksheet.Cells["L63"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["L63"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M62:M63"].Merge = true;
                    worksheet.Cells["M62"].Value = "/";
                    worksheet.Cells["M62"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M62"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M64"].Value = "□合格 □不合格";
                    worksheet.Cells["M64"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M64"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B65:D68"].Merge = true;
                    worksheet.Cells["B65"].Value = "（2）体重测量值范围允差±0.5Kg。";
                    worksheet.Cells["B65"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B65"].Style.WrapText = true;

                    worksheet.Cells["E65:F65"].Merge = true;
                    worksheet.Cells["E65"].Value = "25kg标准砝码测试值";
                    worksheet.Cells["E65"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E65"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["G65:L65"].Merge = true;
                    worksheet.Cells["G65"].Value = "老化前： ______ kg      老化后： ______ kg";
                    worksheet.Cells["G65"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G65"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M65:M66"].Merge = true;
                    worksheet.Cells["M65"].Value = "□合格 □不合格";
                    worksheet.Cells["M65"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M65"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E66:F66"].Merge = true;
                    worksheet.Cells["E66"].Value = "25kg标准砝码误差值";
                    worksheet.Cells["E66"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E66"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["G66:L66"].Merge = true;
                    worksheet.Cells["G66"].Value = "老化前： ______ kg      老化后： ______ kg";
                    worksheet.Cells["G66"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G66"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E67:F67"].Merge = true;
                    worksheet.Cells["E67"].Value = "50kg标准砝码测试值";
                    worksheet.Cells["E67"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E67"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["G67:L67"].Merge = true;
                    worksheet.Cells["G67"].Value = "老化前： ______ kg      老化后： ______ kg";
                    worksheet.Cells["G67"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G67"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["M67:M68"].Merge = true;
                    worksheet.Cells["M67"].Value = "□合格 □不合格";
                    worksheet.Cells["M67"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M67"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["E68:F68"].Merge = true;
                    worksheet.Cells["E68"].Value = "50kg标准砝码误差值";
                    worksheet.Cells["E68"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E68"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["G68:L68"].Merge = true;
                    worksheet.Cells["G68"].Value = "老化前： ______ kg      老化后： ______ kg";
                    worksheet.Cells["G68"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G68"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(69).Height = 65.75;

                    worksheet.Cells["A69"].Value = "包装检验";
                    worksheet.Cells["A69"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A69"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["A69"].Value = "包装检验";
                    worksheet.Cells["A69"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A69"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B69:F69"].Merge = true;
                    worksheet.Cells["B69"].Value = "a)包装应完整、牢固、无破损；\nb)与装箱清单核查，物品齐全；\nc)标签或铭牌信息清晰、完整、正确，包括产品名称、规格型号、生产日期、产品编号、使用期限等。";
                    worksheet.Cells["B69"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B69"].Style.WrapText = true;

                    worksheet.Cells["G69:L69"].Merge = true;
                    worksheet.Cells["G69"].Value = "a)整机包装完整、牢固、无破损；\nb)装箱清单物品齐全、标签、铭牌完整、正确；\nc)铭牌信息清晰、完整、正确。";
                    worksheet.Cells["G69"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G69"].Style.WrapText = true;

                    worksheet.Cells["M69"].Value = "√合格 □不合格";
                    worksheet.Cells["M69"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["M69"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(70).Height = 28.5;

                    worksheet.Cells["A70"].Value = "备    注";
                    worksheet.Cells["A70"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A70"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B70:M70"].Merge = true;

                    worksheet.Row(71).Height = 28.5;

                    worksheet.Cells["A71"].Value = "检 验 结 论";
                    worksheet.Cells["A71"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A71"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B71:M71"].Merge = true;
                    worksheet.Cells["B71"].Value = "□合格      □不合格";
                    worksheet.Cells["B71"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B71"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(72).Height = 28.5;

                    worksheet.Cells["A72"].Value = "质 检 员";
                    worksheet.Cells["A72"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A72"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B72:F72"].Merge = true;

                    worksheet.Cells["G72"].Value = "日   期";
                    worksheet.Cells["G72"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G72"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["H72:M72"].Merge = true;

                    worksheet.Row(73).Height = 28.5;

                    worksheet.Cells["A73"].Value = "批 准 人";
                    worksheet.Cells["A73"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A73"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B73:F73"].Merge = true;

                    worksheet.Cells["G73"].Value = "日   期";
                    worksheet.Cells["G73"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G73"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["H73:M73"].Merge = true;

                    //create a SaveFileDialog instance with some properties
                    SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                    saveFileDialog1.Title = "导出";
                    saveFileDialog1.Filter = "Excel files|*.xlsx|All files|*.*";
                    saveFileDialog1.FileName = m_ProductSerialNumber + ".xlsx";

                    //check if user clicked the save button
                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        //Get the FileInfo
                        FileInfo fi = new FileInfo(saveFileDialog1.FileName);
                        //write the file to the disk
                        excelPackage.SaveAs(fi);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                }

            }
        }

    }
}
