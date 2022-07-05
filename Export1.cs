using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Drawing;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace comTest
{
    class Export1
    {
        public string m_MachineType;
        public string m_SerialNumber;
        public string m_ProductSerialNumber;
        public float[][][] m_Measurements;
        public float[][][] m_Errors;
        public string[] m_BodyStr;
        public string[] m_BodyAngleStr;
        public string[] m_RateStr;
        public float[] m_TestWeightLevels;
        public float[][] m_TestWeightResults;
        public string[] m_WeightPositionLevels;
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
                    excelPackage.Workbook.Properties.Title = "临床营养检测分析仪（AiNST-CNDS）生产批记录";
                    excelPackage.Workbook.Properties.Subject = "AINST";
                    excelPackage.Workbook.Properties.Created = nowDate;

                    //Create the WorkSheet
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("记录单");

                    worksheet.Column(1).Width = 8;
                    worksheet.Column(2).Width = 12;
                    worksheet.Column(3).Width = 15;
                    worksheet.Column(4).Width = 10;
                    worksheet.Column(5).Width = 10;
                    worksheet.Column(6).Width = 12;
                    worksheet.Column(7).Width = 10;
                    worksheet.Column(8).Width = 10;

                    worksheet.Cells["A1:H34"].Style.Font.Size = 10;
                    worksheet.Cells["A1:H34"].Style.Font.Name = "宋体";
                    worksheet.Cells["A1:H34"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A1:H34"].Style.Border.Top.Color.SetColor(Color.Black);
                    worksheet.Cells["A1:H34"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A1:H34"].Style.Border.Right.Color.SetColor(Color.Black);
                    worksheet.Cells["A1:H34"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A1:H34"].Style.Border.Bottom.Color.SetColor(Color.Black);
                    worksheet.Cells["A1:H34"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    worksheet.Cells["A1:H34"].Style.Border.Left.Color.SetColor(Color.Black);

                    worksheet.Row(1).Height = 15;
                    worksheet.Cells["A1:H1"].Merge = true;
                    worksheet.Cells["A1"].Style.Font.Size = 10.5f;
                    worksheet.Cells["A1"].Value = "编号：CX0901-R01                                          版本：V1.0";

                    worksheet.Row(2).Height = 21;
                    worksheet.Cells["A2:H2"].Merge = true;
                    worksheet.Cells["A2"].Style.Font.Size = 16;
                    worksheet.Cells["A2"].Style.Font.Bold = true;
                    worksheet.Cells["A2"].Value = "临床营养检测分析仪（AiNST-CNDS）";
                    worksheet.Cells["A2"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Row(3).Height = 21;
                    worksheet.Cells["A3:H3"].Merge = true;
                    worksheet.Cells["A3"].Style.Font.Size = 16;
                    worksheet.Cells["A3"].Style.Font.Bold = true;
                    worksheet.Cells["A3"].Value = "生产批记录";
                    worksheet.Cells["A3"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                    worksheet.Row(4).Height = 17.25;
                    worksheet.Cells["A4:H4"].Merge = true;
                    worksheet.Cells["A4"].Style.Font.Size = 10.5f;
                    worksheet.Cells["A4"].Value = string.Format("编号：{0}", m_SerialNumber);
                    worksheet.Cells["A4"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

                    worksheet.Row(5).Height = 25;

                    worksheet.Cells["A5"].Value = "工序号";
                    worksheet.Cells["A5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A5"].Style.Font.Bold = true;

                    worksheet.Cells["B5"].Value = "工序名称";
                    worksheet.Cells["B5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["B5"].Style.Font.Bold = true;

                    worksheet.Cells["C5:H5"].Merge = true;
                    worksheet.Cells["C5"].Value = "内容简要描述";
                    worksheet.Cells["C5"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C5"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C5"].Style.Font.Bold = true;

                    worksheet.Row(6).Height = 25;
                    worksheet.Row(7).Height = 50;

                    worksheet.Cells["C6:C7"].Merge = true;
                    worksheet.Cells["C6"].Value = "工装分类";
                    worksheet.Cells["C6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["D6:H6"].Merge = true;
                    worksheet.Cells["D6"].Value = "测量项目（单位Ω）";
                    worksheet.Cells["D6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["D7"].Value = "RA\n（右上肢）";
                    worksheet.Cells["D7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D7"].Style.WrapText = true;

                    worksheet.Cells["E7"].Value = "LA\n（左上肢）";
                    worksheet.Cells["E7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E7"].Style.WrapText = true;

                    worksheet.Cells["F7"].Value = "TT\n（躯干）";
                    worksheet.Cells["F7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F7"].Style.WrapText = true;

                    worksheet.Cells["G7"].Value = "RL\n（右下肢）";
                    worksheet.Cells["G7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G7"].Style.WrapText = true;

                    worksheet.Cells["H7"].Value = "LL\n（左下肢）";
                    worksheet.Cells["H7"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H7"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H7"].Style.WrapText = true;

                    string[] devices = { "校准工装档一", "校准工装档二", "校准工装档三", "校准工装档四", "校准工装档五", "校准工装档六"};
                    double[] ra_la_min = { 712.5, 646, 532, 408.5, 250, 190 };
                    double[] ra_la_max = { 787.5, 714, 588, 451.5, 315, 210 };
                    double[] tt_min = { 45, 38.7, 32.4, 27, 18, 9 };
                    double[] tt_max = { 55, 47.3, 39.6, 33, 22, 11 };
                    double[] rl_ll_min = { 487.5, 342, 285, 237.5, 190, 142.5 };
                    double[] rl_ll_max = { 532.5, 378, 315, 262.5, 210, 157.5 };

                    int row1 = 8;
                    for (int deviceIndex = 0; deviceIndex < devices.Length; deviceIndex++)
                    {
                        float[][] measureValues = m_Measurements[deviceIndex];

                        float[] ra_values = measureValues[1];
                        float[] la_values = measureValues[4];
                        float[] tt_values = measureValues[2];
                        float[] rl_values = measureValues[0];
                        float[] ll_values = measureValues[3];

                        int deviceRow = row1 + 2 * deviceIndex;
                        worksheet.Row(deviceRow).Height = 25;
                        worksheet.Row(deviceRow + 1).Height = 25;

                        string cc = string.Format("C{0}", deviceRow);
                        worksheet.Cells[string.Format("C{0}:C{1}", deviceRow, deviceRow + 1)].Merge = true;
                        worksheet.Cells[cc].Value = devices[deviceIndex];
                        worksheet.Cells[cc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[cc].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[cc].Style.Font.Bold = true;

                        string dc = string.Format("D{0}", deviceRow);
                        worksheet.Cells[string.Format("D{0}:E{0}", deviceRow)].Merge = true;
                        worksheet.Cells[dc].Value = string.Format("{0}～{1}", ra_la_min[deviceIndex], ra_la_max[deviceIndex]);
                        worksheet.Cells[dc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[dc].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[dc].Style.Font.Bold = true;

                        if (ra_values.Length > 0)
                        {
                            string dc1 = string.Format("D{0}", deviceRow + 1);
                            worksheet.Cells[dc1].Value = ra_values[0];
                            worksheet.Cells[dc1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[dc1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        if (la_values.Length > 0)
                        {
                            string ec1 = string.Format("E{0}", deviceRow + 1);
                            worksheet.Cells[ec1].Value = la_values[0];
                            worksheet.Cells[ec1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[ec1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        string fc = string.Format("F{0}", deviceRow);
                        worksheet.Cells[fc].Value = string.Format("{0}～{1}", tt_min[deviceIndex], tt_max[deviceIndex]);
                        worksheet.Cells[fc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[fc].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[fc].Style.Font.Bold = true;

                        if (tt_values.Length > 0)
                        {
                            string fc1 = string.Format("F{0}", deviceRow + 1);
                            worksheet.Cells[fc1].Value = tt_values[0];
                            worksheet.Cells[fc1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[fc1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        string gc = string.Format("G{0}", deviceRow);
                        worksheet.Cells[string.Format("G{0}:H{0}", deviceRow)].Merge = true;
                        worksheet.Cells[gc].Value = string.Format("{0}～{1}", rl_ll_min[deviceIndex], rl_ll_max[deviceIndex]);
                        worksheet.Cells[gc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[gc].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[gc].Style.Font.Bold = true;

                        if (rl_values.Length > 0)
                        {
                            string gc1 = string.Format("G{0}", deviceRow + 1);
                            worksheet.Cells[gc1].Value = rl_values[0];
                            worksheet.Cells[gc1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[gc1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        if (ll_values.Length > 0)
                        {
                            string hc1 = string.Format("H{0}", deviceRow + 1);
                            worksheet.Cells[hc1].Value = ll_values[0];
                            worksheet.Cells[hc1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[hc1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                    }

                    worksheet.Row(20).Height = 75;

                    worksheet.Cells["C20:G20"].Merge = true;
                    worksheet.Cells["C20"].Value = "（9）将阻抗采集工装1连接到临床营养检测分析仪手脚电极上，LA左手、RA右手、LL左脚、RL右脚，将开关旋钮旋至第七档；点击“测量阻抗值”按钮，信息窗会显示阻抗检测信息、阻抗检测完成、阻抗读取成功，读取阻抗值应符合要求。记录5K频率测值";
                    worksheet.Cells["C20"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C20"].Style.WrapText = true;

                    worksheet.Row(21).Height = 75;

                    worksheet.Cells["C21:G21"].Merge = true;
                    worksheet.Cells["C21"].Value = "（10）将阻抗采集工装1连接到临床营养检测分析仪手脚电极上，LA左手、RA右手、LL左脚、RL右脚，将开关旋钮旋至第八档；点击图“测量阻抗值”按钮，等待，信息窗会显示阻抗检测信息、阻抗检测完成、阻抗读取成功，读取阻抗值应符合要求。记录5K频率测值";
                    worksheet.Cells["C21"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C21"].Style.WrapText = true;

                    worksheet.Row(22).Height = 25;
                    worksheet.Row(23).Height = 50;

                    worksheet.Cells["C22:C23"].Merge = true;
                    worksheet.Cells["C22"].Value = "工装分类";
                    worksheet.Cells["C22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C22"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["D22:H22"].Merge = true;
                    worksheet.Cells["D22"].Value = "测量项目（单位Ω）";
                    worksheet.Cells["D22"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D22"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["D23"].Value = "RA\n（右上肢）";
                    worksheet.Cells["D23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D23"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D23"].Style.WrapText = true;

                    worksheet.Cells["E23"].Value = "LA\n（左上肢）";
                    worksheet.Cells["E23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["E23"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E23"].Style.WrapText = true;

                    worksheet.Cells["F23"].Value = "TT\n（躯干）";
                    worksheet.Cells["F23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F23"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F23"].Style.WrapText = true;

                    worksheet.Cells["G23"].Value = "RL\n（右下肢）";
                    worksheet.Cells["G23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["G23"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["G23"].Style.WrapText = true;

                    worksheet.Cells["H23"].Value = "LL\n（左下肢）";
                    worksheet.Cells["H23"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["H23"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["H23"].Style.WrapText = true;

                    double[] ra_la_min1 = { 589, 370.5 };
                    double[] ra_la_max1 = { 651, 409.5 };
                    double[] tt_min1 = { 42.5, 13.5 };
                    double[] tt_max1 = { 51.7, 16.5 };
                    double[] rl_ll_min1 = { 446.5, 493.5 };
                    double[] rl_ll_max1 = { 171, 189 };

                    int row2 = 24;
                    for (int deviceIndex = 0; deviceIndex < ra_la_min1.Length; deviceIndex++)
                    {
                        int deviceIndex1 = deviceIndex + 6;

                        float[][] measureValues = m_Measurements[deviceIndex1];

                        float[] ra_values = measureValues[1];
                        float[] la_values = measureValues[4];
                        float[] tt_values = measureValues[2];
                        float[] rl_values = measureValues[0];
                        float[] ll_values = measureValues[3];

                        int deviceRow = row2 + 2 * deviceIndex;
                        worksheet.Row(deviceRow).Height = 25;
                        worksheet.Row(deviceRow + 1).Height = 25;

                        string dc = string.Format("D{0}", deviceRow);
                        worksheet.Cells[string.Format("D{0}:E{0}", deviceRow)].Merge = true;
                        worksheet.Cells[dc].Value = string.Format("{0}～{1}", ra_la_min1[deviceIndex], ra_la_max1[deviceIndex]);
                        worksheet.Cells[dc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[dc].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[dc].Style.Font.Bold = true;

                        if (ra_values.Length > 0)
                        {
                            string dc1 = string.Format("D{0}", deviceRow + 1);
                            worksheet.Cells[dc1].Value = ra_values[0];
                            worksheet.Cells[dc1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[dc1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        if (la_values.Length > 0)
                        {
                            string ec1 = string.Format("E{0}", deviceRow + 1);
                            worksheet.Cells[ec1].Value = la_values[0];
                            worksheet.Cells[ec1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[ec1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        string fc = string.Format("F{0}", deviceRow);
                        worksheet.Cells[fc].Value = string.Format("{0}～{1}", tt_min1[deviceIndex], tt_max1[deviceIndex]);
                        worksheet.Cells[fc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[fc].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[fc].Style.Font.Bold = true;

                        if (tt_values.Length > 0)
                        {
                            string fc1 = string.Format("F{0}", deviceRow + 1);
                            worksheet.Cells[fc1].Value = tt_values[0];
                            worksheet.Cells[fc1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[fc1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        string gc = string.Format("G{0}", deviceRow);
                        worksheet.Cells[string.Format("G{0}:H{0}", deviceRow)].Merge = true;
                        worksheet.Cells[gc].Value = string.Format("{0}～{1}", rl_ll_min1[deviceIndex], rl_ll_max1[deviceIndex]);
                        worksheet.Cells[gc].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                        worksheet.Cells[gc].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        worksheet.Cells[gc].Style.Font.Bold = true;

                        if (rl_values.Length > 0)
                        {
                            string gc1 = string.Format("G{0}", deviceRow + 1);
                            worksheet.Cells[gc1].Value = rl_values[0];
                            worksheet.Cells[gc1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[gc1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }

                        if (ll_values.Length > 0)
                        {
                            string hc1 = string.Format("H{0}", deviceRow + 1);
                            worksheet.Cells[hc1].Value = ll_values[0];
                            worksheet.Cells[hc1].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                            worksheet.Cells[hc1].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        }
                    }

                    worksheet.Cells["C24:C27"].Merge = true;
                    worksheet.Cells["C24"].Value = "阻抗采集工装1";
                    worksheet.Cells["C24"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C24"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["A6:A27"].Merge = true;
                    worksheet.Cells["A6"].Value = "步骤二";
                    worksheet.Cells["A6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B6:B27"].Merge = true;
                    worksheet.Cells["B6"].Value = "阻抗校准";
                    worksheet.Cells["B6"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B6"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(28).Height = 50;

                    worksheet.Cells["C28:G28"].Merge = true;
                    worksheet.Cells["C28"].Value = "（1）将电脑连接的串口线从采集板拔下，把临床营养检测分析仪的ARM主板的串口线连接到采集板上";
                    worksheet.Cells["C28"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C28"].Style.WrapText = true;

                    worksheet.Row(29).Height = 50;

                    worksheet.Cells["C29:G29"].Merge = true;
                    worksheet.Cells["C29"].Value = "（2）打开临床营养检测分析仪的软件，选中任一患者，将25Kg砝码放于临床营养检测仪的底座上，查看体重值，应满足误差±0.5Kg的要求";
                    worksheet.Cells["C29"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C29"].Style.WrapText = true;

                    worksheet.Row(30).Height = 50;

                    worksheet.Cells["C30:G30"].Merge = true;
                    worksheet.Cells["C30"].Value = "（3）将60Kg和120Kg砝码放于临床营养检测仪的底座上，查看体重值，应满足误差±0.5Kg的要求";
                    worksheet.Cells["C30"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C30"].Style.WrapText = true;

                    worksheet.Row(31).Height = 25;

                    worksheet.Cells["C31"].Value = "25Kg";
                    worksheet.Cells["C31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["C31"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C31"].Style.Font.Bold = true;

                    worksheet.Cells["D31:E31"].Merge = true;
                    worksheet.Cells["D31"].Value = "60Kg";
                    worksheet.Cells["D31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["D31"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["D31"].Style.Font.Bold = true;

                    worksheet.Cells["F31:G31"].Merge = true;
                    worksheet.Cells["F31"].Value = "120Kg";
                    worksheet.Cells["F31"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["F31"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["F31"].Style.Font.Bold = true;

                    worksheet.Row(32).Height = 25;

                    worksheet.Cells["D32:E32"].Merge = true;

                    worksheet.Cells["F32:G32"].Merge = true;

                    worksheet.Cells["A28:A32"].Merge = true;
                    worksheet.Cells["A28"].Value = "步骤三";
                    worksheet.Cells["A28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A28"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B28:B32"].Merge = true;
                    worksheet.Cells["B28"].Value = "体重检测";
                    worksheet.Cells["B28"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B28"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Row(33).Height = 50;

                    worksheet.Cells["A33"].Value = "/";
                    worksheet.Cells["A33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["A33"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["B33"].Value = "后清场";
                    worksheet.Cells["B33"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                    worksheet.Cells["B33"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    worksheet.Cells["C33:G33"].Merge = true;
                    worksheet.Cells["C33"].Value = "生产区已清洁，无本次生产遗留物，设备归位，填写《清场记录》";
                    worksheet.Cells["C33"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["C33"].Style.WrapText = true;

                    worksheet.Row(34).Height = 50;

                    worksheet.Cells["A34:D34"].Merge = true;
                    worksheet.Cells["A34"].Value = "生产人员/日期：";
                    worksheet.Cells["A34"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["A34"].Style.WrapText = true;

                    worksheet.Cells["E34:H34"].Merge = true;
                    worksheet.Cells["E34"].Value = "现场QA/日期：";
                    worksheet.Cells["E34"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    worksheet.Cells["E34"].Style.WrapText = true;

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
