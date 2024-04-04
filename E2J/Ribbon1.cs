using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace Excel2Json
{
    public partial class Ribbon1
    {
        // Ribbon加载时的事件处理程序
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        // 导出客户端数据的按钮点击事件处理程序
        /// <summary>
        /// 导出客户端Json非标准
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportClient_Click(object sender, RibbonControlEventArgs e)
        {

           
             ExportData("c",false);  // 导出客户端Json,fei
        }

        
        /// <summary>
        /// 导出服务端Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportServer_Click(object sender, RibbonControlEventArgs e)
        {
          
            string stringToRemove = "_e";


            // 获取源工作簿
            Excel.Workbook sourceWorkbook = Globals.ThisAddIn.Application.ActiveWorkbook;

            // 获取当前活动工作表
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;
            string sourceSheetName = activeWorksheet.Name;
            int sourceSheetIndex = activeWorksheet.Index;

            // 获取使用范围
            Excel.Range usedRange = activeWorksheet.UsedRange;

            // 弹出保存文件对话框让用户选择保存路径
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                FileName = Path.GetFileName(Globals.ThisAddIn.Application.ActiveWorkbook.FullName), // 默认文件名为源文件名剔除指定字符串之后的名字
                DefaultExt = ".xlsx",
                Filter = "Excel files (*.xlsx)|*.xlsx",
                Title = "保存为"
            };

            if (saveFileDialog.ShowDialog() != DialogResult.OK)
            {
                return; // 用户取消操作
            }

            // 目标文件路径
            string targetWorkbookPath = saveFileDialog.FileName;

            // 打开或创建目标工作簿
            Workbook targetWorkbook=null;
            bool workbookWasOpen = false;

            foreach (Workbook wb in Globals.ThisAddIn.Application.Workbooks)
            {
                if (wb.FullName == targetWorkbookPath)
                {
                    targetWorkbook = wb;
                    workbookWasOpen = true;
                    break;
                }
            }

            if (!workbookWasOpen)
            {
                if (File.Exists(targetWorkbookPath))
                {
                    targetWorkbook = Globals.ThisAddIn.Application.Workbooks.Open(targetWorkbookPath);
                }
                else
                {
                    targetWorkbook = Globals.ThisAddIn.Application.Workbooks.Add();
                }
            }


            // 遍历源工作簿中的所有工作表,并进行表页签同步
            foreach (Excel.Worksheet sourceWorksheet in sourceWorkbook.Sheets)
            {
                //取得页签的中文部分名称
                var name = sourceWorksheet.Name.Split('_')[0].Trim();

                // 获取或创建目标工作簿中的工作表
                Excel.Worksheet tarWorksheet = GetOrCreateWorksheet(targetWorkbook, name);

                // 移动到与源工作簿中相同的位置
                tarWorksheet.Move(Before: targetWorkbook.Sheets[sourceWorksheet.Index]);
            }

            // 检查目标工作簿中是否存在与当前源工作表同名的工作表
            var ssname = sourceSheetName.Split('_')[0].Trim();
            Worksheet targetWorksheet = GetOrCreateWorksheet(targetWorkbook, ssname);

            // 如果存在，清除现有工作表内容
            targetWorksheet.Cells.Clear();


            // 将服务端数据复制到目标工作表
            int targetRow = 1; // 目标工作表的起始行
           
            for (int row = 2; row <= usedRange.Rows.Count; row++)
            {
               
                string rowDataType = (usedRange.Cells[row, 1].Value2?.ToString().Trim() ?? "cs").ToLower();
                if (rowDataType != "s" && rowDataType != "cs" && rowDataType != null && rowDataType == "#") continue; // 如果不是服务端数据，跳过这行

                int targetCol = 1;// 目标工作表的起始列
                for (int col = 2; col <= usedRange.Columns.Count; col++)
                {
                    string colDataType = (usedRange.Cells[1, col].Value2?.ToString().Trim() ?? "cs").ToLower();
                    if (colDataType != "s" && colDataType != "cs" && colDataType != null || colDataType == "#") continue; // 如果不是目标类型数据，跳过这列

                    string fieldName = usedRange.Cells[3, col].Value2?.ToString().Trim(); // 获取字段名
                    if (fieldName == null || fieldName == "")
                    {
                        MessageBox.Show($"字段名为空。位置：行{3}，列{col}");
                        Globals.ThisAddIn.Application.Goto(usedRange.Cells[3, col]);
                        return; // 如果字段名不存在，返回
                    }

                    object fieldValue = usedRange.Cells[row, col].Value2; // 获取字段值

                    if (fieldValue == null || fieldValue.ToString().Trim() == "")
                    {
                        MessageBox.Show($"数据内容为空。位置：行{row}，列{col}");
                        Globals.ThisAddIn.Application.Goto(usedRange.Cells[row, col]);
                        return;
                    }
                    // 复制单元格内容
                    targetWorksheet.Cells[targetRow, targetCol].Value = usedRange.Cells[row, col].Value;


                    targetCol++;
                }
                targetRow++;
                
            }

            // 保存并关闭目标工作簿
            if (!workbookWasOpen)
            {
                targetWorkbook.SaveAs(targetWorkbookPath);
                targetWorkbook.Close();
            }
            else
            {
                targetWorkbook.Save();
            }

            // 释放COM对象
            Marshal.ReleaseComObject(targetWorksheet);
            Marshal.ReleaseComObject(targetWorkbook);
            MessageBox.Show("服务端Excel导出完成！");


        }

        /// <summary>
        /// 导出标准的客户端json
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            ExportData("c");
        }
        /// <summary>
        /// 导出服务端Json
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnExportServerJson_Click(object sender, RibbonControlEventArgs e)
        {
            ExportData("s");//导出服务端Json
        }


       

        // 导出数据
        private void ExportData(string type,bool biaozhun=true)
        {

            // 获取当前活动工作表
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

            // 获取使用范围
            Excel.Range usedRange = activeWorksheet.UsedRange;

            // 构建JSON数据结构
            var jsonData = new List<Dictionary<string, object>>();

            // 遍历使用范围的所有行
            for (int row = 4; row <= usedRange.Rows.Count; row++) // 从第三行开始，跳过备注行
            {
                string rowDataType = (usedRange.Cells[row, 1].Value2?.ToString().Trim() ?? "cs").ToLower();
                if (rowDataType != type && rowDataType != "cs" && rowDataType != null || rowDataType == "#") continue; // 如果不是目标类型数据，跳过这行

                // 遍历列
                Dictionary<string, object> rowData = new Dictionary<string, object>();
                for (int col = 2; col <= usedRange.Columns.Count; col++) // 从第二列开始
                {
                    string colDataType = (usedRange.Cells[1, col].Value2?.ToString().Trim() ?? "cs").ToLower();
                    if (colDataType != type && colDataType != "cs" && colDataType != null || colDataType == "#") continue; // 如果不是目标类型数据，跳过这列

                    string fieldName = usedRange.Cells[3, col].Value2?.ToString().Trim(); // 获取字段名
                    if (fieldName == null || fieldName == "")
                    {
                        MessageBox.Show($"字段名为空。位置：行{3}，列{col}");
                        Globals.ThisAddIn.Application.Goto(usedRange.Cells[3, col]);
                        return; // 如果字段名不存在，返回
                    }

                    object fieldValue = usedRange.Cells[row, col].Text.Trim(); // 获取字段值

                    if (fieldValue == null || fieldValue.ToString() == "")
                    {
                        if (colDataType != "c")
                        {
                            MessageBox.Show($"数据内容为空。位置：行{row}，列{col},列类型：{colDataType}");
                            Globals.ThisAddIn.Application.Goto(usedRange.Cells[row, col]);
                            return;
                        }
                        else
                        {
                            continue;
                        }
                        
                    }

                    rowData.Add(fieldName, fieldValue);
                }

                if (rowData.Count > 0)
                {
                    jsonData.Add(rowData); // 只有当rowData不为空时才添加到jsonData
                }
            }

            // 将JSON数据转换为字符串
            string jsonString = JsonConvert.SerializeObject(jsonData, Formatting.Indented);

           var titlelist= activeWorksheet.Name.Split('_');
            var title = titlelist[0].Trim();
            if (titlelist.Length > 1)
            {

                title= titlelist[1].Trim();
            }

            if (!biaozhun)
            {
                jsonString = "{\r\n  \"Config\": {\r\n    \"-xmlns:xsi\": \"http://www.w3.org/2001/XMLSchema-instance\",\r\n    \"" + title + "\":"+ jsonString+ "  }\r\n}";
            }

            // 弹出保存文件对话框
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "JSON files (*.json)|*.json";
            saveFileDialog.FileName = title;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // 保存文件
                System.IO.File.WriteAllText(saveFileDialog.FileName, jsonString);
                MessageBox.Show("导出完成");
            }
        }
        // 创建或获取工作表
        Excel.Worksheet GetOrCreateWorksheet(Excel.Workbook workbook, string sheetName)
        {
            foreach (Excel.Worksheet sheet in workbook.Sheets)
            {
                if (sheet.Name == sheetName)
                {
                    return sheet;
                }
            }

            Excel.Worksheet newSheet = workbook.Sheets.Add(After: workbook.Sheets[workbook.Sheets.Count]);
            newSheet.Name = sheetName;
            return newSheet;
        }
    }
}
