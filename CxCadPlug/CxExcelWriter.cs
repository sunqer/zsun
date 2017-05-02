using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.IO;
using System.Windows.Forms;

// microsoft 引用
using Microsoft.Office.Interop;

namespace CxCadPlug
{
    /// <summary>
    /// 类名：CxExcelWriter
    /// 说明：本类用于将读取到的实体属性进行写Excel操作
    /// 作者： sun
    /// 日期： 2016-09-24
    /// </summary>
    class CxExcelWriter
    {

        private string m_strFilePath;
        private Microsoft.Office.Interop.Excel.Application m_pApp;
        private Microsoft.Office.Interop.Excel.Workbook m_pWb;

        public ProgressBar ProgressBar { get; set; }
        public List<string> DefaultColumn { get; set; } 

        public CxExcelWriter()
        { 
        }

        /// <summary>
        /// 创建一个新的 Excel.Application 对象
        /// </summary>
        public void Create()
        {
            m_pApp = new Microsoft.Office.Interop.Excel.Application();
            m_pWb = m_pApp.Workbooks.Add(true);
        }

        /// <summary>
        /// 打开一个Excel文档
        /// </summary>
        public void Open(string strFileName)
        {
            m_strFilePath = strFileName;

            if (!File.Exists(strFileName))
            {
                Create();
                m_pWb.SaveAs(strFileName);
                return;
            }

            m_pApp = new Microsoft.Office.Interop.Excel.Application();
            m_pWb = m_pApp.Workbooks.Add(strFileName);
        }

        /// <summary>
        /// 根据名称获取一个工作表
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public Microsoft.Office.Interop.Excel.Worksheet GetSheet(string strName)
        {
            return (Microsoft.Office.Interop.Excel.Worksheet)m_pWb.Worksheets[strName];
        }

        /// <summary>
        /// 激活一个工作表
        /// </summary>
        /// <param name="strName"></param>
        public void ActiveSheet(string strName)
        { 
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)m_pWb.Worksheets[strName];
            ws.Activate();
        }

        /// <summary>
        /// 判断指定名称的表单是否存在
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        private bool HasSheet(string strName)
        {
            Microsoft.Office.Interop.Excel.Worksheet ws;
            
            for (int i = 1; i <= m_pWb.Worksheets.Count; ++i)
            {
                ws = (Microsoft.Office.Interop.Excel.Worksheet)m_pApp.Worksheets[i];
                if (ws.Name == strName)
                {
                    return true;
                }
            }
                return false;
        }

        /// <summary>
        /// 根据名称添加一个工作表
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public Microsoft.Office.Interop.Excel.Worksheet AddSheet(string strName)
        {
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)m_pWb.Worksheets.Add();
            ws.Name = strName;
            return ws;
        }

        /// <summary>
        /// 根据名称删除一个工作表
        /// </summary>
        /// <param name="strName"></param>
        /// <returns></returns>
        public void DelSheet(string strName)
        {
            ((Microsoft.Office.Interop.Excel.Worksheet)m_pWb.Worksheets[strName]).Delete();
        }

        public Microsoft.Office.Interop.Excel.Worksheet ReNameSheet(string strOldName, string strNewName)
        {
            Microsoft.Office.Interop.Excel.Worksheet ws = (Microsoft.Office.Interop.Excel.Worksheet)m_pWb.Worksheets[strOldName];
            if (ws != null)
            {
                ws.Name = strNewName;
            }
            return ws;
        }

        /// <summary>
        /// 将内存中的数据表格插入到Excel工作表中
        /// </summary>
        /// <param name="dt">内存数据表</param>
        /// <param name="wsName">要插入的表名称</param>
        /// <param name="startX">起始行,这个参数使用ref传值，因为添加标题后行会增加</param>
        /// <param name="startY">起始列</param>
        public void InsertTable(ref System.Data.DataTable dt, string wsName, ref int startRow, int startCol)
        {
            Microsoft.Office.Interop.Excel.Worksheet ws;
            if (HasSheet(wsName))
            {
                ws = GetSheet(wsName);
            }
            else 
            {
                ws = AddSheet(wsName);
            }

            ActiveSheet(wsName);

            if (startRow == 1)
            {
                // 将DataTable的列名导入Excel的第一行
                int columnIdx = 1;
                ws.Cells[startRow, columnIdx++] = "序号";

                foreach (System.Data.DataColumn col in dt.Columns)
                {
                    ws.Cells[startRow, columnIdx++] = col.ColumnName;
                }

                // 加入额外需要的空列
                foreach (string col in DefaultColumn)
                {
                    ws.Cells[startRow, columnIdx++] = col;
                }

                Microsoft.Office.Interop.Excel.Range titleRange = ws.Range[ws.Cells[startRow, startCol], ws.Cells[startRow, columnIdx]];
                titleRange.Font.Bold = true;   //加粗显示

                startRow++;
            }

            for (int i = 0; i < dt.Rows.Count; ++i)
            {
                ws.Cells[startRow + i, startCol] = i.ToString();   // 第一列加入序号信息

                for (int j = 0; j < dt.Columns.Count; ++j)
                {
                    ws.Cells[startRow + i, startCol + 1 + j] = dt.Rows[i][j].ToString();
                }
                ProgressBar.Value++;
            }
        }

        /// <summary>
        /// 保存工作表
        /// </summary>
        /// <returns></returns>
        public bool Save()
        {
            if (m_strFilePath == "")
            {
                return false;
            }
            else
            {
                try
                {
                    m_pWb.Save();
                    return true;
                }
                catch(Exception ex)
                {
                    return false;
                }
            }
        }
        /// <summary>
        /// 关闭Excel
        /// </summary>
        public void Close()
        {
            m_pWb.Close(true);
            m_pApp.Quit();
            m_pWb = null;
            m_pApp = null;
            GC.Collect();
        }
    }
}
