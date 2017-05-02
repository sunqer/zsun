using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Threading;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;

namespace CxCadPlug
{
    public partial class export2ExcelFrm : Form
    {
        private CxCadReader mCadReader;
        private List<string> mPointTableColumn;
        private List<string> mLineTableColumn;
        private List<string> mNoteTableColumn;

        public export2ExcelFrm()
        {
            InitializeComponent();

            mCadReader = null;

            ReadLayer();

            ReadCfgFile();

            this.compareCmb.ItemHeight = 19;
            this.typeCmb.ItemHeight = 19;
            this.compareCmb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.typeCmb.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.compareCmb.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.compareCmb_DrawItem);
            this.typeCmb.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.typeCmb_DrawItem);
        }

        /*
         * 自定义下拉框的高度
         * 首先设置一个较大的 ItemHeight 值，比如 20；
         * 然后设置 ComboBox 的 DrawMode 为 OwnerDrawVariable；
         * 然后在 DrawItem 事件中实现如何代码
         */
        private void compareCmb_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            e.DrawBackground();
            e.DrawFocusRectangle();
            e.Graphics.DrawString(compareCmb.Items[e.Index].ToString(), e.Font, new SolidBrush(e.ForeColor), e.Bounds.X, e.Bounds.Y + 3);
        }
        /// <summary>
        /// 下拉框的重绘事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void typeCmb_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0)
            {
                return;
            }

            e.DrawBackground();
            e.DrawFocusRectangle();
            e.Graphics.DrawString(typeCmb.Items[e.Index].ToString(), e.Font, new SolidBrush(e.ForeColor), e.Bounds.X, e.Bounds.Y + 3);
        }


        /// <summary>
        /// 得到当前用户选择需要导出的图层名称
        /// </summary>
        /// <param name="layers">传出参数，选择的图层名称列表</param>
        public void GetSelectedLayers(ref List<string> layers)
        {
            for (int idx = 0; idx < layerChb.Items.Count; ++idx)
            {
                if (layerChb.GetItemChecked(idx))
                {
                    layers.Add(layerChb.Items[idx].ToString());
                }
            }
        }

        /// <summary>
        /// 获取当前激活视图中的所有图层信息
        /// </summary>
        private void ReadLayer()
        {
            // 获取当前文档,开启事务管理器
            Document acDoc;
            acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acDoc.TransactionManager.StartTransaction())
            {
                // 读取图层表的信息, 并添加到下拉框中
                LayerTable acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                LayerTableRecord acLyrTblRec;

                foreach (ObjectId acObjId in acLyrTbl)
                {
                    acLyrTblRec = acTrans.GetObject(acObjId, OpenMode.ForRead) as LayerTableRecord;
                    layerChb.Items.Add(acLyrTblRec.Name);
                }
            }
        }

        /// <summary>
        /// 读取本地的图层信息和对照表信息的配置文件
        /// </summary>
        private void ReadCfgFile()
        {
            XmlDocument xmlDoc = new XmlDocument();
            string strFilePath = GetSefPath() + "/CxCfg.xml";
            try
            {
                xmlDoc.Load(strFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "读取错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            typeCmb.Items.Add("---请选择管网类型---");
            typeCmb.SelectedIndex = 0;

            // 使用xpath表达式选择layer子节点显示在下拉框中
            XmlNodeList layerNodes = xmlDoc.SelectNodes("//layers/layer");
            if (null != layerNodes)
            {
                foreach (XmlNode layerNode in layerNodes)
                {
                    //通过Attributes获得属性名为name的属性
                    typeCmb.Items.Add(layerNode.Attributes["name"].Value);
                }
            }

            compareCmb.Items.Add("---请选择对照表---");
            compareCmb.SelectedIndex = 0;
            InitCompare("");

            XmlNodeList compareNodes = xmlDoc.SelectNodes("//compare/table");
            if (null != compareNodes)
            {
                foreach (XmlNode compareNode in compareNodes)
                {
                    //通过Attributes获得属性名为name的属性
                    compareCmb.Items.Add(compareNode.Attributes["name"].Value);
                }
            }

            // 3. 读取excel所需要附加的额外的列
            XmlNodeList pointColumn = xmlDoc.SelectNodes("//defaultColumn/point/column");
            if (null != pointColumn)
            {
                mPointTableColumn = new List<string>();
                foreach (XmlNode colNode in pointColumn)
                {
                    mPointTableColumn.Add(colNode.Attributes["name"].Value);
                }
            }
            XmlNodeList lineColumn = xmlDoc.SelectNodes("//defaultColumn/line/column");
            if (null != lineColumn)
            {
                mLineTableColumn = new List<string>();
                foreach (XmlNode colNode in lineColumn)
                {
                    mLineTableColumn.Add(colNode.Attributes["name"].Value);
                }
            }
            XmlNodeList textColumn = xmlDoc.SelectNodes("//defaultColumn/note/column");
            if (null != textColumn)
            {
                mNoteTableColumn = new List<string>();
                foreach (XmlNode colNode in textColumn)
                {
                    mNoteTableColumn.Add(colNode.Attributes["name"].Value);
                }
            }
        }

        /// <summary>
        /// 获取这个动态链接库的位置
        /// </summary>
        /// <returns></returns>
        private string GetSefPath()
        {
            string strDllPath = this.GetType().Assembly.CodeBase;
            int start = 8;                 // 去除file:///  
            int end = strDllPath.LastIndexOf('/');// 去除文件名xxx.dll及文件名前的/  
            strDllPath = strDllPath.Substring(start, end - start);
            return strDllPath;
        }

        /// <summary>
        /// 读取实体对照表的数据
        /// </summary>
        private void InitCompare(string strComName)
        {
            if (null == mCadReader)
            {
                mCadReader = new CxCadReader();
            }

            mCadReader.ClearCompare();

            if (strComName == "")
            {
                return;
            }

            XmlDocument xmlDoc = new XmlDocument();
            try
            {
                xmlDoc.Load(strComName);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "读取错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 使用xpath表达式选择block对照信息存放在Hash中
            XmlNodeList blockNodes = xmlDoc.SelectNodes("//blocks/block");
            if (null != blockNodes)
            {
                foreach (XmlNode blockNode in blockNodes)
                {
                    mCadReader.AddCompare(blockNode.Attributes["id"].Value, blockNode.Attributes["name"].Value);
                }
            }
        }

        /// <summary>
        /// 用于打开文件对话框，选择导出Except数据的存放位置
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void selectBtn_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDlg = new SaveFileDialog();
            saveDlg.Filter = "Excel 工作簿|*.xls;*.xlsx|所有文件|*.*";
            saveDlg.FilterIndex = 0;
            saveDlg.RestoreDirectory = true;
            string strPath = "";
            strPath = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Name;
            saveDlg.FileName = System.IO.Path.GetFileNameWithoutExtension(strPath); ;

            if (saveDlg.ShowDialog() == DialogResult.OK)
            {
                exportTxb.Text = saveDlg.FileName;
            }
        }

        /// <summary>
        /// 取得导出文件的路径
        /// </summary>
        /// <returns></returns>
        private string getExportPath()
        {
            string strfileName = exportTxb.Text;

            if (!File.Exists(strfileName))
            {
            }
            return strfileName;
        }

        /// <summary>
        /// 确认导出按钮事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void okBtn_Click(object sender, EventArgs e)
        {
            Export();
        }
        private delegate void cadDelegate();
        private void Export()
        {
            CxExcelWriter excelWriter = new CxExcelWriter();
            List<string> selectedLayers = new List<string>();
            System.Data.DataTable lineTable;
            System.Data.DataTable pointTable;
            System.Data.DataTable txtTable;
            int lineStartRow = 1;    // 默认从第一行开始
            int pointStartRow = 1;
            int textStartRow = 1;
            int startColumn = 1;
            int currIdx = 0;
            string strDiap = "";

            mCadReader.DecimalNumber = Convert.ToInt32(decimalSpin.Value);
            mCadReader.BelongLayer = typeCmb.SelectedIndex <= 0 ? "未知" : typeCmb.SelectedItem.ToString();

            GetSelectedLayers(ref selectedLayers);
            excelWriter.Open(getExportPath());

            statusPgBar.Maximum = selectedLayers.Count;

            foreach (string strLayer in selectedLayers)
            {
                strDiap = string.Format("共 {0} 个图层，当前正在导出第{1}个（{2}）", selectedLayers.Count, ++currIdx, strLayer);
                excelWriter.ProgressBar = statusPgBar;
                mCadReader.GetLayerEntyties(strLayer);
                lineTable = mCadReader.GetLineTable();
                pointTable = mCadReader.GetPointTable();
                txtTable = mCadReader.GetTextTable();

                statusPgBar.Value = 0;
                statusPgBar.Maximum = lineTable.Rows.Count + pointTable.Rows.Count + txtTable.Rows.Count;
                statusLabel.Text = strDiap + "[线表]";
                excelWriter.DefaultColumn = mLineTableColumn;
                excelWriter.InsertTable(ref lineTable, "线表", ref lineStartRow, startColumn);
                statusLabel.Text = strDiap + "[点表]";
                excelWriter.DefaultColumn = mPointTableColumn;
                excelWriter.InsertTable(ref pointTable, "点表", ref pointStartRow, startColumn);
                statusLabel.Text = strDiap + "[注记表]";
                excelWriter.DefaultColumn = mNoteTableColumn;
                excelWriter.InsertTable(ref txtTable, "注记表", ref textStartRow, startColumn);

                lineStartRow += lineTable.Rows.Count;
                pointStartRow += pointTable.Rows.Count;
                textStartRow += txtTable.Rows.Count;
            }

            excelWriter.Save();
            excelWriter.Close();

            MessageBox.Show("文件已经成功导出到Excel表格", "导出成功", MessageBoxButtons.OK, MessageBoxIcon.None);
            Close();
        }

        /// <summary>
        /// 取消按钮事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void cancelBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void compareCmb_SelectedIndexChanged(object sender, EventArgs e)
        {
            string strTableName = compareCmb.SelectedIndex <= 0 ? "" : compareCmb.SelectedItem.ToString();
            if (strTableName != "")
            {
                strTableName += ".xml";
                InitCompare(strTableName);
            }
        }
    }
}
