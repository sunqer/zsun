using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Collections;

using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;

namespace CxCadPlug
{
    /// <summary>
    /// 类名：CxCadReader
    /// 说明：本类用于从当前激活的cad视图中获取实体的属性
    /// 单位： ZBCX
    /// 作者： sun
    /// 日期： 2016-09-24
    /// </summary>
    class CxCadReader
    {
        private System.Data.DataTable mPointTable;     // 实体中点对象表记录
        private System.Data.DataTable mLineTable;      // 线对象记录表
        private System.Data.DataTable mTextTable;      // 文本对象集合
        private Hashtable mCompareList;                // 块实体对照表
        private Transaction mAcTrans;

        public const string STARTPOS = "起始点";
        public const string ENDPOS = "终止点";
        public const string POSITION = "坐标";
        public const string BLOCKNAME = "块名称";
        public const string ATTRIBUTE = "属性";
        public const string COLOR = "颜色";
        public const string TXTCONTEXT = "文字内容";
        public const string TXTHEIGHT = "文字高度";
        public const string TXTROTATE = "文字转角";
        public const string LAYER = "所属管网";

        // 定义属性，设置所属图层名称
        public string BelongLayer { get; set; }
        // 属性 设置和获取小数位数
        public int DecimalNumber { get; set; }
        // 属性 一个图层的实体总数
        public int SelectedEntityCount { get; set; }
        // 属性 需要导出的实体个数
        public int ExportEntityCount { get; set; }

        public CxCadReader()
        {
            mPointTable = null;
            mLineTable = null;
            mTextTable = null;
            mCompareList = new Hashtable();
        }

        public System.Data.DataTable GetPointTable()
        {
            if (null == mPointTable)
            {
                mPointTable = new System.Data.DataTable("pointTable");
                mPointTable.Columns.Add(POSITION, typeof(string));     // 位置(x,y,z)
                mPointTable.Columns.Add(BLOCKNAME, typeof(string));    // 块名称
                mPointTable.Columns.Add(ATTRIBUTE, typeof(string));    // 特性信息
                mPointTable.Columns.Add(LAYER, typeof(string));        // 图层
            }

            return mPointTable;
        }

        public System.Data.DataTable GetLineTable()
        {
            if (null == mLineTable)
            {
                mLineTable = new System.Data.DataTable("lineTable");
                mLineTable.Columns.Add(STARTPOS, typeof(string));      // 起始点
                mLineTable.Columns.Add(ENDPOS, typeof(string));        // 终止点
                mLineTable.Columns.Add(POSITION, typeof(string));      // 位置(x,y,z;...)
                mLineTable.Columns.Add(COLOR, typeof(string));         // 颜色
                mLineTable.Columns.Add(ATTRIBUTE, typeof(string));     // 特性信息
                mLineTable.Columns.Add(LAYER, typeof(string));         // 图层
            }

            return mLineTable;
        }

        public System.Data.DataTable GetTextTable()
        {
            if (null == mTextTable)
            {
                mTextTable = new System.Data.DataTable("textTable");
                mTextTable.Columns.Add(TXTCONTEXT, typeof(string));     // 文字内容
                mTextTable.Columns.Add(POSITION, typeof(string));       // 基点位置
                mTextTable.Columns.Add(TXTHEIGHT, typeof(string));      // 文字高度
                mTextTable.Columns.Add(TXTROTATE, typeof(string));      // 文字转角
                mTextTable.Columns.Add(LAYER, typeof(string));          // 图层
            }

            return mTextTable;
        }

        public void AddCompare(string key, string value)
        {
            mCompareList.Add(key, value);
        }

        public void ClearCompare()
        {
            mCompareList.Clear();
        }

        /// <summary>
        /// 得到一个图层对象的所有实体要素
        /// </summary>
        /// <param name="strName">图层名称</param>
        /// <param name="pEntyties">实体对象列表</param>
        public void GetLayerEntyties(string strName)
        {
            // 获取当前激活视图的对象
            Document acDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // 开启事务
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // 查找指定图层名称的图层
                mAcTrans = acTrans;
                LayerTable acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;

                if (acLyrTbl.Has(strName))
                {
                    // 构造一个选择集，加入当前图层的所有对象
                    TypedValue[] values = { new TypedValue((int)DxfCode.LayerName, strName)};
                    SelectionFilter filter = new SelectionFilter(values);
                    PromptSelectionResult selRes = acDoc.Editor.SelectAll(filter);

                    if (selRes.Status == PromptStatus.OK)
                    {
                        DBObject acDbObj;
                        ObjectId[] ids = selRes.Value.GetObjectIds();
                        SelectedEntityCount = ids.Length;
                        ExportEntityCount = 0;

                        Clear();   // 清空上一个图层的数据,重新写入新图层数据

                        foreach (ObjectId acObjId in ids)
                        {
                            acDbObj = acTrans.GetObject(acObjId, OpenMode.ForRead);
                            string s2 = acDbObj.GetType().Name;
                            switch (acDbObj.GetType().Name)
                            {
                                case "DBText":
                                    AppendText(acDbObj as DBText); break;
                                case "MText":
                                    AppendMText(acDbObj as MText); break;
                                case "Polyline":
                                    AppendPolyline(acDbObj as Polyline); break;
                                case "Line":
                                    AppendLine(acDbObj as Line); break;
                                case "DBPoint":
                                    AppendPoint(acDbObj as DBPoint); break;
                                case "BlockReference":
                                    AppendBlock(acDbObj as BlockReference); break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// 清空内存表中的数据记录。
        /// </summary>
        private void Clear()
        {
            if (null != mPointTable)
            {
                mPointTable.Clear();
            }
            if (null != mLineTable)
            {
                mLineTable.Clear();
            }
            if (null != mTextTable)
            {
                mTextTable.Clear();
            }
        }

        /// <summary>
        /// 往要生成的点表中插入一条记录
        /// </summary>
        /// <param name="pointEnt">点实体类型</param>
        private void AppendPoint(DBPoint pointEnt)
        {
            System.Data.DataRow dr = GetPointTable().NewRow();
            dr[POSITION] = PntToString(pointEnt.Position);
            dr[BLOCKNAME] = "Point";
            dr[LAYER] = BelongLayer;
            mPointTable.Rows.Add(dr);
        }

        /// <summary>
        /// 添加块参照对象属性
        /// </summary>
        /// <param name="blockEnt"></param>
        private void AppendBlock(BlockReference blockEnt)
        {
            System.Data.DataRow dr = GetPointTable().NewRow();
            dr[POSITION] = PntToString(blockEnt.Position);

            BlockTableRecord blR = (BlockTableRecord)mAcTrans.GetObject(blockEnt.BlockTableRecord, OpenMode.ForRead);
            dr[BLOCKNAME] = blR.Name;

            if (mCompareList.Contains(blR.Name))
            {
                // 如果在对照表中,则替换为对照配置
                dr[BLOCKNAME] = mCompareList[blR.Name].ToString();
            }

            string strAttribute = "";
            if (blockEnt.AttributeCollection.Count != 0)
            {
                System.Collections.IEnumerator bRefEnum = blockEnt.AttributeCollection.GetEnumerator();
                AttributeReference aRef;
                ObjectId aId;
                while (bRefEnum.MoveNext())
                {
                    aId = (ObjectId)bRefEnum.Current;   // 这一句非常重要
                    aRef = (AttributeReference)mAcTrans.GetObject(aId, OpenMode.ForRead, false, true);

                    strAttribute += aRef.Tag +":"+ aRef.TextString + ";";
                }
            }
            dr[ATTRIBUTE] = strAttribute;
            dr[LAYER] = BelongLayer;
            mPointTable.Rows.Add(dr);
        }

        /// <summary>
        /// 往要生成的线表中插入一条记录
        /// </summary>
        /// <param name="lineEnt">两点线</param>
        private void AppendLine(Line lineEnt)
        {
            System.Data.DataRow dr = GetLineTable().NewRow();
            string strColor = lineEnt.Color.ColorValue.Name;
            string strStart = PntToString(lineEnt.StartPoint);
            string strEnd = PntToString(lineEnt.EndPoint);
            dr[STARTPOS] = strStart;
            dr[ENDPOS] = strEnd;
            dr[POSITION] = strStart + strEnd;
            dr[COLOR] = strColor.Remove(0, 2).Insert(0, "#");
            dr[LAYER] = BelongLayer;
            mLineTable.Rows.Add(dr);
        }

        /// <summary>
        /// 添加一个多线要素
        /// </summary>
        /// <param name="polylineEnt"></param>
        private void AppendPolyline(Polyline polylineEnt)
        {
            string strPos = "";

            for (int idx = 0; idx < polylineEnt.NumberOfVertices; ++idx)
            {
                strPos += PntToString(polylineEnt.GetPoint3dAt(idx));
            }

            string strColor = polylineEnt.Color.ColorValue.Name;

            System.Data.DataRow dr = GetLineTable().NewRow();
            dr[STARTPOS] = PntToString(polylineEnt.StartPoint);
            dr[ENDPOS] = PntToString(polylineEnt.EndPoint);
            dr[POSITION] = strPos;
            dr[COLOR] = strColor.Remove(0, 2).Insert(0, "#");
            dr[LAYER] = BelongLayer;
            mLineTable.Rows.Add(dr);
        }

        /// <summary>
        /// 往要生成的注记表中插入一条记录
        /// </summary>
        /// <param name="txtEnt">简单单行文本</param>
        private void AppendText(DBText txtEnt)
        {
            System.Data.DataRow dr = GetTextTable().NewRow();
            dr[TXTCONTEXT] = txtEnt.TextString;
            dr[POSITION] = PntToString(txtEnt.Position);
            dr[TXTHEIGHT] = txtEnt.Height.ToString();
            dr[TXTROTATE] = txtEnt.Rotation.ToString();
            dr[LAYER] = BelongLayer;
            mTextTable.Rows.Add(dr);

            if (txtEnt.TextStyleName.Contains("@"))
            {
                // 如果是带有@前导的字符，说明是竖着放的
            }
        }

        /// <summary>
        /// 插入一个换行文本对象
        /// </summary>
        /// <param name="mTxtEnt"></param>
        private void AppendMText(MText mTxtEnt)
        {
            System.Data.DataRow dr = GetTextTable().NewRow();
            dr[TXTCONTEXT] = mTxtEnt.Text;
            dr[POSITION] = PntToString(mTxtEnt.Location);
            dr[TXTHEIGHT] = mTxtEnt.TextHeight.ToString();
            dr[TXTROTATE] = mTxtEnt.Rotation.ToString();
            dr[LAYER] = BelongLayer;
            mTextTable.Rows.Add(dr);
        }

        /// <summary>
        /// 修改坐标字符串的格式
        /// </summary>
        /// <param name="strPos"></param>
        private string PntToString(Point3d pnt3d)
        {
            string strNum = string.Format("f{0}", DecimalNumber);
            string strCoord = pnt3d.X.ToString(strNum);
            strCoord += "," + pnt3d.Y.ToString(strNum) + ",";
            strCoord += pnt3d.Z != 0.0 ? pnt3d.Z.ToString(strNum) : pnt3d.Z.ToString();
            strCoord += ";";
            return strCoord;
        }
    }
}
