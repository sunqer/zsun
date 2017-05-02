using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Windows.ToolPalette;
using Autodesk.AutoCAD.Windows;

namespace CxCadPlug
{
    public class CxCadPlug : Autodesk.AutoCAD.Runtime.IExtensionApplication
    {
        ContextMenuExtension m_pContextMenu;  // 右键扩展菜单

        /// <summary>
        /// 实现继承父类的虚函数, 实现插件的初始化功能,NETLOAD之后会自动加载该方法
        /// </summary>
        public void Initialize()
        {
            AddContextMenu();

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage("CX 导出EXCEL插件加载完成。");
        }

        /// <summary>
        /// 卸载方法
        /// </summary>
        public void Terminate()
        {
            RemoveContextMenu();
        }

        /// <summary>
        /// 新建添加一个右键菜单项
        /// </summary>
        private void AddContextMenu()
        {
            m_pContextMenu = new ContextMenuExtension();
            m_pContextMenu.Title = "导出";

            Autodesk.AutoCAD.Windows.MenuItem pMenuItem;
            pMenuItem = new Autodesk.AutoCAD.Windows.MenuItem("导出到Excel");

            // 关联菜单的时间处理函数
            pMenuItem.Click += CxExportMenu_OnClick;
            m_pContextMenu.MenuItems.Add(pMenuItem);
            Application.AddDefaultContextMenuExtension(m_pContextMenu);
        }

        /// <summary>
        /// 移除右键菜单
        /// </summary>
        private void RemoveContextMenu()
        {
            if (null != m_pContextMenu)
            {
                Application.RemoveDefaultContextMenuExtension(m_pContextMenu);
                m_pContextMenu = null;
            }
        }

        /// <summary>
        /// 右键菜单的实现，弹出一个用户界面让选择导出数据的格式
        /// </summary>
        /// <param name="o">消息发送对象</param>
        /// <param name="e">事件相关参数</param>
        private void CxExportMenu_OnClick(object o, EventArgs e)
        {
            export2ExcelFrm export2ExcelFrm = new export2ExcelFrm();
            Autodesk.AutoCAD.ApplicationServices.Application.ShowModalDialog(export2ExcelFrm);
        }
    }
}
