using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using Microsoft.Office.Core;

namespace Herramientas
{
    public partial class ThisAddIn
    {

        public static Office.IRibbonUI e_ribbon;
        public static Selection e_seleccion;
        public static int e_count;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            
            // Valida si hubo cmabios a la agenda
            //this.Application.WindowSelectionChange += new EApplication_WindowSelectionChangeEventHandler(commandBars_OnUpdate);
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new barra_matrix();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {

        }

        #region Código generado por VSTO

        /// <summary>
        /// Método necesario para admitir el Diseñador. No se puede modificar
        /// el contenido del método con el editor de código.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion

        private void Application_WindowSelectionChange(Selection Sel)
        {
            try
            {
                e_ribbon.Invalidate();

                var s = Sel.Type;
                if (Sel.Type == PpSelectionType.ppSelectionShapes
                    && Sel.ShapeRange[1].Type != MsoShapeType.msoPlaceholder)
                {

                    e_count = Sel.ShapeRange.Count;
                    e_seleccion = Sel;
                }
                else
                {
                    e_count = 0;
                    e_seleccion = null;
                }

            }
            catch (Exception)
            {

                throw;
            }
        }

        /*
        /// <summary>
        /// 
        /// </summary>
        public void commandBars_OnUpdate(Selection Sel)
        {

                e_ribbon.Invalidate();
                barra_matrix.validarEventoCambioAgenda();
            
        }
        */

    }
}
