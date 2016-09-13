using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Globalization;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Diagnostics;
using System.Drawing;
using Herramientas.Clases;

#region Instrucciones Llamada y ejecucion XML
// TODO:  Siga estos pasos para habilitar el elemento (XML) de la cinta de opciones:

// 1: Copie el siguiente bloque de código en la clase ThisAddin, ThisWorkbook o ThisDocument.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new barra_matrix();
//  }

// 2. Cree métodos de devolución de llamada en el área "Devolución de llamadas de la cinta de opciones" de esta clase para controlar acciones del usuario,
//    como hacer clic en un botón. Nota: si ha exportado esta cinta de opciones desde el diseñador de la cinta de opciones,
//    mueva el código de los controladores de eventos a los métodos de devolución de llamada y modifique el código para que funcione con el
//    modelo de programación de extensibilidad de la cinta de opciones (RibbonX).

// 3. Asigne atributos a las etiquetas de control del archivo XML de la cinta de opciones para identificar los métodos de devolución de llamada apropiados en el código.  

// Para obtener más información, vea la documentación XML de la cinta de opciones en la Ayuda de Visual Studio Tools para Office.

#endregion

namespace Herramientas
{
    [ComVisible(true)]
    public class barra_matrix : Office.IRibbonExtensibility
    {
        /*Constructor*/
        public barra_matrix()
        {

        }
        
        #region Declaraciones y Load

        // Referencia a la entidad que representa la barra de botones
        public Office.IRibbonUI ribbon;
        
        // Relación a instancia de la clase métodos
        private static Clases.metodos csmet = Clases.metodos.Instance;

        // Referencia a la presentación y a la slide activa
        PowerPoint.Presentation pptActiva;
        PowerPoint.Slide slideActiva;

        // Ancho y alto de la slide
        private int ancho;
        private int alto;

        //Obtiene la imagen de la clase metodos (csmet).
        public Bitmap GetImage(Office.IRibbonControl control)
        {
            return csmet.getImage(control.Id);
        }

        #endregion

        #region Miembros de IRibbonExtensibility

        // Inicializa el ribbon a partir del archivo xml
        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("Herramientas.barra_matrix.xml");
        }

        #endregion

        #region Devoluciones de llamada de la cinta de opciones
        //Cree aquí métodos de devolución de llamada. Para obtener más información sobre los métodos de devolución de llamada, visite http://go.microsoft.com/fwlink/?LinkID=271226.

        // Evento se ejecuta al cargar el ribbon
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
            ThisAddIn.e_ribbon = ribbonUI;
        }

        //Este método orquesta la lógica de todos los botones de la barra
        public void btnGeneral(Office.IRibbonControl control)
        {
            //Identificador y valor del bot[on (control)
            string valor = string.Empty;
            string ctrl = control.Id;

            // Valida los easterEggs asociados a botones
            validarEasterEgg(control);

            // LOG - siempre lleva un nombre fijo + un timestamp (contador) para evitar archivos repetidos
            string path = csmet.CONFIG["LOGFILE_PATH"] + csmet.CONFIG["LOGFILE_FILE"] + csmet.CONFIG["LOGFILE_COUNTER"] + ".txt";
            FileInfo f = new FileInfo(path);
            long s1 = f.Length;
            
            // Si el log supera cierto tamaño se divide en múltiples archivos
            if (s1 > 1024000)
            {
                // Actualiza el contador
                string counterAnt = csmet.CONFIG["LOGFILE_COUNTER"];
                csmet.CONFIG.Remove("LOGFILE_COUNTER");
                csmet.CONFIG.Add("LOGFILE_COUNTER", DateTime.Now.ToString("yyyyMMddHHmm"));

                try
                {
                    // Guarda la configuración
                    csmet.guardarConfig();
                }
                catch (Exception e)
                {
                    // En caso de error al guardar la configuración sigue usando el log anterior
                    csmet.CONFIG.Remove("LOGFILE_COUNTER");
                    csmet.CONFIG.Add("LOGFILE_COUNTER", counterAnt);
                }

                path = csmet.CONFIG["LOGFILE_PATH"] + csmet.CONFIG["LOGFILE_FILE"] + csmet.CONFIG["LOGFILE_COUNTER"] + ".txt";
            }

            // Si el archivo de log no existe crea un nuevo archivo
            if (!File.Exists(path))
            {
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("USER;DATE;CONTROL;VALOR");
                }
            }

            // Guarda información del control utilizado en formato csv
            using (StreamWriter sw = File.AppendText(path))
            {
                sw.WriteLine(Environment.MachineName + ";" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ";" + ctrl + ";" + valor);
            }

            // Ejecuta la acción según cada id
            switch (control.Id)
            {
                case "btnAlto":
                case "btnAncho":
                    string opc = control.Tag;
                    ajustaTamano(opc);
                    flagBoton(1); // Flag botón se usa para determinar el easter egg de usar todos los controles al menos una vez.
                    break;
                case "btnTake":
                    insertTakeaway();
                    flagBoton(10);
                    break;
                case "btnOp1":
                case "btnOp2":
                case "btnOp3":
                case "btnOp4":
                case "btnOp5":
                    valor = control.Tag;
                    string[] nVal = valor.Split('-');
                    insertCuadroTextoBullet(nVal[0], nVal[1], nVal[2]);
                    flagBoton(100);
                    break;
                case "btnD_pd_D":
                case "btnD_Pr_D":
                case "btnD_Il_D":
                case "btnD_EJ_D":
                case "btnD_Cn_D":
                case "btnD_Bk_D":
                case "btnD_Ne_D":
                case "btnD_Ot_D":
                case "btnD_pd_I":
                case "btnD_Pr_I":
                case "btnD_Il_I":
                case "btnD_EJ_I":
                case "btnD_Cn_I":
                case "btnD_Bk_I":
                case "btnD_Ne_I":
                case "btnD_Ot_I":
                    valor = control.Id;
                    insertDisclaimer(valor);
                    flagBoton(1000);
                    break;
                case "btnGhostSimple":
                    insertGhostSimple();
                    flagBoton(10000);
                    break;

                case "btnSmartGhost":
                    insertGhostSmart();
                    flagBoton(100000);
                    break;

                case "btnDarformato":
                    darFormatoBulletMatrix();
                    flagBoton(1000000);
                    break;
                case "btnConfig":
                    ConfigForm configMenu = new ConfigForm();
                    configMenu.ShowDialog();
                    break;
                case "btnNotaAlPie":
                    insertNotaAlPie();
                    flagBoton(10000000);
                    break;
                case "btnCopyPos":
                    copyPosiciones();
                    flagBoton(100000000);
                    break;
                case "btnPastePos":
                    pastePosiciones();
                    flagBoton(1000000000);
                    break;
                case "btnEmpalmaV":
                    empalmarVertical();
                    flagBoton(10000000000);
                    break;
                case "btnEmpalmaH":
                    empalmarHorizontal();
                    flagBoton(100000000000);
                    break;
                case "btnDestacar":
                    destacar();
                    flagBoton(1000000000000);
                    break;
                case "btnSombrear":
                    sombrear();
                    flagBoton(10000000000000);
                    break;
                case "btnCrearAgenda":
                    crearAgenda();
                    flagBoton(100000000000000);
                    break;
                case "btnEliminarAgenda":
                    eliminarAgenda(DialogResult.Yes);
                    flagBoton(100000000000000);
                    break;
                case "btnActualizarAgenda":
                    //actualizarAgenda();
                    flagBoton(100000000000000);
                    break;
                case "btnFraseDia":
                    fraseDelDia();
                    flagBoton(1000000000000000);
                    break;
                case "btnLinkDia":
                    linkDelDia();
                    flagBoton(1000000000000000);
                    break;
                case "btnAddEspaciadoH":
                    ajustarEspaciado(1, 1);
                    flagBoton(10000000000000000);
                    break;
                case "btnSubEspaciadoH":
                    ajustarEspaciado(1, 2);
                    flagBoton(10000000000000000);
                    break;
                case "btnAddEspaciadoV":
                    ajustarEspaciado(2, 1);
                    flagBoton(10000000000000000);
                    break;
                case "btnSubEspaciadoV":
                    ajustarEspaciado(2, 2);
                    flagBoton(10000000000000000);
                    break;
                case "btnSeleccionarSim":
                    seleccionarSimilares();
                    flagBoton(100000000000000000);
                    break;
                case "btncn3":
                    insertarCirculosNumerados(3, false);
                    flagBoton(1000000000000000000);
                    break;
                case "btncn10":
                    insertarCirculosNumerados(10, false);
                    flagBoton(1000000000000000000);
                    break;
                case "btncniii":
                    insertarCirculosNumerados(3, true);
                    flagBoton(1000000000000000000);
                    break;
                case "btncnx":
                    insertarCirculosNumerados(10,true);
                    flagBoton(1000000000000000000);
                    break;
                case "btnCajas2":
                    insertCajas(2);
                    flagBoton(1000000000000000000);
                    break;
                case "btnCajas3":
                    insertCajas(3);
                    flagBoton(1000000000000000000);
                    break;
                case "btnCajas4":
                    insertCajas(4);
                    flagBoton(1000000000000000000);
                    break;
                case "btnCallout":
                    insertarCallout();
                    break;
                case "btnTitulo":
                    insertTitulo();
                    break;
                case "btnIgual":
                    insertEstados(1);
                    break;
                case "btnCheckBlanco":
                    insertEstados(2);
                    break;
                case "btnCheck":
                    insertEstados(3);
                    break;
                case "btnExclamacion":
                    insertEstados(4);
                    break;
                case "btnCruz":
                    insertEstados(5);
                    break;
                case "btnCalendario4":
                    DateTime primerDia = csmet.ShowDialogNum("Insertar la fecha del primer lunes que se mostrará", "Calendario");
                    if(primerDia != DateTime.MinValue)
                        insertCalendario(4, primerDia);
                    break;
                case "btnCalendario5":
                    primerDia = csmet.ShowDialogNum("Insertar el día del primer lunes que se mostrará", "Calendario");
                    insertCalendario(5, primerDia);
                    break;
                case "btnChangePos":
                    cambiarPosiciones();
                    break;
                case "btnAbrirTemplate":
                    abrirTemplate();
                    break;
                case "btnMostrarHistoria":
                    mostrarHistoria();
                    break;
                case "btnTrofeos":
                    verTrofeos();
                    break;
            }
        }

        #endregion

        #region Easter eggs

        // Contadores para easter eggs que se activan al repetir acciones
        private int contadorEmpalma = 0;
        private int contadorTakeaway = 0;

        // Valida easter eggs asociados a utilizar los distintos botones (controles)
        public void validarEasterEgg(Office.IRibbonControl control)
        {
            string ctrl = control.Id;
            if ((csmet.CONFIG["Alineacion_integ1"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")) && ctrl.StartsWith("btnEmpalma"))
            {
                contadorEmpalma++;
                if (contadorEmpalma == 3)
                {
                    GeekEncontrado ge = new GeekEncontrado();
                    ge.inicializar("Nicolás Palma", "integ1", "Nicolás se unió al Equipo Extra Geek en 2015, liderando el frente Mekko y la relación con David Goldstein.", "presionar Empalmar tres veces seguidas", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                    ge.ShowDialog();
                    csmet.CONFIG.Remove("Alineacion_integ1");
                    csmet.CONFIG.Add("Alineacion_integ1", "1");
                    csmet.guardarConfig();
                    contadorEmpalma = 0;
                }
            }
            else if ((csmet.CONFIG["Alineacion_integ9"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")) && ctrl.StartsWith("btnTake"))
            {
                contadorTakeaway++;
                if (contadorTakeaway == 2)
                {
                    GeekEncontrado ge = new GeekEncontrado();
                    ge.inicializar("Jaime Siles", "integ9", "Jaime ha sido el sponsor natural del Equipo Extra Geek desde sus inicios, aportando valiosos inputs para el desarrollo de los proyectos.", "presionar Takeaway dos veces seguidas", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                    ge.ShowDialog();
                    csmet.CONFIG.Remove("Alineacion_integ9");
                    csmet.CONFIG.Add("Alineacion_integ9", "1");
                    csmet.guardarConfig();
                    contadorTakeaway = 0;
                }
            }
            else
            {
                contadorEmpalma = 0;
                contadorTakeaway = 0;
            }


        }

        // Valida si se usaron todos los botones
        public void flagBoton(long id)
        {
            if ((csmet.CONFIG["Alineacion_integ12"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
            {
                string usados = csmet.CONFIG["Botones_usados"];

                string actual = id + "";
                int pos = actual.Length - 1;

                if(usados.Substring(pos,1).Equals("0"))
                {
                    usados = usados.Substring(0, pos) + "1" + usados.Substring(pos + 1);
                    csmet.CONFIG.Remove("Botones_usados");
                    csmet.CONFIG.Add("Botones_usados",usados);
                    csmet.guardarConfig();
                }

                if (usados.Equals("1111111111111111111"))
                {
                    GeekEncontrado ge = new GeekEncontrado();
                    ge.inicializar("Gabriel Adriazola", "integ12", "Gabriel se unió al Equipo Extra Geek en 2015. Desde entonces, ha sido un continuo proveedor de Tips de la Semana.", "usar todas las funciones al menos una vez", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                    ge.ShowDialog();
                    csmet.CONFIG.Remove("Alineacion_integ12");
                    csmet.CONFIG.Add("Alineacion_integ12", "1");
                    csmet.guardarConfig();
                }
            }
        }

        // Inicializa el Form de trofeos
        public void verTrofeos()
        {
            Trofeos t = new Trofeos();
            t.inicializar();
            t.ShowDialog();
        }

        #endregion

        #region Metodos y Funciones Barra de Herramientas Matrix.

        #region Frase, link del día, template y story

        // Muestra la frase del día
        public void fraseDelDia()
        {

            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            slideActiva = csmet.getDiapositiva(this.pptActiva);
            int dia = DateTime.Now.DayOfYear % int.Parse(csmet.CONFIG["FRASEDELDIA_COUNT"]);
            string frase = csmet.darFrase(dia);
            MessageBox.Show(frase);

            string ultimavez = csmet.CONFIG["Ult_fdd"];
            long numFrases = long.Parse(csmet.CONFIG["num_fdd"]);
            if(!ultimavez.Equals(DateTime.Now.ToString("yyyyMMdd")))
            {
                numFrases++;
                csmet.CONFIG.Remove("Ult_fdd");
                csmet.CONFIG.Add("Ult_fdd", DateTime.Now.ToString("yyyyMMdd"));
                csmet.CONFIG.Remove("num_fdd");
                csmet.CONFIG.Add("num_fdd", numFrases + "");
                csmet.guardarConfig();

                if ((csmet.CONFIG["Alineacion_integ11"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")) && numFrases >= 5)
                {
                    GeekEncontrado ge = new GeekEncontrado();
                        ge.inicializar("Agustín Dagnino", "integ11", "Agustín se unió al Equipo Extra Geek en 2016 tras resolver exitosamente el Recruiting Brainteaser de las 12 bolas y la balanza.", "abrir la Frase del Día en 5 días diferentes", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                        ge.ShowDialog();
                        csmet.CONFIG.Remove("Alineacion_integ11");
                        csmet.CONFIG.Add("Alineacion_integ11", "1");
                        csmet.guardarConfig();
                    
                }
            }
        }

        // Muestra el link del día
        public void linkDelDia()
        {

            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            slideActiva = csmet.getDiapositiva(this.pptActiva);
            int dia = DateTime.Now.DayOfYear % int.Parse(csmet.CONFIG["LINKDELDIA_COUNT"]);
            string frase = csmet.darLink(dia);

            LinkMessageBox lmb = new LinkMessageBox();
            lmb.inicializar(frase, "Link del día");
            lmb.ShowDialog();


            string ultimavez = csmet.CONFIG["Ult_ldd"];
            long numFrases = long.Parse(csmet.CONFIG["num_ldd"]);
            if (!ultimavez.Equals(DateTime.Now.ToString("yyyyMMdd")))
            {
                numFrases++;
                csmet.CONFIG.Remove("Ult_ldd");
                csmet.CONFIG.Add("Ult_ldd", DateTime.Now.ToString("yyyyMMdd"));
                csmet.CONFIG.Remove("num_ldd");
                csmet.CONFIG.Add("num_ldd", numFrases + "");
                csmet.guardarConfig();

                if ((csmet.CONFIG["Alineacion_integ13"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")) && numFrases >= 5)
                {
                    GeekEncontrado ge = new GeekEncontrado();
                    ge.inicializar("Andrés Oksemberg", "integ13", "Andrés se unió al Equipo Extra Geek en 2016 tras resolver exitosamente el Recruiting Brainteaser de las 12 bolas y la balanza.", "abrir el Link del Día en 5 días diferentes", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                    ge.ShowDialog();
                    csmet.CONFIG.Remove("Alineacion_integ13");
                    csmet.CONFIG.Add("Alineacion_integ13", "1");
                    csmet.guardarConfig();

                }
            }
        }

        // Abre la presentación del template
        public void abrirTemplate()
        {
            var app = new PowerPoint.Application();
            var pres = app.Presentations;
            try
            {
                var file = pres.Open(@"H:\04 Template\01 Nuevo Template MatrixConsulting v1.0.pptx", MsoTriState.msoTrue, MsoTriState.msoTrue);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(""+e);
            }
        }

        // Muestra la historia creada por los títulos de las diapositivas
        public void mostrarHistoria()
        {
            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            string historia = "";
            int id = 0;

            foreach (Slide sl in pptActiva.Slides)
            {
                if(sl.SlideShowTransition.Hidden.Equals(MsoTriState.msoFalse))
                {
                    id++;
                    try
                    {
                        historia = historia + "\n" + id + "- " + sl.Shapes.Title.TextFrame.TextRange.Text;
                    }
                    catch (Exception e)
                    { }
                }
            }
            csmet.CustomMessageBox(historia,"Títulos de la presentación");
        }

        #endregion

        #region Agenda

        // Código de agenda comentado para futura referencia
        
        // Atributos de apoyo para administrar la agenda
        /*
        public static string LABEL_ITEM_AGENDA = "SECCION_";
        public static string LABEL_TITULO_AGENDA = "TITULO_AGENDA_";
        public static string LABEL_MARCADOR_AGENDA = "MARCADOR_AGENDA_";
        public static string LABEL_AGENDA = "Agenda_";
        private static List<AgendaClase> agendas = null;
        public static Slide agenda = null;
        private static IList<Agenda> agendaForm = null;
        private static PowerPoint.Shape tituloAgenda = null;
        private static Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape> itemsAgenda = new Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape>();
        private static Dictionary<string, float[]> posItemsAgenda = new Dictionary<string, float[]>();
        private static bool horaItemsAgenda;
        private static int numItems;
        private static bool warningSubAgenda = false;
        private static bool agendaCargada = false;
        private static Random rnd = new Random();
        private static bool tienePortada = true;
        private static PowerPoint.Shape linea = null;
        private static int idAgenda = 0;
        private static int cantidadAgendas = 0;
         */

        private static List<AgendaClase> agendas = new List<AgendaClase> { };
        private static int idAgenda = 0;
        private static int cantidadAgendas = 0;

        /*
        // Cambia el título de la agenda
        public void actualizarTituloAgenda(String nTitulo)
        {
            try
            {
                agenda.Shapes.Title.TextFrame.TextRange.Text = nTitulo;
            }
            catch(Exception e)
            {
                tituloAgenda.TextFrame.TextRange.Text = nTitulo;
            }

            if (nTitulo.ToLower().StartsWith("objetivos de la reunión") && (csmet.CONFIG["Alineacion_integ8"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
            {
                GeekEncontrado ge = new GeekEncontrado();
                ge.inicializar("Alex Valenzuela", "integ8", "Alex se unió al Equipo Extra Geek en 2015. Impulsó el Tip de la Semana y participó en el comité de desarrollo de la Barra Matrix.", "crear Agenda con título \"Objetivos de la reunión\"", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                ge.ShowDialog();
                csmet.CONFIG.Remove("Alineacion_integ8");
                csmet.CONFIG.Add("Alineacion_integ8", "1");
                csmet.guardarConfig();
            }

            pptActiva.Application.StartNewUndoEntry();
        }

        // Identifica si existe una slide con el nombre de la agenda base (LABEL_AGENDA + "0") y crea la referencia.     
        private static bool cargarAgenda()
        {
            try {
                if (!agendaCargada && agenda == null)
                {
                    Presentation pptActiva = Globals.ThisAddIn.Application.ActivePresentation;
                    foreach (Slide sl in pptActiva.Slides)
                        if (sl.Name.Equals(LABEL_AGENDA + "0"))
                        {
                            agenda = sl;
                            foreach (PowerPoint.Shape sh in agenda.Shapes)
                            {
                                if (sh.Name.StartsWith(LABEL_ITEM_AGENDA))
                                {
                                    itemsAgenda.Add(sh.Name, sh);
                                    float[] pos = { sh.Top, sh.Left, sh.Width, sh.Height };
                                    posItemsAgenda.Add(sh.Name, pos);
                                    numItems++;
                                }
                                else if (sh.Name.StartsWith(LABEL_TITULO_AGENDA))
                                {
                                    tituloAgenda = sh;
                                }
                                else if (sh.Name.Equals(LABEL_MARCADOR_AGENDA))
                                {
                                    tienePortada = false;
                                }

                            }
                        }

                    agendaCargada = true;
                }

                bool res = false;
                if(agenda != null)
                {
                    try
                    {
                        int idAgenda = agenda.SlideIndex;
                        res = true;
                    }
                    catch (Exception e) {

                        // Aquí en caso de error se eliminaba la agenda - ver código comentado.
                        MessageBox.Show("Hubo un error al cargar la agenda.");
                        
                        DialogResult respuesta = MessageBox.Show("Al parecer se ha borrado la diapositiva base de la agenda.\n¿Quieres eliminar el resto de las diapositivas de agenda?.", "Eliminar agenda", MessageBoxButtons.YesNo);
                        eliminarAgenda(respuesta);
                    }
                }

                return res;

            }catch(Exception e)
            {
                return false;
            }
        }

        // Valida los cambios a las diapositivas de agenda
        public static void validarEventoCambioAgenda()
        {
            Presentation pptActiva = Globals.ThisAddIn.Application.ActivePresentation;
            Slide slideActiva = null;

            try { slideActiva = pptActiva.Slides[pptActiva.Application.ActiveWindow.Selection.SlideRange.SlideNumber]; } catch (Exception e) { }

            if (slideActiva != null && slideActiva.Name.StartsWith(barra_matrix.LABEL_AGENDA))
            {
                bool cambioSlide = false;
                
                if (!slideActiva.Name.Equals(barra_matrix.LABEL_AGENDA + "0"))
                {
                    for (int i = 0; i < numItems; i++)
                    {
                        Microsoft.Office.Interop.PowerPoint.Shape sh = null;
                        try { sh = slideActiva.Shapes[LABEL_ITEM_AGENDA + (i + 1)];
                        } catch (Exception e) {
                            cambioSlide = true;
                        }
                        
                        if (sh != null && sh.Top != posItemsAgenda[LABEL_ITEM_AGENDA + (i + 1)][0])
                            cambioSlide = true;
                        if (sh != null && sh.Left != posItemsAgenda[LABEL_ITEM_AGENDA + (i + 1)][1])
                            cambioSlide = true;
                        if (sh != null && sh.Width != posItemsAgenda[LABEL_ITEM_AGENDA + (i + 1)][2])
                            cambioSlide = true;
                        if (sh != null && sh.Height != posItemsAgenda[LABEL_ITEM_AGENDA + (i + 1)][3])
                            cambioSlide = true;
                    }

                    if (cambioSlide && !warningSubAgenda && !slideActiva.Name.Equals(LABEL_AGENDA + "0"))
                    {
                        System.Windows.Forms.MessageBox.Show("No debes modificar las diapositivas de una sección. Has click en \"Crear/Actualizar Agenda\" para modificar la agenda principal y luego has click en el botón \"Actualizar agenda\" para hacer efectivos los cambios.");
                        warningSubAgenda = true;
                    }
                }
                else
                {

                    for (int i = 0; i < numItems; i++)
                    {
                        Microsoft.Office.Interop.PowerPoint.Shape sh = null;
                        try
                        {
                            sh = slideActiva.Shapes[LABEL_ITEM_AGENDA + (i + 1)];
                        }
                        catch (Exception e)
                        {
                            cambioSlide = true;
                            eliminarItemAgenda(i+1,false);
                        }
                    }
                }
            }
        }

        // Elimina las diapositivas de agenda
        public static void eliminarAgenda(DialogResult respuesta)
        {
            Presentation pptActiva = Globals.ThisAddIn.Application.ActivePresentation;
            Object thisLock = new Object();

            agenda = null;
            itemsAgenda = new Dictionary<string, PowerPoint.Shape>();
            posItemsAgenda = new Dictionary<string, float[]>();
            tituloAgenda = null;
            numItems = 0;

            for (int i = 1; i <= pptActiva.Slides.Count;)
            {
                Slide sl = pptActiva.Slides[i];
                bool cambiarNombre = false;
                bool eliminado = false;
                if (sl.Name.StartsWith(LABEL_AGENDA))
                {
                    if (respuesta.Equals(DialogResult.Yes))
                    {
                        sl.Delete();
                        eliminado = true;
                    }
                    else
                    {
                        cambiarNombre = true;
                    }
                }
                if (cambiarNombre)
                {
                    sl.Name = "noName_"+i;
                }
                if (!eliminado)
                {
                    i++;
                }
            }

            pptActiva.Application.StartNewUndoEntry();
        }

        // Crea la diapositiva base con sus elementos por default
        public void XcrearAgenda()
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                if (!cargarAgenda())
                {
                    //Crea una nueva agenda
                    bool creada = false;
                    try
                    {
                        CustomLayout estilo = pptActiva.SlideMaster.CustomLayouts[Int32.Parse(csmet.CONFIG["Agenda_slide_style_1"])];
                        barra_matrix.agenda = pptActiva.Slides.AddSlide(slideActiva.SlideIndex + Int32.Parse(csmet.CONFIG["Agenda_slide_insertAfter"]), estilo);
                        creada = true;
                    }
                    catch (Exception e)
                    {
                        try
                        {
                            CustomLayout estilo = pptActiva.SlideMaster.CustomLayouts[Int32.Parse(csmet.CONFIG["Agenda_slide_style_2"])];
                            agenda = pptActiva.Slides.AddSlide(slideActiva.SlideIndex + Int32.Parse(csmet.CONFIG["Agenda_slide_insertAfter"]), estilo);
                            creada = true;
                        }
                        catch (Exception e2)
                        {
                            try
                            {
                                CustomLayout estilo = pptActiva.SlideMaster.CustomLayouts[Int32.Parse(csmet.CONFIG["Agenda_slide_style_3"])];
                                agenda = pptActiva.Slides.AddSlide(slideActiva.SlideIndex + Int32.Parse(csmet.CONFIG["Agenda_slide_insertAfter"]), estilo);
                                creada = true;
                            }
                            catch (Exception e3)
                            {
                            }
                        }
                    }


                    if (!creada)
                    {
                        MessageBox.Show("No fue posible crear la agenda. Contacte al administrador.\nError: El formato de la diapositiva no existe.");
                        return;
                    }

                    agenda.Name = LABEL_AGENDA + idAgenda;
                    idAgenda++;
                    cantidadAgendas++;
                    agenda.Select();

                    // Añade Titulo de Agenda
                    string strTituloAgenda = csmet.CONFIG["Agenda_title"];
                    horaItemsAgenda = csmet.CONFIG["Agenda_hora"].Equals("YES");
                    try
                    {
                        agenda.Shapes.Title.Name = LABEL_TITULO_AGENDA;
                        agenda.Shapes.Title.TextFrame.TextRange.Text = strTituloAgenda;
                    }
                    catch (Exception e2)
                    {
                        // Cuadro titulo
                        float topTitle = float.Parse(csmet.CONFIG["Agenda_title_top"], CultureInfo.InvariantCulture);
                        float leftTitle = float.Parse(csmet.CONFIG["Agenda_title_left"], CultureInfo.InvariantCulture);
                        float widthTitle = float.Parse(csmet.CONFIG["Agenda_title_width"], CultureInfo.InvariantCulture);
                        float heightTitle = float.Parse(csmet.CONFIG["Agenda_title_height"], CultureInfo.InvariantCulture);
                        float marginTitleTop = float.Parse(csmet.CONFIG["Agenda_title_margin_top"], CultureInfo.InvariantCulture);
                        float marginTitleBottom = float.Parse(csmet.CONFIG["Agenda_title_margin_bottom"], CultureInfo.InvariantCulture);
                        float marginTitleLeft = float.Parse(csmet.CONFIG["Agenda_title_margin_left"], CultureInfo.InvariantCulture);
                        float marginTitleRight = float.Parse(csmet.CONFIG["Agenda_title_margin_right"], CultureInfo.InvariantCulture);
                        tituloAgenda = agenda.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftTitle, topTitle, widthTitle, heightTitle);
                        tituloAgenda.TextFrame.TextRange.Font.Size = 16;
                    }

                    
                    //OptionPane oP = new OptionPane();
                    //oP.inicializar("¿Deseas incluir la hora de cada ítem en la agenda?","Incluir hora en agenda");
                    //oP.ShowDialog();
                    //horaItemsAgenda = oP.darResultado();
                    

                    agendaForm = new Agenda();
                    agendaForm.inicializar(this, agenda, strTituloAgenda);

                    // Añade linea que une pelotas
                    linea = agenda.Shapes.AddLine(csmet.centimetersToPoints(2.45), csmet.centimetersToPoints(5.36), csmet.centimetersToPoints(2.45), csmet.centimetersToPoints(5.36)); //Última medida es el alto de las pelotas
                    linea.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("gris"));
                    linea.Line.Weight = float.Parse("5");
                    linea.ZOrder(MsoZOrderCmd.msoSendToBack);
                    linea.Name = "lineaTemplate";

                    // Añade los ítems de la agenda
                    for (int i = 0; i < Int32.Parse(csmet.CONFIG["Agenda_numItems"]); i++)
                    {
                        agregarItemAgenda(true);
                    }

                    tienePortada = csmet.CONFIG["Agenda_portada"].Equals("YES");

                    if (!tienePortada)
                    {

                        Microsoft.Office.Interop.PowerPoint.Shape sel = agenda.Shapes[LABEL_ITEM_AGENDA + "1"];

                        Microsoft.Office.Interop.PowerPoint.ShapeRange modificar = sel.GroupItems.Range("item1" + "_1");
                        modificar.TextFrame.TextRange.Font.Bold = MsoTriState.msoCTrue;

                        modificar = sel.GroupItems.Range("item1" + "_3");
                        modificar.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("rojo"));
                        modificar.ZOrder(MsoZOrderCmd.msoBringToFront);

                    }
                    agendaCargada = true;
                }
                else
                    agenda.Select();
                
                agendaForm = new Agenda();
                agendaForm.inicializar(this, agenda, csmet.CONFIG["Agenda_title"]);
                
                foreach (Microsoft.Office.Interop.PowerPoint.Shape sh in agenda.Shapes)
                {
                    if (sh.Name.StartsWith(LABEL_ITEM_AGENDA))
                    {
                        agendaForm.agregarItem();
                    }
                }
                agendaForm.Show();

                pptActiva.Application.StartNewUndoEntry();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Hubo un problema al crear la agenda:\n"+"("+ex.Message+")");
            }

        }
        */

        public void crearAgenda()
        {
            if(cantidadAgendas == 0)
            {
                idAgenda++;
                cantidadAgendas++;
                try
                {
                    AgendaClase nueva = new AgendaClase(this, Globals.ThisAddIn.Application.ActivePresentation, idAgenda);
                    agendas.Add(nueva);
                }
                catch (Exception e)
                {
                    idAgenda--;
                    cantidadAgendas--;
                }
            }
            else
            {
                SeleccionAgenda seleccion = new SeleccionAgenda(agendas);
                seleccion.Show();
                string agendaSeleccionada = "nada";
                seleccion.Disposed += (sender, e) =>
                {
                    try
                    {
                        agendaSeleccionada = seleccion.seleccionado;
                        if (agendaSeleccionada.Equals("Nueva"))
                        {
                            idAgenda++;
                            cantidadAgendas++;
                            try
                            {
                                AgendaClase nueva = new AgendaClase(this, Globals.ThisAddIn.Application.ActivePresentation, idAgenda);
                                agendas.Add(nueva);
                            }
                            catch (Exception e1)
                            {
                                idAgenda--;
                                cantidadAgendas--;
                            }
                        }
                        else
                        {
                            foreach (AgendaClase a in agendas)
                            {
                                if (a.titulo.Equals(agendaSeleccionada))
                                {
                                    a.crearAgenda();
                                    break;
                                }
                            }
                        }
                    }
                    catch (Exception e1) { }
                };

            }
        }

        public void eliminarAgenda(DialogResult resultado)
        {
            SeleccionAgenda seleccion = new SeleccionAgenda(agendas, "Eliminar agenda");
            seleccion.Show();
            string agendaSeleccionada = "nada";
            seleccion.Disposed += (sender, e) =>
            {
                try
                {
                    agendaSeleccionada = seleccion.seleccionado;
                    foreach (AgendaClase a in agendas)
                    {
                        if (a.titulo.Equals(agendaSeleccionada))
                        {
                            DialogResult confirmacion = MessageBox.Show("Está seguro que quiere borrar la agenda \"" + agendaSeleccionada + "\"?", "Eliminar agenda", MessageBoxButtons.OKCancel);
                            if (confirmacion.Equals(DialogResult.OK))
                            {
                                cantidadAgendas--;
                                agendas.Remove(a);
                                a.eliminarAgenda(resultado);
                            }
                        }
                    }
                }
                catch (Exception e1) { }
            };
        }

        /*
        // Agrega un ítem a la agenda creando el objeto
        public void agregarItemAgenda(bool creacion)
        {

            Microsoft.Office.Interop.PowerPoint.Shape grp = null;

            // Añade ítem 1
            float topTitle = float.Parse(csmet.CONFIG["Agenda_item_top"], CultureInfo.InvariantCulture);
            float heightTitle = float.Parse(csmet.CONFIG["Agenda_item_height"], CultureInfo.InvariantCulture);
            float marginTitle = float.Parse(csmet.CONFIG["Agenda_item_margin"], CultureInfo.InvariantCulture);
            float widthTitle = float.Parse(csmet.CONFIG["Agenda_item_width"], CultureInfo.InvariantCulture);
            float leftTitle = float.Parse(csmet.CONFIG["Agenda_item_left"], CultureInfo.InvariantCulture);
            int index = itemsAgenda.Count + 1;
            Microsoft.Office.Interop.PowerPoint.Shape item = null;
            item = agenda.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, leftTitle, topTitle + (index - 1) * float.Parse(csmet.CONFIG["Agenda_item_gapToNext"], CultureInfo.InvariantCulture), widthTitle, heightTitle);
            item.Name = "item" + index + "_1";
            item.TextFrame.TextRange.Text = "Ítem " + index;
            item.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
            item.TextFrame.TextRange.Font.Size = float.Parse(csmet.CONFIG["Agenda_item_fontSize"], CultureInfo.InvariantCulture);
            item.TextFrame.TextRange.Font.Name = csmet.CONFIG["MainFont"];
            item.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft; //alineacion texto centro
            item.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = PpBaselineAlignment.ppBaselineAlignCenter;

            //item.TextFrame2.TextRange.ParagraphFormat.LeftIndent = float.Parse(csmet.CONFIG["Agenda_item_leftIndent"], CultureInfo.InvariantCulture);
            //item.TextFrame2.TextRange.ParagraphFormat.FirstLineIndent = float.Parse(csmet.CONFIG["Agenda_item_leftIndentFirstLine"], CultureInfo.InvariantCulture);

            heightTitle = float.Parse(csmet.CONFIG["Agenda_item_hora_height"], CultureInfo.InvariantCulture);
            marginTitle = float.Parse(csmet.CONFIG["Agenda_item_hora_margin"], CultureInfo.InvariantCulture);
            widthTitle = float.Parse(csmet.CONFIG["Agenda_item_hora_width"], CultureInfo.InvariantCulture);
            leftTitle = float.Parse(csmet.CONFIG["Agenda_item_hora_left"], CultureInfo.InvariantCulture);

            // Diseño de la agenda
            float heightPelota = float.Parse(csmet.CONFIG["Agenda_pelota_height"], CultureInfo.InvariantCulture);
            float widthPelota = float.Parse(csmet.CONFIG["Agenda_pelota_width"], CultureInfo.InvariantCulture);
            float leftPelota = float.Parse(csmet.CONFIG["Agenda_pelota_left"], CultureInfo.InvariantCulture);
            PowerPoint.Shape pelota = agenda.Shapes.AddShape(MsoAutoShapeType.msoShapeOval, leftPelota, topTitle + (index - 1) * float.Parse(csmet.CONFIG["Agenda_item_gapToNext"], CultureInfo.InvariantCulture), widthPelota, heightPelota);
            pelota.Line.Visible = MsoTriState.msoFalse;
            pelota.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("grisSombra"));
            pelota.Name = "item" + index + "_3";

            // Agrandar linea
            if (index > 1 && index <= 5)
                linea.Height = linea.Height + 80;

            Array arrayItem = new string[] { "item" + index + "_1", "item" + index + "_3" };

            if (horaItemsAgenda)
            {
                Microsoft.Office.Interop.PowerPoint.Shape itemHora = agenda.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, leftTitle, topTitle + index * float.Parse(csmet.CONFIG["Agenda_item_gapToNext"], CultureInfo.InvariantCulture), widthTitle, heightTitle);
                itemHora.TextFrame.TextRange.Text = "00:00";
                itemHora.Name = "item" + index + "_2";
                itemHora.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
                itemHora.TextFrame.TextRange.Font.Size = float.Parse(csmet.CONFIG["Agenda_item_hora_fontSize"], CultureInfo.InvariantCulture);
                itemHora.TextFrame.TextRange.Font.Name = csmet.CONFIG["MainFont"];
                itemHora.TextFrame.TextRange.ParagraphFormat.Bullet.Type = PpBulletType.ppBulletNone;
                itemHora.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft; //alineacion texto centro
                itemHora.TextFrame.MarginRight = marginTitle;
                itemHora.TextFrame.MarginLeft = marginTitle;

                arrayItem = new string[] { "item" + index + "_1", "item" + index + "_2", "item" + index + "_3" };
            }

            grp = agenda.Shapes.Range(arrayItem).Group();

            agregarItemAgenda(grp, -1);

            pptActiva.Application.StartNewUndoEntry();
        }

        // Agrega un item a la agenda dada la figura (shape) que lo representa
        public void agregarItemAgenda(Microsoft.Office.Interop.PowerPoint.Shape nuevoItem, int index)
        {
            if (index == -1)
                index = itemsAgenda.Count + 1;

            for (int i = itemsAgenda.Count; i >= index; i--)
            {
                Microsoft.Office.Interop.PowerPoint.Shape item = itemsAgenda[LABEL_ITEM_AGENDA + i];
                MessageBox.Show(item.Name);
                item.Name = LABEL_ITEM_AGENDA + (i + 1);
                MessageBox.Show("Ahora " + item.Name);
            }

            nuevoItem.Name = LABEL_ITEM_AGENDA + (index);
            itemsAgenda.Add(LABEL_ITEM_AGENDA + (index), nuevoItem);
            float[] posItem = { nuevoItem.Top, nuevoItem.Left, nuevoItem.Width, nuevoItem.Height };
            numItems++;
            posItemsAgenda.Add(LABEL_ITEM_AGENDA + (index), posItem);
            agendaForm.agregarItem();

            // Distribuir items en espacio asignado a la agenda
            if (numItems > 5)
            {
                float gap_aux = (4 * float.Parse(csmet.CONFIG["Agenda_item_gapToNext"], CultureInfo.InvariantCulture)) / (index-1);
                for (int i = 2; i <= index; i++)
                {
                    Microsoft.Office.Interop.PowerPoint.Shape modificar = agenda.Shapes[LABEL_ITEM_AGENDA + i];
                    modificar.Top = float.Parse(csmet.CONFIG["Agenda_item_top"], CultureInfo.InvariantCulture) + (i-1) * gap_aux;
                }
            }
        }

        // Elimina un ítem de la agenda
        public static void eliminarItemAgenda(int index, bool eliminarShape)
        {
            Presentation pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            if (index == 1)
            {
                System.Windows.Forms.MessageBox.Show("No se puede eliminar el primer item");
            }
            else
            {
                Microsoft.Office.Interop.PowerPoint.Shape item = itemsAgenda[LABEL_ITEM_AGENDA + index];
                if (eliminarShape)
                {
                    item.Delete();
                }
                itemsAgenda.Remove(LABEL_ITEM_AGENDA + (index));
                posItemsAgenda.Remove(LABEL_ITEM_AGENDA + (index));

                //Acortar linea
                if(numItems <= 5)
                    linea.Height = linea.Height - 80;

                try
                {
                    //Cambia el id de la diapositiva correspondiente si esta existe
                    Slide slideAgenda = pptActiva.Slides[LABEL_AGENDA + index];
                    slideAgenda.Name = LABEL_AGENDA + "eliminar" + rnd.Next(999);
                }
                catch (Exception e)
                { }

                for (int i = index + 1; i <= itemsAgenda.Count + 1; i++)
                {
                    item = itemsAgenda[LABEL_ITEM_AGENDA + i];
                    item.Name = LABEL_ITEM_AGENDA + (i - 1);
                    foreach (Microsoft.Office.Interop.PowerPoint.Shape modificar in item.GroupItems)
                    {
                        int largo = modificar.Name.Length;
                        string antiguo = "" + i;
                        string nuevo = "" + (i - 1);
                        modificar.Name = modificar.Name.Substring(0, modificar.Name.Length-1).Replace(antiguo, nuevo) + modificar.Name.Substring(modificar.Name.Length-1);
                    }
                    itemsAgenda.Remove(LABEL_ITEM_AGENDA + i);
                    itemsAgenda.Add(LABEL_ITEM_AGENDA + (i - 1), item);

                    try
                    {
                        //Cambia el id de la diapositiva correspondiente si esta existe
                        Slide slideAgenda = pptActiva.Slides[LABEL_AGENDA + (i)];
                        slideAgenda.Name = LABEL_AGENDA + (i - 1);
                    }
                    catch (Exception e)
                    { }

                    float topTitle = float.Parse(csmet.CONFIG["Agenda_item_top"], CultureInfo.InvariantCulture);
                    float gapToNext = float.Parse(csmet.CONFIG["Agenda_item_gapToNext"], CultureInfo.InvariantCulture);
                    item.Top = topTitle + ((i - 1) - 1) * gapToNext;

                    float[] pos = posItemsAgenda[LABEL_ITEM_AGENDA + i];
                    pos[0] = item.Top;
                    posItemsAgenda.Remove(LABEL_ITEM_AGENDA + i);
                    posItemsAgenda.Add(LABEL_ITEM_AGENDA + (i - 1), pos);
                }

                numItems--;

                //Redistribuir items restantes
                if (numItems > 4)
                {
                    float gap_aux = (4 * float.Parse(csmet.CONFIG["Agenda_item_gapToNext"], CultureInfo.InvariantCulture)) / (numItems - 1);
                    for (int i = 2; i <= numItems; i++)
                    {
                        Microsoft.Office.Interop.PowerPoint.Shape modificar = agenda.Shapes[LABEL_ITEM_AGENDA + i];
                        modificar.Top = float.Parse(csmet.CONFIG["Agenda_item_top"], CultureInfo.InvariantCulture) + (i-1) * gap_aux;
                    }
                }

                if (agendaForm != null)
                    agendaForm.eliminarItem();

                pptActiva.Application.StartNewUndoEntry();
            }
        }

        // Actualiza todas las diapositivas de agenda según la agenda base.
        public void actualizarAgenda()
        {
            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;
            
            int indexBase = agenda.SlideIndex;
            int[] indexPosiciones = new int[itemsAgenda.Count];
            int maxIndex = indexBase + 1;

            for (int i = (tienePortada?1:2); i <= itemsAgenda.Count; i++)
            {
                bool eliminada = false;
                foreach (Slide sl in pptActiva.Slides)
                {
                    if (sl.Name.Equals(LABEL_AGENDA + i) && !eliminada)
                    {
                        indexPosiciones[i - 1] = sl.SlideIndex;
                        if (sl.SlideIndex >= maxIndex)
                            maxIndex = sl.SlideIndex + 1;

                        sl.Delete();
                        eliminada = true;
                    }
                }

                if (!eliminada)
                {
                    indexPosiciones[i - 1] = maxIndex;
                    maxIndex++;
                }


                Slide newSlide = pptActiva.Slides[pptActiva.Slides[indexBase].Duplicate().SlideIndex];
                newSlide.Name = LABEL_AGENDA + (i);

                marcarItem(newSlide, i);

                newSlide.MoveTo(indexPosiciones[i-1]);
            }

            foreach (Slide sl in pptActiva.Slides)
            {
                if (sl.Name.StartsWith(LABEL_AGENDA + "eliminar"))
                {
                    sl.Delete();
                }
            }
            

        }

        public void marcarItem(Slide slide, int numero)
        {

            Microsoft.Office.Interop.PowerPoint.Shape selAnt = slide.Shapes[LABEL_ITEM_AGENDA + 1];
            Microsoft.Office.Interop.PowerPoint.ShapeRange modificarAnt = selAnt.GroupItems.Range("item" + 1 + "_1");
            modificarAnt.TextFrame.TextRange.Font.Bold = MsoTriState.msoFalse;

            modificarAnt = selAnt.GroupItems.Range("item" + 1 + "_3");
            modificarAnt.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("grisSombra"));

            selAnt.ZOrder(MsoZOrderCmd.msoBringToFront);

            Microsoft.Office.Interop.PowerPoint.Shape sel = slide.Shapes[LABEL_ITEM_AGENDA + numero];

            Microsoft.Office.Interop.PowerPoint.ShapeRange modificar = sel.GroupItems.Range("item" + numero + "_1");
            modificar.TextFrame.TextRange.Font.Bold = MsoTriState.msoCTrue;

            modificar = sel.GroupItems.Range("item" + numero + "_3");
            modificar.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("rojo"));

            sel.ZOrder(MsoZOrderCmd.msoBringToFront);
        }
        */
        #endregion

        #region Destacar y sombrear
            
        // Agrega un shape rectangular con contorno de color a las figuras seleccionadas y una convención para escribir el foco.
        public void destacar()
        {

            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    float left = 99999;
                    float top = 99999;
                    float right = 0;
                    float bottom = 0;

                    foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                    {
                        if (sel.Left < left)
                            left = sel.Left;

                        if (sel.Top < top)
                            top = sel.Top;

                        if (sel.Left + sel.Width > right)
                            right = sel.Left + sel.Width;

                        if (sel.Top + sel.Height > bottom)
                            bottom = sel.Top + sel.Height;
                    }


                    PowerPoint.Shape box = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left - 5, top - 5, right - left + 10, bottom - top + 10);
                    box.Fill.Visible = MsoTriState.msoFalse;
                    box.Line.DashStyle = MsoLineDashStyle.msoLineDash;
                    box.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Highlight_lineColor"]));
                    box.Line.Weight = float.Parse(csmet.CONFIG["Highlight_lineWidth"], CultureInfo.InvariantCulture);


                    PowerPoint.Shape label = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, float.Parse(csmet.CONFIG["Highlight_label_left"],CultureInfo.InvariantCulture), float.Parse(csmet.CONFIG["Highlight_label_top"], CultureInfo.InvariantCulture), float.Parse(csmet.CONFIG["Highlight_label_width"], CultureInfo.InvariantCulture), float.Parse(csmet.CONFIG["Highlight_label_height"], CultureInfo.InvariantCulture));
                    label.Fill.Visible = MsoTriState.msoFalse;
                    label.Line.DashStyle = MsoLineDashStyle.msoLineDash;
                    label.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Highlight_lineColor"]));
                    label.Line.Weight = float.Parse(csmet.CONFIG["Highlight_lineWidth"], CultureInfo.InvariantCulture);
                    label.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Highlight_background"]));
                    label.Fill.Transparency = float.Parse("0.50", CultureInfo.InvariantCulture); // Pasar a metodos

                    label.TextFrame.TextRange.Text = "Foco";
                    label.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    label.TextFrame.MarginBottom = 0;
                    label.TextFrame.MarginRight = 0;
                    label.TextFrame.MarginTop = 0;
                    label.TextFrame.MarginLeft = float.Parse(csmet.CONFIG["Highlight_text_marginLeft"], CultureInfo.InvariantCulture);
                    label.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoFalse;
                    label.TextFrame.TextRange.Font.Size = float.Parse(csmet.CONFIG["Highlight_text_fontSize"], CultureInfo.InvariantCulture);
                    label.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Highlight_text_fontColor"]));
                    label.TextFrame.TextRange.Font.Name = csmet.CONFIG["MainFont"];
                    label.TextFrame.WordWrap = MsoTriState.msoFalse;


                    label.TextFrame.TextRange.Select();


                }
                else
                {
                    MessageBox.Show("Debe seleccionar al menos un objeto", "Información");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        // Agrega un shape rectangular sobre las figuras seleccionadas
        public void sombrear()
        {

            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    float left = 99999;
                    float top = 99999;
                    float right = 0;
                    float bottom = 0;

                    foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                    {
                        if (sel.Left < left)
                            left = sel.Left;

                        if (sel.Top < top)
                            top = sel.Top;

                        if (sel.Left + sel.Width > right)
                            right = sel.Left + sel.Width;

                        if (sel.Top + sel.Height > bottom)
                            bottom = sel.Top + sel.Height;
                    }


                    PowerPoint.Shape shadow = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, left - 5, top - 5, right - left + 10, bottom - top + 10);
                    shadow.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(Color.White);
                    shadow.Fill.Transparency = float.Parse("0.30", CultureInfo.InvariantCulture);
                    shadow.Line.Visible = MsoTriState.msoFalse;

                }
                else
                {
                    MessageBox.Show("Debe seleccionar al menos un objeto", "Información");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        #endregion

        #region Insertar Elementos

        // Arreglo estático para referenciar los números romanos rápidamente
        private static string[] numerosRomanos = { "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X" };

        // Inserta círculos numerados sea con números arabigos o romanos
        public void insertarCirculosNumerados(int num, bool romanos)
        {
            if (romanos && num > numerosRomanos.Length) {
                MessageBox.Show("Ha seleccionado una cantidad muy alta");
                return;
            }

            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            slideActiva = csmet.getDiapositiva(this.pptActiva);

            // Dimensiones de los círculos
            float width = float.Parse(csmet.CONFIG["CN_width"], CultureInfo.InvariantCulture);
            float height = float.Parse(csmet.CONFIG["CN_height"], CultureInfo.InvariantCulture);
            float top = float.Parse(csmet.CONFIG["CN_top"], CultureInfo.InvariantCulture);
            float left = float.Parse(csmet.CONFIG["CN_left"], CultureInfo.InvariantCulture);
            float gap = float.Parse(csmet.CONFIG["CN_gap"], CultureInfo.InvariantCulture);
            
            for (int i = 1; i <= num; i++)
            {
                // Inserta el círuclo y edita el texto
                PowerPoint.Shape cn = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeOval, left, top + i*gap, width, height);
                cn.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["CN_background"]));
                cn.Line.Style = Office.MsoLineStyle.msoLineSingle;
                cn.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["CN_lineColor"]));
                cn.Line.Weight = float.Parse(csmet.CONFIG["CN_lineWidth"], CultureInfo.InvariantCulture);

                cn.TextFrame.TextRange.Paragraphs().ParagraphFormat.Alignment = (PowerPoint.PpParagraphAlignment)Int32.Parse(csmet.CONFIG["CN_aligment"]);
                cn.TextFrame.TextRange.Paragraphs().Font.Size = float.Parse(csmet.CONFIG["CN_Font_size"], CultureInfo.InvariantCulture);
                cn.TextFrame.TextRange.Paragraphs().Font.Bold = (Office.MsoTriState)Int32.Parse(csmet.CONFIG["CN_Font_bold"]);
                cn.TextFrame.TextRange.Paragraphs().Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["CN_foreColor"]));
                cn.TextFrame.TextRange.Paragraphs().Font.Name = csmet.CONFIG["MainFont"];
                cn.TextFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
                cn.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                cn.TextFrame.WordWrap = MsoTriState.msoFalse;

                if (romanos)
                    cn.TextFrame.TextRange.Text = numerosRomanos[(i - 1)];
                else
                    cn.TextFrame.TextRange.Text = i + "";


                cn.Select((i == 1?MsoTriState.msoTrue:MsoTriState.msoFalse));
            }
            
            pptActiva.Application.StartNewUndoEntry();

        }

        // Inserta el takeaway
        public void insertTakeaway()
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                /*Tamaño Diapositiva*/
                ancho = (int)slideActiva.CustomLayout.Width;
                alto = (int)slideActiva.CustomLayout.Height;

                if (!csmet.getValidaGrupoShape(slideActiva, "Grupo Conclusion"))
                {
                    /*Linea superior */
                    float widthFormaLinea = float.Parse(csmet.CONFIG["Takeaway_linea_width"], CultureInfo.InvariantCulture);
                    float heightFormaLinea = float.Parse(csmet.CONFIG["Takeaway_linea_height"], CultureInfo.InvariantCulture);
                    float topLinea = float.Parse(csmet.CONFIG["Takeaway_linea_top"], CultureInfo.InvariantCulture);
                    float leftLinea = float.Parse(csmet.CONFIG["Takeaway_linea_left"], CultureInfo.InvariantCulture);

                    /*Añade Linea superior*/
                    PowerPoint.Shape takeL = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftLinea, topLinea, widthFormaLinea, heightFormaLinea);
                    takeL.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Takeaway_linea_background"]));
                    takeL.Name = "linea";
                    takeL.Line.Visible = MsoTriState.msoFalse;

                    /*Rectangulo superior */
                    float widthFormaRectangulo = float.Parse(csmet.CONFIG["Takeaway_rectangulo_width"], CultureInfo.InvariantCulture);
                    float heightFormaRectangulo = float.Parse(csmet.CONFIG["Takeaway_rectangulo_height"], CultureInfo.InvariantCulture);
                    float topRectangulo = float.Parse(csmet.CONFIG["Takeaway_rectangulo_top"], CultureInfo.InvariantCulture);
                    float leftRectangulo = float.Parse(csmet.CONFIG["Takeaway_rectangulo_left"], CultureInfo.InvariantCulture);
                    
                    /*Añade Rectangulo superior*/
                    PowerPoint.Shape takeR = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo, topRectangulo, widthFormaRectangulo, heightFormaRectangulo);
                    takeR.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Takeaway_rectangulo_background"]));
                    takeR.Name = "rectangulo";
                    takeR.Line.Visible = MsoTriState.msoFalse;

                    /*Rectangulo inferior */
                    float widthFormaCuadrado = float.Parse(csmet.CONFIG["Takeaway_cuadrado_width"], CultureInfo.InvariantCulture);
                    float heightFormaCuadrado = float.Parse(csmet.CONFIG["Takeaway_cuadrado_height"], CultureInfo.InvariantCulture);
                    float topCuadrado = float.Parse(csmet.CONFIG["Takeaway_cuadrado_top"], CultureInfo.InvariantCulture);
                    float leftCuadrado = float.Parse(csmet.CONFIG["Takeaway_cuadrado_left"], CultureInfo.InvariantCulture);

                    /*Añade Rectangulo inferior*/
                    PowerPoint.Shape takeC = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftCuadrado, topCuadrado, widthFormaCuadrado, heightFormaCuadrado);
                    takeC.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Takeaway_cuadrado_background"]));
                    takeC.Name = "cuadrado";
                    takeC.Line.Visible = MsoTriState.msoFalse;

                    /*Cuadro Conclusion*/
                    float topConclu = float.Parse(csmet.CONFIG["Takeaway_text_top"], CultureInfo.InvariantCulture);
                    float leftConclu = float.Parse(csmet.CONFIG["Takeaway_text_left"], CultureInfo.InvariantCulture);
                    float widthConclu = float.Parse(csmet.CONFIG["Takeaway_text_width"], CultureInfo.InvariantCulture);
                    float heightConclu = float.Parse(csmet.CONFIG["Takeaway_text_height"], CultureInfo.InvariantCulture);
                    float marginConclu = float.Parse(csmet.CONFIG["Takeaway_text_margin"], CultureInfo.InvariantCulture);

                    /*Añade Cuadro de Conclusión*/
                    PowerPoint.Shape take = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftConclu, topConclu, widthConclu, heightConclu);
                    take.Fill.Visible = MsoTriState.msoFalse;
                    take.Line.Visible = MsoTriState.msoFalse;
                    take.Name = "conclu";
                    take.TextFrame.TextRange.Text = "Conclusión";
                    take.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Takeaway_text_foreColor"]));
                    take.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                    take.TextFrame.TextRange.Font.Size = float.Parse(csmet.CONFIG["Takeaway_text_fontSize"], CultureInfo.InvariantCulture);
                    take.TextFrame.TextRange.Font.Name = csmet.CONFIG["MainFont"];
                    take.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = PpBaselineAlignment.ppBaselineAlignCenter;
                    take.TextFrame.TextRange.ParagraphFormat.Bullet.Type = PpBulletType.ppBulletNone;
                    take.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft; //alineacion texto centro
                    take.TextFrame.MarginRight = marginConclu;
                    take.TextFrame.MarginLeft = marginConclu;
                    take.TextFrame.MarginTop = 0; // Pasar a diccionario
                    take.TextFrame.MarginBottom = 0; // Pasar a diccionario

                    /*Agrupacion*/
                    string[] ArrayTake = new string[] { "linea", "rectangulo", "cuadrado", "conclu" };
                    PowerPoint.Shape grpo = slideActiva.Shapes.Range(ArrayTake).Group();
                    grpo.Name = "Grupo Conclusion";
                    grpo.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

                    /*Propiedades Grupo*/
                    grpo.Top = float.Parse(csmet.CONFIG["Takeaway_linea_top"], CultureInfo.InvariantCulture);
                    grpo.Left = float.Parse(csmet.CONFIG["Takeaway_linea_left"], CultureInfo.InvariantCulture);

                    /* Animación */
                    if (csmet.CONFIG["Takeaway_animate"].Equals("YES"))
                        slideActiva.TimeLine.MainSequence.AddEffect(grpo, MsoAnimEffect.msoAnimEffectAppear);


                    take.Select();
                    take.TextFrame.TextRange.Select();
                    pptActiva.Application.StartNewUndoEntry();

                }
                else
                {
                    MessageBox.Show("Ya existe un Takeaway en esta lámina.", "Insertar Takeaway");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        // Inserta cajas
        public void insertCajas(int num)
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                /*Tamaño Diapositiva*/
                ancho = (int)slideActiva.CustomLayout.Width;
                alto = (int)slideActiva.CustomLayout.Height;

                /*Cuadro Texto 1*/
                float topTexto1 = float.Parse(csmet.CONFIG["Cajas_texto1_top"], CultureInfo.InvariantCulture);
                float topTexto2 = float.Parse(csmet.CONFIG["Cajas_texto2_top"], CultureInfo.InvariantCulture);
                float leftTexto = float.Parse(csmet.CONFIG["Cajas_texto_left"], CultureInfo.InvariantCulture);
                float widthTexto = float.Parse(csmet.CONFIG["Cajas_texto_width"], CultureInfo.InvariantCulture);
                float heightTexto = float.Parse(csmet.CONFIG["Cajas_texto_height"], CultureInfo.InvariantCulture);

                /*Añade Cuadro Texto 1*/
                PowerPoint.Shape takeT1 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftTexto, topTexto1, widthTexto, heightTexto);
                takeT1.Fill.Visible = MsoTriState.msoFalse;
                takeT1.Line.Visible = MsoTriState.msoFalse;
                takeT1.Name = "dimension1";
                takeT1.TextFrame.TextRange.Text = "Dimensión 1:";
                takeT1.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_texto_foreColor"]));
                takeT1.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                takeT1.TextFrame.TextRange.Font.Size = float.Parse(csmet.CONFIG["Cajas_texto_fontSize"], CultureInfo.InvariantCulture);
                takeT1.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                takeT1.TextFrame.MarginLeft = 0; //Pasar a diccionario
                takeT1.TextFrame.MarginRight = 0; //Pasar a diccionario


                /*Añade Cuadro Texto 2*/
                PowerPoint.Shape takeT2 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftTexto, topTexto2, widthTexto, heightTexto);
                takeT2.Fill.Visible = MsoTriState.msoFalse;
                takeT2.Line.Visible = MsoTriState.msoFalse;
                takeT2.Name = "dimension2";
                takeT2.TextFrame.TextRange.Text = "Dimensión 2:";
                takeT2.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_texto_foreColor"]));
                takeT2.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                takeT2.TextFrame.TextRange.Font.Size = float.Parse(csmet.CONFIG["Cajas_texto_fontSize"], CultureInfo.InvariantCulture);
                takeT2.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;
                takeT2.TextFrame.MarginLeft = 0; //Pasar a diccionario
                takeT2.TextFrame.MarginRight = 0; //Pasar a diccionario

                /*Cajas */
                float heightRectangulo = float.Parse(csmet.CONFIG["Cajas_height"], CultureInfo.InvariantCulture);
                float topRectangulo = float.Parse(csmet.CONFIG["Cajas_top"], CultureInfo.InvariantCulture);
                float leftRectangulo = float.Parse(csmet.CONFIG["Cajas_left"], CultureInfo.InvariantCulture);

                /*Cuadro de texto de contenido*/
                float heightContenido = float.Parse(csmet.CONFIG["Cajas_contenido_height"], CultureInfo.InvariantCulture);
                float topContenido = float.Parse(csmet.CONFIG["Cajas_texto2_top"], CultureInfo.InvariantCulture);

                /*Formato de cajas de contenido*/
                PowerPoint.Shape modelo = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo, topContenido, heightRectangulo, heightContenido);
                modelo.Fill.Visible = MsoTriState.msoFalse;
                modelo.Line.Visible = MsoTriState.msoFalse;
                modelo.TextFrame.TextRange.Text = "Contenido";
                modelo.TextFrame.TextRange.Font.Size = float.Parse("12", CultureInfo.InvariantCulture);
                modelo.TextFrame.TextRange.ParagraphFormat.Bullet.Type = PpBulletType.ppBulletUnnumbered;
                modelo.TextFrame.TextRange.Font.Bold = MsoTriState.msoFalse;
                modelo.PickUp();
                modelo.Delete();

                switch(num)
                {
                    case 2:
                        {
                            float widthRectangulo = float.Parse(csmet.CONFIG["Cajas2_width"], CultureInfo.InvariantCulture);
                            float spanRectangulo = float.Parse(csmet.CONFIG["Cajas2_span"], CultureInfo.InvariantCulture);

                            /*Añade Rectangulo superior*/
                            PowerPoint.Shape take1 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take1.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take1.TextFrame.TextRange.Text = "Caja 1";
                            take1.Name = "rectangulo1";
                            PowerPoint.Shape take2 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo + widthRectangulo + spanRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take2.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take2.TextFrame.TextRange.Text = "Caja 2";
                            take2.Name = "rectangulo2";

                            /*Añade cuadro de contenido*/
                            PowerPoint.Shape contenido1 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido1.Fill.Visible = MsoTriState.msoFalse;
                            contenido1.Line.Visible = MsoTriState.msoFalse;
                            contenido1.TextFrame.TextRange.Text = "Contenido";
                            contenido1.Apply();
                            PowerPoint.Shape contenido2 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo + widthRectangulo + spanRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido2.Fill.Visible = MsoTriState.msoFalse;
                            contenido2.Line.Visible = MsoTriState.msoFalse;
                            contenido2.TextFrame.TextRange.Text = "Contenido";
                            contenido2.Apply();

                            /*Agrupacion*/
                            string[] ArrayRectangulo2 = new string[] { "rectangulo1", "rectangulo2"};
                            PowerPoint.Shape grpo2 = slideActiva.Shapes.Range(ArrayRectangulo2).Group();
                            grpo2.Name = "Grupo Rectangulo2";
                            grpo2.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                            grpo2.Line.Visible = MsoTriState.msoFalse;
                            grpo2.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_foreColor"]));
                            grpo2.Ungroup();
                            break;
                        }
                    case 3:
                        {
                            float widthRectangulo = float.Parse(csmet.CONFIG["Cajas3_width"], CultureInfo.InvariantCulture);
                            float spanRectangulo = float.Parse(csmet.CONFIG["Cajas3_span"], CultureInfo.InvariantCulture);

                            /*Añade Rectangulo superior*/
                            PowerPoint.Shape take1 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take1.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take1.TextFrame.TextRange.Text = "Caja 1";
                            take1.Name = "rectangulo1";
                            PowerPoint.Shape take2 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo + widthRectangulo + spanRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take2.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take2.TextFrame.TextRange.Text = "Caja 2";
                            take2.Name = "rectangulo2";
                            PowerPoint.Shape take3 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo + 2 * widthRectangulo + 2 * spanRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take3.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take3.TextFrame.TextRange.Text = "Caja 3";
                            take3.Name = "rectangulo3";

                            /*Añade cuadro de contenido*/
                            PowerPoint.Shape contenido1 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido1.Fill.Visible = MsoTriState.msoFalse;
                            contenido1.Line.Visible = MsoTriState.msoFalse;
                            contenido1.TextFrame.TextRange.Text = "Contenido";
                            contenido1.Apply();
                            PowerPoint.Shape contenido2 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo + widthRectangulo + spanRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido2.Fill.Visible = MsoTriState.msoFalse;
                            contenido2.Line.Visible = MsoTriState.msoFalse;
                            contenido2.TextFrame.TextRange.Text = "Contenido";
                            contenido2.Apply();
                            PowerPoint.Shape contenido3 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo + 2* widthRectangulo + 2*spanRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido3.Fill.Visible = MsoTriState.msoFalse;
                            contenido3.Line.Visible = MsoTriState.msoFalse;
                            contenido3.TextFrame.TextRange.Text = "Contenido";
                            contenido3.Apply();

                            /*Agrupacion*/
                            string[] ArrayRectangulo3 = new string[] { "rectangulo1", "rectangulo2", "rectangulo3" };
                            PowerPoint.Shape grpo3 = slideActiva.Shapes.Range(ArrayRectangulo3).Group();
                            grpo3.Name = "Grupo Rectangulo3";
                            grpo3.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                            grpo3.Line.Visible = MsoTriState.msoFalse;
                            grpo3.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_foreColor"]));
                            grpo3.Ungroup();
                            break;
                        }
                    case 4:
                        {
                            float widthRectangulo = float.Parse(csmet.CONFIG["Cajas4_width"], CultureInfo.InvariantCulture);
                            float spanRectangulo = float.Parse(csmet.CONFIG["Cajas4_span"], CultureInfo.InvariantCulture);

                            /*Añade Rectangulo superior*/
                            PowerPoint.Shape take1 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take1.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take1.TextFrame.TextRange.Text = "Caja 1";
                            take1.Name = "rectangulo1";
                            PowerPoint.Shape take2 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo + widthRectangulo + spanRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take2.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take2.TextFrame.TextRange.Text = "Caja 2";
                            take2.Name = "rectangulo2";
                            PowerPoint.Shape take3 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo + 2 * widthRectangulo + 2 * spanRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take3.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take3.TextFrame.TextRange.Text = "Caja 3";
                            take3.Name = "rectangulo3";
                            PowerPoint.Shape take4 = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftRectangulo + 3 * widthRectangulo + 3 * spanRectangulo, topRectangulo, widthRectangulo, heightRectangulo);
                            take4.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_background"]));
                            take4.TextFrame.TextRange.Text = "Caja 4";
                            take4.Name = "rectangulo4";

                            /*Añade cuadro de contenido*/
                            PowerPoint.Shape contenido1 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido1.Fill.Visible = MsoTriState.msoFalse;
                            contenido1.Line.Visible = MsoTriState.msoFalse;
                            contenido1.TextFrame.TextRange.Text = "Contenido";
                            contenido1.Apply();
                            PowerPoint.Shape contenido2 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo + widthRectangulo + spanRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido2.Fill.Visible = MsoTriState.msoFalse;
                            contenido2.Line.Visible = MsoTriState.msoFalse;
                            contenido2.TextFrame.TextRange.Text = "Contenido";
                            contenido2.Apply();
                            PowerPoint.Shape contenido3 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo + 2 * widthRectangulo + 2 * spanRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido3.Fill.Visible = MsoTriState.msoFalse;
                            contenido3.Line.Visible = MsoTriState.msoFalse;
                            contenido3.TextFrame.TextRange.Text = "Contenido";
                            contenido3.Apply();
                            PowerPoint.Shape contenido4 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftRectangulo + 3 * widthRectangulo + 3 * spanRectangulo, topContenido, widthRectangulo, heightContenido);
                            contenido4.Fill.Visible = MsoTriState.msoFalse;
                            contenido4.Line.Visible = MsoTriState.msoFalse;
                            contenido4.TextFrame.TextRange.Text = "Contenido";
                            contenido4.Apply();

                            /*Agrupacion*/
                            string[] ArrayRectangulo4 = new string[] { "rectangulo1", "rectangulo2", "rectangulo3", "rectangulo4" };
                            PowerPoint.Shape grpo4 = slideActiva.Shapes.Range(ArrayRectangulo4).Group();
                            grpo4.Name = "Grupo Rectangulo4";
                            grpo4.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                            grpo4.Line.Visible = MsoTriState.msoFalse;
                            grpo4.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Cajas_foreColor"]));
                            grpo4.Ungroup();
                            break;
                        }

                }

                pptActiva.Application.StartNewUndoEntry();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        // Inserta la nota al pie
        public void insertNotaAlPie()
        {
            try
            {
                PowerPoint.Shape CuadroTexto = null;

                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                float[] sizeCuadroTexto = csmet.getSizeNotaAlPie();
                float top = sizeCuadroTexto[0];
                float left = sizeCuadroTexto[1];
                float wid = sizeCuadroTexto[2];
                float hei = sizeCuadroTexto[3];

                CuadroTexto = slideActiva.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, wid, hei);
                CuadroTexto.TextFrame.TextRange.Text = "Fuente: Experiencia MatrixConsulting";

                //Aplica formato por defecto a todo el cuadro
                CuadroTexto.TextFrame.TextRange.Paragraphs().Font.Name = csmet.CONFIG["MainFont"]; //al utilizar Paragraphs(), indica todas las lineas de texto
                CuadroTexto.TextFrame.TextRange.Paragraphs().ParagraphFormat.Alignment = (PowerPoint.PpParagraphAlignment)Int32.Parse(csmet.CONFIG["NotaAlPie_aligment"]);
                CuadroTexto.TextFrame.TextRange.Paragraphs().Font.Size = float.Parse(csmet.CONFIG["NotaAlPie_Font_size"], CultureInfo.InvariantCulture);
                CuadroTexto.TextFrame.TextRange.Paragraphs().Font.Bold = (Office.MsoTriState)Int32.Parse(csmet.CONFIG["NotaAlPie_Font_bold"]);
                CuadroTexto.TextFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
                CuadroTexto.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorBottom;
                CuadroTexto.Height = hei;
                CuadroTexto.TextFrame.MarginBottom = float.Parse(csmet.CONFIG["NotaAlPie_MarginBottom"], CultureInfo.InvariantCulture);
                CuadroTexto.TextFrame.MarginLeft = float.Parse(csmet.CONFIG["NotaAlPie_MarginLeft"], CultureInfo.InvariantCulture);
                CuadroTexto.TextFrame.MarginRight = float.Parse(csmet.CONFIG["NotaAlPie_MarginRight"], CultureInfo.InvariantCulture);
                CuadroTexto.TextFrame.MarginTop = float.Parse(csmet.CONFIG["NotaAlPie_MarginTop"], CultureInfo.InvariantCulture);

                CuadroTexto.TextFrame.TextRange.Select();
                pptActiva.Application.StartNewUndoEntry();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        // Inserta un callout de conclusión
        public void insertarCallout()
        {
            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            slideActiva = csmet.getDiapositiva(this.pptActiva);

            /*Tamaño Diapositiva*/
            ancho = (int)slideActiva.CustomLayout.Width;
            alto = (int)slideActiva.CustomLayout.Height;

            /*Caja Callout */
            float widthFormaCallout = float.Parse(csmet.CONFIG["Callout_width"], CultureInfo.InvariantCulture);
            float heightFormaCallout = float.Parse(csmet.CONFIG["Callout_height"], CultureInfo.InvariantCulture);
            float topCallout = float.Parse(csmet.CONFIG["Callout_top"], CultureInfo.InvariantCulture);
            float leftCallout = float.Parse(csmet.CONFIG["Callout_left"], CultureInfo.InvariantCulture);

            /*Añade Caja Callout*/
            PowerPoint.Shape take = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftCallout, topCallout, 2*widthFormaCallout, heightFormaCallout);
            take.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Callout_background"]));
            take.Name = "calloutCaja";
            take.TextFrame.TextRange.Paragraphs().Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Callout_foreColor"]));
            take.TextFrame.TextRange.Paragraphs().Font.Name = csmet.CONFIG["MainFont"];
            take.TextFrame.TextRange.Font.Bold = MsoTriState.msoFalse;
            take.TextFrame.MarginLeft = float.Parse("5.67", CultureInfo.InvariantCulture);
            take.TextFrame.MarginRight = float.Parse("5.67", CultureInfo.InvariantCulture);
            take.Line.Visible = MsoTriState.msoTrue;
            take.Line.Style = Office.MsoLineStyle.msoLineSingle;
            take.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Callout_lineColor"]));
            take.Line.Weight = float.Parse(csmet.CONFIG["Callout_lineWidth"], CultureInfo.InvariantCulture);

            /*Linea Callout */
            float widthLineaCallout = float.Parse(csmet.CONFIG["Callout_width"], CultureInfo.InvariantCulture);
            float heightLineaCallout = float.Parse(csmet.CONFIG["Callout_height"], CultureInfo.InvariantCulture);
            float topLineaCallout = float.Parse(csmet.CONFIG["Callout_top"], CultureInfo.InvariantCulture);
            float leftLineaCallout = float.Parse(csmet.CONFIG["Callout_left"], CultureInfo.InvariantCulture);

            /*Añade Linea Callout*/
            PowerPoint.Shape takeL = slideActiva.Shapes.AddLine(leftCallout, topCallout + heightFormaCallout / 2, leftCallout - widthLineaCallout, topCallout + heightFormaCallout / 2);
            takeL.Name = "calloutLinea";
            takeL.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Callout_lineColor"]));
            takeL.Line.Weight = float.Parse(csmet.CONFIG["Callout_lineWidth"], CultureInfo.InvariantCulture);
            takeL.Line.EndArrowheadStyle = Office.MsoArrowheadStyle.msoArrowheadOval; //Cambiar al diccionario
            takeL.Line.EndArrowheadLength = Office.MsoArrowheadLength.msoArrowheadLengthMedium; //Cambiar al diccionario
            takeL.Line.EndArrowheadWidth = Office.MsoArrowheadWidth.msoArrowheadWidthMedium; //Cambiar al diccionario

            takeL.ConnectorFormat.BeginConnect(take, 2);

            /*Agrupacion*/
            string[] ArrayCallout = new string[] { "calloutCaja", "calloutLinea" };
            PowerPoint.Shape grpoCallout = slideActiva.Shapes.Range(ArrayCallout).Group();
            grpoCallout.Name = "Grupo Callout";
            grpoCallout.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

        }

        // Inserta un cuadro de texto con el formato de bullets
        public void insertCuadroTextoBullet(string n1, string n2, string n3)
        {
            try
            {
                PowerPoint.Shape CuadroTexto = null;

                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                float[] sizeCuadroTexto = csmet.getSizeCuadroTextoBullet();
                float top = sizeCuadroTexto[0];
                float left = sizeCuadroTexto[1];
                float wid = sizeCuadroTexto[2];
                float hei = sizeCuadroTexto[3];

                CuadroTexto = slideActiva.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, wid, hei);
                CuadroTexto.TextFrame.TextRange.Text = "Texto Primer Nivel\rTexto Segundo Nivel\rTexto Tercer Nivel";

                //Aplica formato por defecto a todo el cuadro
                CuadroTexto.TextFrame.TextRange.Paragraphs().Font.Name = csmet.getFontNameCuadroTextoBullets(); //al utilizar Paragraphs(), indica todas las lineas de texto

                //Párrafo 1
                csmet.formatBulletLevel(1, CuadroTexto.TextFrame.TextRange.Paragraphs(1), CuadroTexto.TextFrame2.TextRange.Paragraphs.Item(1), float.Parse(n1, CultureInfo.InvariantCulture), n1);
                //Párrafo 2
                csmet.formatBulletLevel(2, CuadroTexto.TextFrame.TextRange.Paragraphs(2), CuadroTexto.TextFrame2.TextRange.Paragraphs.Item(2), float.Parse(n2, CultureInfo.InvariantCulture), n1);
                //Párrafo 3
                csmet.formatBulletLevel(3, CuadroTexto.TextFrame.TextRange.Paragraphs(3), CuadroTexto.TextFrame2.TextRange.Paragraphs.Item(3), float.Parse(n3, CultureInfo.InvariantCulture), n1);

                CuadroTexto.TextFrame.TextRange.Font.Bold = MsoTriState.msoFalse;

                CuadroTexto.TextFrame.TextRange.Select();
                pptActiva.Application.StartNewUndoEntry();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        // Da el formato a los bullets según la configuración
        public void darFormatoBulletMatrix()
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;
                slideActiva = csmet.getDiapositiva(this.pptActiva);

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                    {
                        var formato = sel.TextFrame.TextRange;
                        var nivel = sel.TextFrame;

                        /*Tamaño y Formato Fuentes*/
                        formato.Paragraphs().Font.Name = csmet.getFontNameCuadroTextoBullets();

                        //El tamaño de letra más cercano de los predeterminados
                        int indexTamanoLetraSimilar = csmet.identificarTamanoLetra(formato.Paragraphs(1).Font.Size);
                        string[] tamanhos = csmet.CONFIG["Bllts_tamanos"].Split(';');

                        if ((float.Parse(tamanhos[indexTamanoLetraSimilar], CultureInfo.InvariantCulture) > 2 && csmet.CONFIG["Bllts_ajustarTam"].Equals("YES")) || csmet.CONFIG["Bllts_ajustarTam"].Equals("NO"))
                        {
                            /*******/
                            string fontPrimerNivel = tamanhos[indexTamanoLetraSimilar];

                            for (int i = 1; i <= formato.Paragraphs().Count; i++)
                            {
                                string fontNivel = tamanhos[indexTamanoLetraSimilar + formato.Paragraphs(i).IndentLevel - 1];
                                csmet.adjustFormatBulletLevel(formato.Paragraphs(i).IndentLevel, formato.Paragraphs(i), sel.TextFrame2.TextRange.Paragraphs.Item(i), float.Parse(fontNivel, CultureInfo.InvariantCulture), fontPrimerNivel);

                            }

                        }
                        else
                            MessageBox.Show("El tamaño de letra es muy chico para ejecutar esta función: " + tamanhos[indexTamanoLetraSimilar], "Dar formato");
                    }

                    pptActiva.Application.StartNewUndoEntry();
                }
                else
                {
                    //mensaje que debe seleccionar un objeto
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Hubo un error. Por favor contacte al administrador del sistema: \n" + ex.Message, "Dar formato");
            }
        }

        // Inserta el disclaimer según el tipo
        public void insertDisclaimer(string tipo)
        {
            try
            {

                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                float beginx = 0;

                string item = string.Empty;

                /* Texto Predeterminado----largo fijo ya que se descuadra la lectura del lbltexto en pptCustom Matrix*/
                if (tipo.StartsWith("btnD_Ot"))
                {
                    item = csmet.ShowDialog("Ingrese el texto del disclaimer", "Otro disclaimer");
                }
                else
                    item = csmet.CONFIG["Disclaimer_lbl_" + tipo.Substring(0, tipo.Length - 2)];

                if (!item.Equals(""))
                {
                    /*PPT Activa*/
                    slideActiva = csmet.getDiapositiva(this.pptActiva);

                    ancho = (int)slideActiva.CustomLayout.Width;
                    alto = (int)slideActiva.CustomLayout.Height;

                    float topTexto = float.Parse(csmet.CONFIG["Disclaimer_top"], CultureInfo.InvariantCulture);
                    float heightTexto = float.Parse(csmet.CONFIG["Disclaimer_height"], CultureInfo.InvariantCulture);

                    /*Insert Texto*/
                    PowerPoint.Shape lblTexto = slideActiva.Shapes.AddLabel(
                        Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, topTexto, 0, heightTexto);//float width no es necesario ya que toma el ancho de la palabra
                    lblTexto.TextFrame.TextRange.Text = item;

                    /*Propiedades Texto*/
                    lblTexto.TextFrame.MarginLeft = 0;
                    lblTexto.TextFrame.MarginRight = 0;
                    lblTexto.TextFrame.TextRange.Font.Name = csmet.CONFIG["MainFont"];
                    lblTexto.TextFrame.TextRange.Font.Size = float.Parse(csmet.CONFIG["Disclaimer_font_size"], CultureInfo.InvariantCulture);
                    lblTexto.TextFrame.TextRange.Font.Bold = (Microsoft.Office.Core.MsoTriState)Int32.Parse(csmet.CONFIG["Disclaimer_font_bold"]);
                    lblTexto.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("grisPlomo")); //Pasar a metodos
                    lblTexto.Left = ancho - lblTexto.Width - float.Parse(csmet.CONFIG["Disclaimer_gap"], CultureInfo.InvariantCulture); //texto alineado derecha                
                    lblTexto.TextFrame.MarginLeft = 0;
                    lblTexto.TextFrame.MarginRight = 0;
                    lblTexto.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;

                    /*margen lineas segun texto insertado*/
                    beginx = ancho - lblTexto.Width - float.Parse(csmet.CONFIG["Disclaimer_gap"], CultureInfo.InvariantCulture);

                    /*Insert linea inferior*/
                    float toplineaInf = float.Parse(csmet.CONFIG["Disclaimer_linf_top"], CultureInfo.InvariantCulture);
                    PowerPoint.Shape lineaInferior = slideActiva.Shapes.AddLine(beginx, toplineaInf, beginx + lblTexto.Width, toplineaInf);

                    /*Propiedades Linea Inferior*/
                    lineaInferior.ShapeStyle = (Office.MsoShapeStyleIndex)Int32.Parse(csmet.CONFIG["Disclaimer_line_shape"]);
                    lineaInferior.Line.Weight = float.Parse(csmet.CONFIG["Disclaimer_line_width"], CultureInfo.InvariantCulture);//Grosor linea        
                    lineaInferior.Line.Style = (Office.MsoLineStyle)Int32.Parse(csmet.CONFIG["Disclaimer_line_style"]);//Tipo de Linea 'Simple'
                    lineaInferior.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("rojo"));

                    string[] myRangeArray = new string[] { lblTexto.Name, lineaInferior.Name };

                    slideActiva.Shapes.Range(myRangeArray).Align(Office.MsoAlignCmd.msoAlignCenters, Office.MsoTriState.msoFalse);
                    PowerPoint.Shape grpo = slideActiva.Shapes.Range(myRangeArray).Group();

                    if (tipo.EndsWith("I"))
                        grpo.Left = float.Parse(csmet.CONFIG["Disclaimer_gap"], CultureInfo.InvariantCulture);

                    if (item.Length > 50)
                    {
                        MessageBox.Show("Tay loco?!! Bueno...es tu problema", "Comentario ExtraGeek...", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                    pptActiva.Application.StartNewUndoEntry();

                    // easter egg
                    if (item.ToLower().Equals("extra geek") && (csmet.CONFIG["Alineacion_integ3"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
                    {
                        GeekEncontrado ge = new GeekEncontrado();
                        ge.inicializar("Ricardo Mariño", "integ3", "Ricardo es miembro fundador y líder del Equipo Extra Geek. Lideró el desarrollo de la Barra Matrix, entre otros proyectos del equipo.", "escribir \"EXTRA GEEK\" en un Disclaimer", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                        ge.ShowDialog();
                        csmet.CONFIG.Remove("Alineacion_integ3");
                        csmet.CONFIG.Add("Alineacion_integ3", "1");
                        csmet.guardarConfig();
                    }
                    else if (item.ToLower().Equals("conceptual") && (csmet.CONFIG["Alineacion_integ4"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
                    {
                        GeekEncontrado ge = new GeekEncontrado();
                        ge.inicializar("Juan Domingo Pau", "integ4", "Juando es miembro fundador del Equipo Extra Geek. Impulsó la idea de construir la barra Matrix y el frente de Gadgets para la oficina.", "escribir \"CONCEPTUAL\" en un Disclaimer", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                        ge.ShowDialog();
                        csmet.CONFIG.Remove("Alineacion_integ4");
                        csmet.CONFIG.Add("Alineacion_integ4", "1");
                        csmet.guardarConfig();
                    }
                    
                }
                // Easter egg
                else if (item.Equals("") && (csmet.CONFIG["Alineacion_integ10"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
                {
                    GeekEncontrado ge = new GeekEncontrado();
                    ge.inicializar("Jürgen Uhlmann", "integ10", "Jürgen se unió al Equipo Extra Geek en 2015. Desde entonces, ha liderado el Hoborable Comité Inquisidor.", "insertar un Disclaimer con texto en blanco", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                    ge.ShowDialog();
                    csmet.CONFIG.Remove("Alineacion_integ10");
                    csmet.CONFIG.Add("Alineacion_integ10", "1");
                    csmet.guardarConfig();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        // Inserta un titulo con linea debajo
        public void insertTitulo()
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                /*Tamaño Diapositiva*/
                ancho = (int)slideActiva.CustomLayout.Width;
                alto = (int)slideActiva.CustomLayout.Height;

                /*Linea inferior */
                float widthFormaLinea = float.Parse(csmet.CONFIG["Titulo_linea_width"], CultureInfo.InvariantCulture);
                float heightFormaLinea = float.Parse(csmet.CONFIG["Titulo_linea_height"], CultureInfo.InvariantCulture);
                float topLinea = float.Parse(csmet.CONFIG["Titulo_linea_top"], CultureInfo.InvariantCulture);
                float leftLinea = float.Parse(csmet.CONFIG["Titulo_linea_left"], CultureInfo.InvariantCulture);

                /*Añade Linea inferior*/
                PowerPoint.Shape take = slideActiva.Shapes.AddLine(leftLinea, topLinea, widthFormaLinea+leftLinea, topLinea);
                take.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Titulo_linea_color"]));
                take.Line.Weight = float.Parse(csmet.CONFIG["Titulo_linea_weight"], CultureInfo.InvariantCulture);
                take.Name = "lineaTitulo";

                /*Cuadro Texto*/
                float topTexto = float.Parse(csmet.CONFIG["Titulo_text_top"], CultureInfo.InvariantCulture);
                float leftTexto = float.Parse(csmet.CONFIG["Titulo_text_left"], CultureInfo.InvariantCulture);
                float widthTexto = float.Parse(csmet.CONFIG["Titulo_text_width"], CultureInfo.InvariantCulture);
                float heightTexto = float.Parse(csmet.CONFIG["Titulo_text_height"], CultureInfo.InvariantCulture);

                /*Añade Cuadro de Conclusión*/
                PowerPoint.Shape takeT = slideActiva.Shapes.AddShape(Office.MsoAutoShapeType.msoShapeRectangle, leftTexto, topTexto, widthTexto, heightTexto);
                takeT.Fill.Visible = MsoTriState.msoFalse;
                takeT.Line.Visible = MsoTriState.msoFalse;
                takeT.Name = "textoTitulo";
                takeT.TextFrame.TextRange.Text = "Título";
                takeT.TextFrame.TextRange.Paragraphs().Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix(csmet.CONFIG["Titulo_text_foreColor"]));
                takeT.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                takeT.TextFrame.TextRange.Font.Size = float.Parse(csmet.CONFIG["Titulo_text_fontSize"], CultureInfo.InvariantCulture);
                takeT.TextFrame.TextRange.Font.Name = csmet.CONFIG["MainFont"];
                takeT.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = PpBaselineAlignment.ppBaselineAlignCenter;
                takeT.TextFrame.TextRange.ParagraphFormat.Bullet.Type = PpBulletType.ppBulletNone;
                takeT.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft; //alineacion texto centro
                takeT.TextFrame.MarginLeft = float.Parse(csmet.CONFIG["Titulo_text_marginLeft"], CultureInfo.InvariantCulture);
                takeT.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("grisPlomo")); //Pasar a metodos

                /*Agrupacion*/
                string[] ArrayTitulo = new string[] { "lineaTitulo", "textoTitulo" };
                PowerPoint.Shape grpoTitulo = slideActiva.Shapes.Range(ArrayTitulo).Group();
                grpoTitulo.Name = "Grupo Titulo";
                grpoTitulo.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignLeft;

                /*Propiedades Grupo*/
                grpoTitulo.Top = float.Parse(csmet.CONFIG["Titulo_text_top"], CultureInfo.InvariantCulture);
                grpoTitulo.Left = float.Parse(csmet.CONFIG["Titulo_text_left"], CultureInfo.InvariantCulture);

                takeT.Select();
                takeT.TextFrame.TextRange.Select();
                pptActiva.Application.StartNewUndoEntry();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        // Inserta el icono de estado seleccionado
        public void insertEstados(int num)
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                /*Tamaño Diapositiva*/
                ancho = (int)slideActiva.CustomLayout.Width;
                alto = (int)slideActiva.CustomLayout.Height;

                /*Añade icono*/
                Image icono = (Image)Properties.Resources.icono_igual.Clone();
                switch(num)
                {
                    case 1:
                        icono = (Image)Properties.Resources.icono_igual.Clone();
                        break;
                    case 2:
                        icono = (Image)Properties.Resources.icono_check_blanco.Clone();
                        break;
                    case 3:
                        icono = (Image)Properties.Resources.icono_check.Clone();
                        break;
                    case 4:
                        icono = (Image)Properties.Resources.icono_exclamacion.Clone();
                        break;
                    case 5:
                        icono = (Image)Properties.Resources.icono_cruz.Clone();
                        break;
                }
                /*
                String temporaryFilePath = "icono";
                icono.Save(temporaryFilePath + ".png", System.Drawing.Imaging.ImageFormat.Png);
                PowerPoint.Shape takeT = slideActiva.Shapes.AddPicture(temporaryFilePath + ".png", MsoTriState.msoFalse, MsoTriState.msoTrue, -50, 0, 40, 40);
                */

                String path = Path.Combine(Path.GetTempPath(), "icono.png");
                icono.Save(path);
                PowerPoint.Shape takeT = slideActiva.Shapes.AddPicture(path, MsoTriState.msoFalse, MsoTriState.msoTrue, -50, 0, 40, 40);

                File.Delete(path);
                 
                pptActiva.Application.StartNewUndoEntry();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        // Insertar calendario
        public void insertCalendario(int semanas, DateTime fecha)
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                /*Tamaño Diapositiva*/
                ancho = (int)slideActiva.CustomLayout.Width;
                alto = (int)slideActiva.CustomLayout.Height;

                /*Datos para cálculo de días */
                int num = fecha.Day;
                int mes = fecha.Month;
                int diasMes = DateTime.DaysInMonth(fecha.Year,fecha.Month);
                string mesNombre = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(fecha.ToString("MMMM", CultureInfo.CreateSpecificCulture(("es-ES"))));
                string mesNombreSgte = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(fecha.AddMonths(1).ToString("MMMM", CultureInfo.CreateSpecificCulture(("es-ES"))));

                /*Formato */
                float gap = csmet.centimetersToPoints(0.1); // Pasar a diccionario                    

                /* Etiquetas dias */
                PowerPoint.Shape lunes = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, csmet.centimetersToPoints(2.31), csmet.centimetersToPoints(3.39), csmet.centimetersToPoints(4.23), csmet.centimetersToPoints(0.73)); //Pasar a diccionario
                lunes.TextFrame.TextRange.Text = "Lunes";
                lunes.Name = "lunes";
                lunes.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignCenter;
                lunes.TextFrame.TextRange.Font.Bold = MsoTriState.msoCTrue;

                PowerPoint.ShapeRange martes = lunes.Duplicate();
                martes.TextFrame.TextRange.Text = "Martes";
                martes.Name = "martes";
                martes.Top = lunes.Top;
                martes.Left = lunes.Left + lunes.Width + gap;

                PowerPoint.ShapeRange miercoles = lunes.Duplicate();
                miercoles.TextFrame.TextRange.Text = "Miércoles";
                miercoles.Name = "miercoles";
                miercoles.Top = lunes.Top;
                miercoles.Left = lunes.Left + 2*lunes.Width + 2*gap;

                PowerPoint.ShapeRange jueves = lunes.Duplicate();
                jueves.TextFrame.TextRange.Text = "Jueves";
                jueves.Name = "jueves";
                jueves.Top = lunes.Top;
                jueves.Left = lunes.Left + 3*lunes.Width + 3*gap;

                PowerPoint.ShapeRange viernes = lunes.Duplicate();
                viernes.TextFrame.TextRange.Text = "Viernes";
                viernes.Name = "viernes";
                viernes.Top = lunes.Top;
                viernes.Left = lunes.Left + 4*lunes.Width + 4*gap;

                string[] ArrayEtiquetas = new string[] { "lunes", "martes", "miercoles", "jueves", "viernes" };

                PowerPoint.Shape grupoEtiquetas = slideActiva.Shapes.Range(ArrayEtiquetas).Group();
                grupoEtiquetas.Name = "Grupo Etiquetas";
                grupoEtiquetas.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("grisOscuro"));
                grupoEtiquetas.Fill.Visible = MsoTriState.msoFalse;

                /*Cuadro Día*/
                float topTexto = csmet.centimetersToPoints(4.32);
                float leftTexto = csmet.centimetersToPoints(2.3);
                float widthTexto = csmet.centimetersToPoints(4.23);
                float heightTexto = csmet.centimetersToPoints(2.42); //Por defecto tiene el alto de 5 semanas
                if (semanas == 4)
                    heightTexto = csmet.centimetersToPoints(3.03);

                /*Añade Cuadro Día*/
                PowerPoint.Shape dia1 = slideActiva.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftTexto, topTexto, widthTexto, heightTexto);
                dia1.Line.Visible = MsoTriState.msoFalse;
                dia1.TextFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
                dia1.Fill.BackColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("gris"));
                dia1.Height = heightTexto;
                dia1.Name = "dia1";
                dia1.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                dia1.TextFrame.TextRange.Font.Size = 12; // Pasar a diccionario
                dia1.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;
                dia1.TextFrame.TextRange.ParagraphFormat.BaseLineAlignment = PpBaselineAlignment.ppBaselineAlignTop;

                PowerPoint.ShapeRange dia2 = dia1.Duplicate();
                dia2.Name = "dia2";
                dia2.Top = dia1.Top;
                dia2.Left = dia1.Left + dia1.Width + gap;

                PowerPoint.ShapeRange dia3 = dia2.Duplicate();
                dia3.Name = "dia3";
                dia3.Top = dia2.Top;
                dia3.Left = dia2.Left + dia2.Width + gap;

                PowerPoint.ShapeRange dia4 = dia3.Duplicate();
                dia4.Name = "dia4";
                dia4.Top = dia3.Top;
                dia4.Left = dia3.Left + dia3.Width + gap;

                PowerPoint.ShapeRange dia5 = dia4.Duplicate();
                dia5.Name = "dia5";
                dia5.Top = dia4.Top;
                dia5.Left = dia4.Left + dia4.Width + gap;

                /*Agrupacion*/
                string[] ArraySemana1 = new string[] { "dia1", "dia2", "dia3", "dia4", "dia5" };
                PowerPoint.Shape semana1 = slideActiva.Shapes.Range(ArraySemana1).Group();
                semana1.Name = "Grupo Semana 1";
                semana1.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("negroBase"));

                PowerPoint.ShapeRange semana2 = semana1.Duplicate();
                semana2.Name = "Grupo Semana 2";
                semana2.Top = semana1.Top + semana1.Height + gap;
                semana2.Left = semana1.Left;

                PowerPoint.ShapeRange semana3 = semana2.Duplicate();
                semana3.Name = "Grupo Semana 3";
                semana3.Top = semana2.Top + semana2.Height + gap;
                semana3.Left = semana2.Left;

                PowerPoint.ShapeRange semana4 = semana3.Duplicate();
                semana4.Name = "Grupo Semana 4";
                semana4.Top = semana3.Top + semana3.Height + gap;
                semana4.Left = semana3.Left;

                string[] ArraySemanas;

                if(semanas == 5)
                {
                    PowerPoint.ShapeRange semana5 = semana4.Duplicate();
                    semana5.Name = "Grupo Semana 5";
                    semana5.Top = semana4.Top + semana4.Height + gap;
                    semana5.Left = semana4.Left;

                    ArraySemanas = new string[] { "Grupo Semana 1", "Grupo Semana 2", "Grupo Semana 3", "Grupo Semana 4", "Grupo Semana 5" };
                }
                else
                {
                    ArraySemanas = new string[] { "Grupo Semana 1", "Grupo Semana 2", "Grupo Semana 3", "Grupo Semana 4" };
                }
                
                PowerPoint.Shape todos = slideActiva.Shapes.Range(ArraySemanas).Group();

                int primerDia = num;
                int contadorDias = 1;
                int contadorSemanas = 0;
                int semanaCambio = semanas;
                foreach (PowerPoint.Shape aux in todos.GroupItems)
                {
                    if(primerDia > diasMes)
                    {
                        semanaCambio = contadorSemanas;
                        primerDia = primerDia - diasMes;
                    }
                    aux.TextFrame.TextRange.Text = "" + primerDia;
                    primerDia++;
                    contadorDias++;
                    if(contadorDias == 6)
                    {
                        contadorDias = 1;
                        contadorSemanas++;
                        primerDia++;
                        primerDia++;
                    }
                }

                /*Etiquetas meses */
                PowerPoint.Shape mes1 = slideActiva.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, csmet.centimetersToPoints(1.46), csmet.centimetersToPoints(4.32), csmet.centimetersToPoints(0.73), semanaCambio*heightTexto+(semanaCambio-1)*gap);
                mes1.TextFrame.TextRange.Text = mesNombre;
                mes1.TextFrame.TextRange.Font.Size = 12;
                mes1.TextFrame.TextRange.Font.Bold = MsoTriState.msoCTrue;
                mes1.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("grisOscuro"));
                mes1.TextFrame.TextRange.Font.Color.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("blanco"));
                mes1.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationUpward;
                mes1.Line.Visible = MsoTriState.msoFalse;

                if (semanaCambio < semanas)
                {
                    PowerPoint.ShapeRange mes2 = mes1.Duplicate();
                    mes2.TextFrame.TextRange.Text = mesNombreSgte;
                    mes2.Height = (semanas - semanaCambio) * heightTexto + (semanas - semanaCambio - 1) * gap;
                    mes2.Top = mes1.Top + mes1.Height + gap;
                    mes2.Left = mes1.Left;
                }

                pptActiva.Application.StartNewUndoEntry();

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        #endregion

        #region Transformaciones de elementos

        // Ajusta ancho o alto de las figuras
        public void ajustaTamano(string opcion)
        {
            try
            {
                //Presentacion activa
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                int count = 0;

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    /*Cuenta Elementos seleccionados*/
                    count = pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count;

                    /*Obtiene posicion Inicial segun cantidad de formas seleccionadas*/
                    if (count > 1)
                    {
                        PowerPoint.Shape fpad = pptActiva.Application.ActiveWindow.Selection.ShapeRange[1];

                        if (opcion == "anc")
                        {
                            foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                            {
                                if (sel.Rotation == 270 || sel.Rotation == 90)
                                    sel.Height = (fpad.Rotation == 270 || fpad.Rotation == 90) ? fpad.Height : fpad.Width;
                                else
                                    sel.Width = (fpad.Rotation == 270 || fpad.Rotation == 90) ? fpad.Height : fpad.Width;
                            }
                        }
                        else if (opcion == "alt")
                        {
                            foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                            {
                                if (sel.Rotation == 270 || sel.Rotation == 90)
                                    sel.Width = (fpad.Rotation == 270 || fpad.Rotation == 90) ? fpad.Width : fpad.Height;
                                else
                                    sel.Height = (fpad.Rotation == 270 || fpad.Rotation == 90) ? fpad.Width : fpad.Height;
                            }
                        }

                    }
                    else
                    {
                        MessageBox.Show("Personalizar Mensaje-(No se puede ajustar forma)", "Información");
                    }

                    pptActiva.Application.StartNewUndoEntry();
                }
                else
                {
                    MessageBox.Show("Personalizar Mensaje-(Debe Seleccionar un Item)", "Información");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Personalizar Mensaje-(Debe Seleccionar un Item)", "Información");
            }

        }

        // Empalma las figuras asignando la posición vertical según la altura de la figura anterior
        public void empalmarVertical()
        {
            try
            {
                //Presentacion activa
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                // Ordena las figuras según su posición vertical
                int cuenta = pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count;
                if (cuenta > 1)
                {
                    int inicioOrder = 1;
                    int[] order = new int[cuenta + 1];
                    order[1] = -1;
                    float nextPos = 0;
                    Boolean primer = true;
                    for (int j = 2; j <= cuenta; j++)
                    {
                        PowerPoint.Shape sel = pptActiva.Application.ActiveWindow.Selection.ShapeRange[j];
                        order[j] = -1;
                        int k = inicioOrder;
                        if (pptActiva.Application.ActiveWindow.Selection.ShapeRange[k].Top > sel.Top)
                        {
                            inicioOrder = j;
                            order[j] = k;
                        }
                        else
                        {
                            int anterior = inicioOrder;
                            k = order[inicioOrder];
                            while (k != -1 && pptActiva.Application.ActiveWindow.Selection.ShapeRange[k].Top < sel.Top)
                            {
                                anterior = k;
                                k = order[k];
                            }

                            order[anterior] = j;
                            order[j] = k;

                        }
                    }

                    // Asigna la posición de cada figura
                    int i = inicioOrder;
                    while (i != -1)
                    {
                        PowerPoint.Shape sel = pptActiva.Application.ActiveWindow.Selection.ShapeRange[i];

                        if (!primer)
                        {
                            sel.Top = nextPos;
                        }
                        else
                            primer = false;

                        nextPos = sel.Top + sel.Height + float.Parse(csmet.CONFIG["Empalmar_gapV"], CultureInfo.InvariantCulture);

                        i = order[i];
                    }

                    pptActiva.Application.StartNewUndoEntry();


                }
                else
                {
                    MessageBox.Show("De seleccionar al menos dos ítems", "Información");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Error, contacte al administrador del sistema: " + ex.Message, "Información");
            }

        }

        // Empalma las figuras asignando la posición vertical según la altura de la figura anterior
        public void empalmarHorizontal()
        {
            try
            {
                //Presentacion activa
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                // Organiza las shapes según su posición horizontal
                int cuenta = pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count;
                if (cuenta > 1)
                {
                    int inicioOrder = 1;
                    int[] order = new int[cuenta + 1];
                    order[1] = -1;
                    float nextPos = 0;
                    Boolean primer = true;
                    for (int j = 2; j <= cuenta; j++)
                    {
                        PowerPoint.Shape sel = pptActiva.Application.ActiveWindow.Selection.ShapeRange[j];
                        order[j] = -1;
                        int k = inicioOrder;
                        if (pptActiva.Application.ActiveWindow.Selection.ShapeRange[k].Left > sel.Left)
                        {
                            inicioOrder = j;
                            order[j] = k;
                        }
                        else
                        {
                            int anterior = inicioOrder;
                            k = order[inicioOrder];
                            while (k != -1 && pptActiva.Application.ActiveWindow.Selection.ShapeRange[k].Left < sel.Left)
                            {
                                anterior = k;
                                k = order[k];
                            }

                            order[anterior] = j;
                            order[j] = k;

                        }
                    }

                    // Asigna la posición horizontal
                    int i = inicioOrder;
                    while (i != -1)
                    {
                        PowerPoint.Shape sel = pptActiva.Application.ActiveWindow.Selection.ShapeRange[i];

                        if (!primer)
                        {
                            sel.Left = nextPos;
                        }
                        else
                            primer = false;

                        nextPos = sel.Left + sel.Width + float.Parse(csmet.CONFIG["Empalmar_gapV"], CultureInfo.InvariantCulture);

                        i = order[i];
                    }

                    pptActiva.Application.StartNewUndoEntry();
                }
                else
                {
                    MessageBox.Show("De seleccionar al menos dos ítems", "Información");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Error, contacte al administrador del sistema: " + ex.Message, "Información");
            }

        }

        // Para copiar y pegar posiciones se guarda un atributo de referencia con esquina superior izquierda
        private float posiciones_Left = -99999;
        private float posiciones_Top = -99999;

        // Asigna la posición guardada a todas las figuras seleccionadas
        public void pastePosiciones()
        {
            try
            {
                //Presentacion activa
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {

                    foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                    {
                        sel.Left = posiciones_Left;
                        sel.Top = posiciones_Top;
                    }

                    pptActiva.Application.StartNewUndoEntry();
                }
                else
                {
                    MessageBox.Show("Debe seleccionar al menos un ítem", "Información");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Error, contacte al administrador del sistema: " + ex.Message, "Información");
            }

        }

        // Guarda la posición del objeto que se encuentre más arriba a la izquierda
        public void copyPosiciones()
        {
            try
            {
                //Presentacion activa
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                int count = 0;

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    /*Cuenta Elementos seleccionados*/
                    count = pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count;

                    posiciones_Left = 9999999;
                    posiciones_Top = 9999999;
                    int i = 0;
                    foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                    {
                        if (posiciones_Left > sel.Left)
                            posiciones_Left = sel.Left;

                        if (posiciones_Top > sel.Top)
                            posiciones_Top = sel.Top;
                        i++;
                    }

                    pptActiva.Application.StartNewUndoEntry();
                }
                else
                {
                    MessageBox.Show("Debe Seleccionar un Item", "Información");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Error, contacte al administrador del sistema: " + ex.Message, "Información");
            }

        }

        // Cambia la posición de los dos elementos seleccionados
        public void cambiarPosiciones()
        {
            try
            {
                //Presentacion activa
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count == 2)
                {
                    PowerPoint.Shape shape1 = pptActiva.Application.ActiveWindow.Selection.ShapeRange[1];
                    PowerPoint.Shape shape2 = pptActiva.Application.ActiveWindow.Selection.ShapeRange[2];

                    float aux1Top = shape1.Top;
                    float aux1Left = shape1.Left;

                    float aux2Top = shape2.Top;
                    float aux2Left = shape2.Left;

                    shape1.Top = aux2Top;
                    shape1.Left = aux2Left;

                    shape2.Top = aux1Top;
                    shape2.Left = aux1Left;

                    pptActiva.Application.StartNewUndoEntry();
                }
                else
                {
                    MessageBox.Show("Debe seleccionar al menos un ítem", "Información");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Error, contacte al administrador del sistema: " + ex.Message, "Información");
            }

        }

        // Aumenta o disminuye el espaciado entre los objetos
        public void ajustarEspaciado(int direccion, int aumentarOReducir)
        {
            try
            {
                //Presentacion activa
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                int cuenta = pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count;
                if (cuenta > 1)
                {
                    float[] poss = new float[cuenta + 1];
                    int[] siguiente = new int[cuenta + 1];
                    int minI = -1;
                    float minPos = 999999;
                    float maxPos = -999999;
                    float sumaDims = 0;

                    int i = 1;
                    foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                    {
                        float pos;
                        float dim;
                        if (direccion == 1) pos = sel.Left; else pos = sel.Top;
                        if (direccion == 1) dim = sel.Width; else dim = sel.Height;

                        int j = minI;
                        int jAnt = -1;
                        Boolean encontrado = false;
                        int contador2 = 1;
                        poss[i] = pos;
                        while (contador2 < i && !encontrado)
                        {
                            if (pos < poss[j])
                            {
                                if (jAnt == -1)
                                {
                                    siguiente[i] = minI;
                                    minI = i;
                                    minPos = pos;
                                }
                                else
                                {
                                    siguiente[jAnt] = i;
                                    siguiente[i] = j;
                                    encontrado = true;
                                }
                            }
                            jAnt = j;
                            j = siguiente[j];
                            contador2++;
                        }

                        if (!encontrado)
                        {
                            if (pos < minPos)
                            {
                                siguiente[i] = minI;
                                minI = i;
                                minPos = pos;
                            }
                            else
                            {
                                siguiente[jAnt] = i;
                            }
                        }


                        if (pos + dim > maxPos)
                        {
                            maxPos = pos + dim;
                        }
                        sumaDims += dim;
                        i++;
                    }

                    float multiplicador = float.Parse(csmet.CONFIG["Espaciado_mult"], CultureInfo.InvariantCulture);
                    if (aumentarOReducir != 1)
                        multiplicador = multiplicador * (-1);

                    float espaciadoTotal = (maxPos - minPos - sumaDims) * (1 + multiplicador);
                    float nuevoSumaDims = (maxPos - minPos - espaciadoTotal);

                    if (espaciadoTotal > 0 && nuevoSumaDims > 0)
                    {
                        int contador2 = 1;
                        float nuevaPos = minPos;
                        int j = minI;
                        while (contador2 <= cuenta)
                        {
                            PowerPoint.Shape sel = pptActiva.Application.ActiveWindow.Selection.ShapeRange[j];
                            if (direccion == 1)
                                sel.Left = nuevaPos;
                            else
                                sel.Top = nuevaPos;

                            if (direccion == 1)
                                sel.Width = (sel.Width / sumaDims) * nuevoSumaDims;
                            else
                                sel.Height = (sel.Height / sumaDims) * nuevoSumaDims;

                            j = siguiente[j];
                            if (direccion == 1)
                                nuevaPos = nuevaPos + sel.Width + (espaciadoTotal / (cuenta - 1));
                            else
                                nuevaPos = nuevaPos + sel.Height + (espaciadoTotal / (cuenta - 1));
                            contador2++;
                        }

                        pptActiva.Application.StartNewUndoEntry();
                    }
                    else
                    {
                        MessageBox.Show("No es posible ajustar el espaciado.", "Información");
                    }
                }
                else
                {
                    MessageBox.Show("De seleccionar al menos dos ítems", "Información");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Error, contacte al administrador del sistema: " + ex.Message, "Información");
            }

        }

        #endregion

        #region Seleccionar


        // Parámetros para seleccionar similares
        public Boolean tamanio = false;
        public Boolean colorFondo = false;
        public Boolean colorLinea = false;
        public Boolean tipo = false;

        // Selecciona las figuras similares a la que se encuentre seleccionada
        public void seleccionarSimilares()
        {
            //Presentacion activa
            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            slideActiva = csmet.getDiapositiva(this.pptActiva);

            int cuenta = pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count;
            if (cuenta == 1)
            {
                // Dependiendo de la configuración, se muestra el Form de opciones para seleccionar similares o se asignan los parámetros.
                if (csmet.CONFIG["SeleccionarSimilares_preguntar"].Equals("YES") ? true : false)
                {
                    SeleccionarSimilares selecForm = new SeleccionarSimilares();
                    selecForm.inicializar(this);
                    selecForm.Show();
                }
                else
                {
                    tamanio = csmet.CONFIG["SeleccionarSimilares_tamanio"].Equals("YES") ? true : false;
                    colorFondo = csmet.CONFIG["SeleccionarSimilares_fondo"].Equals("YES") ? true : false;
                    colorLinea = csmet.CONFIG["SeleccionarSimilares_linea"].Equals("YES") ? true : false;
                    tipo = csmet.CONFIG["SeleccionarSimilares_tipo"].Equals("YES") ? true : false;
                    seleccionarSimilares2();
                }
            }
            else
            {
                MessageBox.Show("Debe seleccionar sólo una figura");
            }
        }

        // Una vez asignados los parámetros se procede a la selección
        public void seleccionarSimilares2()
        {
            //Presentacion activa
            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            slideActiva = csmet.getDiapositiva(this.pptActiva);

            PowerPoint.Shape referencia = pptActiva.Application.ActiveWindow.Selection.ShapeRange[1];

            // Recorre todos los shapes y determina si cumple con los parámetros
            foreach (PowerPoint.Shape sh in slideActiva.Shapes)
            {
                bool tam = !tamanio || (referencia.Width == sh.Width && referencia.Height == sh.Height);
                bool fon = !colorFondo || (referencia.Fill.ForeColor.RGB.CompareTo(sh.Fill.ForeColor.RGB) == 0);
                bool lin = !colorLinea || (referencia.Line.ForeColor.RGB.CompareTo(sh.Line.ForeColor.RGB) == 0);
                bool tip = !tipo || (referencia.AutoShapeType.CompareTo(sh.AutoShapeType) == 0);

                if (tam && fon && lin && tip)
                {
                    sh.Select(MsoTriState.msoFalse);
                }

            }

            if (!tamanio && !colorFondo && !colorLinea && !tipo && (csmet.CONFIG["Alineacion_integ7"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
            {
                GeekEncontrado ge = new GeekEncontrado();
                ge.inicializar("Matías Schlotfeldt", "integ7", "Matías se unió al Equipo Extra Geek en 2015. Participó en el comité de desarrollo de la Barra Matrix.", "utilizar Seleccionar Similares sin ninguna opción activada", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                ge.ShowDialog();
                csmet.CONFIG.Remove("Alineacion_integ7");
                csmet.CONFIG.Add("Alineacion_integ7", "1");
                csmet.guardarConfig();
            }

        }

        #endregion

        #region Ghosts 

        // Esta región controla la creación de Ghost Simple y Smart Ghost

        Clases.metodos.Escala escal = new Clases.metodos.Escala();

        public void insertGhostSimple()
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                ancho = (int)slideActiva.CustomLayout.Width;

                int count = 0;

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    /*Cuenta Elementos seleccionados*/
                    count = pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count;

                    string[] myRangeArray = new string[count];
                    int ct = 0;

                    PowerPoint.Shape grpo = null;

                    if ((!csmet.getValidaGrupoShape(slideActiva, "Grupo ghost")))
                    {

                        foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                        {

                            sel.Copy(); //copia original
                            var array = slideActiva.Shapes.Paste(); //pega original con todas las propiedades

                            float borde = sel.Line.Weight;
                            float foreColor = sel.Line.ForeColor.RGB;
                            string tipo = sel.AutoShapeType.ToString();

                            /*borde predeterminado para toda forma*/
                            if (borde > 0)
                                array.Line.Weight = float.Parse("0.25", CultureInfo.InvariantCulture);

                            /*Edicion de propiedades segun Tipo de forma*/
                            if (tipo != "msoShapeMixed")
                            {
                                array.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("grisOscuro"));//Color default
                                array[1].TextFrame.TextRange.Text = string.Empty;
                            }
                            else
                            {
                                array.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("grisOscuro"));//Color default
                                array.Line.BeginArrowheadLength = MsoArrowheadLength.msoArrowheadShort;
                                array.Line.EndArrowheadLength = MsoArrowheadLength.msoArrowheadShort;
                            }

                            /*add formas en arreglo*/
                            array.Name = "shape" + ct + "";
                            myRangeArray[ct] = array.Name;
                            ct += 1;
                            grpo = array[1];//agrega shape cuando se crea ghost de una sola figura pasa de shapesRange{} a shape 

                        }

                        /*Agrupacion*/
                        if (ct > 1)
                            grpo = slideActiva.Shapes.Range(myRangeArray).Group();

                        grpo.Name = "Grupo ghost";

                        float W = float.Parse("53.29", CultureInfo.InvariantCulture);
                        float H = float.Parse("34.30", CultureInfo.InvariantCulture);
                        float T = float.Parse("65.015", CultureInfo.InvariantCulture);
                        float L = ancho-W-float.Parse("14.17", CultureInfo.InvariantCulture);

                        escal = csmet.getEscalaGhost(grpo.Width, grpo.Height, W, H);

                        grpo.Width = escal.Width;
                        grpo.Height = escal.Height;
                        grpo.Top = T + (H - grpo.Height) / 2;
                        grpo.Left = L + (W - grpo.Width) / 2;

                        /*Añade texto fantasma*/
                        float widthTxt = float.Parse("156.25", CultureInfo.InvariantCulture);
                        float heightTxt = float.Parse("21.81", CultureInfo.InvariantCulture);
                        float leftTxt = L - widthTxt - 20;
                        float topTxt = T + (H - heightTxt) / 2;

                        PowerPoint.Shape textFanta = slideActiva.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, leftTxt, topTxt, widthTxt, heightTxt);
                        textFanta.TextFrame.MarginLeft = 0;
                        textFanta.TextFrame.MarginRight = 0;
                        textFanta.TextFrame.TextRange.Text = "Fantasma";
                        textFanta.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;
                        textFanta.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                        textFanta.TextFrame.TextRange.Font.Name = "Lao UI";
                        textFanta.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                        textFanta.TextFrame.TextRange.Font.Size = float.Parse("10.5", CultureInfo.InvariantCulture);

                        textFanta.Top = T + (H - textFanta.Height) / 2;
                        textFanta.TextFrame.TextRange.Select(); // cuadro con texto fantasma queda seleccionado.

                        pptActiva.Application.StartNewUndoEntry();//checkPoint

                    }
                    else
                    {
                        MessageBox.Show("Personalizar Mensaje-(Grupo ghost ya existe)", "Ghost");
                    }

                    if (count == 4 && (csmet.CONFIG["Alineacion_integ5"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
                    {
                       
                            GeekEncontrado ge = new GeekEncontrado();
                            ge.inicializar("Íñigo Fika", "integ5", "Íñigo ha sido sponsor estratégico desde el área de TI. No pierde oportunidad de declararse \"fan\" del Equipo Extra Geek.", "crear un ghost con 4 objetos", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                            ge.ShowDialog();
                            csmet.CONFIG.Remove("Alineacion_integ5");
                            csmet.CONFIG.Add("Alineacion_integ5", "1");
                            csmet.guardarConfig();
                            
                    }
                }
                else
                {
                    MessageBox.Show("Personalizar Mensaje-(Debe seleccionar una forma)", "Ghost");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Personalizar Mensaje-(Debe seleccionar una forma)", "Ghost");
            }
        }

        public void insertGhostSmart()
        {
            try
            {
                pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

                slideActiva = csmet.getDiapositiva(this.pptActiva);

                ancho = (int)slideActiva.CustomLayout.Width;

                int count = 0;

                if (pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count > 0)
                {
                    int cuenta = pptActiva.Application.ActiveWindow.Selection.ShapeRange.Count;
                    List<PowerPoint.Shape> lformas = new List<PowerPoint.Shape>();

                    foreach (PowerPoint.Shape sel in pptActiva.Application.ActiveWindow.Selection.ShapeRange)
                    {
                        if (sel.AutoShapeType != MsoAutoShapeType.msoShapeMixed)
                            count++;

                        lformas.Add(sel);//genera lista de formas seleccionadas.                       
                    }

                    /*Metodo que agrega formas en nuevas PPT*/
                    insertNuevaDiapositiva(pptActiva, count, lformas, ancho);



                    if (cuenta == 4 && (csmet.CONFIG["Alineacion_integ6"].Equals("0") || csmet.CONFIG["ModoGeek"].Equals("1")))
                    {
                        GeekEncontrado ge = new GeekEncontrado();
                        ge.inicializar("Joshua Barten", "integ6", "Joshua es miembro fundador del Equipo Extra Geek. Lideró el desarrollo del curso de Macros.", "crear un Smart Ghost con 4 objetos", csmet.darGeeksEncontrados(), csmet.darTotalGeeks());
                        ge.ShowDialog();
                        csmet.CONFIG.Remove("Alineacion_integ6");
                        csmet.CONFIG.Add("Alineacion_integ6", "1");
                        csmet.guardarConfig();
                    }
                }
                else
                {
                    MessageBox.Show("Personalizar Mensaje-(Ninguna PPT Creada)", "Información");
                }

            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                MessageBox.Show("Personalizar Mensaje-(Debe seleccionar 2 o mas formas)", "Información");
            }
        }


        public void insertNuevaDiapositiva(PowerPoint.Presentation pptPreso, int count, List<PowerPoint.Shape> formas, int ancho)
        {
            /*Metodo que añade una nueva Hoja o Presentacion segun como sea llamado*/

            PowerPoint.CustomLayout pptLayout;
            PowerPoint.Slide newSlide = null;

            pptLayout = default(PowerPoint.CustomLayout);

            //CustomLayouts._Index(7) = PPT en blanco
            //_Index(6) = PPT Solo titulo

            if ((pptPreso.SlideMaster.CustomLayouts._Index(6) == null))
                pptLayout = pptPreso.SlideMaster.CustomLayouts._Index(1);
            else
                pptLayout = pptPreso.SlideMaster.CustomLayouts._Index(6);

            slideActiva = csmet.getDiapositiva(pptPreso);

            int index_ppt = slideActiva.SlideIndex + 1;

            /*Obtiene posicion Inicial segun cantidad de formas seleccionadas*/
            for (int i = 0; i < count; i++)//Nuevas PPT
            {
                /*inserta ppt desde Origen*/
                newSlide = pptPreso.Slides.AddSlide((index_ppt), pptLayout);

                int pinta = 0;
                string[] myRangeArray = new string[formas.Count];
                int ct = 0;

                PowerPoint.Shape grpo = null;
                string textoForma = string.Empty;

                foreach (PowerPoint.Shape item in formas)// inserta formas
                {
                    /*Agrega shape 'forma'*/
                    float borde = item.Line.Weight;
                    float foreColor = item.Line.ForeColor.RGB;
                    string tipo = item.AutoShapeType.ToString();

                    item.Copy();
                    var array = newSlide.Shapes.Paste();
                    //PowerPoint.Shape array = newSlide.Shapes.AddShape(item.AutoShapeType, item.Left, item.Top, item.Width, item.Height);
                    if (borde > 0)
                        array[1].Line.Weight = float.Parse("0.25", CultureInfo.InvariantCulture);

                    /*Color Default*/
                    if (tipo != "msoShapeMixed")
                        array.Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("gris"));//Color default                                            
                    else
                    {
                        array.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("gris"));//Color default
                        array.Line.BeginArrowheadLength = MsoArrowheadLength.msoArrowheadShort;
                        array.Line.EndArrowheadLength = MsoArrowheadLength.msoArrowheadShort;
                    }


                    array[1].Name = "shape" + ct + "";
                    myRangeArray[ct] = array[1].Name;
                    ct += 1;
                    grpo = array[1];

                    if (tipo != "msoShapeMixed")
                    {
                        /*Asigna color de fondo segun orden de figuras y ppt*/
                        if (i == pinta)
                        {
                            array[1].Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("rojo")); //Color imagen segun orden de ppt
                            textoForma = item.TextFrame.TextRange.Text;
                            array[1].TextFrame.TextRange.Text = "";

                        }
                        else
                        {

                            array[1].Fill.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("gris")); //Color imagen segun orden de ppt                            
                            array[1].TextFrame.TextRange.Text = "";
                        }
                    }

                    if (tipo == "msoShapeMixed")
                        pinta += 0;
                    else
                        pinta += 1;
                }

                /*Agrupacion*/
                if (ct > 1)
                    grpo = newSlide.Shapes.Range(myRangeArray).Group();


                float W = float.Parse("53.29", CultureInfo.InvariantCulture);
                float H = float.Parse("34.30", CultureInfo.InvariantCulture);
                float T = float.Parse("84.12", CultureInfo.InvariantCulture);
                float L = float.Parse("658.45", CultureInfo.InvariantCulture);

                escal = csmet.getEscalaGhost(grpo.Width, grpo.Height, W, H);

                grpo.Width = escal.Width;
                grpo.Height = escal.Height;
                grpo.Top = T + (H - grpo.Height) / 2;
                grpo.Left = L + (W - grpo.Width) / 2;


                float widthTxt = float.Parse("156.25", CultureInfo.InvariantCulture);
                float heightTxt = float.Parse("21.81", CultureInfo.InvariantCulture);
                float leftTxt = ancho - widthTxt - W - 20;
                float topTxt = T + (H - heightTxt) / 2;

                PowerPoint.Shape textFanta = newSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, leftTxt, topTxt, widthTxt, heightTxt);
                textFanta.TextFrame.MarginLeft = 0;
                textFanta.TextFrame.MarginRight = 0;
                textFanta.TextFrame.TextRange.Text = textoForma;
                textFanta.TextFrame.TextRange.ParagraphFormat.Alignment = PpParagraphAlignment.ppAlignRight;
                textFanta.TextFrame.VerticalAnchor = MsoVerticalAnchor.msoAnchorMiddle;
                textFanta.TextFrame.TextRange.Font.Name = "Verdana";
                textFanta.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoCTrue;
                textFanta.TextFrame.TextRange.Font.Size = float.Parse("10.5", CultureInfo.InvariantCulture);

                textFanta.Top = T + (H - textFanta.Height) / 2;

                index_ppt += 1; //paso a siguiente PPT

                pptActiva.Application.StartNewUndoEntry();
            }
        }

        #endregion

        #endregion

        #region Aplicaciones auxiliares

        // Obtiene el texto de un recurso (usado actualmente para la lectura del xml del ribbon)
        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        #region Eventos

        // Determina si los controles del ribbon estarán o no habilitados
        public bool GetEnabled(IRibbonControl control)
        {

            if (control.Id == "btnGhostSimple" && ThisAddIn.e_seleccion == null)
                return false;
            else if (control.Id == "btnSmartGhost" && ThisAddIn.e_seleccion == null)
                return false;
            else if (control.Id == "btnDarformato" && ThisAddIn.e_seleccion == null)
                return false;
            else if (control.Id == "btnAncho" && ThisAddIn.e_seleccion == null)
                return false;
            else if (control.Id == "btnSeleccionarSim" && (ThisAddIn.e_seleccion == null || ThisAddIn.e_count != 1))
                return false;
            else if ((control.Id == "btnAddEspaciadoV" || control.Id == "btnSubEspaciadoV" || control.Id == "btnAddEspaciadoH" || control.Id == "btnSubEspaciadoH") && (ThisAddIn.e_seleccion == null || ThisAddIn.e_count < 2))
                return false;
            else if ((control.Id == "btnCopyPos" || control.Id == "btnDestacar" || control.Id == "btnSombrear") && ThisAddIn.e_seleccion == null)
                return false;
            else if (control.Id == "btnPastePos" && (ThisAddIn.e_seleccion == null || posiciones_Left == -99999))
                return false;
            else if (control.Id == "btnAlto" || control.Id == "btnAncho" || control.Id == "btnEmpalmaV" || control.Id == "btnEmpalmaH")
            {
                if (ThisAddIn.e_seleccion == null || ThisAddIn.e_count < 2)
                    return false;
                else
                    return true;
            }
            else if (control.Id == "btnChangePos")
            {
                if (ThisAddIn.e_seleccion == null || ThisAddIn.e_count != 2)
                    return false;
                else
                    return true;
            }

            else if (control.Id == "btnActualizarAgenda" || control.Id == "btnEliminarAgenda")
            {
                foreach (AgendaClase a in agendas)
                {
                    if (!a.cargarAgenda())
                        return false;
                }
                return true;
            }
            else
                return true;
        }

        #endregion

    }//end class
}//end namespace
