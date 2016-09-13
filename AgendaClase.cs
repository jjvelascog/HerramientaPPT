using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Office = Microsoft.Office.Core;
using System.Globalization;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Windows.Forms;

namespace Herramientas
{
    public class AgendaClase
    {
        // Atributos de apoyo para administrar la agenda
        public string LABEL_ITEM_AGENDA = "SECCION_";
        public string LABEL_TITULO_AGENDA = "TITULO_AGENDA_";
        public string LABEL_MARCADOR_AGENDA = "MARCADOR_AGENDA_";
        public string LABEL_AGENDA = "Agenda_";
        public Slide agenda = null;
        public int id = 0;
        public string titulo = "Agenda";
        private Agenda agendaForm = null;
        private PowerPoint.Shape tituloAgenda = null;
        private Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape> itemsAgenda = new Dictionary<string, Microsoft.Office.Interop.PowerPoint.Shape>();
        private Dictionary<string, float[]> posItemsAgenda = new Dictionary<string, float[]>();
        private static bool horaItemsAgenda;
        private int numItems;
        private static bool warningSubAgenda = false;
        private bool agendaCargada = false;
        private static Random rnd = new Random();
        private bool tienePortada = true;
        private PowerPoint.Shape linea = null;
        private static barra_matrix barraMatrix;

        // Relación a instancia de la clase métodos
        private static Clases.metodos csmet = Clases.metodos.Instance;
        // Referencia a la presentación y a la slide activa
        PowerPoint.Presentation pptActiva;
        PowerPoint.Slide slideActiva;

        public AgendaClase(barra_matrix barra, Presentation activePresentation, int idGeneral)
        {
            id = idGeneral;
            barraMatrix = barra;
            pptActiva = activePresentation;
            titulo = titulo + id;
            LABEL_ITEM_AGENDA = LABEL_ITEM_AGENDA + id + "_";
            LABEL_TITULO_AGENDA = LABEL_TITULO_AGENDA + id;
            LABEL_AGENDA = LABEL_AGENDA + id + "_";
            crearAgenda();
        }

        // Cambia el título de la agenda
        public void actualizarTituloAgenda(String nTitulo)
        {
            titulo = nTitulo;
            try
            {
                agenda.Shapes.Title.TextFrame.TextRange.Text = nTitulo;
            }
            catch (Exception e)
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

            //pptActiva.Application.StartNewUndoEntry();
        }

        // Identifica si existe una slide con el nombre de la agenda base (LABEL_AGENDA + "0") y crea la referencia.     
        public bool cargarAgenda()
        {
            try
            {
                if (!agendaCargada && agenda == null)
                {
                    Presentation pptActiva = Globals.ThisAddIn.Application.ActivePresentation;
                    foreach (Slide sl in pptActiva.Slides)
                        if (sl.Name.Equals(LABEL_AGENDA + 1))
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
                if (agenda != null)
                {
                    try
                    {
                        int idAgenda = agenda.SlideIndex;
                        res = true;
                    }
                    catch (Exception e)
                    {

                        // Aquí en caso de error se eliminaba la agenda - ver código comentado.
                        MessageBox.Show("Hubo un error al cargar la agenda.");

                        DialogResult respuesta = MessageBox.Show("Al parecer se ha borrado la diapositiva base de la agenda.\n¿Quieres eliminar el resto de las diapositivas de agenda?.", "Eliminar agenda", MessageBoxButtons.YesNo);
                        eliminarAgenda(respuesta);
                    }
                }

                return res;

            }
            catch (Exception e)
            {
                return false;
            }
        }

        // Elimina las diapositivas de agenda
        public void eliminarAgenda(DialogResult respuesta)
        {
            Presentation pptActiva = Globals.ThisAddIn.Application.ActivePresentation;
            Object thisLock = new Object();

            agenda = null;
            itemsAgenda = new Dictionary<string, PowerPoint.Shape>();
            posItemsAgenda = new Dictionary<string, float[]>();
            tituloAgenda = null;
            numItems = 0;

            for (int i = 1; i <= pptActiva.Slides.Count; )
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
                    sl.Name = "noName_" + i;
                }
                if (!eliminado)
                {
                    i++;
                }
            }

            //pptActiva.Application.StartNewUndoEntry();
        }

        // Crea la diapositiva base con sus elementos por default
        public void crearAgenda()
        {
            try
            {
                slideActiva = csmet.getDiapositiva(this.pptActiva);

                if (!cargarAgenda())
                {
                    //Crea una nueva agenda
                    bool creada = false;
                    try
                    {
                        CustomLayout estilo = pptActiva.SlideMaster.CustomLayouts[Int32.Parse(csmet.CONFIG["Agenda_slide_style_1"])];
                        agenda = pptActiva.Slides.AddSlide(slideActiva.SlideIndex + Int32.Parse(csmet.CONFIG["Agenda_slide_insertAfter"]), estilo);
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

                    agenda.Name = LABEL_AGENDA + 1;
                    agenda.Select();

                    // Cuadro titulo
                    float topTitle = float.Parse(csmet.CONFIG["Agenda_title_top"], CultureInfo.InvariantCulture);
                    float leftTitle = float.Parse(csmet.CONFIG["Agenda_title_left"], CultureInfo.InvariantCulture);
                    float widthTitle = float.Parse(csmet.CONFIG["Agenda_title_width"], CultureInfo.InvariantCulture);
                    float heightTitle = float.Parse(csmet.CONFIG["Agenda_title_height"], CultureInfo.InvariantCulture);
                    float marginTitleTop = float.Parse(csmet.CONFIG["Agenda_title_margin_top"], CultureInfo.InvariantCulture);
                    float marginTitleBottom = float.Parse(csmet.CONFIG["Agenda_title_margin_bottom"], CultureInfo.InvariantCulture);
                    float marginTitleLeft = float.Parse(csmet.CONFIG["Agenda_title_margin_left"], CultureInfo.InvariantCulture);
                    float marginTitleRight = float.Parse(csmet.CONFIG["Agenda_title_margin_right"], CultureInfo.InvariantCulture);

                    // Añade Titulo de Agenda
                    string strTituloAgenda = titulo; // csmet.CONFIG["Agenda_title"];
                    horaItemsAgenda = csmet.CONFIG["Agenda_hora"].Equals("YES");
                    try
                    {
                        agenda.Shapes.Title.Name = LABEL_TITULO_AGENDA;
                        agenda.Shapes.Title.TextFrame.TextRange.Text = strTituloAgenda;
                    }
                    catch (Exception e2)
                    {
                        tituloAgenda = agenda.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, leftTitle, topTitle, widthTitle, heightTitle);
                        tituloAgenda.TextFrame.TextRange.Font.Size = 16;
                    }

                    /*
                    OptionPane oP = new OptionPane();
                    oP.inicializar("¿Deseas incluir la hora de cada ítem en la agenda?","Incluir hora en agenda");
                    oP.ShowDialog();
                    horaItemsAgenda = oP.darResultado();
                    */

                    agendaForm = new Agenda();
                    agendaForm.inicializar(barraMatrix, agenda, strTituloAgenda, this);

                    // Añade linea que une pelotas
                    linea = agenda.Shapes.AddLine(csmet.centimetersToPoints(2.45), csmet.centimetersToPoints(5.36), csmet.centimetersToPoints(2.45), csmet.centimetersToPoints(5.36)); //Última medida es el alto de las pelotas
                    linea.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(csmet.getColorRGB_Matrix("gris"));
                    linea.Line.Weight = float.Parse("5");
                    linea.ZOrder(MsoZOrderCmd.msoSendToBack);
                    linea.Name = "lineaTemplate" + id;

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
                agendaForm.inicializar(barraMatrix, agenda, titulo, this);

                foreach (Microsoft.Office.Interop.PowerPoint.Shape sh in agenda.Shapes)
                {
                    if (sh.Name.StartsWith(LABEL_ITEM_AGENDA))
                    {
                        agendaForm.agregarItem();
                    }
                }
                agendaForm.Show();

                //pptActiva.Application.StartNewUndoEntry();

            }
            catch (Exception ex)
            {
                MessageBox.Show("Hubo un problema al crear la agenda:\n" + "(" + ex.Message + ")");
            }

        }

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
                float gap_aux = (4 * float.Parse(csmet.CONFIG["Agenda_item_gapToNext"], CultureInfo.InvariantCulture)) / (index - 1);
                for (int i = 2; i <= index; i++)
                {
                    Microsoft.Office.Interop.PowerPoint.Shape modificar = agenda.Shapes[LABEL_ITEM_AGENDA + i];
                    modificar.Top = float.Parse(csmet.CONFIG["Agenda_item_top"], CultureInfo.InvariantCulture) + (i - 1) * gap_aux;
                }
            }
        }

        // Elimina un ítem de la agenda
        public void eliminarItemAgenda(int index, bool eliminarShape)
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
                if (numItems <= 5)
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
                        modificar.Name = modificar.Name.Substring(0, modificar.Name.Length - 1).Replace(antiguo, nuevo) + modificar.Name.Substring(modificar.Name.Length - 1);
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
                        modificar.Top = float.Parse(csmet.CONFIG["Agenda_item_top"], CultureInfo.InvariantCulture) + (i - 1) * gap_aux;
                    }
                }

                if (agendaForm != null)
                    agendaForm.eliminarItem();

                //pptActiva.Application.StartNewUndoEntry();
            }
        }

        // Actualiza todas las diapositivas de agenda según la agenda base.
        public void actualizarAgenda()
        {
            pptActiva = Globals.ThisAddIn.Application.ActivePresentation;

            int indexBase = agenda.SlideIndex;
            int[] indexPosiciones = new int[itemsAgenda.Count];
            int maxIndex = indexBase + 1;

            for (int i = (tienePortada ? 1 : 2); i <= itemsAgenda.Count; i++)
            {
                bool eliminada = false;
                foreach (Slide sl in pptActiva.Slides)
                {
                    //MessageBox.Show(sl.Name + " = " + LABEL_AGENDA + i);
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

                newSlide.MoveTo(indexPosiciones[i - 1]);
            }

            foreach (Slide sl in pptActiva.Slides)
            {
                if (sl.Name.StartsWith(LABEL_AGENDA + "eliminar"))
                {
                    sl.Delete();
                }
            }
            foreach (Slide sl1 in pptActiva.Slides)
            {
                if (sl1.Name.StartsWith(LABEL_AGENDA + "eliminar"))
                {
                    sl1.Delete(); 
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
    }
}
