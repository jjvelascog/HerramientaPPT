using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.IO;
using System.Windows.Forms;
using System.Globalization;

namespace Herramientas.Clases
{
    class metodos
    {
        public Dictionary<string,string> CONFIG = new Dictionary<string, string>();
        public static string PATH_TO_FILE = "C:\\ExtraGeek\\MATRIXppt\\config.txt";
        public static string SUBPATH = "C:\\ExtraGeek\\MATRIXppt";
        public static string VERSION = "1.0.0.0.5.1";

        private static metodos instance;

        private metodos()
        {

            if (!File.Exists(PATH_TO_FILE))
            {
                System.IO.Directory.CreateDirectory(SUBPATH);
                cargarValoresPorDefecto();
            }
            else
            {
                foreach (var row in File.ReadAllLines(PATH_TO_FILE))
                    CONFIG.Add(row.Split('=')[0], string.Join("=", row.Split('=').Skip(1).ToArray()));
            }
            
            if(!CONFIG.ContainsKey("version") || !CONFIG["version"].Equals(VERSION))
            {
                CONFIG = new Dictionary<string, string>();
                cargarValoresPorDefecto();
            }

            //Logger
            string path = CONFIG["LOGFILE_PATH"] + CONFIG["LOGFILE_FILE"] + CONFIG["LOGFILE_COUNTER"]+".txt";
            // This text is added only once to the file.
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("USER;DATE;CONTROL;VALOR");
                }
            }

            //FrasesYLinkdelDia
            path = CONFIG["FRASEDELDIA_PATH"];
            // This text is added only once to the file.
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("\"La experiencia es un maestro feroz, pero está claro que te hace aprender.\" Clive Staples Lewis");
                    sw.WriteLine("\"Supongo que solo aquéllos que no hacen nada están libres cometer errores.\" Joseph Conrad");
                    sw.WriteLine("\"Estoy tan orgulloso de lo que no hacemos como de lo que hacemos.\" Steve Jobs");
                    sw.WriteLine("\"Hay una fuerza motriz más poderosa que el vapor, la electricidad y la energía atómica: la voluntad.\" Albert Einstein");
                    sw.WriteLine("\"La posibilidad de realizar un sueño es lo que hace que la vida sea interesante.\" Paulo Coelho");
                    sw.WriteLine("\"A los mayores les gustan las cifras. Cuando se les habla de un nuevo amigo, jamás preguntan sobre lo esencial del mismo. Nunca se les ocurre preguntar: \"¿Qué tono tiene su voz? ¿Qué juegos prefiere? ¿Le gusta coleccionar mariposas?\" Pero en cambio preguntan: \"¿Qué edad tiene? ¿Cuántos hermanos? ¿Cuánto pesa? ¿Cuánto gana su padre?\" Solamente con estos detalles creen conocerle.\" El Principito");
                    sw.WriteLine("\"Todos somos muy ignorantes. Lo que ocurre es que no todos ignoramos las mismas cosas.\" Albert Einstein");
                    sw.WriteLine("\"Cuando se innova, se corre el riesgo de cometer errores. Es mejor admitirlo rápidamente y continuar con otra innovación.\" Steve Jobs");
                    sw.WriteLine("\"El éxito es un pésimo maestro que seduce a la gente a pensar que no puede perder.\" Bill Gates");
                    sw.WriteLine("\"Si yo ordenara -decía frecuentemente-, si yo ordenara a un general que se transformara en ave marina y el general no me obedeciese, la culpa no sería del general, sino mía.\" El Principito");
                    sw.WriteLine("\"Tus clientes más descontentos son tu mayor fuente de aprendizaje.\" Bill Gates");
                    sw.WriteLine("\"Es mejor ser pirata que alistarse en la marina.\" Steve Jobs");
                    sw.WriteLine("\"Un emprendedor ve oportunidades allá donde otros solo ven problemas.\" Michael Gerber");
                    sw.WriteLine("\"Las personas no son recordadas por el número de veces que fracasan, sino por el número de veces que tienen éxito.\" Thomas Alva Edison");
                    sw.WriteLine("\"Es un error capital teorizar antes de poseer datos. Uno comienza a alterar los hechos para encajarlos en las teorías, en lugar de encajar las teorías en los hechos.\" Sherlock Holmes");

                }
            }

            path = CONFIG["LINKDELDIA_PATH"];
            // This text is added only once to the file.
            if (!File.Exists(path))
            {
                // Create a file to write to.
                using (StreamWriter sw = File.CreateText(path))
                {
                    sw.WriteLine("http://www.matrixconsulting.com/");
                    sw.WriteLine("http://www.siliconvalley.com/");
                    sw.WriteLine("http://www.bcentral.cl/");
                    sw.WriteLine("http://www.ted.com/");
                    sw.WriteLine("http://www.similarsites.com/");
                    sw.WriteLine("http://www.woorank.com/");
                }
            }

        }

        public static metodos Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new metodos();
                }
                return instance;
            }
        }


        /*clases personalizadas*/

        /// <summary>
        /// Clase datos de tamaño a escala de formas ghost y smartghost
        /// </summary>
        public class Escala
        {
            public float Width { get; set; }
            public float Height { get; set; }
        }

        public string darFrase(int frase)
        {
            return File.ReadAllLines(CONFIG["FRASEDELDIA_PATH"])[frase];
        }

        public string darLink(int link)
        {
            return File.ReadAllLines(CONFIG["LINKDELDIA_PATH"])[link];
        }

        /// <summary>
        /// Devuelve colores predeterminados Matrix Consulting {rojo,rojoOscuro,negro,gris,grisOscuro,amarillo,negroBase}
        /// </summary>
        /// <param name="color"></param>
        /// <returns></returns>
        public Color getColorRGB_Matrix(string color)
        {

            Color clr = Color.FromArgb(0, 0, 0); //Default

            switch (color)
            {
                case "rojo":
                    clr = Color.FromArgb(204, 0, 0);
                    break;
                case "rosadoOscuro":
                    clr = Color.FromArgb(255, 130, 130);
                    break;
                case "rosadoMedio":
                    clr = Color.FromArgb(255, 190, 190);
                    break;
                case "rosadoClaro":
                    clr = Color.FromArgb(255, 230, 230);
                    break;
                case "negro":
                    clr = Color.FromArgb(60, 60, 60);
                    break;
                case "gris":
                    clr = Color.FromArgb(235, 235, 235);
                    break;
                case "grisSombra":
                    clr = Color.FromArgb(200, 200, 200);
                    break;
                case "grisOscuro":
                    clr = Color.FromArgb(150, 150, 150);
                    break;
                case "grisPlomo":
                    clr = Color.FromArgb(50, 50, 50);
                    break;
                case "amarillo":
                    clr = Color.FromArgb(255, 192, 0);
                    break;
                case "negroBase":
                    clr = Color.FromArgb(0, 0, 0);
                    break;
                case "blanco":
                    clr = Color.FromArgb(255, 255, 255);
                    break;
            }

            return clr;

        }

        /// <summary>
        /// Obtiene diapositiva Activa
        /// </summary>
        /// <param name="pptActiva"></param>
        /// <returns></returns>
        public PowerPoint.Slide getDiapositiva(PowerPoint.Presentation pptActiva)
        {
            PowerPoint.Slides slides; //objeto todas las diapo
            PowerPoint.Slide slide; //objeto una diapo                

            slides = pptActiva.Slides;
            slide = slides[pptActiva.Application.ActiveWindow.Selection.SlideRange.SlideNumber];

            return slide;
        }

        /// <summary>
        /// Escala medida se imagenes para crear Ghost en la parte superior y este no se deforme
        /// </summary>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <returns></returns>
        public Escala getEscalaGhost(float width, float height, float W, float H)
        {

            Escala esc = new Escala();

            if (height / H < width / W)
            {
                esc.Width = W;
                esc.Height = height / width * W;
            }
            else
            {
                esc.Width = width / height * H;
                esc.Height = H;
            }

            return esc;

        }

        /// <summary>
        /// Valida que no exista el shapeName creado
        /// </summary>
        /// <param name="activa"></param>
        /// <param name="nomShape"></param>
        /// <returns></returns>
        public bool getValidaGrupoShape(PowerPoint.Slide activa, string nomShape)
        {
            bool ret = false;

            foreach (PowerPoint.Shape item in activa.Shapes)
            {
                if (item.Name == nomShape)
                {
                    ret = true;
                }
            }

            return ret;
        }

        /// <summary>
        /// Obtiene imaganes para los iconos en la barra de herramientas Matrix
        /// </summary>
        /// <param name="control"></param>
        /// <returns></returns>
        public Bitmap getImage(string control)
        {
            switch (control)
            {
                case "men_disclaimer":
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
                    return new Bitmap(Properties.Resources.Disclaimer);
                    
                case "btnGhostSimple":
                case "grpGhosts":
                    return new Bitmap(Properties.Resources.ghost_28x28_matrix_gris);

                case "btnSmartGhost":
                    return new Bitmap(Properties.Resources.ghost_smart_28x28_matrix_2);

                case "btnAlto":
                    return new Bitmap(Properties.Resources.base_icon_28x28_matrix_mismo_alto);

                case "grpAjustar":
                case "btnAncho":
                    return new Bitmap(Properties.Resources.base_icon_28x28_matrix_mismo_ancho);
                    
                case "grpBullets":
                case "men_CreaBullet":
                case "btnOp3":
                    return new Bitmap(Properties.Resources.base_icon_28x28_matrix_insertar_texto);

                case "grpPredet":
                case "btnTake":
                    return new Bitmap(Properties.Resources.Conclusion);

                case "btnDarformato":
                    return new Bitmap(Properties.Resources.base_icon_28x28_matrix_dar_formato);
                case "btnConfig":
                case "grpConfig":
                    return new Bitmap(Properties.Resources.btnConfig);
                case "btnEmpalmaV":
                    return new Bitmap(Properties.Resources.empalmarV);
                case "btnEmpalmaH":
                    return new Bitmap(Properties.Resources.empalmarH);
                case "btnNotaAlPie":
                    return new Bitmap(Properties.Resources.Fuente);
                case "btnCopyPos":
                    return new Bitmap(Properties.Resources.copiarPosiciones);
                case "btnPastePos":
                    return new Bitmap(Properties.Resources.pegarPosiciones);
                case "btnDestacar":
                case "grpDestacarSombrear":
                    return new Bitmap(Properties.Resources.destacar);
                case "btnSombrear":
                    return new Bitmap(Properties.Resources.sombrear);
                case "btnCrearAgenda":
                case "grpAgenda":
                    return new Bitmap(Properties.Resources.Agenda);
                case "btnActualizarAgenda":
                    return new Bitmap(Properties.Resources.actAgenda);
                case "btnEliminarAgenda":
                    return new Bitmap(Properties.Resources.BorrarAgenda);
                case "btnFraseDia":
                case "grpFraseYLinkDelDia":
                    return new Bitmap(Properties.Resources.frase);
                case "btnLinkDia":
                    return new Bitmap(Properties.Resources.link);
                case "grpAjustarEspaciado":
                case "btnAddEspaciadoV":
                    return new Bitmap(Properties.Resources.espaciadoAV);
                case "btnSubEspaciadoV":
                    return new Bitmap(Properties.Resources.espaciadoSV);
                case "btnAddEspaciadoH":
                    return new Bitmap(Properties.Resources.espaciadoAH);
                case "btnSubEspaciadoH":
                    return new Bitmap(Properties.Resources.espaciadoSH);
                case "btnSeleccionarSim":
                case "grpSeleccionar":
                    return new Bitmap(Properties.Resources.similares);
                case "btncn1":
                case "btncn2":
                case "btncn4":
                case "btncn8":
                case "btncni":
                case "btncnii":
                case "btncniv":
                case "btncnviii":
                    return new Bitmap(Properties.Resources.cn);
                case "men_cn":
                    return new Bitmap(Properties.Resources.Circulos);
                case "btnCallout":
                    return new Bitmap(Properties.Resources.callout);
                case "btnTitulo":
                    return new Bitmap(Properties.Resources.Titulo);
                case "men_cajas":
                    return new Bitmap(Properties.Resources.Cajas);
                case "btnCajas2":
                    return new Bitmap(Properties.Resources.cajas2);
                case "btnCajas3":
                    return new Bitmap(Properties.Resources.Cajas);
                case "btnCajas4":
                    return new Bitmap(Properties.Resources.cajas4);
                case "men_estados":
                    return new Bitmap(Properties.Resources.Estados);
                case "btnIgual":
                    return new Bitmap(Properties.Resources.icono_igual);
                case "btnCheckBlanco":
                    return new Bitmap(Properties.Resources.icono_check_blanco);
                case "btnCheck":
                    return new Bitmap(Properties.Resources.icono_check);
                case "btnExclamacion":
                    return new Bitmap(Properties.Resources.icono_exclamacion);
                case "btnCruz":
                    return new Bitmap(Properties.Resources.icono_cruz);
                case "men_calendario":
                    return new Bitmap(Properties.Resources.calendario);
                case "btnChangePos":
                    return new Bitmap(Properties.Resources.cambiarPosiciones);
                case "btnMostrarHistoria":
                    return new Bitmap(Properties.Resources.historia);
                case "btnAbrirTemplate":
                    return new Bitmap(Properties.Resources.template);
                case "btnTrofeos":
                    return new Bitmap(Properties.Resources.extrageek_peq);
            }
            return null;

        }


        /* Métodos para configurar formatos */


        public float[] getSizeCuadroTextoBullet()
        {
            float[] sizeCuadroTexto = new float[4];
            sizeCuadroTexto[0] = float.Parse(CONFIG["Bllts_Top"], CultureInfo.InvariantCulture);
            sizeCuadroTexto[1] = float.Parse(CONFIG["Bllts_Left"], CultureInfo.InvariantCulture);
            sizeCuadroTexto[2] = float.Parse(CONFIG["Bllts_Width"], CultureInfo.InvariantCulture);
            sizeCuadroTexto[3] = float.Parse(CONFIG["Bllts_Height"], CultureInfo.InvariantCulture);
            return sizeCuadroTexto;
        }

        public float[] getSizeNotaAlPie()
        {
            float[] sizeCuadroTexto = new float[4];
            sizeCuadroTexto[0] = float.Parse(CONFIG["NotaAlPie_Top"], CultureInfo.InvariantCulture);
            sizeCuadroTexto[1] = float.Parse(CONFIG["NotaAlPie_Left"], CultureInfo.InvariantCulture);
            sizeCuadroTexto[2] = float.Parse(CONFIG["NotaAlPie_Width"], CultureInfo.InvariantCulture);
            sizeCuadroTexto[3] = float.Parse(CONFIG["NotaAlPie_Height"], CultureInfo.InvariantCulture);
            return sizeCuadroTexto;
        }

        internal void formatBulletLevel(int level, PowerPoint.TextRange paragraph, TextRange2 paragraph2, float fontSize, string refMargen)
        {
            paragraph2.ParagraphFormat.IndentLevel = level;
            paragraph.ParagraphFormat.Alignment = (PowerPoint.PpParagraphAlignment)Int32.Parse(CONFIG["Bllts_aligment_" + level]);
            paragraph.ParagraphFormat.Bullet.Font.Name = CONFIG["Bllts_bulletFont_" + level];
            if (level == 1)
            {
                paragraph.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
            }
            else
            {
                paragraph.ParagraphFormat.Bullet.Character = (char)Int32.Parse(CONFIG["Bllts_bulletChar_" + level]);         //marlett
            }
            paragraph.Font.Size = fontSize;
            float lI = float.Parse(CONFIG["Bllts_margin_" + refMargen + "_Left_" + level], CultureInfo.InvariantCulture);
            paragraph2.ParagraphFormat.LeftIndent = lI > 100? lI/10:lI;
            paragraph2.ParagraphFormat.FirstLineIndent = Int32.Parse(CONFIG["Bllts_firstLineIndent_" + level]);
            float sB = float.Parse(CONFIG["Bllts_SpaceBefore_" + level], CultureInfo.InvariantCulture);
            paragraph2.ParagraphFormat.SpaceBefore = sB > 100? sB/10: sB; //0.15
        }


        internal void adjustFormatBulletLevel(int level, PowerPoint.TextRange paragraph, TextRange2 paragraph2, float fontSize, string refMargen)
        {
            paragraph2.ParagraphFormat.IndentLevel = level;
            paragraph.ParagraphFormat.Alignment = (PowerPoint.PpParagraphAlignment)Int32.Parse(CONFIG["Bllts_aligment_" + level]);
            paragraph.ParagraphFormat.Bullet.Font.Name = CONFIG["Bllts_bulletFont_" + level];            

            if (level == 1)
            {
                paragraph.ParagraphFormat.Bullet.Type = PowerPoint.PpBulletType.ppBulletUnnumbered;
            }
            else
            {
                paragraph.ParagraphFormat.Bullet.Character = (char)Int32.Parse(CONFIG["Bllts_bulletChar_" + level]);
            }

            if(CONFIG["Bllts_ajustarTam"].Equals("YES"))
                paragraph.Font.Size = fontSize;

            if (CONFIG["Bllts_ajustarSangria"].Equals("YES"))
            {
                float lI = float.Parse(CONFIG["Bllts_margin_" + refMargen + "_Left_" + level], CultureInfo.InvariantCulture);
                paragraph2.ParagraphFormat.LeftIndent = lI > 100?lI/10:lI;
                paragraph2.ParagraphFormat.FirstLineIndent = Int32.Parse(CONFIG["Bllts_firstLineIndent_" + level]);
            }

            if (CONFIG["Bllts_ajustarEspaciado"].Equals("YES"))
            {
                float sB = float.Parse(CONFIG["Bllts_SpaceBefore_" + level], CultureInfo.InvariantCulture);
                paragraph2.ParagraphFormat.SpaceBefore = sB > 100? sB/10:sB; 
            }
        }

        internal string getFontNameCuadroTextoBullets()
        {
            return CONFIG["MainFont"];
        }

        internal int identificarTamanoLetra(float tamano)
        {

            string[] tamanhos = CONFIG["Bllts_tamanos"].Split(';');

            //Identificar tamaño del primer nivel

            int j = 0;
            while (float.Parse(tamanhos[j], CultureInfo.InvariantCulture) > tamano)
            {
                j++;
            }

            return j;
        }

        /* Input dialog */
        public string ShowDialog(string text, string caption)
        {
            Form prompt = new Form();
            prompt.Width = 400;
            prompt.Height = 135;
            prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
            prompt.Text = caption;
            prompt.StartPosition = FormStartPosition.CenterScreen;
            Label textLabel = new Label() { Left = 50, Top = 10, Text = text, Width = 300 };
            TextBox textBox = new TextBox() { Left = 50, Top = 30, Width = 300 };
            Button confirmation = new Button() { Text = "Ok", Left = 250, Width = 100, Top = 60, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : "";
        }

        /* Input dialog (number) */
        public DateTime ShowDialogNum(string text, string caption)
        {
            Form prompt = new Form();
            prompt.Width = 400;
            prompt.Height = 160;
            prompt.FormBorderStyle = FormBorderStyle.FixedDialog;
            prompt.Text = caption;
            prompt.StartPosition = FormStartPosition.CenterScreen;
            Label textLabel = new Label() { Left = 50, Top = 10, Text = text, Width = 300 };
            //NumericUpDown inputBox = new NumericUpDown () { Left = 50, Top = 30, Width = 300 };
            DateTimePicker inputBox = new DateTimePicker() { Left = 50, Top = 30, Width = 300, Value =  GetNextWeekday(DateTime.Today, DayOfWeek.Monday)};
            Button confirmation = new Button() { Text = "Ok", Left = 250, Width = 100, Top = 60, DialogResult = DialogResult.OK};
            confirmation.Click += (sender, e) => {
                    prompt.Close();
            };

            inputBox.ValueChanged += (sender, e) =>
            {
                if (inputBox.Value.DayOfWeek != DayOfWeek.Monday)
                {
                    System.Windows.Forms.MessageBox.Show("Debe seleccionar un lunes");
                    inputBox.Value = GetNextWeekday(DateTime.Today, DayOfWeek.Monday);
                }
            };

            prompt.Controls.Add(inputBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);

            return prompt.ShowDialog() == DialogResult.OK ? inputBox.Value : DateTime.MinValue;
        }

        /* message box con altura máxima */
        public void CustomMessageBox(string text, string caption)
        {
            Form form1 = new Form() { Height = 500, Width = 600};
            Button button1 = new Button() { Text = "Ok", Left = 450, Width = 100, Top = 400, DialogResult = DialogResult.OK };
            Panel panel = new Panel() { Top = 10, Left = 10, Width = 580, Height = 380, AutoScroll = true, Padding = new Padding(20)};
            Label texto = new Label() { Top = 0, Width = 0, AutoSize = true, Text = text, MaximumSize = new Size(550, 0), Font = new Font("Lao UI", 10) };   

            form1.Text = caption;

            form1.FormBorderStyle = FormBorderStyle.FixedDialog;
            form1.AcceptButton = button1;
            form1.StartPosition = FormStartPosition.CenterScreen;

            form1.Controls.Add(button1);
            form1.Controls.Add(panel);
            panel.Controls.Add(texto);

            form1.ShowDialog();

            button1.Click += (sender, e) =>
            {
                form1.Dispose();
            };
        }

        /* Obtener lunes siguiente */
        public static DateTime GetNextWeekday(DateTime start, DayOfWeek day)
        {
            // The (... + 7) % 7 ensures we end up with a value in the range [0, 6]
            int daysToAdd = ((int)day - (int)start.DayOfWeek + 7) % 7;
            return start.AddDays(daysToAdd);
        }

        /* Metodo para guardar configuración */

        private void cargarValoresPorDefecto()
        {
            //Versión
            CONFIG.Add("version", VERSION);

            //Logger
            CONFIG.Add("LOGFILE_PATH", SUBPATH);
            CONFIG.Add("LOGFILE_FILE", "\\Log_");
            CONFIG.Add("LOGFILE_COUNTER", DateTime.Now.ToString("yyyyMMddHHmm"));

            //Fuente de texto principal
            CONFIG.Add("MainFont", "Lao UI");

            //Frase del día y link del día
            CONFIG.Add("FRASEDELDIA_PATH", SUBPATH + "\\fdd");
            CONFIG.Add("FRASEDELDIA_COUNT", "15");
            CONFIG.Add("LINKDELDIA_PATH", SUBPATH + "\\ldd");
            CONFIG.Add("LINKDELDIA_COUNT", "6");

            //Bullets ajustar formato
            CONFIG.Add("Bllts_ajustarTam", "YES");
            CONFIG.Add("Bllts_ajustarSangria", "YES");
            CONFIG.Add("Bllts_ajustarEspaciado", "YES");

            CONFIG.Add("Bllts_aligment_1", "1");
            CONFIG.Add("Bllts_aligment_2", "1");
            CONFIG.Add("Bllts_aligment_3", "1");
            CONFIG.Add("Bllts_bulletFont_1", "Lao UI");
            CONFIG.Add("Bllts_bulletFont_2", "Lao UI");
            CONFIG.Add("Bllts_bulletFont_3", "Lao UI");
            CONFIG.Add("Bllts_bulletChar_1", "152");
            CONFIG.Add("Bllts_bulletChar_2", "8722"); // 45
            CONFIG.Add("Bllts_bulletChar_3", "9642"); // 9642
            CONFIG.Add("Bllts_firstLineIndent_1", "-13"); //13
            CONFIG.Add("Bllts_firstLineIndent_2", "-13"); //13
            CONFIG.Add("Bllts_firstLineIndent_3", "-13"); //13

            float erIndet = float.Parse("22.5", CultureInfo.InvariantCulture); //valor inicial de margen bullet
            CONFIG.Add("Bllts_margin_48_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_48_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_48_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_44_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_44_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_44_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_40_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_40_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_40_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_36_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_36_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_36_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_32_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_32_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_32_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_28_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_28_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_28_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_24_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_24_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_24_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_20_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_20_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_20_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_18_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_18_Left_2", "" + (erIndet + 15));
            CONFIG.Add("Bllts_margin_18_Left_3", "" + (erIndet + 35));
            CONFIG.Add("Bllts_margin_16_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_16_Left_2", "" + (erIndet + 6));
            CONFIG.Add("Bllts_margin_16_Left_3", "" + (erIndet + 20));
            CONFIG.Add("Bllts_margin_14_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_14_Left_2", "" + (erIndet + 6));
            CONFIG.Add("Bllts_margin_14_Left_3", "" + (erIndet + 20));
            CONFIG.Add("Bllts_margin_12_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_12_Left_2", "" + (erIndet + float.Parse("5.3", CultureInfo.InvariantCulture))); //10
            CONFIG.Add("Bllts_margin_12_Left_3", "" + (erIndet + 20));
            CONFIG.Add("Bllts_margin_11_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_11_Left_2", "" + (erIndet + 5));
            CONFIG.Add("Bllts_margin_11_Left_3", "" + (erIndet + 20));
            CONFIG.Add("Bllts_margin_10.5_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_10.5_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_10.5_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_10_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_10_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_10_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_9_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_9_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_9_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_8_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_8_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_8_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_7_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_7_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_7_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_6_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_6_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_6_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_5_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_5_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_5_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_4_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_4_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_4_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_3_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_3_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_3_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_2_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_2_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_2_Left_3", "" + (erIndet + 29));
            CONFIG.Add("Bllts_margin_1_Left_1", "" + (erIndet - float.Parse("9.5", CultureInfo.InvariantCulture)));
            CONFIG.Add("Bllts_margin_1_Left_2", "" + (erIndet + 9));
            CONFIG.Add("Bllts_margin_1_Left_3", "" + (erIndet + 29));

            CONFIG.Add("Bllts_Top", "99.8");
            CONFIG.Add("Bllts_Left", "19.7");
            CONFIG.Add("Bllts_Width", "262.9");
            CONFIG.Add("Bllts_Height", "72.3");
            

            CONFIG.Add("Bllts_tamanos", "48;44;40;36;32;28;24;20;18;16;14;12;11;10.5;10;9;8;7;6;5;4;3;2;1");
            
            CONFIG.Add("Bllts_SpaceBefore_1", "6");
            CONFIG.Add("Bllts_SpaceBefore_2", "2");
            CONFIG.Add("Bllts_SpaceBefore_3", "0");


            //Disclaimers
            CONFIG.Add("Disclaimer_lbl_btnD_pd", "PARA DISCUSIÓN");
            CONFIG.Add("Disclaimer_lbl_btnD_Pr", "PRELIMINAR");
            CONFIG.Add("Disclaimer_lbl_btnD_Il", "ILUSTRATIVO");
            CONFIG.Add("Disclaimer_lbl_btnD_EJ", "EJEMPLO");
            CONFIG.Add("Disclaimer_lbl_btnD_Cn", "CONFIDENCIAL");
            CONFIG.Add("Disclaimer_lbl_btnD_Bk", "BACK - UP");
            CONFIG.Add("Disclaimer_lbl_btnD_Ne", "NO EXHAUSTIVO");
            CONFIG.Add("Disclaimer_top", "70.015");
            CONFIG.Add("Disclaimer_height", "23.75");
            CONFIG.Add("Disclaimer_font_size", "14");
            CONFIG.Add("Disclaimer_font_bold", "1");
            CONFIG.Add("Disclaimer_gap", "13.89");
            CONFIG.Add("Disclaimer_lsup_top", "87.78");
            CONFIG.Add("Disclaimer_linf_top", "90.71");
            CONFIG.Add("Disclaimer_line_shape", "10001");
            CONFIG.Add("Disclaimer_line_width", "2.25");
            CONFIG.Add("Disclaimer_line_style", "1");

            // Nota al pie
            CONFIG.Add("NotaAlPie_Top", "517.89");
            CONFIG.Add("NotaAlPie_Left", "0");
            CONFIG.Add("NotaAlPie_Width", "705.83");
            CONFIG.Add("NotaAlPie_Height", "22.11"); 
            CONFIG.Add("NotaAlPie_Font_size", "8");
            CONFIG.Add("NotaAlPie_Font_bold", "0");
            CONFIG.Add("NotaAlPie_aligment", "1");
            CONFIG.Add("NotaAlPie_MarginBottom", "13.04");
            CONFIG.Add("NotaAlPie_MarginTop", "0");
            CONFIG.Add("NotaAlPie_MarginLeft", "14.17");
            CONFIG.Add("NotaAlPie_MarginRight", "5.67");

            //Takeaway
            CONFIG.Add("Takeaway_animate", "YES");

            CONFIG.Add("Takeaway_linea_width", "345.26");
            CONFIG.Add("Takeaway_linea_height", "1.7");
            CONFIG.Add("Takeaway_linea_top", "450.42");
            CONFIG.Add("Takeaway_linea_left", "14.17");
            CONFIG.Add("Takeaway_linea_background", "rojo");
            CONFIG.Add("Takeaway_linea_lineColor", "negroBase");
            CONFIG.Add("Takeaway_linea_lineWidth", "0");

            CONFIG.Add("Takeaway_rectangulo_width", "12.76");
            CONFIG.Add("Takeaway_rectangulo_height", "51.3");
            CONFIG.Add("Takeaway_rectangulo_top", "450.43");
            CONFIG.Add("Takeaway_rectangulo_left", "14.17");
            CONFIG.Add("Takeaway_rectangulo_background", "rojo");
            CONFIG.Add("Takeaway_rectangulo_lineColor", "negroBase");
            CONFIG.Add("Takeaway_rectangulo_lineWidth", "0");

            CONFIG.Add("Takeaway_cuadrado_width", "12.76");
            CONFIG.Add("Takeaway_cuadrado_height", "12.76");
            CONFIG.Add("Takeaway_cuadrado_top", "489.26");
            CONFIG.Add("Takeaway_cuadrado_left", "14.17");
            CONFIG.Add("Takeaway_cuadrado_background", "rosadoMedio");
            CONFIG.Add("Takeaway_cuadrado_lineColor", "negroBase");
            CONFIG.Add("Takeaway_cuadrado_lineWidth", "0");
            
            CONFIG.Add("Takeaway_text_width", "650.55");
            CONFIG.Add("Takeaway_text_height", "38.83");
            CONFIG.Add("Takeaway_text_top", "456.66");
            CONFIG.Add("Takeaway_text_left", "26.93");
            CONFIG.Add("Takeaway_text_margin", "7.09");
            CONFIG.Add("Takeaway_text_foreColor", "grisPlomo");
            
            CONFIG.Add("Takeaway_text_fontSize", "16");

            //Titulo
            CONFIG.Add("Titulo_linea_width", "691.65");
            CONFIG.Add("Titulo_linea_height", "0");
            CONFIG.Add("Titulo_linea_left", "14.17");
            CONFIG.Add("Titulo_linea_top", "117.64");
            CONFIG.Add("Titulo_linea_color", "black");
            CONFIG.Add("Titulo_linea_weight", "1");

            CONFIG.Add("Titulo_text_width", "691.65");
            CONFIG.Add("Titulo_text_height", "22.68");
            CONFIG.Add("Titulo_text_left", "14.17");
            CONFIG.Add("Titulo_text_top", "94.96");
            CONFIG.Add("Titulo_text_foreColor", "black");
            CONFIG.Add("Titulo_text_fontSize", "14");
            CONFIG.Add("Titulo_text_marginLeft", "0");

            //Empalmar
            CONFIG.Add("Empalmar_gapH", "0");
            CONFIG.Add("Empalmar_gapV", "0");

            //Destacar
            CONFIG.Add("Highlight_lineColor", "rojo");
            CONFIG.Add("Highlight_lineWidth", "1.5");
            CONFIG.Add("Highlight_label_width", "22.67");
            CONFIG.Add("Highlight_label_height", "17");
            CONFIG.Add("Highlight_text_marginLeft", "39.68");
            CONFIG.Add("Highlight_text_fontSize", "10");
            CONFIG.Add("Highlight_text_fontColor", "negro");
            CONFIG.Add("Highlight_label_top", "483.12");
            CONFIG.Add("Highlight_label_left", "14.46");
            CONFIG.Add("Highlight_background", "rosadoClaro");

            //Cajas
            CONFIG.Add("Cajas_texto1_top", "99.5");
            CONFIG.Add("Cajas_texto2_top", "151.94");
            CONFIG.Add("Cajas_texto_height", "20.41");
            CONFIG.Add("Cajas_texto_width", "105.16");
            CONFIG.Add("Cajas_texto_foreColor", "grisPlomo");
            CONFIG.Add("Cajas_texto_fontSize", "12");
            CONFIG.Add("Cajas_texto_left", "23.53");
            CONFIG.Add("Cajas_height", "39.12");
            CONFIG.Add("Cajas_left", "138.0472");
            CONFIG.Add("Cajas_top", "99.5");
            CONFIG.Add("Cajas_foreColor", "blanco");
            CONFIG.Add("Cajas2_width", "272.13");
            CONFIG.Add("Cajas2_span", "14.15");
            CONFIG.Add("Cajas3_width", "180");
            CONFIG.Add("Cajas3_span", "9.33");
            CONFIG.Add("Cajas4_width", "134.36");
            CONFIG.Add("Cajas4_span", "7.07");
            CONFIG.Add("Cajas_background", "rojo");
            CONFIG.Add("Cajas_contenido_height", "100");

            //Agenda
            CONFIG.Add("Agenda_numItems", "4");
            CONFIG.Add("Agenda_slide_style_1", "7"); 
            CONFIG.Add("Agenda_slide_style_2", "5"); 
            CONFIG.Add("Agenda_slide_style_3", "2"); 
            CONFIG.Add("Agenda_slide_insertAfter", "1");

            CONFIG.Add("Agenda_title_width", "705.6");
            CONFIG.Add("Agenda_title_height", "36");
            CONFIG.Add("Agenda_title_top", "6.4799");
            CONFIG.Add("Agenda_title_left", "7.2");
            CONFIG.Add("Agenda_title_margin_top", "0");
            CONFIG.Add("Agenda_title_margin_bottom", "0");
            CONFIG.Add("Agenda_title_margin_left", "0");
            CONFIG.Add("Agenda_title_margin_right", "0");
            CONFIG.Add("Agenda_title_fontSize", "30");
            CONFIG.Add("Agenda_title_bold", "NO");
            CONFIG.Add("Agenda_title", "Agenda");

            CONFIG.Add("Agenda_hora", "NO");
            CONFIG.Add("Agenda_portada", "NO"); 

            CONFIG.Add("Agenda_item_width", "581.1");
            CONFIG.Add("Agenda_item_height", "55");
            CONFIG.Add("Agenda_item_top", "117.64");
            CONFIG.Add("Agenda_item_gapToNext", "84.19");
            CONFIG.Add("Agenda_item_left", "124.72");
            CONFIG.Add("Agenda_item_margin", "3.68");
            CONFIG.Add("Agenda_item_leftIndent", "0");
            CONFIG.Add("Agenda_item_leftIndentFirstLine", "36");
            CONFIG.Add("Agenda_item_fontSize", "24");
            CONFIG.Add("Agenda_item_hora_left", "500");
            CONFIG.Add("Agenda_item_hora_width", "100");
            CONFIG.Add("Agenda_item_hora_height", "25.61");
            CONFIG.Add("Agenda_item_hora_margin", "3.68");
            CONFIG.Add("Agenda_item_hora_fontSize", "18");
            CONFIG.Add("Agenda_item_bulletFont", "Marlett");
            CONFIG.Add("Agenda_item_bulletChar", "105");

            CONFIG.Add("Agenda_pelota_width", "37.13");
            CONFIG.Add("Agenda_pelota_height", "37.13");
            CONFIG.Add("Agenda_pelota_top", "114.8");
            CONFIG.Add("Agenda_pelota_left", "50.74");

            CONFIG.Add("Agenda_Highlight_lineColor", "negro");
            CONFIG.Add("Agenda_Highlight_backColor", "blanco");
            CONFIG.Add("Agenda_Highlight_lineWidth", "1");
            CONFIG.Add("Agenda_Highlight_Left", "42.48");
            CONFIG.Add("Agenda_Highlight_GapTop", "12.24");
            CONFIG.Add("Agenda_Highlight_Width", "635.04");
            CONFIG.Add("Agenda_Highlight_Height", "64.8");
            CONFIG.Add("Agenda_Highlight_Shadow_offset", "2.9");
            CONFIG.Add("Agenda_Highlight_Shadow_color", "grisSombra");

            //Frase del día
            CONFIG.Add("FraseDelDia_left", "10");
            CONFIG.Add("FraseDelDia_top", "10");
            CONFIG.Add("FraseDelDia_width", "50");
            CONFIG.Add("FraseDelDia_height", "17");
            CONFIG.Add("FraseDelDia_marginLeft", "0");
            CONFIG.Add("FraseDelDia_fontSize", "11");
            CONFIG.Add("FraseDelDia_fontColor", "negro");

            //Link del día
            CONFIG.Add("LinkDelDia_left", "10");
            CONFIG.Add("LinkDelDia_top", "20");
            CONFIG.Add("LinkDelDia_width", "22.67");
            CONFIG.Add("LinkDelDia_height", "17");
            CONFIG.Add("LinkDelDia_marginLeft", "0");
            CONFIG.Add("LinkDelDia_fontSize", "11");
            CONFIG.Add("LinkDelDia_fontColor", "negro");

            //Espaciado
            CONFIG.Add("Espaciado_mult","0.2");

            //Seleccionar similares
            CONFIG.Add("SeleccionarSimilares_tamanio","YES");
            CONFIG.Add("SeleccionarSimilares_fondo", "NO");
            CONFIG.Add("SeleccionarSimilares_linea", "NO");
            CONFIG.Add("SeleccionarSimilares_tipo", "NO");
            CONFIG.Add("SeleccionarSimilares_preguntar", "YES");

            //Circulos numerados
            CONFIG.Add("CN_width", "28.08");
            CONFIG.Add("CN_height", "28.08");
            CONFIG.Add("CN_left", "-30");
            CONFIG.Add("CN_top", "0");
            CONFIG.Add("CN_gap", "40");
            CONFIG.Add("CN_background", "amarillo");
            CONFIG.Add("CN_lineColor", "blanco");
            CONFIG.Add("CN_foreColor", "negro");
            CONFIG.Add("CN_lineWidth", "2.5");
            CONFIG.Add("CN_Font_size", "12");
            CONFIG.Add("CN_Font_bold", "1");
            CONFIG.Add("CN_aligment", "2");

            //Callout
            CONFIG.Add("Callout_width", "40");
            CONFIG.Add("Callout_height", "50");
            CONFIG.Add("Callout_left", "-100");
            CONFIG.Add("Callout_top", "0");
            CONFIG.Add("Callout_background", "blanco");
            CONFIG.Add("Callout_lineColor", "amarillo");
            CONFIG.Add("Callout_foreColor", "negro");
             CONFIG.Add("Callout_lineWidth", "2.5");
            CONFIG.Add("Callout_Font_size", "11");
            CONFIG.Add("Callout_alignment", "1");

            //Easter eggs (PONER VALORES EN 0 PARA ACTIVAR FUNCION)
            CONFIG.Add("ModoGeek", "2");
            CONFIG.Add("Alineacion_integ1", "2"); //Nico Palma
            CONFIG.Add("Alineacion_integ2", "2");
            CONFIG.Add("Alineacion_integ3", "2"); //R. Marino
            CONFIG.Add("Alineacion_integ4", "2"); //J. Domingo
            CONFIG.Add("Alineacion_integ5", "2"); //I. Fika
            CONFIG.Add("Alineacion_integ6", "2"); //J. Barten
            CONFIG.Add("Alineacion_integ7", "2"); //M.  Schlotfeldt
            CONFIG.Add("Alineacion_integ8", "2"); //A. Valenzuela
            CONFIG.Add("Alineacion_integ9", "2"); //Jaime Siles
            CONFIG.Add("Alineacion_integ10", "2"); //Jürgen Uhlmann
            CONFIG.Add("Alineacion_integ11", "2"); //Agustín Dagnino
            CONFIG.Add("Alineacion_integ12", "2"); //Gabriel Adriazola
            CONFIG.Add("Alineacion_integ13", "2"); //Andrés Oksemberg
            CONFIG.Add("Alineacion_integ14", "2"); 
            CONFIG.Add("Alineacion_integ15", "2"); //Jorge Marín


            CONFIG.Add("Botones_usados", "0000000000000000000");
            CONFIG.Add("Ult_fdd", "");
            CONFIG.Add("num_fdd", "0");
            CONFIG.Add("Ult_ldd", "");
            CONFIG.Add("num_ldd", "0");


            guardarConfig();
        }

        internal void guardarConfig()
        {
            using (StreamWriter file = new StreamWriter(PATH_TO_FILE))
                foreach (var entry in CONFIG)
                    file.WriteLine("{0}={1}", entry.Key, entry.Value);
        }

        public int darGeeksEncontrados()
        {
            int geeksEncontrados = 0;
            for (int i = 1; i <= 15; i++)
            {
                if(i != 2 && i != 14 && CONFIG["Alineacion_integ" + i].Equals("1"))
                    geeksEncontrados++;
            }
            return geeksEncontrados;
        }
        
        public int darTotalGeeks()
        {
            return 13;
        }

        public float centimetersToPoints(double cms)
        {
            return Convert.ToSingle(cms * (72 / 2.54));
        }
    }
}
