using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Proyecto_Final_1
{
    public partial class Form1 : Form
    {
        private readonly HttpClient _httpClient = new HttpClient();
        private const string apiKey = "Colocar la APIKey de OpenAI aquí(no la deje por seguridad)";

        private async void btnInvestigar_Click(object sender, EventArgs e) //Boton de Investigar
        {
            string prompt = txtBuscar.Text;
            if (string.IsNullOrWhiteSpace(prompt)) return;

            string respuesta = await ObtenerRespuesta(prompt);
            txtRespuesta.Text = respuesta;

            if (string.IsNullOrWhiteSpace(prompt) || string.IsNullOrWhiteSpace(respuesta))
            {
                MessageBox.Show("Por favor, ingrese una pregunta y obtenga una respuesta antes de generar el documento.");
                return;
            }

            GuardarEnBaseDeDatos(prompt, respuesta);

            string baseFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "InvestigacionesAI");
            string folderName = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string finalFolder = Path.Combine(baseFolder, folderName);
            Directory.CreateDirectory(finalFolder);

            string wordPath = Path.Combine(finalFolder, "InvestigacionesAI.docx");
            string pptPath = Path.Combine(finalFolder, "InvestigacionesAI.ppt");

            GenerarWord(respuesta, wordPath);
            GenerarPowerPoint(prompt, respuesta, pptPath);

            MessageBox.Show($"Documentos generados en: \n" + finalFolder);
        }



        private async Task<string> ObtenerRespuesta(string prompt) //Obtener Respuesta 
        {
            try
            {
                _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", apiKey);

                var content = new
                {
                    model = "gpt-3.5-turbo",
                    messages = new[]
                    {
                        new { role = "user", content = prompt }
                    }
                };

                var jsonContent = new StringContent(JsonSerializer.Serialize(content), Encoding.UTF8, "application/json");
                var response = await _httpClient.PostAsync("https://api.openai.com/v1/chat/completions", jsonContent);

                if (!response.IsSuccessStatusCode)
                {
                    string errorMsg = await response.Content.ReadAsStringAsync();
                    MessageBox.Show($"Error en la respuesta de la API:\n{response.StatusCode}\n{errorMsg}", "Error API", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return string.Empty;
                }

                var json = await response.Content.ReadAsStringAsync();

                var doc = JsonDocument.Parse(json);
                string result = doc.RootElement
                                       .GetProperty("choices")[0]
                                       .GetProperty("message")
                                       .GetProperty("content")
                                       .GetString();

                return result?.Trim() ?? string.Empty;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ocurrió un error al obtener la respuesta:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return string.Empty;
            }
        }

        private void GuardarEnBaseDeDatos(string prompt, string respuesta) //Guardado en la Base de datos
        {
            string connectionString = "Data Source=ALEJANDROHPVICT\\SQLEXPRESS;Initial Catalog=ProyectoFinal1db;Integrated Security=True";
            string query = "INSERT INTO Investigaciones (Prompt, Respuesta) VALUES (@Pregunta, @Respuesta)";

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@Pregunta", prompt);
                    cmd.Parameters.AddWithValue("@Respuesta", respuesta);
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void GenerarWord(string contenido, string rutaArchivo) //Generar Word
        {
            var wordApp = new Word.Application();
            var documento = wordApp.Documents.Add();
            documento.Content.Text = contenido.Length > 3000 ? contenido.Substring(0, 3000) + "..." : contenido;
            Word.Paragraph parrafo = documento.Content.Paragraphs.Add();
            documento.SaveAs2(rutaArchivo);
            documento.Close();
            wordApp.Quit();
        }

        private void GenerarPowerPoint(string tema, string contenido, string rutaArchivo) //Generar PowerPoint
        {
            var pptApp = new PowerPoint.Application();
            var presentacion = pptApp.Presentations.Add();


            //primera diapositiva (título)
            var slide1 = presentacion.Slides.Add(1, PowerPoint.PpSlideLayout.ppLayoutTitle);
            slide1.Shapes[1].TextFrame.TextRange.Text = "Investigación";
            slide1.Shapes[2].TextFrame.TextRange.Text = tema;

            // Diapositiva 2 - Contenido
            var slide2 = presentacion.Slides.Add(2, PowerPoint.PpSlideLayout.ppLayoutText);
            slide2.Shapes[1].TextFrame.TextRange.Text = "Resultado";
            slide2.Shapes[2].TextFrame.TextRange.Text = contenido.Length > 3000 ? contenido.Substring(0, 3000) + "..." : contenido;

            presentacion.SaveAs(rutaArchivo);
            presentacion.Close();
            pptApp.Quit();
        }







        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
