
using Microsoft.AspNetCore.Mvc.RazorPages;

using ClosedXML.Excel;
using System.Data.SqlClient;
using System.Data;
using Microsoft.AspNetCore.Mvc;
using DocumentFormat.OpenXml.Spreadsheet;

namespace carlos.Pages
{
    public class IndexModel : PageModel
    {
        private readonly ILogger<IndexModel> _logger;

        public class RowData<T>
        {
            public T[] Data { get; set; }
        }

        public IndexModel(ILogger<IndexModel> logger)
        {
            _logger = logger;
        }

        public void OnGet()
        {
            String conecctionString = "Data Source=.\\sqlexpress;Initial Catalog=transpais_carlos;Integrated Security=True";
            //String excelFilePath = "C:\\Users\\52834\\source\\repos\\carlos\\usuarios.xlsx";

           
        }

        public static List<RowData<T>> ReadExcelData<T>(string filePath, List<string> columnHeaders)
        {
            List<RowData<T>> rowDataList = new List<RowData<T>>();

            // Abrir el archivo Excel con ClosedXML
            using (var workbook = new XLWorkbook(filePath))
            {
                // Obtener la primera hoja del libro
                var worksheet = workbook.Worksheet(1);

                // Leer los títulos de las columnas y guardarlos en el array columnHeaders
                int columnCount = worksheet.ColumnsUsed().Count();
                for (int col = 1; col <= columnCount; col++)
                {
                    var columnHeader = worksheet.Cell(1, col).GetValue<string>();
                    columnHeaders.Add(columnHeader);
                }

                // Leer los datos de cada fila y almacenarlos en el array rowDataList
                int rowCount = worksheet.RowsUsed().Count();
                for (int row = 2; row <= rowCount; row++)
                {
                    RowData<T> rowData = new RowData<T>();
                    rowData.Data = new T[columnCount];

                    for (int col = 1; col <= columnCount; col++)
                    {
                        var cellValue = worksheet.Cell(row, col).GetValue<T>();
                        rowData.Data[col - 1] = cellValue;
                    }

                    rowDataList.Add(rowData);
                }
            }

            return rowDataList;
        }

        public IActionResult OnPost(IFormFile archivo)
        {
            if (archivo != null && archivo.Length > 0)
            {
                // Obtener el nombre del archivo y la extensión
                string nombreArchivo = Path.GetFileName(archivo.FileName);
                string extensionArchivo = Path.GetExtension(nombreArchivo);

                // Validar la extensión del archivo (opcional)
                if (extensionArchivo.ToLower() == ".xls" || extensionArchivo.ToLower() == ".xlsx")
                {
                    String conecctionString = "Data Source=.\\sqlexpress;Initial Catalog=transpais_carlos;Integrated Security=True";
                    string rutaGuardado = Path.Combine("C:\\Users\\52834\\source\\repos\\carlos\\xlsx", nombreArchivo);
                    using (var memoryStream = new MemoryStream())
                    {
                        archivo.CopyTo(memoryStream);
                        using (var workbook = new XLWorkbook(memoryStream))
                        {
                            var worksheet = workbook.Worksheet(1);
                            workbook.SaveAs(rutaGuardado);
                        }
                        List<string> columnHeaders = new List<string>();
                        List<RowData<string>> rowDataList = ReadExcelData<string>(rutaGuardado, columnHeaders);

                        DataTable excelData = new DataTable();
                        excelData.Columns.Add("nombre", typeof(string));
                        excelData.Columns.Add("apellido", typeof(string));
                        excelData.Columns.Add("correo", typeof(string));
                        excelData.Columns.Add("sexo", typeof(string));

                        foreach (var rowData in rowDataList)
                        {

                            string nombre = rowData.Data[0].ToString();
                            string apellido = rowData.Data[1].ToString();
                            string correo = rowData.Data[2].ToString();
                            string sexo = rowData.Data[3].ToString();

                            excelData.Rows.Add(nombre, apellido, correo, sexo);
                        }

                        try
                        {
                            using (SqlConnection connection = new SqlConnection(conecctionString))
                            {
                                connection.Open();
                                using (SqlCommand command = new SqlCommand("sp_InsertIntoUsuariosMultiple", connection))
                                {
                                    command.CommandType = CommandType.StoredProcedure;


                                    SqlParameter parameter = command.Parameters.AddWithValue("@ExcelData", excelData);
                                    parameter.SqlDbType = SqlDbType.Structured;
                                    parameter.TypeName = "dbo.YourExcelTableType"; // Reemplaza con el nombre correcto del tipo de tabla personalizado

                                    command.ExecuteNonQuery();
                                }
                            }
                            Console.WriteLine("estamos");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error: " + ex.Message);
                        }

                        return RedirectToPage("UploadSuccess");
                    }
                }
                else
                {
                    // Archivo con extensión no permitida, mostrar mensaje de error o realizar alguna acción adicional
                    Console.WriteLine("Solo se permiten archivos con extensión .txt o .csv");
                }
            }
            else
            {
                // No se seleccionó ningún archivo, mostrar mensaje de error o realizar alguna acción adicional
                Console.WriteLine("Por favor, seleccione un archivo para cargar.");
            }

            // Si ocurre algún error, mostrar nuevamente la página de carga de archivos
            return Page();
        }

    }
}