using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using ExportarArqExcel.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Hosting.Internal;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExportarArqExcel.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class UsuariosController : ControllerBase
    {
        public IHostingEnvironment _hostingEnvironment;

        public UsuariosController(IHostingEnvironment hosting)
        {
            _hostingEnvironment = hosting;
        }

        /// <summary>
        /// Gerar Excel de Usuários
        /// </summary>
        /// <returns>Gera um arquivo .xlsx e disponibiliza o link para download do arquivo pelo navegador</returns>
        [HttpGet("excel")]
        public async Task<IActionResult> GerarExcelUsuarios()
        {
            // Esta propriedade busca o caminho (path) da pasta "wwwroot" para geração do arquivo no final do método
            string webRootPath = _hostingEnvironment.WebRootPath;
            string nomeArquivo = "usuarios.xlsx";
            try
            {
                if (Usuarios.listUsuarios.Count == 0)
                {
                    //Adiciona 5 usuários padrões
                    for (int i = 0; i < 5; i++)
                    {
                        Usuarios usuario = new Usuarios()
                        {
                            Nome = $"Usuario{i + 1}",
                            Senha = "usuario123",
                            Email = $"usuario{i + 1}@email.com",
                            Cpf = "Não consta",
                        };
                        Usuarios.listUsuarios.Add(usuario);
                    }
                }

                // Instancia classe que vai guardar o fluxo dos dados em memória
                MemoryStream ms = new MemoryStream();

                //Using para criação da tabela
                using (FileStream fs = new FileStream(Path.Combine(webRootPath, nomeArquivo), FileMode.Create))
                {

                    // Cria arquivo excel .xlsx (com .xls ocorrem problemas)
                    IWorkbook workbook = new XSSFWorkbook();
                    // Cria um novo plano para o Excel
                    ISheet sheet = workbook.CreateSheet("Usuarios");

                    int rowNumber = 0;

                    // Cria uma nova linha (primeira linha)
                    IRow row = sheet.CreateRow(rowNumber);
                    ICell cell;

                    // Cria uma celula na linha, em seguida, coloca o valor da célula
                    cell = row.CreateCell(0);
                    cell.SetCellValue("Nome");

                    cell = row.CreateCell(1);
                    cell.SetCellValue("Email");

                    cell = row.CreateCell(2);
                    cell.SetCellValue("CPF");

                    rowNumber++;

                    // For para implementação de todos os usuarios da lista
                    for (int i = 0; i < Usuarios.listUsuarios.Count; i++)
                    {
                        // Cria uma nova linha para implementação de novo usuário
                        row = sheet.CreateRow(rowNumber);

                        row.CreateCell(0).SetCellValue(Usuarios.listUsuarios[i].Nome);
                        row.CreateCell(1).SetCellValue(Usuarios.listUsuarios[i].Email);
                        row.CreateCell(2).SetCellValue(Usuarios.listUsuarios[i].Cpf);

                        rowNumber++;
                    }

                    //Tamanho das colunas
                    sheet.SetColumnWidth(0, 40 * 256);
                    sheet.SetColumnWidth(1, 40 * 256);
                    sheet.SetColumnWidth(2, 40 * 256);

                    workbook.Write(fs);
                }

                // Using responsável para gerar o arquivo na pasta wwwroot
                using (FileStream fs = new FileStream(Path.Combine(webRootPath, nomeArquivo), FileMode.Open))
                {
                    await fs.CopyToAsync(ms);
                }

                ms.Position = 0;

                // Método herdado de BaseController que recebe a memoryStream, ContentType(diz que é um arquivo .xlsx) e o nome do arquivo.
                var retorno = File(ms, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nomeArquivo);

                return retorno;
            }
            catch (Exception ex)
            {
                return StatusCode(400, $"Ocorreu um erro na geração do Excel de Usuários / {ex.Message}");
            }
        }
    }
}
