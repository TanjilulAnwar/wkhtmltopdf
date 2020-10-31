using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Wkhtmltopdf.NetCore;

namespace dixi.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForecastController : ControllerBase
    {
        private IGeneratePdf _generatePdf;
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger, IGeneratePdf generatePdf)
        {
            _generatePdf = generatePdf;
            _logger = logger;

        }

        private List<User> users = new List<User>
    {
        new User { Id = 1, Username = "Dolores Abernathy" },
        new User { Id = 2, Username = "Maeve Millay" },
        new User { Id = 3, Username = "Bernard Lowe" },
        new User { Id = 4, Username = "ManIn Black" }
    };
        [HttpGet]
        [Route("csv")]
        public IActionResult Csv()
        {
            var builder = new StringBuilder();
            builder.AppendLine("Id,Username");
            foreach (var user in users)
            {
                builder.AppendLine($"{user.Id},{user.Username}");
            }

            return File(Encoding.UTF8.GetBytes(builder.ToString()), "text/csv", "users.csv");
        }
        [HttpGet]
        [Route("excel")]
        public IActionResult Excel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Users");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Username";
                foreach (var user in users)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = user.Id;
                    worksheet.Cell(currentRow, 2).Value = user.Username;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        "users.xlsx");
                }
            }
        }

        [HttpGet]
        public IActionResult Get()
        {
            //var rng = new Random();
            //return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            //{
            //    Date = DateTime.Now.AddDays(index),
            //    TemperatureC = rng.Next(-20, 55),
            //    Summary = Summaries[rng.Next(Summaries.Length)]
            //})
            //.ToArray();
            var imagePath = Path.Combine(Directory.GetCurrentDirectory(), "Assets", "jeydonx400.jpg");
            var html = new StringBuilder();

            html.Append(@"<!doctype html>
                        <html lang='en'>
                        <head>

                   

                    <meta charset='utf-8'>
                       <meta name = 'viewport' content = 'width=device-width, initial-scale=1.0'>
                            <meta http-equiv='Content - type' content='text / html; charset='utf-8' /><meta charset='UTF-8' />
                                <title> document</title>


                             <style>
 @font-face  { font-family: font-family: 'Hind Siliguri', sans-serif; } 
table {
  font-family: arial, sans-serif;         
  border-collapse: collapse;
  width: 100%;
}

td, th {
  border: 1px solid #dddddd;
  text-align: left;
  padding: 8px;
}

/*tr:nth-child(even) {
  background-color: #dddddd;
}*/
</style>
                        <link href='https://fonts.maateen.me/bangla/font.css' rel='stylesheet'>

                      </head>
                         <body onload='getPdfInfo()'>");

            html.Append(@"<h2>test pdf generation</h2>");
            html.AppendFormat($"<img src='{imagePath}' width=40% />");
           // html.Append(@"<h2>আমরা কাটিয়ে উঠব</h2>");
            html.Append(@"<h2>HTML Table</h2>

<table>
  <tr>
    <th>Company</th>
    <th>Contact</th>
    <th>Country</th>
  </tr>
  <tr>
    <td>Alfreds Futterkiste</td>
    <td>Maria Anders</td>
    <td>Germany</td>
  </tr>
  <tr>
    <td>Centro comercial Moctezuma</td>
    <td>Francisco Chang</td>
    <td>Mexico</td>
  </tr>
  <tr>
    <td>Ernst Handel</td>
    <td>Roland Mendel</td>
    <td>Austria</td>
  </tr>
  <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>

 <tr>
    <td>Island Trading</td>
    <td>Helen Bennett</td>
    <td>UK</td>
  </tr>
  </tr>
  <tr>
    <td>Laughing Bacchus Winecellars</td>
    <td>Yoshi Tannamuri</td>
    <td>Canada</td>
  </tr>
  <tr>
    <td>Magazzini Alimentari Riuniti</td>
    <td>Giovanni Rovelli</td>
    <td>Italy</td>
  </tr>
</table>");





            html.Append(@"<script>var pdfInfo = {};
  var x = document.location.search.substring(1).split('&');
  for (var i in x) { var z = x[i].split('=',2); pdfInfo[z[0]] = unescape(z[1]); }
  function getPdfInfo() {
    var page = pdfInfo.page || 1;
    var pageCount = pdfInfo.topage || 1;
    document.getElementById('pdfkit_page_current').textContent = page;
    document.getElementById('pdfkit_page_count').textContent = pageCount;
  }</script></body></html>");




            var pdf = _generatePdf.GetPDF(html.ToString());
           
            var pdfStreamResult = new MemoryStream();
            pdfStreamResult.Write(pdf, 0, pdf.Length);
            pdfStreamResult.Position = 0;
            return new FileStreamResult(pdfStreamResult, "application /pdf");
        }
    }
}
