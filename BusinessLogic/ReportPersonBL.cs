using BusinessLogic.Models;
using Microsoft.Reporting.WebForms;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BusinessLogic
{
    public class ReportPersonBL
    {
        public MemoryStream GetReport()
        {
            // 欄位是否顯示，一定要用 IEnumerable<T> 的資料型態
            var columnVisibleDto = new List<ReportPersonColumnVisibleDto>
                                   {
                                       new ReportPersonColumnVisibleDto
                                       {
                                           Id = true,
                                           Name = false,
                                           Age = true
                                       }
                                   };

            // 用來顯示欄位名稱，一定要用 IEnumerable<T> 的資料型態
            var columnNameDto = new List<ReportPersonColumnNameDto>
                                {
                                    new ReportPersonColumnNameDto
                                    {
                                        Id = "編號",
                                        Name = "名字",
                                        Age = "年齡"
                                    }
                                };

            // 欄位的值
            var cellDtos = new List<ReportPersonDto>
                           {
                               new ReportPersonDto
                               {
                                   Id = 1,
                                   Name = "A",
                                   Age = 18,
                               },
                               new ReportPersonDto
                               {
                                   Id = 2,
                                   Name = "B",
                                   Age = 19,
                               },
                               new ReportPersonDto
                               {
                                   Id = 3,
                                   Name = "C",
                                   Age = 20,
                               },
                               new ReportPersonDto
                               {
                                   Id = 4,
                                   Name = "D",
                                   Age = 21,
                               },
                               new ReportPersonDto
                               {
                                   Id = 5,
                                   Name = "E",
                                   Age = 22,
                               },
                           };

            var reportViewer = new ReportViewer();
            reportViewer.ProcessingMode = ProcessingMode.Local;
            // reportViewer.LocalReport.ReportPath = $"{Request.MapPath(Request.ApplicationPath)}Report\\報表名稱.rdlc";
            reportViewer.LocalReport.ReportEmbeddedResource = "BusinessLogic.RDLC.ReportPerson.rdlc";
            reportViewer.LocalReport.DataSources.Add(new ReportDataSource("DsReportPersonVisible", columnVisibleDto));
            reportViewer.LocalReport.DataSources.Add(new ReportDataSource("DsReportPersonColumn", columnNameDto));
            reportViewer.LocalReport.DataSources.Add(new ReportDataSource("DsReportPerson", cellDtos));

            var reportStreamBytes = reportViewer.LocalReport.Render("EXCELOPENXML", null, out var mimeType, out var encoding, out var extension, out var streamids, out var warnings);
            var reportStream = new MemoryStream(reportStreamBytes);

            return reportStream;
        }
    }
}