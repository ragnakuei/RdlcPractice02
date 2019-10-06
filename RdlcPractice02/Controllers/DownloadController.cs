using BusinessLogic;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using System.Web.Http;

namespace RdlcPractice01.Controllers
{
    public class DownloadController : ApiController
    {
        [HttpGet, Route("api/Download/Excel")]
        public IHttpActionResult Excel()
        {
            var reportStream = new ReportPersonBL().GetReport();

            var response = new HttpResponseMessage(HttpStatusCode.OK);
            var fileName = "Test.xlsx";
            response.Content = new StreamContent(reportStream);
            response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
            response.Content.Headers.ContentDisposition.FileName = fileName;
            response.Content.Headers.ContentType = new MediaTypeHeaderValue(MimeMapping.GetMimeMapping(fileName));
            response.Content.Headers.ContentLength = reportStream.Length; //告知瀏覽器下載長度
            return ResponseMessage(response);
        }
    }
}