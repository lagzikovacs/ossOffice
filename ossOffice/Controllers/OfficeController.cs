using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace ossOffice.Controllers
{
    public class OfficeController : ApiController
    {
        [HttpPost]
        public ByteArrayResult ExcelToPdf([FromBody] byte[] excel)
        {
            var result = new ByteArrayResult();

            try
            {
                var b = new byte[2];
                b[0] = 65;
                b[1] = 66;
                result.Result = b;
            }
            catch (Exception ex)
            {
                result.Error = ex.InmostMessage();
            }

            return result;
        }
    }
}
