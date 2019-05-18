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
  
            }
            catch (Exception ex)
            {
                result.Error = ex.InmostMessage();
            }

            return result;
        }
    }
}
