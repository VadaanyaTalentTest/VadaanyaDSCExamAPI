using Microsoft.AspNetCore.Mvc;
using System.Net;

namespace VadaanyaTalentTest1.Controllers
{
    public class StatusCodeException : Exception
    {
        public HttpStatusCode StatusCode { get; }

        public StatusCodeException(HttpStatusCode statusCode, string message) : base(message)
        {
            StatusCode = statusCode;
        }
    }

    public class BaseController : ControllerBase
    {
        // Common logic or services can be added here
        protected IActionResult HandleError(Exception ex)
        {
            // Handle the error and return a proper response
            if (ex is StatusCodeException statusCodeEx && statusCodeEx.StatusCode != HttpStatusCode.InternalServerError)
            {
                return StatusCode((int)statusCodeEx.StatusCode, $"{ex.Message}");
            }
            Console.WriteLine(ex.Message);
            return StatusCode(500, $"Some error occurred.");
        }
    }
}
