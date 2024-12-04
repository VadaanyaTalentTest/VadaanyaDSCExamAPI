using Microsoft.AspNetCore.Mvc;

namespace VadaanyaTalentTest1.Controllers
{   

    public class BaseController : ControllerBase
    {
        // Common logic or services can be added here
        protected IActionResult HandleError(Exception ex)
        {
            // Handle the error and return a proper response
            return StatusCode(500, $"Internal server error: {ex.Message}");
        }
    }
}
