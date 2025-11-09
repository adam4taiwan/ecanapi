using Microsoft.AspNetCore.Mvc;
using Print2Engine;
using Ecanapi.Models;
namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class CalculationController : ControllerBase
    {
        private readonly Print2Engine.Print2Engine _engine;

        public CalculationController(Print2Engine.Print2Engine engine)
        {
            _engine = engine;
        }

        [HttpPost]
        public IActionResult Calculate([FromBody] UserInput input)
        {
            if (input == null)
            {
                return BadRequest("輸入資料不能為空。");
            }

            try
            {
                var result = _engine.Process(input);
                return Ok(result);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"內部伺服器錯誤: {ex.Message}");
            }
        }
    }


}