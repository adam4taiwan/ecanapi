using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Ecanapi.Data;
using System.Security.Claims;
using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    [Authorize] // 確保所有客戶操作都需要認證
    public class CustomersController : ControllerBase
    {
        private readonly ApplicationDbContext _context;
        private readonly UserManager<ApplicationUser> _userManager;

        public CustomersController(ApplicationDbContext context, UserManager<ApplicationUser> userManager)
        {
            _context = context;
            _userManager = userManager;
        }

        /// <summary>
        /// 取得目前登入者所有的客戶清單
        /// </summary>
        [HttpGet]
        public async Task<IActionResult> GetCustomers()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (userId == null)
            {
                return Unauthorized();
            }

            var customers = await _context.Customers
                                        .Where(c => c.ApplicationUserId == userId)
                                        .Select(c => new CustomerDto
                                        {
                                            Id = c.Id,
                                            Name = c.Name,
                                            Email = c.Email,
                                            Gender = c.Gender,
                                            BirthDateTime = c.BirthDateTime
                                        })
                                        .ToListAsync();

            return Ok(customers);
        }

        /// <summary>
        /// 新增客戶
        /// </summary>
        [HttpPost]
        public async Task<IActionResult> PostCustomer([FromBody] CustomerDto customerDto)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (userId == null)
            {
                return Unauthorized();
            }

            var customer = new Customer
            {
                Name = customerDto.Name,
                Email = customerDto.Email,
                Gender = customerDto.Gender,
                BirthDateTime = customerDto.BirthDateTime,
                ApplicationUserId = userId
            };

            _context.Customers.Add(customer);
            await _context.SaveChangesAsync();

            return Ok(new CustomerDto
            {
                Id = customer.Id,
                Name = customer.Name,
                Email = customer.Email,
                Gender = customer.Gender,
                BirthDateTime = customer.BirthDateTime
            });
        }

        /// <summary>
        /// 更新客戶
        /// </summary>
        [HttpPut("{id}")]
        public async Task<IActionResult> PutCustomer(int id, [FromBody] CustomerDto customerDto)
        {
            if (id != customerDto.Id)
            {
                return BadRequest(new { Message = "客戶編號不匹配。" });
            }

            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (userId == null)
            {
                return Unauthorized();
            }

            var existingCustomer = await _context.Customers
                                                 .FirstOrDefaultAsync(c => c.Id == id && c.ApplicationUserId == userId);

            if (existingCustomer == null)
            {
                return NotFound(new { Message = "找不到客戶資料。" });
            }

            existingCustomer.Name = customerDto.Name;
            existingCustomer.Email = customerDto.Email;
            existingCustomer.Gender = customerDto.Gender;
            existingCustomer.BirthDateTime = customerDto.BirthDateTime;

            try
            {
                await _context.SaveChangesAsync();
                return Ok(new { Message = "客戶資料更新成功！" });
            }
            catch (DbUpdateConcurrencyException)
            {
                if (!_context.Customers.Any(e => e.Id == id))
                {
                    return NotFound(new { Message = "找不到客戶資料。" });
                }
                throw;
            }
        }

        /// <summary>
        /// 刪除客戶
        /// </summary>
        [HttpDelete("{id}")]
        public async Task<IActionResult> DeleteCustomer(int id)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (userId == null)
            {
                return Unauthorized();
            }

            var customer = await _context.Customers
                                         .FirstOrDefaultAsync(c => c.Id == id && c.ApplicationUserId == userId);
            if (customer == null)
            {
                return NotFound(new { Message = "找不到客戶資料或您沒有權限刪除。" });
            }

            _context.Customers.Remove(customer);
            await _context.SaveChangesAsync();
            return Ok(new { Message = "客戶資料刪除成功！" });
        }
    }

    public class CustomerDto
    {
        public int Id { get; set; }
        [Required]
        public required string Name { get; set; }
        [Required]
        [EmailAddress]
        public required string Email { get; set; }
        [Required]
        public int Gender { get; set; }
        [Required]
        public DateTime BirthDateTime { get; set; }
    }
}
