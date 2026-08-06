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
        /// 搜尋客戶：依姓名或客戶編號模糊比對（管理員使用）
        /// </summary>
        [HttpGet("search")]
        public async Task<IActionResult> SearchCustomers([FromQuery] string q = "")
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (userId == null) return Unauthorized();

            var query = _context.Customers.Where(c => c.ApplicationUserId == userId);
            if (!string.IsNullOrWhiteSpace(q))
                query = query.Where(c => c.Name.Contains(q) || c.CustomerCode.Contains(q));

            var customers = await query
                .OrderByDescending(c => c.CreatedAt)
                .Take(20)
                .Select(c => new CustomerDto
                {
                    Id = c.Id,
                    CustomerCode = c.CustomerCode,
                    Name = c.Name,
                    Email = c.Email,
                    Gender = c.Gender,
                    BirthDateTime = c.BirthDateTime,
                    Notes = c.Notes,
                    CreatedAt = c.CreatedAt,
                })
                .ToListAsync();

            return Ok(customers);
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
                                        .OrderByDescending(c => c.CreatedAt)
                                        .Select(c => new CustomerDto
                                        {
                                            Id = c.Id,
                                            CustomerCode = c.CustomerCode,
                                            Name = c.Name,
                                            Email = c.Email,
                                            Gender = c.Gender,
                                            BirthDateTime = c.BirthDateTime,
                                            Notes = c.Notes,
                                            CreatedAt = c.CreatedAt,
                                        })
                                        .ToListAsync();

            return Ok(customers);
        }

        /// <summary>
        /// 新增客戶（管理員代建）：自動產生客戶編號與虛擬 email
        /// </summary>
        [HttpPost]
        public async Task<IActionResult> PostCustomer([FromBody] CreateCustomerRequest req)
        {
            if (string.IsNullOrWhiteSpace(req.Name))
                return BadRequest(new { Message = "姓名為必填。" });

            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            if (userId == null) return Unauthorized();

            // 客戶編號：建檔時間 yyyyMMddHHmmss，若衝突補流水號
            string code;
            int suffix = 0;
            do
            {
                code = DateTime.UtcNow.AddHours(8).ToString("yyyyMMddHHmmss")
                       + (suffix > 0 ? suffix.ToString() : "");
                suffix++;
            } while (await _context.Customers.AnyAsync(c => c.CustomerCode == code));

            // 虛擬 email（若未提供）
            string email = string.IsNullOrWhiteSpace(req.Email)
                ? $"guest_{code}@yudongzi.tw"
                : req.Email;

            var now = DateTime.UtcNow;
            var customer = new Customer
            {
                CustomerCode = code,
                Name = req.Name.Trim(),
                Email = email,
                Gender = req.Gender,
                BirthDateTime = req.BirthDateTime,
                Notes = req.Notes,
                CreatedAt = now,
                ApplicationUserId = userId
            };

            _context.Customers.Add(customer);
            await _context.SaveChangesAsync();

            return Ok(new CustomerDto
            {
                Id = customer.Id,
                CustomerCode = customer.CustomerCode,
                Name = customer.Name,
                Email = customer.Email,
                Gender = customer.Gender,
                BirthDateTime = customer.BirthDateTime,
                Notes = customer.Notes,
                CreatedAt = customer.CreatedAt,
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
        public string CustomerCode { get; set; } = "";
        public required string Name { get; set; }
        public required string Email { get; set; }
        public int Gender { get; set; }
        public DateTime BirthDateTime { get; set; }
        public string? Notes { get; set; }
        public DateTime CreatedAt { get; set; }
    }

    // 管理員代建客戶的請求格式（email 可選，生日用整數組合）
    public class CreateCustomerRequest
    {
        [Required]
        public required string Name { get; set; }
        public int Gender { get; set; } = 1;
        public DateTime BirthDateTime { get; set; }
        public string? Email { get; set; }
        public string? Notes { get; set; }
    }
}
