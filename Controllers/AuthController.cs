using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Identity;
using Microsoft.AspNetCore.Mvc;
using Ecanapi.Data;
using System.Threading.Tasks;
using System.IdentityModel.Tokens.Jwt;
using System.Text;
using Microsoft.IdentityModel.Tokens;
using System.Security.Claims;
using Microsoft.Extensions.Configuration;
using System.Collections.Generic;

namespace Ecanapi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class AuthController : ControllerBase
    {
        private readonly UserManager<ApplicationUser> _userManager;
        private readonly SignInManager<ApplicationUser> _signInManager;
        private readonly IConfiguration _configuration;

        public AuthController(
            UserManager<ApplicationUser> userManager,
            SignInManager<ApplicationUser> signInManager,
            IConfiguration configuration)
        {
            _userManager = userManager;
            _signInManager = signInManager;
            _configuration = configuration;
        }

        [HttpPost("register")]
        public async Task<IActionResult> Register([FromBody] RegisterModel model)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var user = new ApplicationUser { UserName = model.Email, Email = model.Email, Name = model.Name };
            var result = await _userManager.CreateAsync(user, model.Password);

            if (result.Succeeded)
            {
                return Ok(new { Message = "註冊成功！" });
            }

            return BadRequest(result.Errors);
        }

        /// <summary>
        /// 處理使用者登入請求並返回 JWT 憑證
        /// </summary>
        [HttpPost("login")]
        public async Task<IActionResult> Login([FromBody] LoginModel model)
        {
            if (!ModelState.IsValid)
            {
                return BadRequest(ModelState);
            }

            var user = await _userManager.FindByEmailAsync(model.Email);
            if (user == null)
            {
                return Unauthorized(new { Message = "使用者或密碼錯誤。" });
            }

            var result = await _signInManager.CheckPasswordSignInAsync(user, model.Password, false);

            if (result.Succeeded)
            {
                // 1. 建立宣告 (Claims)
                var authClaims = new List<Claim>
                {
                    new Claim(ClaimTypes.NameIdentifier, user.Id!),
                    new Claim(ClaimTypes.Email, user.Email!)
                };

                // 2. 取得 JWT 密鑰
                var authSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_configuration["JWT:Secret"]!));

                // 3. 建立 JWT Token
                var token = new JwtSecurityToken(
                    issuer: _configuration["JWT:ValidIssuer"],
                    audience: _configuration["JWT:ValidAudience"],
                    expires: DateTime.Now.AddDays(7),
                    claims: authClaims,
                    signingCredentials: new SigningCredentials(authSigningKey, SecurityAlgorithms.HmacSha256)
                );

                // 4. 回傳包含 Token 的 JSON 物件
                return Ok(new
                {
                    token = new JwtSecurityTokenHandler().WriteToken(token),
                    message = "登入成功！"
                });
            }

            return Unauthorized(new { Message = "使用者或密碼錯誤。" });
        }

        [HttpPut("change-password")]
        [Authorize]
        public async Task<IActionResult> ChangePassword([FromBody] ChangePasswordModel model)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var user = await _userManager.FindByIdAsync(userId!);
            if (user == null)
                return Unauthorized(new { message = "使用者不存在" });

            var result = await _userManager.ChangePasswordAsync(user, model.CurrentPassword, model.NewPassword);
            if (result.Succeeded)
                return Ok(new { message = "密碼修改成功！" });

            var error = result.Errors.FirstOrDefault()?.Description ?? "修改失敗";
            return BadRequest(new { message = error });
        }

        /// <summary>取得會員基本資料（含生辰）</summary>
        [HttpGet("profile")]
        [Authorize]
        public async Task<IActionResult> GetProfile()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var user = await _userManager.FindByIdAsync(userId!);
            if (user == null) return Unauthorized();

            return Ok(new
            {
                name = user.Name,
                email = user.Email,
                points = user.Points,
                hasBirthData = user.HasBirthData,
                birthYear = user.BirthYear,
                birthMonth = user.BirthMonth,
                birthDay = user.BirthDay,
                birthHour = user.BirthHour,
                birthMinute = user.BirthMinute,
                birthGender = user.BirthGender,
                dateType = user.DateType ?? "solar",
                chartName = user.ChartName ?? user.Name
            });
        }

        /// <summary>儲存／更新會員生辰資料</summary>
        [HttpPut("profile")]
        [Authorize]
        public async Task<IActionResult> UpdateProfile([FromBody] UpdateProfileModel model)
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var user = await _userManager.FindByIdAsync(userId!);
            if (user == null) return Unauthorized();

            if (!string.IsNullOrEmpty(model.ChartName)) user.ChartName = model.ChartName;
            if (model.BirthYear.HasValue) user.BirthYear = model.BirthYear;
            if (model.BirthMonth.HasValue) user.BirthMonth = model.BirthMonth;
            if (model.BirthDay.HasValue) user.BirthDay = model.BirthDay;
            if (model.BirthHour.HasValue) user.BirthHour = model.BirthHour;
            if (model.BirthMinute.HasValue) user.BirthMinute = model.BirthMinute;
            if (model.BirthGender.HasValue) user.BirthGender = model.BirthGender;
            if (!string.IsNullOrEmpty(model.DateType)) user.DateType = model.DateType;

            var result = await _userManager.UpdateAsync(user);
            if (!result.Succeeded)
                return BadRequest(new { message = "更新失敗", errors = result.Errors });

            return Ok(new { message = "生辰資料已儲存", hasBirthData = user.HasBirthData });
        }
    }

    public class RegisterModel
    {
        public required string Name { get; set; }
        public required string Email { get; set; }
        public required string Password { get; set; }
    }

    public class LoginModel
    {
        public required string Email { get; set; }
        public required string Password { get; set; }
    }

    public class ChangePasswordModel
    {
        public required string CurrentPassword { get; set; }
        public required string NewPassword { get; set; }
    }

    public class UpdateProfileModel
    {
        public string? ChartName { get; set; }
        public int? BirthYear { get; set; }
        public int? BirthMonth { get; set; }
        public int? BirthDay { get; set; }
        public int? BirthHour { get; set; }
        public int? BirthMinute { get; set; }
        public int? BirthGender { get; set; }
        public string? DateType { get; set; }
    }
}