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
using System.Net.Http;
using System.Text.Json;

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

        /// <summary>管理員重設任意帳號密碼（僅限開發測試用）</summary>
        [HttpPost("admin-reset-password")]
        public async Task<IActionResult> AdminResetPassword([FromBody] AdminResetPasswordModel model)
        {
            if (model.AdminKey != "YuDongZi2026")
                return Unauthorized(new { message = "無效的管理員金鑰" });

            var user = await _userManager.FindByEmailAsync(model.Email);
            if (user == null) return NotFound(new { message = "帳號不存在" });

            var token = await _userManager.GeneratePasswordResetTokenAsync(user);
            var result = await _userManager.ResetPasswordAsync(user, token, model.NewPassword);

            if (result.Succeeded)
                return Ok(new { message = $"{model.Email} 密碼已重設成功" });

            return BadRequest(new { message = "重設失敗", errors = result.Errors });
        }

        /// <summary>取得會員基本資料（含生辰）</summary>
        [HttpGet("profile")]
        [Authorize]
        public async Task<IActionResult> GetProfile()
        {
            var userId = User.FindFirstValue(ClaimTypes.NameIdentifier);
            var user = await _userManager.FindByIdAsync(userId!);
            if (user == null) return Unauthorized();

            var adminEmail = _configuration["Admin:Email"];
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
                chartName = user.ChartName ?? user.Name,
                isAdmin = string.Equals(user.Email, adminEmail, StringComparison.OrdinalIgnoreCase),
                hasLineLinked = user.LineUserId != null
            });
        }

        /// <summary>LINE Login OAuth 換取 JWT</summary>
        [HttpPost("line-login")]
        public async Task<IActionResult> LineLogin([FromBody] LineLoginModel model)
        {
            var channelId = _configuration["LineLogin:ChannelId"]!;
            var channelSecret = _configuration["LineLogin:ChannelSecret"]!;
            var redirectUri = !string.IsNullOrEmpty(model.RedirectUri)
                ? model.RedirectUri
                : _configuration["LineLogin:RedirectUri"]!;

            // 1. 用 code 換 access token
            using var http = new HttpClient();
            var tokenParams = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "authorization_code",
                ["code"] = model.Code,
                ["redirect_uri"] = redirectUri,
                ["client_id"] = channelId,
                ["client_secret"] = channelSecret
            });
            var tokenRes = await http.PostAsync("https://api.line.me/oauth2/v2.1/token", tokenParams);
            if (!tokenRes.IsSuccessStatusCode)
                return BadRequest(new { message = "LINE 授權失敗，請重試" });

            var tokenJson = await tokenRes.Content.ReadAsStringAsync();
            var tokenData = JsonDocument.Parse(tokenJson).RootElement;
            if (!tokenData.TryGetProperty("access_token", out var accessTokenElem))
                return BadRequest(new { message = "無法取得 LINE access token" });

            var accessToken = accessTokenElem.GetString()!;

            // 2. 取得 LINE 用戶資料
            http.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            var profileRes = await http.GetAsync("https://api.line.me/v2/profile");
            if (!profileRes.IsSuccessStatusCode)
                return BadRequest(new { message = "無法取得 LINE 用戶資料" });

            var profileJson = await profileRes.Content.ReadAsStringAsync();
            var profile = JsonDocument.Parse(profileJson).RootElement;
            var lineUserId = profile.GetProperty("userId").GetString()!;
            var displayName = profile.GetProperty("displayName").GetString() ?? "LINE 用戶";

            // 3. 找或建立用戶
            var user = await _userManager.FindByLoginAsync("LINE", lineUserId);
            if (user == null)
            {
                // 用 LineUserId 查找
                user = _userManager.Users.FirstOrDefault(u => u.LineUserId == lineUserId);
            }
            if (user == null)
            {
                // 建立新帳號
                var email = $"line_{lineUserId}@line.user";
                user = new ApplicationUser
                {
                    UserName = email,
                    Email = email,
                    Name = displayName,
                    LineUserId = lineUserId,
                    EmailConfirmed = true
                };
                var createResult = await _userManager.CreateAsync(user);
                if (!createResult.Succeeded)
                    return BadRequest(new { message = "建立帳號失敗" });
            }
            else if (user.LineUserId == null)
            {
                user.LineUserId = lineUserId;
                await _userManager.UpdateAsync(user);
            }

            // 4. 產生 JWT
            var authClaims = new List<Claim>
            {
                new Claim(ClaimTypes.NameIdentifier, user.Id!),
                new Claim(ClaimTypes.Email, user.Email!)
            };
            var authSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_configuration["JWT:Secret"]!));
            var token = new JwtSecurityToken(
                issuer: _configuration["JWT:ValidIssuer"],
                audience: _configuration["JWT:ValidAudience"],
                expires: DateTime.Now.AddDays(7),
                claims: authClaims,
                signingCredentials: new SigningCredentials(authSigningKey, SecurityAlgorithms.HmacSha256)
            );

            return Ok(new
            {
                token = new JwtSecurityTokenHandler().WriteToken(token),
                message = "LINE 登入成功！",
                name = user.Name
            });
        }

        /// <summary>Google Login OAuth 換取 JWT</summary>
        [HttpPost("google-login")]
        public async Task<IActionResult> GoogleLogin([FromBody] GoogleLoginModel model)
        {
            var clientId = _configuration["GoogleLogin:ClientId"]!;
            var clientSecret = _configuration["GoogleLogin:ClientSecret"]!;
            var redirectUri = !string.IsNullOrEmpty(model.RedirectUri)
                ? model.RedirectUri
                : "https://myweb.fly.dev/auth/google/callback";

            // 1. 用 code 換 access token
            using var http = new HttpClient();
            var tokenParams = new FormUrlEncodedContent(new Dictionary<string, string>
            {
                ["grant_type"] = "authorization_code",
                ["code"] = model.Code,
                ["redirect_uri"] = redirectUri,
                ["client_id"] = clientId,
                ["client_secret"] = clientSecret
            });
            var tokenRes = await http.PostAsync("https://oauth2.googleapis.com/token", tokenParams);
            if (!tokenRes.IsSuccessStatusCode)
                return BadRequest(new { message = "Google 授權失敗，請重試" });

            var tokenJson = await tokenRes.Content.ReadAsStringAsync();
            var tokenData = JsonDocument.Parse(tokenJson).RootElement;
            if (!tokenData.TryGetProperty("access_token", out var accessTokenElem))
                return BadRequest(new { message = "無法取得 Google access token" });

            var accessToken = accessTokenElem.GetString()!;

            // 2. 取得 Google 用戶資料
            http.DefaultRequestHeaders.Authorization =
                new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", accessToken);
            var profileRes = await http.GetAsync("https://www.googleapis.com/oauth2/v2/userinfo");
            if (!profileRes.IsSuccessStatusCode)
                return BadRequest(new { message = "無法取得 Google 用戶資料" });

            var profileJson = await profileRes.Content.ReadAsStringAsync();
            var profile = JsonDocument.Parse(profileJson).RootElement;
            var googleUserId = profile.GetProperty("id").GetString()!;
            var displayName = profile.TryGetProperty("name", out var nameElem) ? nameElem.GetString() ?? "Google 用戶" : "Google 用戶";
            var googleEmail = profile.TryGetProperty("email", out var emailElem) ? emailElem.GetString() : null;

            // 3. 找或建立用戶（優先用 Google email 查找已有帳號）
            ApplicationUser? user = null;
            if (!string.IsNullOrEmpty(googleEmail))
                user = await _userManager.FindByEmailAsync(googleEmail);

            if (user == null)
                user = _userManager.Users.FirstOrDefault(u => u.GoogleUserId == googleUserId);

            if (user == null)
            {
                var email = googleEmail ?? $"google_{googleUserId}@google.user";
                user = new ApplicationUser
                {
                    UserName = email,
                    Email = email,
                    Name = displayName,
                    GoogleUserId = googleUserId,
                    EmailConfirmed = true
                };
                var createResult = await _userManager.CreateAsync(user);
                if (!createResult.Succeeded)
                    return BadRequest(new { message = "建立帳號失敗" });
            }
            else if (user.GoogleUserId == null)
            {
                user.GoogleUserId = googleUserId;
                await _userManager.UpdateAsync(user);
            }

            // 4. 產生 JWT
            var authClaims = new List<Claim>
            {
                new Claim(ClaimTypes.NameIdentifier, user.Id!),
                new Claim(ClaimTypes.Email, user.Email!)
            };
            var authSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_configuration["JWT:Secret"]!));
            var token = new JwtSecurityToken(
                issuer: _configuration["JWT:ValidIssuer"],
                audience: _configuration["JWT:ValidAudience"],
                expires: DateTime.Now.AddDays(7),
                claims: authClaims,
                signingCredentials: new SigningCredentials(authSigningKey, SecurityAlgorithms.HmacSha256)
            );

            return Ok(new
            {
                token = new JwtSecurityTokenHandler().WriteToken(token),
                message = "Google 登入成功！",
                name = user.Name
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

    public class AdminResetPasswordModel
    {
        public required string AdminKey { get; set; }
        public required string Email { get; set; }
        public required string NewPassword { get; set; }
    }

    public class LineLoginModel
    {
        public required string Code { get; set; }
        public string? RedirectUri { get; set; }
    }

    public class GoogleLoginModel
    {
        public required string Code { get; set; }
        public string? RedirectUri { get; set; }
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