using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;

namespace Ecanapi.Services
{
    public interface IEmailService
    {
        Task SendAdminBookingNotifyAsync(string serviceType, string name, string contactType, string contactInfo, string? notes, string? preferredDate, string? userEmail, int bookingId);
        Task SendClientConfirmationAsync(string toEmail, string name, string serviceType);
        Task SendReportReadyAsync(string toEmail, string toName, string reportTitle, string downloadUrl, string expiryDesc);
        Task SendAdminReportNotifyAsync(string userName, string userEmail, string reportTitle, string reportType, string adminReviewUrl);
    }

    public class EmailService : IEmailService
    {
        private readonly IConfiguration _config;
        private readonly ILogger<EmailService> _logger;

        public EmailService(IConfiguration config, ILogger<EmailService> logger)
        {
            _config = config;
            _logger = logger;
        }

        private bool IsConfigured()
        {
            var host = _config["Smtp:Host"];
            var user = _config["Smtp:User"];
            var pass = _config["Smtp:Password"];
            return !string.IsNullOrEmpty(host) && !string.IsNullOrEmpty(user) && !string.IsNullOrEmpty(pass);
        }

        private async Task SendAsync(string toEmail, string toName, string subject, string htmlBody)
        {
            if (!IsConfigured())
            {
                _logger.LogWarning("SMTP not configured, skipping email to {Email}", toEmail);
                return;
            }

            var message = new MimeMessage();
            var fromEmail = _config["Smtp:User"]!;
            var fromName = _config["Smtp:FromName"] ?? "玉洞子星相古學堂";
            message.From.Add(new MailboxAddress(fromName, fromEmail));
            message.To.Add(new MailboxAddress(toName, toEmail));
            message.Subject = subject;

            var builder = new BodyBuilder { HtmlBody = htmlBody };
            message.Body = builder.ToMessageBody();

            using var client = new SmtpClient();
            var host = _config["Smtp:Host"]!;
            var port = int.TryParse(_config["Smtp:Port"], out var p) ? p : 587;

            await client.ConnectAsync(host, port, SecureSocketOptions.StartTls);
            await client.AuthenticateAsync(_config["Smtp:User"]!, _config["Smtp:Password"]!);
            await client.SendAsync(message);
            await client.DisconnectAsync(true);
        }

        private static string ServiceTypeName(string serviceType) => serviceType switch
        {
            "blessing" => "祈福服務",
            "consultation" => "問事預約",
            _ => serviceType
        };

        public async Task SendAdminBookingNotifyAsync(string serviceType, string name, string contactType, string contactInfo, string? notes, string? preferredDate, string? userEmail, int bookingId)
        {
            var adminEmail = _config["Admin:Email"];
            if (string.IsNullOrEmpty(adminEmail)) return;

            var typeName = ServiceTypeName(serviceType);
            var subject = $"[玉洞子] 新{typeName}登記 - {name} (#{bookingId})";

            var html = $@"
<div style='font-family:sans-serif;max-width:600px;margin:0 auto;'>
  <h2 style='color:#b45309;'>新{typeName}登記</h2>
  <table style='width:100%;border-collapse:collapse;'>
    <tr><td style='padding:8px;background:#fef3c7;font-weight:bold;width:120px;'>單號</td><td style='padding:8px;border-bottom:1px solid #eee;'>#{bookingId}</td></tr>
    <tr><td style='padding:8px;background:#fef3c7;font-weight:bold;'>姓名</td><td style='padding:8px;border-bottom:1px solid #eee;'>{name}</td></tr>
    <tr><td style='padding:8px;background:#fef3c7;font-weight:bold;'>聯絡方式</td><td style='padding:8px;border-bottom:1px solid #eee;'>{contactType}: {contactInfo}</td></tr>
    {(string.IsNullOrEmpty(preferredDate) ? "" : $"<tr><td style='padding:8px;background:#fef3c7;font-weight:bold;'>希望時間</td><td style='padding:8px;border-bottom:1px solid #eee;'>{preferredDate}</td></tr>")}
    {(string.IsNullOrEmpty(notes) ? "" : $"<tr><td style='padding:8px;background:#fef3c7;font-weight:bold;'>備註</td><td style='padding:8px;border-bottom:1px solid #eee;'>{notes}</td></tr>")}
    {(string.IsNullOrEmpty(userEmail) ? "" : $"<tr><td style='padding:8px;background:#fef3c7;font-weight:bold;'>帳號信箱</td><td style='padding:8px;border-bottom:1px solid #eee;'>{userEmail}</td></tr>")}
  </table>
  <p style='margin-top:20px;color:#666;'>請前往 <a href='https://yudongzi.tw/admin/bookings'>後台預約管理</a> 處理。</p>
</div>";

            try
            {
                await SendAsync(adminEmail, "玉洞子", subject, html);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to send admin booking notification");
            }
        }

        public async Task SendReportReadyAsync(string toEmail, string toName, string reportTitle, string downloadUrl, string expiryDesc)
        {
            if (string.IsNullOrEmpty(toEmail)) return;

            var subject = $"【玉洞子星相古學堂】您的{reportTitle}已審閱完成，請點擊下載";

            var html = $@"
<div style='font-family:sans-serif;max-width:600px;margin:0 auto;background:#fffbf0;padding:24px;border-radius:8px;'>
  <h2 style='color:#b45309;margin-bottom:8px;'>命書已審閱完成</h2>
  <p style='color:#333;'>親愛的 {toName}，</p>
  <p style='color:#333;'>您申請的「{reportTitle}」已由玉洞子親自審閱完成，請在有效期限內點擊下方按鈕下載。</p>
  <div style='text-align:center;margin:28px 0;'>
    <a href='{downloadUrl}' style='background:#b45309;color:#fff;text-decoration:none;padding:14px 32px;border-radius:8px;font-size:16px;font-weight:bold;display:inline-block;'>
      點擊下載命書
    </a>
  </div>
  <p style='color:#e67e22;font-size:13px;font-weight:bold;'>下載連結有效期限：{expiryDesc}（72小時）</p>
  <p style='color:#999;font-size:12px;margin-top:16px;'>此為系統自動發送，請勿直接回覆此信。如有問題請透過 LINE 聯繫。</p>
  <p style='color:#b45309;font-weight:bold;margin-top:16px;'>玉洞子星相古學堂</p>
</div>";

            try
            {
                await SendAsync(toEmail, toName, subject, html);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to send report ready email to {Email}", toEmail);
            }
        }

        public async Task SendAdminReportNotifyAsync(string userName, string userEmail, string reportTitle, string reportType, string adminReviewUrl)
        {
            var adminEmail = _config["Admin:Email"];
            if (string.IsNullOrEmpty(adminEmail)) return;

            var subject = $"[玉洞子] 新命書申請 - {userName}：{reportTitle}";

            var html = $@"
<div style='font-family:sans-serif;max-width:600px;margin:0 auto;background:#fffbf0;padding:24px;border-radius:8px;'>
  <h2 style='color:#b45309;'>新命書申請通知</h2>
  <table style='width:100%;border-collapse:collapse;margin-bottom:16px;'>
    <tr><td style='padding:8px;background:#fef3c7;font-weight:bold;width:100px;'>申請人</td><td style='padding:8px;border-bottom:1px solid #eee;'>{userName} ({userEmail})</td></tr>
    <tr><td style='padding:8px;background:#fef3c7;font-weight:bold;'>命書種類</td><td style='padding:8px;border-bottom:1px solid #eee;'>{reportTitle}（{reportType}）</td></tr>
  </table>
  <div style='text-align:center;margin:24px 0;'>
    <a href='{adminReviewUrl}' style='background:#b45309;color:#fff;text-decoration:none;padding:12px 28px;border-radius:8px;font-size:15px;font-weight:bold;display:inline-block;'>
      前往後台審核命書
    </a>
  </div>
  <p style='color:#999;font-size:12px;'>此為系統自動發送。</p>
</div>";

            try { await SendAsync(adminEmail, "玉洞子", subject, html); }
            catch (Exception ex) { _logger.LogError(ex, "Failed to send admin report notification"); }
        }

        public async Task SendClientConfirmationAsync(string toEmail, string name, string serviceType)
        {
            if (string.IsNullOrEmpty(toEmail)) return;

            var typeName = ServiceTypeName(serviceType);
            var subject = $"【玉洞子星相古學堂】您的{typeName}申請已收到";

            var html = $@"
<div style='font-family:sans-serif;max-width:600px;margin:0 auto;'>
  <h2 style='color:#b45309;'>感謝您的申請</h2>
  <p>親愛的 {name}，</p>
  <p>我們已收到您的{typeName}申請，玉洞子將於 <strong>1-2 個工作日內</strong> 透過您登記的聯絡方式與您確認詳情。</p>
  <p>如有任何疑問，歡迎透過 LINE 或網站聯繫我們。</p>
  <br/>
  <p style='color:#b45309;font-weight:bold;'>玉洞子星相古學堂</p>
  <p style='color:#999;font-size:12px;'>此為系統自動發送，請勿直接回覆此信。</p>
</div>";

            try
            {
                await SendAsync(toEmail, name, subject, html);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Failed to send client confirmation email to {Email}", toEmail);
            }
        }
    }
}
