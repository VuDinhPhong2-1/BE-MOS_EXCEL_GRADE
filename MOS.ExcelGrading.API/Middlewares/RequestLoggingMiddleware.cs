// API/Middlewares/RequestLoggingMiddleware.cs
using System.Diagnostics;
using System.Text;

namespace MOS.ExcelGrading.API.Middlewares
{
    public class RequestLoggingMiddleware
    {
        private readonly RequestDelegate _next;
        private readonly ILogger<RequestLoggingMiddleware> _logger;

        public RequestLoggingMiddleware(RequestDelegate next, ILogger<RequestLoggingMiddleware> logger)
        {
            _next = next;
            _logger = logger;
        }

        public async Task InvokeAsync(HttpContext context)
        {
            var stopwatch = Stopwatch.StartNew();

            // ✅ LOG REQUEST
            await LogRequest(context);

            // ✅ LƯU LẠI RESPONSE BODY ĐỂ LOG
            var originalBodyStream = context.Response.Body;
            using var responseBody = new MemoryStream();
            context.Response.Body = responseBody;

            try
            {
                // ✅ TIẾP TỤC XỬ LÝ REQUEST
                await _next(context);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "[REQUEST ERROR] Unhandled exception");
                throw;
            }
            finally
            {
                stopwatch.Stop();

                // ✅ LOG RESPONSE
                await LogResponse(context, stopwatch.ElapsedMilliseconds);

                // ✅ COPY RESPONSE BODY TRỞ LẠI
                await responseBody.CopyToAsync(originalBodyStream);
            }
        }

        private async Task LogRequest(HttpContext context)
        {
            var request = context.Request;

            // ✅ LẤY TOKEN TỪ HEADER
            var authHeader = request.Headers["Authorization"].ToString();
            var token = authHeader.StartsWith("Bearer ")
                ? authHeader.Substring("Bearer ".Length).Trim()
                : null;

            var tokenInfo = !string.IsNullOrEmpty(token)
                ? $"Token: {token.Substring(0, Math.Min(50, token.Length))}... (Length: {token.Length})"
                : "No Token";

            // ✅ LẤY BODY (NẾU CÓ)
            string requestBody = string.Empty;
            if (request.ContentLength > 0 && request.ContentType?.Contains("application/json") == true)
            {
                request.EnableBuffering();
                using var reader = new StreamReader(request.Body, Encoding.UTF8, leaveOpen: true);
                requestBody = await reader.ReadToEndAsync();
                request.Body.Position = 0;
            }

            _logger.LogInformation(
                "\n" +
                "╔════════════════════════════════════════════════════════════════\n" +
                "║ 📥 INCOMING REQUEST\n" +
                "╠════════════════════════════════════════════════════════════════\n" +
                $"║ Method:      {request.Method}\n" +
                $"║ Path:        {request.Path}{request.QueryString}\n" +
                $"║ Scheme:      {request.Scheme}\n" +
                $"║ Host:        {request.Host}\n" +
                $"║ ContentType: {request.ContentType ?? "N/A"}\n" +
                $"║ {tokenInfo}\n" +
                $"║ IP:          {context.Connection.RemoteIpAddress}\n" +
                (string.IsNullOrEmpty(requestBody) ? "" : $"║ Body:        {requestBody}\n") +
                "╚════════════════════════════════════════════════════════════════");
        }

        private async Task LogResponse(HttpContext context, long elapsedMs)
        {
            var response = context.Response;

            // ✅ ĐỌC RESPONSE BODY
            response.Body.Seek(0, SeekOrigin.Begin);
            var responseBody = await new StreamReader(response.Body).ReadToEndAsync();
            response.Body.Seek(0, SeekOrigin.Begin);

            // ✅ TRUNCATE BODY NẾU QUÁ DÀI
            var bodyPreview = responseBody.Length > 500
                ? responseBody.Substring(0, 500) + "..."
                : responseBody;

            var statusEmoji = response.StatusCode switch
            {
                >= 200 and < 300 => "✅",
                >= 400 and < 500 => "⚠️",
                >= 500 => "❌",
                _ => "ℹ️"
            };

            _logger.LogInformation(
                "\n" +
                "╔════════════════════════════════════════════════════════════════\n" +
                $"║ {statusEmoji} RESPONSE\n" +
                "╠════════════════════════════════════════════════════════════════\n" +
                $"║ Status:      {response.StatusCode}\n" +
                $"║ ContentType: {response.ContentType ?? "N/A"}\n" +
                $"║ Duration:    {elapsedMs}ms\n" +
                (string.IsNullOrEmpty(bodyPreview) ? "" : $"║ Body:        {bodyPreview}\n") +
                "╚════════════════════════════════════════════════════════════════");
        }
    }
}
