// API/Middlewares/RequestLoggingMiddleware.cs
using System.Diagnostics;

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
            var isMultipartRequest = context.Request.HasFormContentType;

            // ✅ LOG REQUEST
            LogRequest(context);

            // For multipart/form-data requests (file upload), avoid wrapping response body
            // to reduce the chance of stream side effects while model binding reads form data.
            if (isMultipartRequest)
            {
                try
                {
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
                    LogResponseMetadata(context, stopwatch.ElapsedMilliseconds);
                }

                return;
            }

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
                context.Response.Body = originalBodyStream;
                responseBody.Seek(0, SeekOrigin.Begin);
                await responseBody.CopyToAsync(originalBodyStream);
            }
        }

        private void LogRequest(HttpContext context)
        {
            var request = context.Request;
            var hasAuthHeader = request.Headers.ContainsKey("Authorization");

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
                $"║ AuthHeader:  {(hasAuthHeader ? "Present" : "None")}\n" +
                $"║ IP:          {context.Connection.RemoteIpAddress}\n" +
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

        private void LogResponseMetadata(HttpContext context, long elapsedMs)
        {
            var response = context.Response;
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
                "║ Body:        <skipped for multipart/form-data request>\n" +
                "╚════════════════════════════════════════════════════════════════");
        }
    }
}
