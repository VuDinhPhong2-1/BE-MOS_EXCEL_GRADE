using System.Security.Claims;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Filters;
using MOS.ExcelGrading.API.Extensions;

namespace MOS.ExcelGrading.API.Authorization
{
    /// <summary>
    /// Gắn attribute này lên action (hoặc cả controller) để yêu cầu user phải có
    /// một permission claim cụ thể, thay cho việc viết thủ công:
    ///
    ///     var hasPermission = User.Claims.Any(c => c.Type == "permission" && c.Value == Permissions.X);
    ///     if (!hasPermission) return Forbid();
    ///
    /// Ví dụ dùng:
    ///     [HttpGet]
    ///     [RequirePermission(Permissions.ViewXmlRules)]
    ///     public async Task&lt;IActionResult&gt; GetRuleSets() { ... }
    ///
    /// Mặc định chỉ cần có ÍT NHẤT MỘT trong các permission truyền vào (OR).
    /// Muốn bắt buộc có ĐỦ TẤT CẢ (AND), dùng requireAll: true.
    ///
    ///     [RequirePermission(requireAll: true, Permissions.EditGrades, Permissions.ExportGrades)]
    ///
    /// Lưu ý: attribute này chỉ kiểm tra permission claim, KHÔNG thay thế
    /// [Authorize] — vẫn cần [Authorize] (hoặc [Authorize(Roles = ...)]) ở
    /// controller/action để đảm bảo request đã authenticate trước.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method | AttributeTargets.Class, AllowMultiple = true)]
    public class RequirePermissionAttribute : Attribute, IAsyncAuthorizationFilter
    {
        private readonly string[] _permissions;
        private readonly bool _requireAll;

        public RequirePermissionAttribute(params string[] permissions)
            : this(requireAll: false, permissions)
        {
        }

        public RequirePermissionAttribute(bool requireAll, params string[] permissions)
        {
            _permissions = permissions ?? Array.Empty<string>();
            _requireAll = requireAll;
        }

        public Task OnAuthorizationAsync(AuthorizationFilterContext context)
        {
            if (_permissions.Length == 0)
            {
                return Task.CompletedTask;
            }

            var user = context.HttpContext.User;
            var granted = _requireAll
                ? user.HasAllPermissions(_permissions)
                : user.HasAnyPermission(_permissions);

            if (!granted)
            {
                LogDenied(context, user);
                context.Result = new ForbidResult();
            }

            return Task.CompletedTask;
        }

        private void LogDenied(AuthorizationFilterContext context, ClaimsPrincipal user)
        {
            var loggerFactory = context.HttpContext.RequestServices.GetService(typeof(ILoggerFactory)) as ILoggerFactory;
            var logger = loggerFactory?.CreateLogger("RequirePermission");

            var userId = user.FindFirst(ClaimTypes.NameIdentifier)?.Value ?? string.Empty;
            var username = user.FindFirst(ClaimTypes.Name)?.Value ?? "unknown";
            var required = string.Join(_requireAll ? " & " : " | ", _permissions);

            logger?.LogWarning(
                "[PERMISSION DENIED] User {Username} (ID: {UserId}) không có quyền [{Required}] cho {Path}",
                username, userId, required, context.HttpContext.Request.Path);
        }
    }
}