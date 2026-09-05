using System.Security.Claims;

namespace MOS.ExcelGrading.API.Extensions
{
    public static class ClaimsPrincipalPermissionExtensions
    {
        private const string PermissionClaimType = "permission";

        public static bool HasPermission(this ClaimsPrincipal user, string permission)
        {
            if (string.IsNullOrWhiteSpace(permission))
            {
                return false;
            }

            return user.Claims.Any(c => c.Type == PermissionClaimType && c.Value == permission);
        }

        public static bool HasAnyPermission(this ClaimsPrincipal user, params string[] permissions)
        {
            if (permissions == null || permissions.Length == 0)
            {
                return false;
            }

            var userPermissions = GetPermissionSet(user);
            return permissions.Any(userPermissions.Contains);
        }

        public static bool HasAllPermissions(this ClaimsPrincipal user, params string[] permissions)
        {
            if (permissions == null || permissions.Length == 0)
            {
                return true;
            }

            var userPermissions = GetPermissionSet(user);
            return permissions.All(userPermissions.Contains);
        }

        private static HashSet<string> GetPermissionSet(ClaimsPrincipal user)
        {
            return user.Claims
                .Where(c => c.Type == PermissionClaimType)
                .Select(c => c.Value)
                .ToHashSet(StringComparer.Ordinal);
        }
    }
}