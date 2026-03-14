namespace MOS.ExcelGrading.Core.DTOs
{
    public class UpdateUserPermissionsRequest
    {
        public List<string> Permissions { get; set; } = new List<string>();
    }
}
