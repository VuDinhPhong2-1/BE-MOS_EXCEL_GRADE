namespace MOS.ExcelGrading.Core.Models
{
    public sealed class GraderRouteDescriptor
    {
        public string ExamType { get; init; } = AssignmentExamTypes.OTTH;
        public string Subject { get; init; } = AssignmentFileSubjects.Excel;
        public string ProjectCode { get; init; } = string.Empty;
        public string? GradingApiEndpoint { get; init; }
        public bool IsRuntimeSupported { get; init; }
        public string Family { get; init; } = "OTTH";
    }

    public static class GraderRouteRegistry
    {
        public static bool TryResolve(
            string? examType,
            string? subject,
            string? projectCode,
            string? gradingApiEndpoint,
            out GraderRouteDescriptor descriptor)
        {
            descriptor = new GraderRouteDescriptor();

            var normalizedExamType = AssignmentExamTypes.Normalize(examType);
            var normalizedSubject = AssignmentFileSubjects.Normalize(subject);
            var normalizedEndpoint = string.IsNullOrWhiteSpace(gradingApiEndpoint)
                ? null
                : GradingApiEndpoints.NormalizeEndpoint(gradingApiEndpoint);

            if (!string.IsNullOrWhiteSpace(normalizedEndpoint))
            {
                if (!GradingApiEndpoints.TryExtractSubject(normalizedEndpoint, out var endpointSubject) ||
                    !GradingApiEndpoints.TryExtractProjectNumber(normalizedEndpoint, out var projectNumber))
                {
                    return false;
                }

                normalizedSubject = endpointSubject;
                descriptor = new GraderRouteDescriptor
                {
                    ExamType = normalizedExamType,
                    Subject = normalizedSubject,
                    ProjectCode = NormalizeProjectCode(projectCode, normalizedSubject, projectNumber),
                    GradingApiEndpoint = normalizedEndpoint,
                    IsRuntimeSupported = normalizedExamType == AssignmentExamTypes.OTTH ||
                                         normalizedExamType == AssignmentExamTypes.OnThi,
                    Family = normalizedExamType == AssignmentExamTypes.GMetrix ? "GMetrix" : "OTTH"
                };

                return true;
            }

            if (normalizedExamType == AssignmentExamTypes.GMetrix && !string.IsNullOrWhiteSpace(normalizedSubject))
            {
                descriptor = new GraderRouteDescriptor
                {
                    ExamType = normalizedExamType,
                    Subject = normalizedSubject,
                    ProjectCode = NormalizeProjectCode(projectCode, normalizedSubject, null),
                    GradingApiEndpoint = null,
                    IsRuntimeSupported = false,
                    Family = "GMetrix"
                };
                return true;
            }

            return false;
        }

        private static string NormalizeProjectCode(string? projectCode, string subject, int? projectNumber)
        {
            if (!string.IsNullOrWhiteSpace(projectCode))
            {
                return projectCode.Trim().ToUpperInvariant();
            }

            if (projectNumber.HasValue)
            {
                return $"{subject.ToUpperInvariant()}_P{projectNumber.Value:00}";
            }

            return $"{subject.ToUpperInvariant()}_GENERIC";
        }
    }
}
