using System.ComponentModel.DataAnnotations;

namespace MOS.ExcelGrading.Core.DTOs
{
    public class CreateComputerRoomRequest
    {
        [Required(ErrorMessage = "SchoolId là bắt buộc")]
        public string SchoolId { get; set; } = string.Empty;

        [Required(ErrorMessage = "Tên phòng máy là bắt buộc")]
        [StringLength(100, MinimumLength = 2)]
        public string Name { get; set; } = string.Empty;

        [Range(0, 1000, ErrorMessage = "Số máy học sinh phải >= 0")]
        public int StudentMachineCount { get; set; }

        [Range(0, 50, ErrorMessage = "Số máy giáo viên phải >= 0")]
        public int TeacherMachineCount { get; set; } = 1;

        [Range(0, 1000, ErrorMessage = "Số máy hỏng phải >= 0")]
        public int BrokenMachineCount { get; set; } = 0;

        [StringLength(120)]
        public string? NetSupportStatus { get; set; }

        [StringLength(120)]
        public string? AudioStatus { get; set; }

        [StringLength(120)]
        public string? CoolingStatus { get; set; }

        [StringLength(120)]
        public string? DevicesPoweredOffStatus { get; set; }

        [StringLength(120)]
        public string? SeatingOrderStatus { get; set; }

        [StringLength(120)]
        public string? RoomHygieneStatus { get; set; }
    }

    public class UpdateComputerRoomRequest
    {
        [StringLength(100, MinimumLength = 2)]
        public string? Name { get; set; }

        [Range(0, 1000, ErrorMessage = "Số máy học sinh phải >= 0")]
        public int? StudentMachineCount { get; set; }

        [Range(0, 50, ErrorMessage = "Số máy giáo viên phải >= 0")]
        public int? TeacherMachineCount { get; set; }

        [Range(0, 1000, ErrorMessage = "Số máy hỏng phải >= 0")]
        public int? BrokenMachineCount { get; set; }

        [StringLength(120)]
        public string? NetSupportStatus { get; set; }

        [StringLength(120)]
        public string? AudioStatus { get; set; }

        [StringLength(120)]
        public string? CoolingStatus { get; set; }

        [StringLength(120)]
        public string? DevicesPoweredOffStatus { get; set; }

        [StringLength(120)]
        public string? SeatingOrderStatus { get; set; }

        [StringLength(120)]
        public string? RoomHygieneStatus { get; set; }

        public bool? IsActive { get; set; }
    }

    public class ComputerRoomResponse
    {
        public string Id { get; set; } = string.Empty;
        public string SchoolId { get; set; } = string.Empty;
        public string OwnerId { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public int StudentMachineCount { get; set; }
        public int TeacherMachineCount { get; set; }
        public int BrokenMachineCount { get; set; }
        public int AvailableStudentMachines { get; set; }
        public int TotalMachineCount { get; set; }
        public string TotalMachinesText { get; set; } = string.Empty;
        public string NetSupportStatus { get; set; } = string.Empty;
        public string AudioStatus { get; set; } = string.Empty;
        public string CoolingStatus { get; set; } = string.Empty;
        public string DevicesPoweredOffStatus { get; set; } = string.Empty;
        public string SeatingOrderStatus { get; set; } = string.Empty;
        public string RoomHygieneStatus { get; set; } = string.Empty;
        public bool IsActive { get; set; }
        public DateTime CreatedAt { get; set; }
        public DateTime? UpdatedAt { get; set; }
    }
}
