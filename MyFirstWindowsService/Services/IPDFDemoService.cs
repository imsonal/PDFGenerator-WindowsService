    using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyFirstWindowsService.Services
{
    public interface IPDFDemoService
    {
         string GetRootPath();
         string GenerateSummaryReportBySummaryDetailsAsync();
         Byte[] GenerateGraphImage();
    }
}
