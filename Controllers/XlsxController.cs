using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using PresentationService.Models;
using PresentationService.Services;

namespace PresentationService.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class XlsxController : ControllerBase
    {
        private ICsvService _csvService;
        private INetworkService _networkService;
        public XlsxController(INetworkService networkService, ICsvService csvService)
        {
            _networkService = networkService;
            _csvService = csvService;
        }
        public async Task<IActionResult> Get([FromQuery]RequestParams dashboardParams)
        {
            if (dashboardParams.dashboardList == null)
            {
                return BadRequest();
            }
            var viewDataList = await _networkService.GetSheet(dashboardParams);
            var xlsx = _csvService.GenerateWorkbook(viewDataList);
            return xlsx;
        }
    }
}