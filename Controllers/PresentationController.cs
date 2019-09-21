/*
 * Created By Pratik Khandelwal
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using PresentationService.Models;
using PresentationService.Services;

namespace PresentationService.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class PresentationController : ControllerBase
    {
        private IPowerpointService _powerpointService;
        private INetworkService _networkService;
        public PresentationController (INetworkService networkService, IPowerpointService powerpointService)
        {
            _networkService = networkService;
            _powerpointService = powerpointService;
        }

        // GET api/presentation
        public async Task<IActionResult> Get([FromQuery]RequestParams dashboardParams)
        {
            if (dashboardParams.dashboardList == null)
            {
                return BadRequest();
            }
            var viewDataList = await _networkService.GetSheet(dashboardParams);
            var ppt = _powerpointService.GeneratePPT(viewDataList);
            return ppt;
        }
    }
}
