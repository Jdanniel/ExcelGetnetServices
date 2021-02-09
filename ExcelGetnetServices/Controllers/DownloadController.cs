using ExcelGetnetServices.Entities.Request;
using ExcelGetnetServices.Interfaces;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGetnetServices.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class DownloadController : ControllerBase
    {
        private IDownload _idownload;
        public DownloadController(IDownload idownload)
        {
            _idownload = idownload;
        }
        [HttpGet("LAYOUTMASIVO")]
        public async Task<IActionResult> GetLayoutMasivo([FromQuery] LayoutMasivo request)
        {
            var res = await _idownload.LayoutMasivo(request);
            return File(
                res,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "servicios.xlsx");
        }
        [HttpGet("LAYOUTMASIVO_2")]
        public async Task<IActionResult> GetLayoutMasivo_2([FromQuery] LayoutMasivo2 request)
        {
            var res = await _idownload.LayoutMasivo2(request);
            return File(
                res,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "servicios.xlsx");
        }
        [HttpGet("LAYOUTMASIVO_3")]
        public async Task<IActionResult> GetLayoutMasivo3([FromQuery] LayoutMasivo3 request)
        {
            var res = await _idownload.LayoutMasivo3(request);
            return File(
                res,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "servicios.xlsx");
        }
        [HttpGet("LAYOUTMASIVO_4")]
        public async Task<IActionResult> GetLayoutMasivo4([FromQuery] LayoutMasivo4 request)
        {
            var res = await _idownload.LayoutMasivo4(request);
            return File(
                res,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "servicios.xlsx");
        }
        [HttpGet("LAYOUTMASIVO_USUARIO")]
        public async Task<IActionResult> GetLayoutMasivoUsuario([FromQuery] LayoutMasivoUsuario request)
        {
            var res = await _idownload.LayoutMasivoUsuario(request);
            return File(
                res,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "servicios.xlsx");
        }
        [HttpGet("LAYOUTMASIVO_MIT")]
        public async Task<IActionResult> GetLayoutMasivoMit([FromQuery] LayoutMasivoGetnetMit request)
        {
            var res = await _idownload.LayoutMasivoGetnetMit(request);
            return File(
                res,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "servicios.xlsx");
        }
        [HttpGet("LAYOUTMASIVO_REINGENIERIA")]
        public async Task<IActionResult> GetLayoutMasivoReingenieria([FromQuery] LayoutMasivoReingenieria request)
        {
            var res = await _idownload.LayoutMasivoReingenieria(request);
            return File(
                res,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "servicios.xlsx");
        }
    }
}
