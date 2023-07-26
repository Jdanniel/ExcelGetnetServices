using ExcelGetnetServices.Entities.Request;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelGetnetServices.Interfaces
{
    public interface IDownload
    {
        Task<byte[]> ExcelTest();
        Task<byte[]> LayoutMasivo(LayoutMasivo request);
        Task<byte[]> LayoutMasivo2(LayoutMasivo2 request);
        Task<byte[]> LayoutMasivo3(LayoutMasivo3 request);
        Task<byte[]> LayoutMasivo4(LayoutMasivo4 request);
        Task<byte[]> LayoutMasivoUsuario(LayoutMasivoUsuario request);
        Task<byte[]> LayoutMasivoGetnetMit(LayoutMasivoGetnetMit request);
        Task<byte[]> LayoutMasivoReingenieria(LayoutMasivoReingenieria request);
        Task<byte[]> LayoutMasivoReingenieria2(LayoutMasivoReingenieria2 request);
        Task<byte[]> LayoutConsultaUnidades(ConsultaUnidades consulta);
        Task<string> CreateFileUnidadesExcel4(int idProducto, int idStatusUnidad, int idCliente, int idConectividad, int idSoftware, int isDaniada, int idTipoResponsable, string idResponsable, int idUsuario, string searchText);
    }
}
