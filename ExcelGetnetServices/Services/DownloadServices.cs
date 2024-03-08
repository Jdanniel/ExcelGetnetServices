using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Threading.Tasks;
using ExcelGetnetServices.Interfaces;
using ExcelGetnetServices.Entities.Request;
using Microsoft.Data.SqlClient;
using ExcelGetnetServices.Data;
using Microsoft.EntityFrameworkCore;
using System.Text;
using System.Data;
using ExcelGetnetServices.Entities.StoredProcedures;
using Newtonsoft.Json;
using Microsoft.IdentityModel.Protocols;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Reflection.PortableExecutable;

namespace ExcelGetnetServices.Services
{
    public class DownloadServices : IDownload
    {
        private readonly ELAVONTESTContext _context;

        public DownloadServices(ELAVONTESTContext context)
        {
            _context = context;
        }
        public async Task AddLog(string Layout, string filters, int UserId)
        {
            BdmassiveLayoutLog bdmassiveLayoutLog = new BdmassiveLayoutLog()
            {
                DateDischarge = DateTime.Now,
                FiltersJson = filters,
                Layout = Layout,
                UserId = UserId
            };
            await _context.BdmassiveLayoutLogs.AddAsync(bdmassiveLayoutLog);
            await _context.SaveChangesAsync();
        }
        public async Task<byte[]> ExcelTest()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Servicios");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Id";
                worksheet.Cell(currentRow, 2).Value = "Username";

                worksheet.Cell(1, 1).Value = 1;
                worksheet.Cell(1, 2).Value = "Daniel";

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }
        }
        public async Task<byte[]> LayoutMasivo(LayoutMasivo request)
        {
            var statusarray = request.status_servicio?.Split(',');  // now you have an array of 3 strings
            bool isStatusthree = false;
            string status = "";

            if (statusarray != null)
            {
                for (int i = 0; i < statusarray.Length; i++)
                {
                    if (statusarray[i].Equals("3"))
                    {
                        isStatusthree = true;
                    }
                }
                if (isStatusthree)
                {
                    status = String.Join(",", statusarray);
                    status += ",4,5,13,35";
                }
                if (status.Length == 0)
                {
                    status = request.status_servicio;
                }
            }

            await AddLog("LAYOUTMASIVO", JsonConvert.SerializeObject(request), request.id_usuario);

            var param = new SqlParameter[]
            {
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini != null ? request.fec_ini : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin != null ? request.fec_fin : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROVEEDOR",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proveedor != null ? request.id_proveedor : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@STATUS_SERVICIO",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = status
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_ZONA",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_zona != null ? request.id_zona : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROYECTO",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proyecto
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini_cierre != null ? request.fec_ini_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin_cierre != null ? request.fec_fin_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@SERIE",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.serie
                },
            };
            _context.Database.SetCommandTimeout(4000);
            var data = await _context.SpLayoutMasivos.FromSqlRaw("SP_LAYOUT_MASIVO " +
                    "@FEC_INI, " +
                    "@FEC_FIN, " +
                    "@ID_PROVEEDOR," +
                    "@STATUS_SERVICIO, " +
                    "@ID_ZONA, " +
                    "@ID_PROYECTO, " +
                    "@FEC_INI_CIERRE, " +
                    "@FEC_FIN_CIERRE, " +
                    "@SERIE ", param).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Consulta");
                List<BdColumnsLayout> columns = await _context.BdColumnsLayouts
                        .Where(x => x.Status == true && x.IdLayout == 1)
                        .OrderBy(x => x.Orden)
                        .ToListAsync();

                int column = 1;
                for (int i = 0; i < columns.Count(); i++)
                {
                    worksheet.Cell(1, column).Value = columns[i].DescColumnLayout;
                    column++;
                }

                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.Bold = true;
                worksheet.Column(1).Style.NumberFormat.Format = "@";
                worksheet.Column(3).Style.NumberFormat.Format = "@";
                worksheet.Column(25).Style.NumberFormat.Format = "@";
                worksheet.Column(29).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
                worksheet.Column(65).Style.NumberFormat.Format = "@";
                int celda = 2;
                for (int i = 0; i < data.Count(); i++)
                {
                    worksheet.Cell(celda, 1).Value = data[i].ODT;
                    worksheet.Cell(celda, 2).Value = data[i].DISCOVER;
                    worksheet.Cell(celda, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(celda, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(celda, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(celda, 6).Value = data[i].COLONIA;
                    worksheet.Cell(celda, 7).Value = data[i].POBLACION;
                    worksheet.Cell(celda, 8).Value = data[i].ESTADO;
                    worksheet.Cell(celda, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(celda, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(celda, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(celda, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(celda, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(celda, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(celda, 15).Value = data[i].NIVEL;
                    worksheet.Cell(celda, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(celda, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(celda, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(celda, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(celda, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(celda, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(celda, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(celda, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(celda, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(celda, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(celda, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(celda, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(celda, 28).Value = data[i].MODELO;
                    worksheet.Cell(celda, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(celda, 30).Value = data[i].CAJA;
                    worksheet.Cell(celda, 31).Value = data[i].RFC;
                    worksheet.Cell(celda, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(celda, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(celda, 34).Value = data[i].TIEMPO_EN_ATENDER;
                    worksheet.Cell(celda, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(celda, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(celda, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(celda, 38).Value = data[i].CANAL;
                    worksheet.Cell(celda, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(celda, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(celda, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(celda, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(celda, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(celda, 44).Value = data[i].EMAIL;
                    worksheet.Cell(celda, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(celda, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(celda, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(celda, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(celda, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 51).Value = data[i].VERSIONSW;
                    worksheet.Cell(celda, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(celda, 53).Value = data[i].BASE;
                    worksheet.Cell(celda, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(celda, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(celda, 56).Value = data[i].ZONA;
                    worksheet.Cell(celda, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(celda, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(celda, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(celda, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(celda, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(celda, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(celda, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(celda, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(celda, 65).Value = data[i].CODIGO;
                    worksheet.Cell(celda, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(celda, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(celda, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(celda, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(celda, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(celda, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(celda, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(celda, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(celda, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(celda, 75).Value = data[i].SERIE;
                    worksheet.Cell(celda, 76).Value = data[i].COMENTARIOS?.ToString().Length >= 32767 ? data[i].COMENTARIOS?.ToString().Substring(0, 32766) : data[i].COMENTARIOS?.ToString();
                    celda++;
                }

                foreach (var itemcolumn in worksheet.ColumnsUsed())
                {
                    itemcolumn.AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }
        }
        public async Task<byte[]> LayoutMasivo2(LayoutMasivo2 request)
        {
            var statusarray = request.status_servicio?.Split(',');  // now you have an array of 3 strings
            bool isStatusthree = false;
            string status = "";

            if (statusarray != null)
            {
                for (int i = 0; i < statusarray.Length; i++)
                {
                    if (statusarray[i].Equals("3"))
                    {
                        isStatusthree = true;
                    }
                }
                if (isStatusthree)
                {
                    status = String.Join(",", statusarray);
                    status += ",4,5,13,35";
                }
                if (status.Length == 0)
                {
                    status = request.status_servicio;
                }
            }

            await AddLog("LAYOUTMASIVO_2", JsonConvert.SerializeObject(request), request.id_usuario);

            var param = new SqlParameter[]
            {
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini != null ? request.fec_ini : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin != null ? request.fec_fin : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROVEEDOR",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proveedor != null ? request.id_proveedor : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@STATUS_SERVICIO",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = status
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_ZONA",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_zona != null ? request.id_zona : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROYECTO",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proyecto
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini_cierre != null ? request.fec_ini_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin_cierre != null ? request.fec_fin_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@SERIE",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.serie
                },
            };
            _context.Database.SetCommandTimeout(4000);
            var data = await _context.SpLayoutMasivo2s.FromSqlRaw("SP_LAYOUT_MASIVO " +
                    "@FEC_INI, " +
                    "@FEC_FIN, " +
                    "@ID_PROVEEDOR," +
                    "@STATUS_SERVICIO, " +
                    "@ID_ZONA, " +
                    "@ID_PROYECTO, " +
                    "@FEC_INI_CIERRE, " +
                    "@FEC_FIN_CIERRE, " +
                    "@SERIE ", param).ToListAsync();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Reporte Servicios Por Dia");
                List<BdColumnsLayout> columns = await _context.BdColumnsLayouts
                        .Where(x => x.Status == true && x.IdLayout == 2)
                        .OrderBy(x => x.Orden)
                        .ToListAsync();

                int column = 1;
                for (int i = 0; i < columns.Count(); i++)
                {
                    worksheet.Cell(1, column).Value = columns[i].DescColumnLayout;
                    column++;
                }

                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.Bold = true;
                worksheet.Column(1).Style.NumberFormat.Format = "@";
                worksheet.Column(3).Style.NumberFormat.Format = "@";
                worksheet.Column(25).Style.NumberFormat.Format = "@";
                worksheet.Column(29).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
                worksheet.Column(65).Style.NumberFormat.Format = "@";
                int celda = 2;
                for (int i = 0; i < data.Count(); i++)
                {
                    worksheet.Cell(celda, 1).Value = data[i].ODT;
                    worksheet.Cell(celda, 2).Value = data[i].DISCOVER;
                    worksheet.Cell(celda, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(celda, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(celda, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(celda, 6).Value = data[i].COLONIA;
                    worksheet.Cell(celda, 7).Value = data[i].POBLACION;
                    worksheet.Cell(celda, 8).Value = data[i].ESTADO;
                    worksheet.Cell(celda, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(celda, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(celda, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(celda, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(celda, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(celda, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(celda, 15).Value = data[i].NIVEL;
                    worksheet.Cell(celda, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(celda, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(celda, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(celda, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(celda, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(celda, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(celda, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(celda, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(celda, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(celda, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(celda, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(celda, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(celda, 28).Value = data[i].MODELO;
                    worksheet.Cell(celda, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(celda, 30).Value = data[i].CAJA;
                    worksheet.Cell(celda, 31).Value = data[i].RFC;
                    worksheet.Cell(celda, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(celda, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(celda, 34).Value = data[i].TIEMPO_EN_ATENDER;
                    worksheet.Cell(celda, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(celda, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(celda, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(celda, 38).Value = data[i].CANAL;
                    worksheet.Cell(celda, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(celda, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(celda, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(celda, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(celda, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(celda, 44).Value = data[i].EMAIL;
                    worksheet.Cell(celda, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(celda, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(celda, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(celda, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(celda, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 51).Value = data[i].VERSIONSW;
                    worksheet.Cell(celda, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(celda, 53).Value = data[i].BASE;
                    worksheet.Cell(celda, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(celda, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(celda, 56).Value = data[i].ZONA;
                    worksheet.Cell(celda, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(celda, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(celda, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(celda, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(celda, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(celda, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(celda, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(celda, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(celda, 65).Value = data[i].CODIGO;
                    worksheet.Cell(celda, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(celda, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(celda, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(celda, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(celda, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(celda, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(celda, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(celda, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(celda, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(celda, 75).Value = data[i].SERIE;
                    worksheet.Cell(celda, 76).Value = data[i].COMENTARIOS?.ToString().Length >= 32767 ? data[i].COMENTARIOS?.ToString().Substring(0, 32766) : data[i].COMENTARIOS?.ToString();
                    celda++;
                }

                foreach (var itemcolumn in worksheet.ColumnsUsed())
                {
                    itemcolumn.AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }
        }
        public async Task<byte[]> LayoutMasivo3(LayoutMasivo3 request)
        {
            var statusarray = request.status_servicio?.Split(',');  // now you have an array of 3 strings
            bool isStatusthree = false;
            string status = "";

            if (statusarray != null)
            {
                for (int i = 0; i < statusarray.Length; i++)
                {
                    if (statusarray[i].Equals("3"))
                    {
                        isStatusthree = true;
                    }
                }
                if (isStatusthree)
                {
                    status = String.Join(",", statusarray);
                    status += ",4,5,13,35";
                }
                if (status.Length == 0)
                {
                    status = request.status_servicio;
                }
            }

            await AddLog("LAYOUTMASIVO_3", JsonConvert.SerializeObject(request), request.id_usuario);

            var param = new SqlParameter[]
            {
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_ini != null ? request.fec_ini : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_fin != null ? request.fec_fin : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROVEEDOR",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_proveedor != null ? request.id_proveedor : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@STATUS_SERVICIO",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = status
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_ZONA",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_zona != null ? request.id_zona : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROYECTO",
                    SqlDbType =SqlDbType.Int,
                    Direction =ParameterDirection.Input,
                    Value = request.id_proyecto
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI_CIERRE",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_ini_cierre != null ? request.fec_ini_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN_CIERRE",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_fin_cierre != null ? request.fec_fin_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@SERIE",
                    SqlDbType =SqlDbType.Int,
                    Direction =ParameterDirection.Input,
                    Value = request.serie
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_SERVICIO",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_servicio != null ? request.id_servicio : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_FALLA",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_falla != null ? request.id_falla : ""
                },
            };
            _context.Database.SetCommandTimeout(4000);
            var data = await _context.SpLayoutMasivo3s.FromSqlRaw("SP_LAYOUT_MASIVO_3 " +
                    "@FEC_INI, " +
                    "@FEC_FIN, " +
                    "@ID_PROVEEDOR," +
                    "@STATUS_SERVICIO, " +
                    "@ID_ZONA, " +
                    "@ID_PROYECTO, " +
                    "@FEC_INI_CIERRE, " +
                    "@FEC_FIN_CIERRE, " +
                    "@SERIE," +
                    "@ID_SERVICIO, " +
                    "@ID_FALLA", param).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Reporte Servicios Por Dia");
                List<BdColumnsLayout> columns = await _context.BdColumnsLayouts
                        .Where(x => x.Status == true && x.IdLayout == 3)
                        .OrderBy(x => x.Orden)
                        .ToListAsync();

                int column = 1;
                for(int i = 0; i < columns.Count(); i++)
                {
                    worksheet.Cell(1, column).Value = columns[i].DescColumnLayout;
                    column++;
                }

                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.Bold = true;
                worksheet.Column(1).Style.NumberFormat.Format = "@";
                worksheet.Column(3).Style.NumberFormat.Format = "@";
                worksheet.Column(25).Style.NumberFormat.Format = "@";
                worksheet.Column(29).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
                worksheet.Column(65).Style.NumberFormat.Format = "@";
                worksheet.Column(76).Style.NumberFormat.Format = "@";
                int celda = 2;
                for (int i = 0; i < data.Count(); i++)
                {
                    worksheet.Cell(celda, 1).Value = data[i].ODT;
                    worksheet.Cell(celda, 2).Value = data[i].MI_COMERCIO;
                    worksheet.Cell(celda, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(celda, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(celda, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(celda, 6).Value = data[i].COLONIA;
                    worksheet.Cell(celda, 7).Value = data[i].POBLACION;
                    worksheet.Cell(celda, 8).Value = data[i].ESTADO;
                    worksheet.Cell(celda, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(celda, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(celda, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(celda, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(celda, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(celda, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(celda, 15).Value = data[i].NIVEL;
                    worksheet.Cell(celda, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(celda, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(celda, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(celda, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(celda, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(celda, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(celda, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(celda, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(celda, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(celda, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(celda, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(celda, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(celda, 28).Value = data[i].MODELO;
                    worksheet.Cell(celda, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(celda, 30).Value = data[i].CAJA;
                    worksheet.Cell(celda, 31).Value = data[i].RFC;
                    worksheet.Cell(celda, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(celda, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(celda, 34).Value = data[i].VESTIDURAS_GETNET;
                    worksheet.Cell(celda, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(celda, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(celda, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(celda, 38).Value = data[i].CANAL;
                    worksheet.Cell(celda, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(celda, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(celda, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(celda, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(celda, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(celda, 44).Value = data[i].EMAIL;
                    worksheet.Cell(celda, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(celda, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(celda, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(celda, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(celda, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 51).Value = data[i].DIVISA;
                    worksheet.Cell(celda, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(celda, 53).Value = data[i].BASE;
                    worksheet.Cell(celda, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(celda, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(celda, 56).Value = data[i].ZONA;
                    worksheet.Cell(celda, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(celda, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(celda, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(celda, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(celda, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(celda, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(celda, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(celda, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(celda, 65).Value = data[i].CODIGO;
                    worksheet.Cell(celda, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(celda, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(celda, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(celda, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(celda, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(celda, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(celda, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(celda, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(celda, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(celda, 75).Value = data[i].AREA_CIERRA;
                    worksheet.Cell(celda, 76).Value = data[i].ODT_SALESFORCE?.TrimEnd();
                    worksheet.Cell(celda, 77).Value = data[i].COMENTARIOS?.ToString().Length >= 32767 ? data[i].COMENTARIOS?.ToString().Substring(0, 32766) : data[i].COMENTARIOS?.ToString();
                    celda++;
                }
                
                foreach (var itemcolumn in worksheet.ColumnsUsed())
                {
                    itemcolumn.AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }
        }
        public async Task<byte[]> LayoutMasivo4(LayoutMasivo4 request)
        {
            var statusarray = request.status_servicio?.Split(',');  // now you have an array of 3 strings
            bool isStatusthree = false;
            string status = "";

            if (statusarray != null)
            {
                for (int i = 0; i < statusarray.Length; i++)
                {
                    if (statusarray[i].Equals("3"))
                    {
                        isStatusthree = true;
                    }
                }
                if (isStatusthree)
                {
                    status = String.Join(",", statusarray);
                    status += ",4,5,13,35";
                }
                if (status.Length == 0)
                {
                    status = request.status_servicio;
                }
            }

            await AddLog("LAYOUTMASIVO_4", JsonConvert.SerializeObject(request), request.id_usuario);

            var param = new SqlParameter[]
            {
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini != null ? request.fec_ini : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin != null ? request.fec_fin : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROVEEDOR",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proveedor != null ? request.id_proveedor : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@STATUS_SERVICIO",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = status
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_ZONA",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_zona != null ? request.id_zona : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROYECTO",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proyecto
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini_cierre != null ? request.fec_ini_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin_cierre != null ? request.fec_fin_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@SERIE",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.serie
                },
            };
            _context.Database.SetCommandTimeout(4000);
            var data = await _context.SpLayoutMasivo4s.FromSqlRaw("SP_LAYOUT_MASIVO " +
                    "@FEC_INI, " +
                    "@FEC_FIN, " +
                    "@ID_PROVEEDOR," +
                    "@STATUS_SERVICIO, " +
                    "@ID_ZONA, " +
                    "@ID_PROYECTO, " +
                    "@FEC_INI_CIERRE, " +
                    "@FEC_FIN_CIERRE, " +
                    "@SERIE ", param).ToListAsync();
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Reporte Servicios Por Dia");
                List<BdColumnsLayout> columns = await _context.BdColumnsLayouts
                        .Where(x => x.Status == true && x.IdLayout == 2)
                        .OrderBy(x => x.Orden)
                        .ToListAsync();

                int column = 1;
                for (int i = 0; i < columns.Count(); i++)
                {
                    worksheet.Cell(1, column).Value = columns[i].DescColumnLayout;
                    column++;
                }

                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.Bold = true;
                worksheet.Column(1).Style.NumberFormat.Format = "@";
                worksheet.Column(3).Style.NumberFormat.Format = "@";
                worksheet.Column(25).Style.NumberFormat.Format = "@";
                worksheet.Column(29).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(65).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
                int celda = 2;
                for (int i = 0; i < data.Count(); i++)
                {
                    worksheet.Cell(celda, 1).Value = data[i].ODT;
                    worksheet.Cell(celda, 2).Value = data[i].DISCOVER;
                    worksheet.Cell(celda, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(celda, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(celda, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(celda, 6).Value = data[i].COLONIA;
                    worksheet.Cell(celda, 7).Value = data[i].POBLACION;
                    worksheet.Cell(celda, 8).Value = data[i].ESTADO;
                    worksheet.Cell(celda, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(celda, 11).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(celda, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(celda, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(celda, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(celda, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(celda, 15).Value = data[i].NIVEL;
                    worksheet.Cell(celda, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(celda, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(celda, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(celda, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(celda, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(celda, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(celda, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(celda, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(celda, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(celda, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(celda, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(celda, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(celda, 28).Value = data[i].MODELO;
                    worksheet.Cell(celda, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(celda, 30).Value = data[i].CAJA;
                    worksheet.Cell(celda, 31).Value = data[i].RFC;
                    worksheet.Cell(celda, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(celda, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(celda, 34).Value = data[i].TIEMPO_EN_ATENDER;
                    worksheet.Cell(celda, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(celda, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(celda, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(celda, 38).Value = data[i].CANAL;
                    worksheet.Cell(celda, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(celda, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(celda, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(celda, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(celda, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(celda, 44).Value = data[i].EMAIL;
                    worksheet.Cell(celda, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(celda, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(celda, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(celda, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(celda, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 51).Value = data[i].VERSIONSW;
                    worksheet.Cell(celda, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(celda, 53).Value = data[i].BASE;
                    worksheet.Cell(celda, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(celda, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(celda, 56).Value = data[i].ZONA;
                    worksheet.Cell(celda, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(celda, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(celda, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(celda, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(celda, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(celda, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(celda, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(celda, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(celda, 65).Value = data[i].CODIGO;
                    worksheet.Cell(celda, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(celda, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(celda, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(celda, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(celda, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(celda, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(celda, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(celda, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(celda, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(celda, 75).Value = data[i].SERIE;
                    celda++;
                }

                foreach (var itemcolumn in worksheet.ColumnsUsed())
                {
                    itemcolumn.AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }
        }
        public async Task<byte[]> LayoutMasivoUsuario(LayoutMasivoUsuario request)
        {
            var statusarray = request.status_servicio?.Split(',');  // now you have an array of 3 strings
            bool isStatusthree = false;
            string status = "";

            if (statusarray != null)
            {
                for (int i = 0; i < statusarray.Length; i++)
                {
                    if (statusarray[i].Equals("3"))
                    {
                        isStatusthree = true;
                    }
                }
                if (isStatusthree)
                {
                    status = String.Join(",", statusarray);
                    status += ",4,5,13,35";
                }
                if (status.Length == 0)
                {
                    status = request.status_servicio;
                }
            }

            await AddLog("LAYOUTMASIVO_USUARIO", JsonConvert.SerializeObject(request), request.id_usuario);

            var param = new SqlParameter[]
            {
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_ini != null ? request.fec_ini : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_fin != null ? request.fec_fin : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROVEEDOR",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_proveedor != null ? request.id_proveedor : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@STATUS_SERVICIO",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = status
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_ZONA",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_zona != null ? request.id_zona : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROYECTO",
                    SqlDbType =SqlDbType.Int,
                    Direction =ParameterDirection.Input,
                    Value = request.id_proyecto
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI_CIERRE",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_ini_cierre != null ? request.fec_ini_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN_CIERRE",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_fin_cierre != null ? request.fec_fin_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@SERIE",
                    SqlDbType =SqlDbType.Int,
                    Direction =ParameterDirection.Input,
                    Value = request.serie
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_SERVICIO",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.idservicio != null ? request.idservicio : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_FALLA",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.idfalla != null ? request.idfalla : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@RFC",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.rfc != null ? request.rfc : ""
                },
            };
            _context.Database.SetCommandTimeout(4000);
            try
            {
                List<SpLayoutMasivoUsuario> data = await _context.SpLayoutMasivoUsuarios.FromSqlRaw("SP_LAYOUT_MASIVO_USUARIO_2 " +
                        "@FEC_INI, " +
                        "@FEC_FIN, " +
                        "@STATUS_SERVICIO, " +
                        "@ID_ZONA, " +
                        "@ID_PROYECTO, " +
                        "@FEC_INI_CIERRE, " +
                        "@FEC_FIN_CIERRE, " +
                        "@SERIE, " +
                        "@ID_PROVEEDOR, " +
                        "@ID_SERVICIO, " +
                        "@ID_FALLA, " +
                        "@RFC ", param).ToListAsync();
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Reporte Servicios Por Dia");
                    List<BdColumnsLayout> columns = await _context.BdColumnsLayouts
                            .Where(x => x.Status == true && x.IdLayout == 5)
                            .OrderBy(x => x.Orden)
                            .ToListAsync();

                    int column = 1;
                    for (int i = 0; i < columns.Count(); i++)
                    {
                        worksheet.Cell(1, column).Value = columns[i].DescColumnLayout;
                        column++;
                    }

                    worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                    worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                    worksheet.Range(1, columns.Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                    worksheet.Range(1, columns.Count(), 1, 1).Style.Font.Bold = true;
                    worksheet.Column(1).Style.NumberFormat.Format = "@";
                    worksheet.Column(3).Style.NumberFormat.Format = "@";
                    worksheet.Column(25).Style.NumberFormat.Format = "@";
                    worksheet.Column(29).Style.NumberFormat.Format = "@";
                    worksheet.Column(30).Style.NumberFormat.Format = "@";
                    worksheet.Column(46).Style.NumberFormat.Format = "@";
                    worksheet.Column(47).Style.NumberFormat.Format = "@";
                    worksheet.Column(48).Style.NumberFormat.Format = "@";
                    worksheet.Column(49).Style.NumberFormat.Format = "@";
                    worksheet.Column(50).Style.NumberFormat.Format = "@";
                    worksheet.Column(65).Style.NumberFormat.Format = "@";
                    worksheet.Column(76).Style.NumberFormat.Format = "@";
                    int celda = 2;
                    for (int i = 0; i < data.Count(); i++)
                    {
                        worksheet.Cell(celda, 1).Value = data[i].ODT;
                        worksheet.Cell(celda, 2).Value = data[i].MI_COMERCIO;
                        worksheet.Cell(celda, 3).Value = data[i].AFILIACION;
                        worksheet.Cell(celda, 4).Value = data[i].COMERCIO;
                        worksheet.Cell(celda, 5).Value = data[i].DIRECCION;
                        worksheet.Cell(celda, 6).Value = data[i].COLONIA;
                        worksheet.Cell(celda, 7).Value = data[i].POBLACION;
                        worksheet.Cell(celda, 8).Value = data[i].ESTADO;
                        worksheet.Cell(celda, 9).Style.NumberFormat.Format = "@";
                        worksheet.Cell(celda, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                        worksheet.Cell(celda, 10).Style.NumberFormat.Format = "@";
                        worksheet.Cell(celda, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                        worksheet.Cell(celda, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                               .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                               new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                               Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                        worksheet.Cell(celda, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                               .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                               new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                               Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                        worksheet.Cell(celda, 13).Value = data[i].TELEFONO;
                        worksheet.Cell(celda, 14).Value = data[i].TIPO_COMERCIO_2;
                        worksheet.Cell(celda, 15).Value = data[i].NIVEL;
                        worksheet.Cell(celda, 16).Value = data[i].TIPO_SERVICIO;
                        worksheet.Cell(celda, 17).Value = data[i].SUB_TIPO_SERVICIO;
                        worksheet.Cell(celda, 18).Value = data[i].CRITERIO_CAMBIO;
                        worksheet.Cell(celda, 19).Value = data[i].ID_TECNICO;
                        worksheet.Cell(celda, 20).Value = data[i].PROVEEDOR;
                        worksheet.Cell(celda, 21).Value = data[i].ESTATUS_SERVICIO;
                        worksheet.Cell(celda, 22).Style.NumberFormat.Format = "@";
                        worksheet.Cell(celda, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                        worksheet.Cell(celda, 23).Style.NumberFormat.Format = "@";
                        worksheet.Cell(celda, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                        worksheet.Cell(celda, 24).Style.NumberFormat.Format = "@";
                        worksheet.Cell(celda, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                        worksheet.Cell(celda, 25).Value = data[i].CODIGO_POSTAL;
                        worksheet.Cell(celda, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                               .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                               new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                               Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                        worksheet.Cell(celda, 27).Value = data[i].CONECTIVIDAD;
                        worksheet.Cell(celda, 28).Value = data[i].MODELO;
                        worksheet.Cell(celda, 29).Value = data[i].EQUIPO;
                        worksheet.Cell(celda, 30).Value = data[i].CAJA;
                        worksheet.Cell(celda, 31).Value = data[i].RFC;
                        worksheet.Cell(celda, 32).Value = data[i].RAZON_SOCIAL;
                        worksheet.Cell(celda, 33).Value = data[i].HORAS_VENCIDAS;
                        worksheet.Cell(celda, 34).Value = data[i].OTRO_CLIENTE;
                        worksheet.Cell(celda, 35).Value = data[i].SLA_FIJO;
                        worksheet.Cell(celda, 36).Value = data[i].CODIGO_AFILIACION;
                        worksheet.Cell(celda, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                        worksheet.Cell(celda, 38).Value = data[i].CANAL;
                        worksheet.Cell(celda, 39).Value = data[i].AFILIACION_AMEX;
                        worksheet.Cell(celda, 40).Value = data[i].IDAMEX;
                        worksheet.Cell(celda, 41).Value = data[i].PRODUCTO;
                        worksheet.Cell(celda, 42).Value = data[i].MOTIVO_CANCELACION;
                        worksheet.Cell(celda, 43).Value = data[i].MOTIVO_RECHAZO;
                        worksheet.Cell(celda, 44).Value = data[i].EMAIL;
                        worksheet.Cell(celda, 45).Value = data[i].ROLLOS_A_INSTALAR;
                        worksheet.Cell(celda, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                        worksheet.Cell(celda, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                        worksheet.Cell(celda, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                        worksheet.Cell(celda, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                        worksheet.Cell(celda, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                        worksheet.Cell(celda, 51).Value = data[i].DIVISA;
                        worksheet.Cell(celda, 52).Value = data[i].CARGADOR;
                        worksheet.Cell(celda, 53).Value = data[i].BASE;
                        worksheet.Cell(celda, 54).Value = data[i].ROLLO_ENTREGADOS;
                        worksheet.Cell(celda, 55).Value = data[i].CABLE_CORRIENTE;
                        worksheet.Cell(celda, 56).Value = data[i].ZONA;
                        worksheet.Cell(celda, 57).Value = data[i].MODELO_INSTALADO;
                        worksheet.Cell(celda, 58).Value = data[i].MODELO_TERMINAL_SALE;
                        worksheet.Cell(celda, 59).Value = data[i].CORREO_EJECUTIVO;
                        worksheet.Cell(celda, 60).Value = data[i].RECHAZO;
                        worksheet.Cell(celda, 61).Value = data[i].CONTACTO1;
                        worksheet.Cell(celda, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                        worksheet.Cell(celda, 63).Value = data[i].TID_AMEX_CIERRE;
                        worksheet.Cell(celda, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                        worksheet.Cell(celda, 65).Value = data[i].CODIGO;
                        worksheet.Cell(celda, 66).Value = data[i].TIENE_AMEX;
                        worksheet.Cell(celda, 67).Value = data[i].ACT_REFERENCIAS;
                        worksheet.Cell(celda, 68).Value = data[i].TIPO_A_B;
                        worksheet.Cell(celda, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                        worksheet.Cell(celda, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                        worksheet.Cell(celda, 71).Value = data[i].AREA_CARGA;
                        worksheet.Cell(celda, 72).Value = data[i].ALTA_POR;
                        worksheet.Cell(celda, 73).Value = data[i].TIPO_CARGA;
                        worksheet.Cell(celda, 74).Value = data[i].CERRADO_POR;
                        worksheet.Cell(celda, 75).Value = data[i].AREA_CIERRA;
                        worksheet.Cell(celda, 76).Value = data[i].ODT_SALESFORCE?.TrimEnd();
                        worksheet.Cell(celda, 77).Value = data[i].COMENTARIOS?.ToString().Length >= 32767 ? data[i].COMENTARIOS?.ToString().Substring(0, 32766) : data[i].COMENTARIOS?.ToString();
                        worksheet.Cell(celda, 78).Value = "";
                        worksheet.Cell(celda, 79).Value = "";
                        worksheet.Cell(celda, 80).Value = "";
                        /*
                        worksheet.Cell(celda, 81).Value = data[i].INDUSTRIA;
                        worksheet.Cell(celda, 82).Value = data[i].MSI_SANTANDER;
                        worksheet.Cell(celda, 83).Value = data[i].MSI_PROSA;
                        worksheet.Cell(celda, 84).Value = data[i].AMEX_PLAN;
                        worksheet.Cell(celda, 85).Value = data[i].QPS;
                        worksheet.Cell(celda, 86).Value = data[i].CARNET;
                        worksheet.Cell(celda, 87).Value = data[i].APLICATIVO;*/



                        worksheet.Cell(celda, 81).Value = data[i].TIPO_DE_CONFIGURACION_MIT;
                        worksheet.Cell(celda, 82).Value = data[i].NegotiationType;
                        //worksheet.Cell(celda, 83).Style.NumberFormat.Format = "@";
                        //worksheet.Cell(celda, 83).Value = data[i].CASO_SF;
                        celda++;
                    }

                    foreach (var itemcolumn in worksheet.ColumnsUsed())
                    {
                        itemcolumn.AdjustToContents();
                    }

                    using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();

                        return content;
                    }
                }
            }
            catch(InvalidCastException ex)
            {
                System.Console.WriteLine(ex);
                return null;
            }


        }
        public async Task<byte[]> LayoutMasivoGetnetMit(LayoutMasivoGetnetMit request)
        {
            var statusarray = request.status_servicio?.Split(',');  // now you have an array of 3 strings
            bool isStatusthree = false;
            string status = "";

            if (statusarray != null)
            {
                for (int i = 0; i < statusarray.Length; i++)
                {
                    if (statusarray[i].Equals("3"))
                    {
                        isStatusthree = true;
                    }
                }
                if (isStatusthree)
                {
                    status = String.Join(",", statusarray);
                    status += ",4,5,13,35";
                }
                if (status.Length == 0)
                {
                    status = request.status_servicio;
                }
            }

            await AddLog("LAYOUTMASIVO_MIT", JsonConvert.SerializeObject(request), request.id_usuario);

            var param = new SqlParameter[]
            {
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_ini != null ? request.fec_ini : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_fin != null ? request.fec_fin : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROVEEDOR",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_proveedor != null ? request.id_proveedor : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@STATUS_SERVICIO",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = status
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_ZONA",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_zona != null ? request.id_zona : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROYECTO",
                    SqlDbType =SqlDbType.Int,
                    Direction =ParameterDirection.Input,
                    Value = request.id_proyecto
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI_CIERRE",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_ini_cierre != null ? request.fec_ini_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN_CIERRE",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.fec_fin_cierre != null ? request.fec_fin_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@SERIE",
                    SqlDbType =SqlDbType.Int,
                    Direction =ParameterDirection.Input,
                    Value = request.serie
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_SERVICIO",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_servicio != null ? request.id_servicio : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_FALLA",
                    SqlDbType =SqlDbType.VarChar,
                    Size = 50,
                    Direction =ParameterDirection.Input,
                    Value = request.id_falla != null ? request.id_falla : ""
                },
            };
            _context.Database.SetCommandTimeout(4000);
            var data = await _context.SpLayoutMasivoGetnetMits.FromSqlRaw("SP_LAYOUT_MASIVO_GETNET_MIT " +
                    "@FEC_INI, " +
                    "@FEC_FIN, " +
                    "@ID_PROVEEDOR," +
                    "@STATUS_SERVICIO, " +
                    "@ID_ZONA, " +
                    "@ID_PROYECTO, " +
                    "@FEC_INI_CIERRE, " +
                    "@FEC_FIN_CIERRE, " +
                    "@SERIE," +
                    "@ID_SERVICIO, " +
                    "@ID_FALLA", param).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Consulta");
                List<BdColumnsLayout> columns = await _context.BdColumnsLayouts
                        .Where(x => x.Status == true && x.IdLayout == 6)
                        .OrderBy(x => x.Orden)
                        .ToListAsync();

                int column = 1;
                for (int i = 0; i < columns.Count(); i++)
                {
                    worksheet.Cell(1, column).Value = columns[i].DescColumnLayout;
                    column++;
                }

                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.Bold = true;
                worksheet.Column(1).Style.NumberFormat.Format = "@";
                worksheet.Column(3).Style.NumberFormat.Format = "@";
                worksheet.Column(25).Style.NumberFormat.Format = "@";
                worksheet.Column(29).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
                worksheet.Column(65).Style.NumberFormat.Format = "@";
                worksheet.Column(76).Style.NumberFormat.Format = "@";
                worksheet.Column(96).Style.NumberFormat.Format = "@";
                int celda = 2;
                for (int i = 0; i < data.Count(); i++)
                {
                    worksheet.Cell(celda, 1).Value = data[i].ODT;
                    worksheet.Cell(celda, 2).Value = data[i].MI_COMERCIO;
                    worksheet.Cell(celda, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(celda, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(celda, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(celda, 6).Value = data[i].COLONIA;
                    worksheet.Cell(celda, 7).Value = data[i].POBLACION;
                    worksheet.Cell(celda, 8).Value = data[i].ESTADO;
                    worksheet.Cell(celda, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(celda, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(celda, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(celda, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(celda, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(celda, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(celda, 15).Value = data[i].NIVEL;
                    worksheet.Cell(celda, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(celda, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(celda, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(celda, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(celda, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(celda, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(celda, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(celda, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(celda, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(celda, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(celda, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(celda, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(celda, 28).Value = data[i].MODELO;
                    worksheet.Cell(celda, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(celda, 30).Value = data[i].CAJA;
                    worksheet.Cell(celda, 31).Value = data[i].RFC;
                    worksheet.Cell(celda, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(celda, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(celda, 34).Value = data[i].VESTIDURAS_GETNET;
                    worksheet.Cell(celda, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(celda, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(celda, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(celda, 38).Value = data[i].CANAL;
                    worksheet.Cell(celda, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(celda, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(celda, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(celda, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(celda, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(celda, 44).Value = data[i].EMAIL;
                    worksheet.Cell(celda, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(celda, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(celda, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(celda, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(celda, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 51).Value = data[i].DIVISA;
                    worksheet.Cell(celda, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(celda, 53).Value = data[i].BASE;
                    worksheet.Cell(celda, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(celda, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(celda, 56).Value = data[i].ZONA;
                    worksheet.Cell(celda, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(celda, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(celda, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(celda, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(celda, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(celda, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(celda, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(celda, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(celda, 65).Value = data[i].CODIGO;
                    worksheet.Cell(celda, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(celda, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(celda, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(celda, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(celda, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(celda, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(celda, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(celda, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(celda, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(celda, 75).Value = data[i].AREA_CIERRA;
                    worksheet.Cell(celda, 76).Value = data[i].ODT_SALESFORCE?.TrimEnd();
                    worksheet.Cell(celda, 77).Value = data[i].COMENTARIOS?.ToString().Length >= 32767 ? data[i].COMENTARIOS?.ToString().Substring(0, 32766) : data[i].COMENTARIOS?.ToString();
                    worksheet.Cell(celda, 78).Value = data[i].DESC_GIRO;
                    worksheet.Cell(celda, 79).Value = "";
                    worksheet.Cell(celda, 80).Value = data[i].MCC;
                    worksheet.Cell(celda, 81).Value = data[i].INDUSTRIA;
                    worksheet.Cell(celda, 82).Value = data[i].MSI_SANTANDER;
                    worksheet.Cell(celda, 83).Value = data[i].MSI_PROSA;
                    worksheet.Cell(celda, 84).Value = data[i].AMEX_PLAN;
                    worksheet.Cell(celda, 85).Value = data[i].QPS;
                    worksheet.Cell(celda, 86).Value = data[i].CARNET;
                    worksheet.Cell(celda, 87).Value = data[i].APLICATIVO;
                    worksheet.Cell(celda, 88).Value = data[i].TIPO_DE_CONFIGURACION_MIT;
                    worksheet.Cell(celda, 89).Value = data[i].NegotiationType;
                    worksheet.Cell(celda, 90).Value = data[i].GLOBALIZADOR;
                    worksheet.Cell(celda, 91).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 91).Value = data[i].IATA_AEROLINEA;
                    worksheet.Cell(celda, 92).Value = data[i].IATA_MATRIZ;
                    //worksheet.Cell(celda, 93).Value = data[i].IATA;
                    //worksheet.Cell(celda, 94).Value = data[i].AEROLINEA_RP3;
                    worksheet.Cell(celda, 93).Value = "";
                    worksheet.Cell(celda, 94).Value = data[i].CAMPANA;
                    //worksheet.Cell(celda, 96).Value = data[i].CASO_SF;
                    celda++;
                }

                foreach (var itemcolumn in worksheet.ColumnsUsed())
                {
                    itemcolumn.AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }

        }
        public async Task<byte[]> LayoutConsultaUnidades(ConsultaUnidades consulta)
        {
            var columHeaders = new string[]
            {
                "REPORTE",
                "PROVEEDOR",
                "AFILIACION",
                "SERIE",
                "CONTACLESS SI/NO",
                "NUMERO DE VECES QUE ESTUVO EN REPARACION",
                "ESTADO",
                "ZONA",
                "CODIGO POSTAL",
                "ULTIMA DIRECCION",
                "APLICATIVO",
                "CONECTIVIDAD",
                "PRODUCTO",
                "CATEGORIA",
                "PRECIO INICIO",
                "MODELO",
                "FACTURA_NO",
                "FECHA_FACTURA",
                "COSTO CON LA DEPRECIACION",
                "ESTATUS",
                "FECHA_ASIGNACION",
                "OBSERVACION/MOTIVO DE ENAJENACION",
                "ODT",
                "FECHA ATENCION PROVEEDOR"
            };

            await AddLog("LAYOUTMASIVO_UNIDADES", JsonConvert.SerializeObject(consulta), consulta.idusuario);

            List<SpGetConsultaUnidades> data = await _context.SpGetConsultaUnidades.FromSqlRaw("EXEC SP_GET_CONSULTA_UNIDADES " +
                "@p0,@p1,@p2,@p3,@p4,@p5,@p6,@p7,@p8,@p9,@p10,@p11,@p12",
                    consulta.search_text,
                    consulta.desc_unidad,
                    consulta.idcategoria,
                    consulta.idproveedor,
                    consulta.isdaniada,
                    consulta.idresponsable,
                    consulta.idtiporesponsable,
                    consulta.idaplicativo,
                    consulta.idconectividad,
                    consulta.idcliente,
                    consulta.idproducto,
                    consulta.isnueva,
                    consulta.idusuario).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Unidades");

                int column = 1;

                for (int i = 0; i < columHeaders.Count(); i++)
                {
                    worksheet.Cell(1, column).Value = columHeaders[i];
                    column++;
                }

                int celda = 2;

                for (int i = 0; i < data.Count(); i++)
                {
                    worksheet.Cell(celda, 1).Value = data[i].ODT;
                    worksheet.Cell(celda, 1).Value = data[i].REPORTE;
                    worksheet.Cell(celda, 2).Value = data[i].PROVEEDOR_2;
                    worksheet.Cell(celda, 3).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 3).Value = data[i].NO_AFILIACION == null ? "" : data[i].NO_AFILIACION;
                    worksheet.Cell(celda, 4).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 4).Value = data[i].NO_SERIE;
                    worksheet.Cell(celda, 5).Value = data[i].CONTACLESS;
                    worksheet.Cell(celda, 6).Value = data[i].NO_REPARACION;
                    worksheet.Cell(celda, 7).Value = data[i].ESTADO_RESPONSABLE;
                    worksheet.Cell(celda, 8).Value = data[i].ZONA_RESPONSABLE;
                    worksheet.Cell(celda, 9).Value = data[i].CP;
                    worksheet.Cell(celda, 10).Value = data[i].DIRECCION;
                    worksheet.Cell(celda, 11).Value = data[i].APLICATIVO;
                    worksheet.Cell(celda, 12).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(celda, 13).Value = data[i].PRODUCTO;
                    worksheet.Cell(celda, 14).Value = data[i].CATEGORIA_2;
                    worksheet.Cell(celda, 15).Value = data[i].PRECIO_INICIAL;
                    worksheet.Cell(celda, 16).Value = data[i].MODELO;
                    worksheet.Cell(celda, 17).Value = data[i].NO_FACTURA;
                    worksheet.Cell(celda, 18).Style.NumberFormat.Format = "dd/mm/yyyy hh:mm:ss";
                    worksheet.Cell(celda, 18).Value = data[i].FECHA_FACTURA;
                    worksheet.Cell(celda, 19).Value = data[i].COSTO_DEPRECIACION;
                    worksheet.Cell(celda, 20).Value = data[i].STATUS_UNIDAD;
                    worksheet.Cell(celda, 21).Style.NumberFormat.Format = "dd/mm/yyyy hh:mm:ss";
                    worksheet.Cell(celda, 21).Value = data[i].FECHA_ASIGNACION;
                    worksheet.Cell(celda, 22).Value = data[i].OBSERVACIONES;
                    worksheet.Cell(celda, 23).Value = data[i].ODT;
                    worksheet.Cell(celda, 24).Style.NumberFormat.Format = "dd/mm/yyyy hh:mm:ss";
                    worksheet.Cell(celda, 24).Value = data[i].FECHA_ATENCION_PROVEEDOR;
                    celda++;
                }

                foreach (var itemcolumn in worksheet.ColumnsUsed())
                {
                    itemcolumn.AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }
        }
        public async Task<byte[]> LayoutMasivoReingenieria(LayoutMasivoReingenieria request)
        {
            var statusarray = request.status_servicio?.Split(',');  // now you have an array of 3 strings
            bool isStatusthree = false;
            string status = "";

            if (statusarray != null)
            {
                for (int i = 0; i < statusarray.Length; i++)
                {
                    if (statusarray[i].Equals("3"))
                    {
                        isStatusthree = true;
                    }
                }
                if (isStatusthree)
                {
                    status = String.Join(",", statusarray);
                    status += ",4,5,13,35";
                }
                if (status.Length == 0)
                {
                    status = request.status_servicio;
                }
            }

            await AddLog("LAYOUTMASIVO_REINGENIERIA", JsonConvert.SerializeObject(request), request.id_usuario);

            var param = new SqlParameter[]
            {
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini != null ? request.fec_ini : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin != null ? request.fec_fin : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROVEEDOR",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proveedor != null ? request.id_proveedor : "-1"
                },
                new SqlParameter()
                {
                    ParameterName = "@STATUS_SERVICIO",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = status
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_ZONA",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_zona != null ? request.id_zona : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROYECTO",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proyecto != 0 ? request.id_proyecto : -1
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini_cierre != null ? request.fec_ini_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin_cierre != null ? request.fec_fin_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@SERIE",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.serie != 0 ? request.serie : -1
                },
            };
            _context.Database.SetCommandTimeout(4000);
            var data = await _context.SpLayoutMasivoReingenierias.FromSqlRaw("SP_LAYOUT_MASIVO_REINGENIERIA " +
                    "@FEC_INI, " +
                    "@FEC_FIN, " +
                    "@ID_PROVEEDOR," +
                    "@STATUS_SERVICIO, " +
                    "@ID_ZONA, " +
                    "@ID_PROYECTO, " +
                    "@FEC_INI_CIERRE, " +
                    "@FEC_FIN_CIERRE, " +
                    "@SERIE ", param).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Consulta");
                List<BdColumnsLayout> columns = await _context.BdColumnsLayouts
                        .Where(x => x.Status == true && x.IdLayout == 7)
                        .OrderBy(x => x.Orden)
                        .ToListAsync();

                int column = 1;
                for (int i = 0; i < columns.Count(); i++)
                {
                    worksheet.Cell(1, column).Value = columns[i].DescColumnLayout;
                    column++;
                }

                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.Bold = true;
                worksheet.Column(1).Style.NumberFormat.Format = "@";
                worksheet.Column(3).Style.NumberFormat.Format = "@";
                worksheet.Column(25).Style.NumberFormat.Format = "@";
                worksheet.Column(29).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
                worksheet.Column(65).Style.NumberFormat.Format = "@";
                int celda = 2;
                for (int i = 0; i < data.Count(); i++)
                {
                    worksheet.Cell(celda, 1).Value = data[i].ODT;
                    worksheet.Cell(celda, 2).Value = data[i].DISCOVER;
                    worksheet.Cell(celda, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(celda, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(celda, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(celda, 6).Value = data[i].COLONIA;
                    worksheet.Cell(celda, 7).Value = data[i].POBLACION;
                    worksheet.Cell(celda, 8).Value = data[i].ESTADO;
                    worksheet.Cell(celda, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(celda, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(celda, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(celda, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(celda, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(celda, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(celda, 15).Value = data[i].NIVEL;
                    worksheet.Cell(celda, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(celda, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(celda, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(celda, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(celda, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(celda, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(celda, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(celda, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(celda, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(celda, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(celda, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(celda, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(celda, 28).Value = data[i].MODELO;
                    worksheet.Cell(celda, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(celda, 30).Value = data[i].CAJA;
                    worksheet.Cell(celda, 31).Value = data[i].RFC;
                    worksheet.Cell(celda, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(celda, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(celda, 34).Value = data[i].VESTIDURAS_GETNET ;
                    worksheet.Cell(celda, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(celda, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(celda, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(celda, 38).Value = data[i].CANAL;
                    worksheet.Cell(celda, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(celda, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(celda, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(celda, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(celda, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(celda, 44).Value = data[i].EMAIL;
                    worksheet.Cell(celda, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(celda, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(celda, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(celda, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(celda, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(celda, 51).Value = data[i].VERSIONSW;
                    worksheet.Cell(celda, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(celda, 53).Value = data[i].BASE;
                    worksheet.Cell(celda, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(celda, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(celda, 56).Value = data[i].ZONA;
                    worksheet.Cell(celda, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(celda, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(celda, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(celda, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(celda, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(celda, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(celda, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(celda, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(celda, 65).Value = data[i].CODIGO;
                    worksheet.Cell(celda, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(celda, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(celda, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(celda, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(celda, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(celda, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(celda, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(celda, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(celda, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(celda, 75).Value = data[i].SERIE;
                    worksheet.Cell(celda, 76).Value = data[i].COMENTARIOS?.ToString().Length >= 32767 ? data[i].COMENTARIOS?.ToString().Substring(0, 32766) : data[i].COMENTARIOS?.ToString();
                    worksheet.Cell(celda, 77).Value = data[i].ELIMINADOR;
                    worksheet.Cell(celda, 78).Value = data[i].TAPA;
                    worksheet.Cell(celda, 79).Value = data[i].CONECTIVIDAD_INSTALADA;
                    worksheet.Cell(celda, 80).Value = data[i].CONECTIVIDAD_RETIRADA;
                    worksheet.Cell(celda, 81).Value = data[i].APLICATIVO_INSTALADO;
                    worksheet.Cell(celda, 82).Value = data[i].APLICATIVO_RETIRADO;
                    worksheet.Cell(celda, 83).Value = data[i].CARRIER_INSTALADO;
                    worksheet.Cell(celda, 84).Value = data[i].TIPO_DE_CONFIGURACION_MIT;
                    worksheet.Cell(celda, 85).Value = data[i].NegotiationType;
                    worksheet.Cell(celda, 86).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 86).Value = data[i].CASO_SF;
                    celda++;
                }

                foreach (var itemcolumn in worksheet.ColumnsUsed())
                {
                    itemcolumn.AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }
        }
        public async Task<byte[]> LayoutMasivoReingenieria2(LayoutMasivoReingenieria2 request)
        {
            var statusarray = request.status_servicio?.Split(',');  // now you have an array of 3 strings
            bool isStatusthree = false;
            string status = "";

            if (statusarray != null)
            {
                for (int i = 0; i < statusarray.Length; i++)
                {
                    if (statusarray[i].Equals("3"))
                    {
                        isStatusthree = true;
                    }
                }
                if (isStatusthree)
                {
                    status = String.Join(",", statusarray);
                    status += ",4,5,13,35";
                }
                if (status.Length == 0)
                {
                    status = request.status_servicio;
                }
            }

            await AddLog("LAYOUTMASIVO_REINGENIERIA2", JsonConvert.SerializeObject(request), request.id_usuario);

            var param = new SqlParameter[]
            {
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini != null ? request.fec_ini : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin != null ? request.fec_fin : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROVEEDOR",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proveedor != null ? request.id_proveedor : "-1"
                },
                new SqlParameter()
                {
                    ParameterName = "@STATUS_SERVICIO",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = status
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_ZONA",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_zona != null ? request.id_zona : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_PROYECTO",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.id_proyecto != 0 ? request.id_proyecto : -1
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_INI_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_ini_cierre != null ? request.fec_ini_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@FEC_FIN_CIERRE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.fec_fin_cierre != null ? request.fec_fin_cierre : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@SERIE",
                    SqlDbType = SqlDbType.Int,
                    Direction = ParameterDirection.Input,
                    Value = request.serie != 0 ? request.serie : -1
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_SERVICIO",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_servicio != null ? request.id_servicio : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@ID_FALLA",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.id_falla != null ? request.id_falla : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@NO_SERIE",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 70,
                    Direction = ParameterDirection.Input,
                    Value = request.noserie != null ? request.noserie : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@AFILIACION",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.afiliacion != null ? request.afiliacion : ""
                },
                new SqlParameter()
                {
                    ParameterName = "@RFC",
                    SqlDbType = SqlDbType.VarChar,
                    Size = 50,
                    Direction = ParameterDirection.Input,
                    Value = request.rfc != null ? request.rfc : ""
                },
            };
            _context.Database.SetCommandTimeout(4000);
            var data = await _context.SpLayoutMasivoReingenierias.FromSqlRaw("SP_LAYOUT_MASIVO_REINGENIERIA_2 " +
                    "@FEC_INI, " +
                    "@FEC_FIN, " +
                    "@ID_PROVEEDOR," +
                    "@STATUS_SERVICIO, " +
                    "@ID_ZONA, " +
                    "@ID_PROYECTO, " +
                    "@FEC_INI_CIERRE, " +
                    "@FEC_FIN_CIERRE, " +
                    "@SERIE, " +
                    "@ID_SERVICIO," +
                    "@ID_FALLA," +
                    "@NO_SERIE, " +
                    "@AFILIACION, " +
                    "@RFC ", param).ToListAsync();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Consulta");
                List<BdColumnsLayout> columns = await _context.BdColumnsLayouts
                        .Where(x => x.Status == true && x.IdLayout == 7)
                        .OrderBy(x => x.Orden)
                        .ToListAsync();

                int column = 1;
                for (int i = 0; i < columns.Count(); i++)
                {
                    worksheet.Cell(1, column).Value = columns[i].DescColumnLayout;
                    column++;
                }

                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                worksheet.Range(1, columns.Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                worksheet.Range(1, columns.Count(), 1, 1).Style.Font.Bold = true;
                worksheet.Column(1).Style.NumberFormat.Format = "@";
                worksheet.Column(3).Style.NumberFormat.Format = "@";
                worksheet.Column(25).Style.NumberFormat.Format = "@";
                worksheet.Column(29).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
                worksheet.Column(65).Style.NumberFormat.Format = "@";
                int celda = 2;
                for (int i = 0; i < data.Count(); i++)
                {
                    worksheet.Cell(celda, 1).Value = data[i].ODT;
                    worksheet.Cell(celda, 2).Value = data[i].AFILIACION;
                    worksheet.Cell(celda, 3).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(celda, 4).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(celda, 5).Value = data[i].CONCLUSIONES;
                    worksheet.Cell(celda, 6).Value = data[i].PROVEEDOR;
                    worksheet.Cell(celda, 7).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 7).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(celda, 8).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(celda, 9).Value = data[i].NUM_SERIE_TERMINAL_ENTRA;
                    worksheet.Cell(celda, 10).Value = data[i].NUM_SERIE_SIM_ENTRA;
                    worksheet.Cell(celda, 11).Value = data[i].APLICATIVO_INSTALADO;
                    worksheet.Cell(celda, 12).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(celda, 13).Value = data[i].CONECTIVIDAD_INSTALADA;
                    worksheet.Cell(celda, 14).Value = data[i].NUM_SERIE_TERMINAL_SALE;
                    worksheet.Cell(celda, 15).Value = data[i].NUM_SERIE_SIM_SALE;
                    worksheet.Cell(celda, 16).Value = data[i].APLICATIVO_RETIRADO;
                    worksheet.Cell(celda, 17).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(celda, 18).Value = data[i].CONECTIVIDAD_RETIRADA;
                    worksheet.Cell(celda, 19).Value = data[i].ZONA;
                    worksheet.Cell(celda, 20).Value = data[i].PRODUCTO;
                    worksheet.Cell(celda, 21).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(celda, 22).Value = data[i].MODELO;
                    worksheet.Cell(celda, 23).Value = data[i].DISCOVER;
                    worksheet.Cell(celda, 24).Value = data[i].COMERCIO;
                    worksheet.Cell(celda, 25).Value = data[i].DIRECCION;
                    worksheet.Cell(celda, 26).Value = data[i].COLONIA;
                    worksheet.Cell(celda, 27).Value = data[i].POBLACION;
                    worksheet.Cell(celda, 28).Value = data[i].ESTADO;
                    worksheet.Cell(celda, 29).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 29).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(celda, 30).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 30).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(celda, 31).Value = data[i].DESCRIPCION;
                    worksheet.Cell(celda, 32).Value = data[i].OBSERVACIONES;
                    worksheet.Cell(celda, 33).Value = data[i].TELEFONO;
                    worksheet.Cell(celda, 34).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(celda, 35).Value = data[i].NIVEL;
                    worksheet.Cell(celda, 36).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(celda, 37).Value = data[i].ID_TECNICO;
                    worksheet.Cell(celda, 38).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 38).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(celda, 39).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 39).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(celda, 40).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(celda, 41).Value = data[i].EQUIPO;
                    worksheet.Cell(celda, 42).Value = data[i].CAJA;
                    worksheet.Cell(celda, 43).Value = data[i].RFC;
                    worksheet.Cell(celda, 44).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(celda, 45).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(celda, 46).Value = data[i].VESTIDURAS_GETNET;
                    worksheet.Cell(celda, 47).Value = data[i].SLA_FIJO;
                    worksheet.Cell(celda, 48).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(celda, 49).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(celda, 50).Value = data[i].CANAL;
                    worksheet.Cell(celda, 51).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(celda, 52).Value = data[i].IDAMEX;
                    worksheet.Cell(celda, 53).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(celda, 54).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(celda, 55).Value = data[i].EMAIL;
                    worksheet.Cell(celda, 56).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(celda, 57).Value = data[i].NUM_SERIE_TERMINAL_MTO;
                    worksheet.Cell(celda, 58).Value = data[i].VERSIONSW;
                    worksheet.Cell(celda, 59).Value = data[i].CARGADOR;
                    worksheet.Cell(celda, 60).Value = data[i].BASE;
                    worksheet.Cell(celda, 61).Value = data[i].BATERIA;
                    worksheet.Cell(celda, 62).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(celda, 63).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(celda, 64).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(celda, 65).Value = data[i].RECHAZO;
                    worksheet.Cell(celda, 66).Value = data[i].CONTACTO1;
                    worksheet.Cell(celda, 67).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(celda, 68).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(celda, 69).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(celda, 70).Value = data[i].CODIGO;
                    worksheet.Cell(celda, 71).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(celda, 72).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(celda, 73).Value = data[i].TIPO_A_B;
                    worksheet.Cell(celda, 74).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(celda, 75).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(celda, 76).Value = data[i].AREA_CARGA;
                    worksheet.Cell(celda, 77).Value = data[i].ALTA_POR;
                    worksheet.Cell(celda, 78).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(celda, 79).Value = data[i].CERRADO_POR;
                    worksheet.Cell(celda, 80).Value = data[i].SERIE;
                    worksheet.Cell(celda, 81).Value = data[i].COMENTARIOS;
                    worksheet.Cell(celda, 82).Value = data[i].ELIMINADOR;
                    worksheet.Cell(celda, 83).Value = data[i].TAPA;
                    worksheet.Cell(celda, 84).Value = data[i].CARRIER_INSTALADO;
                    worksheet.Cell(celda, 85).Value = data[i].TIPO_DE_CONFIGURACION_MIT;
                    worksheet.Cell(celda, 86).Value = data[i].NegotiationType;
                    worksheet.Cell(celda, 87).Style.NumberFormat.Format = "@";
                    worksheet.Cell(celda, 87).Value = data[i].CASO_SF;
                    worksheet.Cell(celda, 88).Value = data[i].CAUSA_PENDIENTE_INVENTARIO;

                    celda++;
                }

                foreach (var itemcolumn in worksheet.ColumnsUsed())
                {
                    itemcolumn.AdjustToContents();
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    return content;
                }
            }
        }
        public async Task<string> CreateFileUnidadesExcel4(int idProducto, int idStatusUnidad, int idCliente, int idConectividad, int idSoftware, int isDaniada, int idTipoResponsable, string idResponsable, int idUsuario, string searchText)
        {
            try
            {

                string conn = "Server=192.168.101.3;Persist Security Info=True;connect timeout=400000;Database=MIC;User Id=api-excelservices-prod;Password=YmWASCIz4S0Kp6nxdHaWiTlQd;TrustServerCertificate=True;";
                
                using (XLWorkbook workbook = new XLWorkbook())
                {

                    var worksheet = workbook.Worksheets.Add("Consulta");

                    worksheet.Row(1).Height = 20;
                    worksheet.Row(1).Style.Font.Bold = true;

                    worksheet.Cell(1, 1).Value = "NO. UNIDAD";
                    worksheet.Cell(1, 2).Value = "CLIENTE";
                    worksheet.Cell(1, 3).Value = "ESTATUS CLIENTE";
                    worksheet.Cell(1, 4).Value = "MARCA";
                    worksheet.Cell(1, 5).Value = "MODELO";
                    worksheet.Cell(1, 6).Value = "CONECTIVIDAD";
                    worksheet.Cell(1, 7).Value = "APLICATIVO";
                    worksheet.Cell(1, 8).Value = "NO. SERIE";
                    worksheet.Cell(1, 9).Value = "REGION";
                    worksheet.Cell(1, 10).Value = "ZONA";
                    worksheet.Cell(1, 11).Value = "RESPONSABLE";
                    worksheet.Cell(1, 12).Value = "NO. AFILIACION";
                    worksheet.Cell(1, 13).Value = "ESTATUS RESPONSABLE UNIDAD";
                    worksheet.Cell(1, 14).Value = "TIPO RESPONSABLE";
                    worksheet.Cell(1, 15).Value = "NO. TARIMA";
                    worksheet.Cell(1, 16).Value = "NO. INVENTARIO";
                    worksheet.Cell(1, 17).Value = "NO. POSICION INVENTARIO";
                    worksheet.Cell(1, 18).Value = "DANADA";
                    worksheet.Cell(1, 19).Value = "PROPIEDAD DEL CLIENTE";
                    worksheet.Cell(1, 20).Value = "NO. SIM";
                    worksheet.Cell(1, 21).Value = "CARRIER";
                    worksheet.Cell(1, 22).Value = "SKU";
                    worksheet.Cell(1, 23).Value = "NIVEL DIAGNOSTICO";
                    worksheet.Cell(1, 24).Value = "ESTATUS";
                    worksheet.Cell(1, 25).Value = "COMENTARIOS";
                    worksheet.Cell(1, 26).Value = "FECHA ULTIMO MOVIMIENTO";
                    worksheet.Cell(1, 27).Value = "ULTIMO MOVIMIENTO USUARIO";
                    worksheet.Cell(1, 28).Value = "ID. ENVIO";
                    worksheet.Cell(1, 29).Value = "NO. GUIA";
                    worksheet.Cell(1, 30).Value = "RESPONSABLE ORIGEN";
                    worksheet.Cell(1, 31).Value = "ESTATUS RESPONSABLE ORIGEN";
                    worksheet.Cell(1, 32).Value = "REGION ORIGEN";
                    worksheet.Cell(1, 33).Value = "ZONA ORIGEN";
                    worksheet.Cell(1, 34).Value = "RESPONSABLE DESTINO";
                    worksheet.Cell(1, 35).Value = "ESTATUS RESPONSABLE DESTINO";
                    worksheet.Cell(1, 36).Value = "REGION DESTINO";
                    worksheet.Cell(1, 37).Value = "ZONA DESTINO";
                    worksheet.Cell(1, 38).Value = "COSTO";
                    worksheet.Cell(1, 39).Value = "IS NUEVA";

                    if (idCliente == 85)
                    {
                        worksheet.Cell(1, 40).Value = "FOLIO TELMEX";
                        worksheet.Cell(1, 41).Value = "PLACA";
                        worksheet.Cell(1, 42).Value = "VIGENCIA";
                    }
                    worksheet.Range(1, worksheet.ColumnsUsed().Count(), 1, 1).Style.Fill.BackgroundColor = XLColor.FromArgb(0, 128, 255);
                    worksheet.Range(1, worksheet.ColumnsUsed().Count(), 1, 1).Style.Fill.PatternType = XLFillPatternValues.Solid;
                    worksheet.Range(1, worksheet.ColumnsUsed().Count(), 1, 1).Style.Font.FontColor = XLColor.White;
                    worksheet.Range(1, worksheet.ColumnsUsed().Count(), 1, 1).Style.Font.Bold = true;

                    using (SqlConnection connection = new SqlConnection())
                    {
                        connection.ConnectionString = conn;

                        using (SqlCommand cmd = new SqlCommand())
                        {
                            cmd.CommandTimeout = 1200;

                            // cmd.CommandText = "SP_GET_UNIDADES_EXCEL_4";

                            cmd.CommandText = "SP_GET_UNIDADES_EXCEL_4_v1";
                            cmd.CommandType = System.Data.CommandType.StoredProcedure;
                            cmd.Parameters.Add("@ID_PRODUCTO", System.Data.SqlDbType.Int).Value = idProducto;
                            cmd.Parameters.Add("@ID_STATUS_UNIDAD", System.Data.SqlDbType.Int).Value = idStatusUnidad;
                            cmd.Parameters.Add("@ID_CLIENTE", System.Data.SqlDbType.Int).Value = idCliente;
                            cmd.Parameters.Add("@ID_CONECTIVIDAD", System.Data.SqlDbType.Int).Value = idConectividad;
                            cmd.Parameters.Add("@ID_SOFTWARE", System.Data.SqlDbType.Int).Value = idSoftware;
                            cmd.Parameters.Add("@IS_DANIADA", System.Data.SqlDbType.Int).Value = isDaniada;
                            cmd.Parameters.Add("@ID_TIPO_RESPONSABLE", System.Data.SqlDbType.Int).Value = idTipoResponsable;
                            cmd.Parameters.Add("@ID_RESPONSABLE", System.Data.SqlDbType.VarChar).Value = idResponsable;
                            cmd.Parameters.Add("@ID_USUARIO", System.Data.SqlDbType.Int).Value = idUsuario;
                            cmd.Parameters.Add("@SEARCH_TEXT", System.Data.SqlDbType.VarChar).Value = searchText != null ? searchText : "";

                            cmd.Connection = connection;

                            connection.Open();

                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                int recordIndex = 2;

                                while (reader.Read())
                                {
                                    worksheet.Cell(recordIndex, 1).Value = reader.GetInt32(0);
                                    worksheet.Cell(recordIndex, 2).Value = reader.IsDBNull(10) ? String.Empty : reader.GetString(10);//10
                                    worksheet.Cell(recordIndex, 3).Value = reader.IsDBNull(26) ? String.Empty : reader.GetString(26);//26
                                    worksheet.Cell(recordIndex, 4).Value = reader.IsDBNull(59) ? String.Empty : reader.GetString(59);
                                    worksheet.Cell(recordIndex, 5).Value = reader.IsDBNull(56) ? String.Empty : reader.GetString(56);//30
                                    worksheet.Cell(recordIndex, 6).Value = reader.IsDBNull(54) ? String.Empty : reader.GetString(54);
                                    worksheet.Cell(recordIndex, 7).Value = reader.IsDBNull(55) ? String.Empty : reader.GetString(55);
                                    worksheet.Cell(recordIndex, 8).Style.NumberFormat.Format = "@";
                                    worksheet.Cell(recordIndex, 8).Value = reader.IsDBNull(12) ? String.Empty : reader.GetString(12);
                                    worksheet.Cell(recordIndex, 9).Value = reader.IsDBNull(30) ? String.Empty : reader.GetString(30);
                                    worksheet.Cell(recordIndex, 10).Value = reader.IsDBNull(31) ? String.Empty : reader.GetString(31);
                                    worksheet.Cell(recordIndex, 11).Value = reader.IsDBNull(32) ? String.Empty : reader.GetString(32);
                                    worksheet.Cell(recordIndex, 12).Value = reader.IsDBNull(36) ? String.Empty : reader.GetString(36);
                                    worksheet.Cell(recordIndex, 13).Value = reader.IsDBNull(35) ? String.Empty : reader.GetString(35);
                                    worksheet.Cell(recordIndex, 14).Value = reader.IsDBNull(53) ? String.Empty : reader.GetString(53);
                                    worksheet.Cell(recordIndex, 15).Value = reader.IsDBNull(13) ? String.Empty : reader.GetString(13);
                                    worksheet.Cell(recordIndex, 16).Value = reader.IsDBNull(14) ? String.Empty : reader.GetString(14);
                                    worksheet.Cell(recordIndex, 17).Value = reader.IsDBNull(16) ? String.Empty : reader.GetString(16);
                                    worksheet.Cell(recordIndex, 18).Value = reader.IsDBNull(37) ? String.Empty : reader.GetString(37);
                                    worksheet.Cell(recordIndex, 19).Value = reader.IsDBNull(38) ? String.Empty : reader.GetString(38);
                                    worksheet.Cell(recordIndex, 20).Value = reader.IsDBNull(19) ? String.Empty : reader.GetString(19);
                                    worksheet.Cell(recordIndex, 21).Value = reader.IsDBNull(20) ? String.Empty : reader.GetString(20);
                                    worksheet.Cell(recordIndex, 22).Value = reader.IsDBNull(49) ? String.Empty : reader.GetString(49);
                                    worksheet.Cell(recordIndex, 23).Value = reader.IsDBNull(50) ? String.Empty : reader.GetString(50);
                                    if (reader.GetInt32(17) == 10 && reader.GetInt32(11) == 85)
                                    {
                                        worksheet.Cell(recordIndex, 24).Value = "Como Dato";
                                    }
                                    else if (reader.GetInt32(17) == 30 && reader.GetInt32(11) == 85)
                                    {
                                        worksheet.Cell(recordIndex, 24).Value = "Facturable";

                                    }
                                    else
                                    {
                                        worksheet.Cell(recordIndex, 24).Value = reader.GetString(60);
                                    }
                                    worksheet.Cell(recordIndex, 25).Value = reader.IsDBNull(18) ? String.Empty : reader.GetString(18);
                                    // worksheet.Cell(recordIndex, 26).Value = reader["ULTIMA_FECHA"] +" " + reader["ULTIMA_HORA"];
                                    worksheet.Cell(recordIndex, 26).Value = reader.IsDBNull(51) ? String.Empty : reader.GetString(51);
                                    worksheet.Cell(recordIndex, 27).Value = reader.IsDBNull(52) ? String.Empty : reader.GetString(52);
                                    worksheet.Cell(recordIndex, 28).Value = reader.IsDBNull(39) ? 0 : reader.GetInt32(39);
                                    worksheet.Cell(recordIndex, 29).Value = reader.IsDBNull(40) ? String.Empty : reader.GetString(40);
                                    worksheet.Cell(recordIndex, 30).Value = reader.IsDBNull(41) ? String.Empty : reader.GetString(41);
                                    worksheet.Cell(recordIndex, 31).Value = reader.IsDBNull(33) ? String.Empty : reader.GetString(33);
                                    worksheet.Cell(recordIndex, 32).Value = reader.IsDBNull(43) ? String.Empty : reader.GetString(43);
                                    worksheet.Cell(recordIndex, 33).Value = reader.IsDBNull(44) ? String.Empty : reader.GetString(44);
                                    worksheet.Cell(recordIndex, 34).Value = reader.IsDBNull(42) ? String.Empty : reader.GetString(42);
                                    worksheet.Cell(recordIndex, 35).Value = reader.IsDBNull(34) ? String.Empty : reader.GetString(34);
                                    worksheet.Cell(recordIndex, 36).Value = reader.IsDBNull(45) ? String.Empty : reader.GetString(45);
                                    worksheet.Cell(recordIndex, 37).Value = reader.IsDBNull(46) ? String.Empty : reader.GetString(46);
                                    worksheet.Cell(recordIndex, 38).Style.NumberFormat.Format = "@";
                                    worksheet.Cell(recordIndex, 38).Value = reader.IsDBNull(57) ? 0.00 : reader.GetDecimal(57);
                                    worksheet.Cell(recordIndex, 39).Value = reader.IsDBNull(5) ? 0 : reader.GetInt32(5);
                                    if (reader.GetInt32(11) == 85)
                                    {
                                        worksheet.Cell(recordIndex, 40).Value = reader.IsDBNull(21) ? String.Empty : reader.GetString(21);
                                        worksheet.Cell(recordIndex, 41).Value = reader.IsDBNull(22) ? String.Empty : reader.GetString(22);
                                        if (reader.GetInt32(23) == 1)
                                        {
                                            worksheet.Cell(recordIndex, 42).Value = reader.IsDBNull(24) ? String.Empty : reader.GetString(24);
                                        }
                                        else
                                        {
                                            worksheet.Cell(recordIndex, 42).Value = "EN VENTA";
                                        }
                                    }

                                    recordIndex++;
                                }
                                connection.Close();
                            }

                        }
                    }
                    /*
                    foreach (var itemcolumn in worksheet.ColumnsUsed())
                    {
                        itemcolumn.AdjustToContents();
                    }*/
                    string excelName = "ConsultaUnidades" + DateTime.Now.Day + DateTime.Now.Month + DateTime.Now.Year + DateTime.Now.Minute + DateTime.Now.Millisecond;
                    workbook.SaveAs(@$"C:\inetpub\wwwroot\MIC3\UNIDADES_2\ARCHIVOS\{excelName}.xlsx");
                    return excelName;
                    /*using (var stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        var content = stream.ToArray();

                        return content;
                    }*/
                }
            }
            catch (Exception e)
            {
                await AddLog("LAYOUTMASIVO_REINGENIERIA", e.Message, 0);
                return $"Error:{e.Message}";
            }
        }
    }
}
