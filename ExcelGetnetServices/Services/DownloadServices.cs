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

namespace ExcelGetnetServices.Services
{
    public class DownloadServices : IDownload
    {
        private readonly ELAVONTESTContext _context;

        public DownloadServices(ELAVONTESTContext context)
        {
            _context = context;
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
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
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
                    worksheet.Cell(celda, 76).Value = data[i].COMENTARIOS;
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
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
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
                    worksheet.Cell(celda, 76).Value = data[i].COMENTARIOS;
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
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
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
                    worksheet.Cell(celda, 77).Value = data[i].COMENTARIOS;
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
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
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
            };
            List<SpLayoutMasivoUsuario> data = await _context.SpLayoutMasivoUsuarios.FromSqlRaw("SP_LAYOUT_MASIVO_USUARIO " +
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
                    "@ID_FALLA", param).ToListAsync();
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
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
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
                    worksheet.Cell(celda, 77).Value = data[i].COMENTARIOS;
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
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
                worksheet.Column(50).Style.NumberFormat.Format = "@";
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
                    worksheet.Cell(celda, 77).Value = data[i].COMENTARIOS;
                    worksheet.Cell(celda, 78).Value = data[i].DESC_GIRO;
                    worksheet.Cell(celda, 79).Value = "";
                    worksheet.Cell(celda, 80).Value = "";
                    worksheet.Cell(celda, 81).Value = data[i].MCC;
                    worksheet.Cell(celda, 82).Value = data[i].INDUSTRIA;
                    worksheet.Cell(celda, 83).Value = data[i].MSI_SANTANDER;
                    worksheet.Cell(celda, 84).Value = data[i].MSI_PROSA;
                    worksheet.Cell(celda, 85).Value = data[i].AMEX_PLAN;
                    worksheet.Cell(celda, 86).Value = data[i].QPS;
                    worksheet.Cell(celda, 87).Value = data[i].CARNET;
                    worksheet.Cell(celda, 88).Value = data[i].APLICATIVO;
                    worksheet.Cell(celda, 89).Value = data[i].TIPO_DE_CONFIGURACION_MIT;
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
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(30).Style.NumberFormat.Format = "@";
                worksheet.Column(46).Style.NumberFormat.Format = "@";
                worksheet.Column(47).Style.NumberFormat.Format = "@";
                worksheet.Column(48).Style.NumberFormat.Format = "@";
                worksheet.Column(49).Style.NumberFormat.Format = "@";
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
                    worksheet.Cell(celda, 76).Value = data[i].COMENTARIOS;
                    worksheet.Cell(celda, 77).Value = data[i].ELIMINADOR;
                    worksheet.Cell(celda, 78).Value = data[i].TAPA;
                    worksheet.Cell(celda, 79).Value = data[i].CONECTIVIDAD_INSTALADA;
                    worksheet.Cell(celda, 80).Value = data[i].CONECTIVIDAD_RETIRADA;
                    worksheet.Cell(celda, 81).Value = data[i].APLICATIVO_INSTALADO;
                    worksheet.Cell(celda, 82).Value = data[i].APLICATIVO_RETIRADO;
                    worksheet.Cell(celda, 83).Value = data[i].CARRIER_INSTALADO;
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
    }
}
