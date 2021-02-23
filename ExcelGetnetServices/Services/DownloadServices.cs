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

                for (int i = 2; i < data.Count(); i++)
                {
                    worksheet.Cell(i, 1).Value = data[i].ODT;
                    worksheet.Cell(i, 2).Value = data[i].DISCOVER;
                    worksheet.Cell(i, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(i, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(i, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(i, 6).Value = data[i].COLONIA;
                    worksheet.Cell(i, 7).Value = data[i].POBLACION;
                    worksheet.Cell(i, 8).Value = data[i].ESTADO;
                    worksheet.Cell(i, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(i, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(i, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(i, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(i, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(i, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(i, 15).Value = data[i].NIVEL;
                    worksheet.Cell(i, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(i, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(i, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(i, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(i, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(i, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(i, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(i, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(i, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(i, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(i, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(i, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(i, 28).Value = data[i].MODELO;
                    worksheet.Cell(i, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(i, 30).Value = data[i].CAJA;
                    worksheet.Cell(i, 31).Value = data[i].RFC;
                    worksheet.Cell(i, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(i, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(i, 34).Value = data[i].TIEMPO_EN_ATENDER;
                    worksheet.Cell(i, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(i, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(i, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(i, 38).Value = data[i].CANAL;
                    worksheet.Cell(i, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(i, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(i, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(i, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(i, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(i, 44).Value = data[i].EMAIL;
                    worksheet.Cell(i, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(i, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(i, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(i, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(i, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 51).Value = data[i].VERSIONSW;
                    worksheet.Cell(i, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(i, 53).Value = data[i].BASE;
                    worksheet.Cell(i, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(i, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(i, 56).Value = data[i].ZONA;
                    worksheet.Cell(i, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(i, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(i, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(i, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(i, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(i, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(i, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(i, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(i, 65).Value = data[i].CODIGO;
                    worksheet.Cell(i, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(i, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(i, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(i, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(i, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(i, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(i, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(i, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(i, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(i, 75).Value = data[i].SERIE;
                    worksheet.Cell(i, 76).Value = data[i].COMENTARIOS;
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

                for (int i = 2; i < data.Count(); i++)
                {
                    worksheet.Cell(i, 1).Value = data[i].ODT;
                    worksheet.Cell(i, 2).Value = data[i].DISCOVER;
                    worksheet.Cell(i, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(i, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(i, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(i, 6).Value = data[i].COLONIA;
                    worksheet.Cell(i, 7).Value = data[i].POBLACION;
                    worksheet.Cell(i, 8).Value = data[i].ESTADO;
                    worksheet.Cell(i, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(i, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(i, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(i, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(i, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(i, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(i, 15).Value = data[i].NIVEL;
                    worksheet.Cell(i, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(i, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(i, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(i, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(i, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(i, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(i, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(i, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(i, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(i, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(i, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(i, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(i, 28).Value = data[i].MODELO;
                    worksheet.Cell(i, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(i, 30).Value = data[i].CAJA;
                    worksheet.Cell(i, 31).Value = data[i].RFC;
                    worksheet.Cell(i, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(i, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(i, 34).Value = data[i].TIEMPO_EN_ATENDER;
                    worksheet.Cell(i, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(i, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(i, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(i, 38).Value = data[i].CANAL;
                    worksheet.Cell(i, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(i, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(i, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(i, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(i, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(i, 44).Value = data[i].EMAIL;
                    worksheet.Cell(i, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(i, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(i, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(i, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(i, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 51).Value = data[i].VERSIONSW;
                    worksheet.Cell(i, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(i, 53).Value = data[i].BASE;
                    worksheet.Cell(i, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(i, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(i, 56).Value = data[i].ZONA;
                    worksheet.Cell(i, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(i, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(i, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(i, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(i, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(i, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(i, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(i, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(i, 65).Value = data[i].CODIGO;
                    worksheet.Cell(i, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(i, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(i, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(i, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(i, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(i, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(i, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(i, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(i, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(i, 75).Value = data[i].SERIE;
                    worksheet.Cell(i, 76).Value = data[i].COMENTARIOS;
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

                for (int i = 2; i < data.Count(); i++)
                {
                    worksheet.Cell(i, 1).Value = data[i].ODT;
                    worksheet.Cell(i, 2).Value = data[i].MI_COMERCIO;
                    worksheet.Cell(i, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(i, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(i, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(i, 6).Value = data[i].COLONIA;
                    worksheet.Cell(i, 7).Value = data[i].POBLACION;
                    worksheet.Cell(i, 8).Value = data[i].ESTADO;
                    worksheet.Cell(i, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(i, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(i, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(i, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(i, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(i, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(i, 15).Value = data[i].NIVEL;
                    worksheet.Cell(i, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(i, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(i, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(i, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(i, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(i, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(i, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(i, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(i, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(i, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(i, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(i, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(i, 28).Value = data[i].MODELO;
                    worksheet.Cell(i, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(i, 30).Value = data[i].CAJA;
                    worksheet.Cell(i, 31).Value = data[i].RFC;
                    worksheet.Cell(i, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(i, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(i, 34).Value = data[i].VESTIDURAS_GETNET;
                    worksheet.Cell(i, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(i, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(i, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(i, 38).Value = data[i].CANAL;
                    worksheet.Cell(i, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(i, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(i, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(i, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(i, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(i, 44).Value = data[i].EMAIL;
                    worksheet.Cell(i, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(i, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(i, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(i, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(i, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 51).Value = data[i].DIVISA;
                    worksheet.Cell(i, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(i, 53).Value = data[i].BASE;
                    worksheet.Cell(i, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(i, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(i, 56).Value = data[i].ZONA;
                    worksheet.Cell(i, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(i, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(i, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(i, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(i, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(i, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(i, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(i, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(i, 65).Value = data[i].CODIGO;
                    worksheet.Cell(i, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(i, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(i, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(i, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(i, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(i, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(i, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(i, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(i, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(i, 75).Value = data[i].AREA_CIERRA;
                    worksheet.Cell(i, 76).Value = data[i].ODT_SALESFORCE?.TrimEnd();
                    worksheet.Cell(i, 77).Value = data[i].COMENTARIOS;
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

                for (int i = 2; i < data.Count(); i++)
                {
                    worksheet.Cell(i, 1).Value = data[i].ODT;
                    worksheet.Cell(i, 2).Value = data[i].DISCOVER;
                    worksheet.Cell(i, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(i, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(i, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(i, 6).Value = data[i].COLONIA;
                    worksheet.Cell(i, 7).Value = data[i].POBLACION;
                    worksheet.Cell(i, 8).Value = data[i].ESTADO;
                    worksheet.Cell(i, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(i, 11).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(i, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(i, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(i, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(i, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(i, 15).Value = data[i].NIVEL;
                    worksheet.Cell(i, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(i, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(i, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(i, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(i, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(i, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(i, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(i, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(i, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(i, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(i, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(i, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(i, 28).Value = data[i].MODELO;
                    worksheet.Cell(i, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(i, 30).Value = data[i].CAJA;
                    worksheet.Cell(i, 31).Value = data[i].RFC;
                    worksheet.Cell(i, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(i, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(i, 34).Value = data[i].TIEMPO_EN_ATENDER;
                    worksheet.Cell(i, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(i, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(i, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(i, 38).Value = data[i].CANAL;
                    worksheet.Cell(i, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(i, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(i, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(i, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(i, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(i, 44).Value = data[i].EMAIL;
                    worksheet.Cell(i, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(i, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(i, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(i, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(i, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 51).Value = data[i].VERSIONSW;
                    worksheet.Cell(i, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(i, 53).Value = data[i].BASE;
                    worksheet.Cell(i, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(i, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(i, 56).Value = data[i].ZONA;
                    worksheet.Cell(i, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(i, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(i, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(i, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(i, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(i, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(i, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(i, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(i, 65).Value = data[i].CODIGO;
                    worksheet.Cell(i, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(i, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(i, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(i, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(i, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(i, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(i, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(i, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(i, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(i, 75).Value = data[i].SERIE;
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

                for (int i = 2; i < data.Count(); i++)
                {
                    worksheet.Cell(i, 1).Value = data[i].ODT;
                    worksheet.Cell(i, 2).Value = data[i].MI_COMERCIO;
                    worksheet.Cell(i, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(i, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(i, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(i, 6).Value = data[i].COLONIA;
                    worksheet.Cell(i, 7).Value = data[i].POBLACION;
                    worksheet.Cell(i, 8).Value = data[i].ESTADO;
                    worksheet.Cell(i, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(i, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(i, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(i, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(i, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(i, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(i, 15).Value = data[i].NIVEL;
                    worksheet.Cell(i, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(i, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(i, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(i, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(i, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(i, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(i, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(i, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(i, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(i, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(i, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(i, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(i, 28).Value = data[i].MODELO;
                    worksheet.Cell(i, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(i, 30).Value = data[i].CAJA;
                    worksheet.Cell(i, 31).Value = data[i].RFC;
                    worksheet.Cell(i, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(i, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(i, 34).Value = data[i].VESTIDURAS_GETNET;
                    worksheet.Cell(i, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(i, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(i, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(i, 38).Value = data[i].CANAL;
                    worksheet.Cell(i, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(i, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(i, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(i, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(i, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(i, 44).Value = data[i].EMAIL;
                    worksheet.Cell(i, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(i, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(i, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(i, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(i, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 51).Value = data[i].DIVISA;
                    worksheet.Cell(i, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(i, 53).Value = data[i].BASE;
                    worksheet.Cell(i, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(i, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(i, 56).Value = data[i].ZONA;
                    worksheet.Cell(i, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(i, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(i, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(i, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(i, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(i, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(i, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(i, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(i, 65).Value = data[i].CODIGO;
                    worksheet.Cell(i, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(i, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(i, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(i, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(i, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(i, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(i, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(i, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(i, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(i, 75).Value = data[i].AREA_CIERRA;
                    worksheet.Cell(i, 76).Value = data[i].ODT_SALESFORCE?.TrimEnd();
                    worksheet.Cell(i, 77).Value = data[i].COMENTARIOS;
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

                for (int i = 2; i < data.Count(); i++)
                {
                    worksheet.Cell(i, 1).Value = data[i].ODT;
                    worksheet.Cell(i, 2).Value = data[i].MI_COMERCIO;
                    worksheet.Cell(i, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(i, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(i, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(i, 6).Value = data[i].COLONIA;
                    worksheet.Cell(i, 7).Value = data[i].POBLACION;
                    worksheet.Cell(i, 8).Value = data[i].ESTADO;
                    worksheet.Cell(i, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(i, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(i, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(i, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(i, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(i, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(i, 15).Value = data[i].NIVEL;
                    worksheet.Cell(i, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(i, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(i, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(i, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(i, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(i, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(i, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(i, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(i, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(i, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(i, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(i, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(i, 28).Value = data[i].MODELO;
                    worksheet.Cell(i, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(i, 30).Value = data[i].CAJA;
                    worksheet.Cell(i, 31).Value = data[i].RFC;
                    worksheet.Cell(i, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(i, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(i, 34).Value = data[i].VESTIDURAS_GETNET;
                    worksheet.Cell(i, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(i, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(i, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(i, 38).Value = data[i].CANAL;
                    worksheet.Cell(i, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(i, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(i, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(i, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(i, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(i, 44).Value = data[i].EMAIL;
                    worksheet.Cell(i, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(i, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(i, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(i, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(i, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 51).Value = data[i].DIVISA;
                    worksheet.Cell(i, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(i, 53).Value = data[i].BASE;
                    worksheet.Cell(i, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(i, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(i, 56).Value = data[i].ZONA;
                    worksheet.Cell(i, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(i, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(i, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(i, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(i, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(i, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(i, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(i, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(i, 65).Value = data[i].CODIGO;
                    worksheet.Cell(i, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(i, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(i, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(i, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(i, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(i, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(i, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(i, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(i, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(i, 75).Value = data[i].AREA_CIERRA;
                    worksheet.Cell(i, 76).Value = data[i].ODT_SALESFORCE?.TrimEnd();
                    worksheet.Cell(i, 77).Value = data[i].COMENTARIOS;
                    worksheet.Cell(i, 78).Value = data[i].DESC_GIRO;
                    worksheet.Cell(i, 79).Value = "";
                    worksheet.Cell(i, 80).Value = "";
                    worksheet.Cell(i, 81).Value = data[i].MCC;
                    worksheet.Cell(i, 82).Value = data[i].INDUSTRIA;
                    worksheet.Cell(i, 83).Value = data[i].MSI_SANTANDER;
                    worksheet.Cell(i, 84).Value = data[i].MSI_PROSA;
                    worksheet.Cell(i, 85).Value = data[i].AMEX_PLAN;
                    worksheet.Cell(i, 86).Value = data[i].QPS;
                    worksheet.Cell(i, 87).Value = data[i].CARNET;
                    worksheet.Cell(i, 88).Value = data[i].APLICATIVO;
                    worksheet.Cell(i, 89).Value = data[i].TIPO_DE_CONFIGURACION_MIT;

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

                for (int i = 2; i < data.Count(); i++)
                {
                    worksheet.Cell(i, 1).Value = data[i].ODT;
                    worksheet.Cell(i, 2).Value = data[i].DISCOVER;
                    worksheet.Cell(i, 3).Value = data[i].AFILIACION;
                    worksheet.Cell(i, 4).Value = data[i].COMERCIO;
                    worksheet.Cell(i, 5).Value = data[i].DIRECCION;
                    worksheet.Cell(i, 6).Value = data[i].COLONIA;
                    worksheet.Cell(i, 7).Value = data[i].POBLACION;
                    worksheet.Cell(i, 8).Value = data[i].ESTADO;
                    worksheet.Cell(i, 9).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 9).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA);
                    worksheet.Cell(i, 10).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 10).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_VENCIMIENTO);
                    worksheet.Cell(i, 11).Value = String.IsNullOrEmpty(data[i].DESCRIPCION) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].DESCRIPCION)));
                    worksheet.Cell(i, 12).Value = String.IsNullOrEmpty(data[i].OBSERVACIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].OBSERVACIONES)));
                    worksheet.Cell(i, 13).Value = data[i].TELEFONO;
                    worksheet.Cell(i, 14).Value = data[i].TIPO_COMERCIO_2;
                    worksheet.Cell(i, 15).Value = data[i].NIVEL;
                    worksheet.Cell(i, 16).Value = data[i].TIPO_SERVICIO;
                    worksheet.Cell(i, 17).Value = data[i].SUB_TIPO_SERVICIO;
                    worksheet.Cell(i, 18).Value = data[i].CRITERIO_CAMBIO;
                    worksheet.Cell(i, 19).Value = data[i].ID_TECNICO;
                    worksheet.Cell(i, 20).Value = data[i].PROVEEDOR;
                    worksheet.Cell(i, 21).Value = data[i].ESTATUS_SERVICIO;
                    worksheet.Cell(i, 22).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 22).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ATENCION_PROVEEDOR);
                    worksheet.Cell(i, 23).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 23).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_CIERRE_SISTEMA);
                    worksheet.Cell(i, 24).Style.NumberFormat.Format = "@";
                    worksheet.Cell(i, 24).Value = String.Format("{0:dd/MM/yyyy HH:mm:ss}", data[i].FECHA_ALTA_SISTEMA);
                    worksheet.Cell(i, 25).Value = data[i].CODIGO_POSTAL;
                    worksheet.Cell(i, 26).Value = String.IsNullOrEmpty(data[i].CONCLUSIONES) ? "" : Encoding.ASCII.GetString(Encoding
                                                           .Convert(Encoding.GetEncoding("ISO-8859-8"), Encoding.GetEncoding(Encoding.ASCII.EncodingName,
                                                           new EncoderReplacementFallback(string.Empty), new DecoderExceptionFallback()),
                                                           Encoding.GetEncoding("ISO-8859-8").GetBytes(data[i].CONCLUSIONES)));
                    worksheet.Cell(i, 27).Value = data[i].CONECTIVIDAD;
                    worksheet.Cell(i, 28).Value = data[i].MODELO;
                    worksheet.Cell(i, 29).Value = data[i].EQUIPO;
                    worksheet.Cell(i, 30).Value = data[i].CAJA;
                    worksheet.Cell(i, 31).Value = data[i].RFC;
                    worksheet.Cell(i, 32).Value = data[i].RAZON_SOCIAL;
                    worksheet.Cell(i, 33).Value = data[i].HORAS_VENCIDAS;
                    worksheet.Cell(i, 34).Value = data[i].VESTIDURAS_GETNET ;
                    worksheet.Cell(i, 35).Value = data[i].SLA_FIJO;
                    worksheet.Cell(i, 36).Value = data[i].CODIGO_AFILIACION;
                    worksheet.Cell(i, 37).Value = data[i].TELEFONOS_EN_CAMPO;
                    worksheet.Cell(i, 38).Value = data[i].CANAL;
                    worksheet.Cell(i, 39).Value = data[i].AFILIACION_AMEX;
                    worksheet.Cell(i, 40).Value = data[i].IDAMEX;
                    worksheet.Cell(i, 41).Value = data[i].PRODUCTO;
                    worksheet.Cell(i, 42).Value = data[i].MOTIVO_CANCELACION;
                    worksheet.Cell(i, 43).Value = data[i].MOTIVO_RECHAZO;
                    worksheet.Cell(i, 44).Value = data[i].EMAIL;
                    worksheet.Cell(i, 45).Value = data[i].ROLLOS_A_INSTALAR;
                    worksheet.Cell(i, 46).Value = data[i].NUM_SERIE_TERMINAL_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 47).Value = data[i].NUM_SERIE_TERMINAL_SALE?.TrimEnd();
                    worksheet.Cell(i, 48).Value = data[i].NUM_SERIE_TERMINAL_MTO?.TrimEnd();
                    worksheet.Cell(i, 49).Value = data[i].NUM_SERIE_SIM_SALE?.TrimEnd();
                    worksheet.Cell(i, 50).Value = data[i].NUM_SERIE_SIM_ENTRA?.TrimEnd();
                    worksheet.Cell(i, 51).Value = data[i].VERSIONSW;
                    worksheet.Cell(i, 52).Value = data[i].CARGADOR;
                    worksheet.Cell(i, 53).Value = data[i].BASE;
                    worksheet.Cell(i, 54).Value = data[i].ROLLO_ENTREGADOS;
                    worksheet.Cell(i, 55).Value = data[i].CABLE_CORRIENTE;
                    worksheet.Cell(i, 56).Value = data[i].ZONA;
                    worksheet.Cell(i, 57).Value = data[i].MODELO_INSTALADO;
                    worksheet.Cell(i, 58).Value = data[i].MODELO_TERMINAL_SALE;
                    worksheet.Cell(i, 59).Value = data[i].CORREO_EJECUTIVO;
                    worksheet.Cell(i, 60).Value = data[i].RECHAZO;
                    worksheet.Cell(i, 61).Value = data[i].CONTACTO1;
                    worksheet.Cell(i, 62).Value = data[i].ATIENDE_EN_COMERCIO;
                    worksheet.Cell(i, 63).Value = data[i].TID_AMEX_CIERRE;
                    worksheet.Cell(i, 64).Value = data[i].AFILIACION_AMEX_CIERRE;
                    worksheet.Cell(i, 65).Value = data[i].CODIGO;
                    worksheet.Cell(i, 66).Value = data[i].TIENE_AMEX;
                    worksheet.Cell(i, 67).Value = data[i].ACT_REFERENCIAS;
                    worksheet.Cell(i, 68).Value = data[i].TIPO_A_B;
                    worksheet.Cell(i, 69).Value = data[i].DIRECCION_ALTERNA_COMERCIO;
                    worksheet.Cell(i, 70).Value = data[i].CANTIDAD_ARCHIVOS;
                    worksheet.Cell(i, 71).Value = data[i].AREA_CARGA;
                    worksheet.Cell(i, 72).Value = data[i].ALTA_POR;
                    worksheet.Cell(i, 73).Value = data[i].TIPO_CARGA;
                    worksheet.Cell(i, 74).Value = data[i].CERRADO_POR;
                    worksheet.Cell(i, 75).Value = data[i].SERIE;
                    worksheet.Cell(i, 76).Value = data[i].COMENTARIOS;
                    worksheet.Cell(i, 77).Value = data[i].ELIMINADOR;
                    worksheet.Cell(i, 78).Value = data[i].TAPA;
                    worksheet.Cell(i, 79).Value = data[i].CONECTIVIDAD_INSTALADA;
                    worksheet.Cell(i, 80).Value = data[i].CONECTIVIDAD_RETIRADA;
                    worksheet.Cell(i, 81).Value = data[i].APLICATIVO_INSTALADO;
                    worksheet.Cell(i, 82).Value = data[i].APLICATIVO_RETIRADO;
                    worksheet.Cell(i, 83).Value = data[i].CARRIER_INSTALADO;
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
