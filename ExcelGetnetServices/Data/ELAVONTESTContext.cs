using System;
using ExcelGetnetServices.Entities.StoredProcedures;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata;

#nullable disable

namespace ExcelGetnetServices.Data
{
    public partial class ELAVONTESTContext : DbContext
    {
        public ELAVONTESTContext()
        {
            Database.SetCommandTimeout(150000);
        }

        public ELAVONTESTContext(DbContextOptions<ELAVONTESTContext> options)
            : base(options)
        {
        }

        public virtual DbSet<BdAr> BdArs { get; set; }
        public virtual DbSet<BdColumnsLayout> BdColumnsLayouts { get; set; }
        public virtual DbSet<CLayout> CLayouts { get; set; }
        public virtual DbSet<SpLayoutMasivo3> SpLayoutMasivo3s { get; set; }
        public virtual DbSet<SpLayoutMasivo> SpLayoutMasivos { get; set; }
        public virtual DbSet<SpLayoutMasivo2> SpLayoutMasivo2s { get; set; }
        public virtual DbSet<SpLayoutMasivo4> SpLayoutMasivo4s { get; set; }
        public virtual DbSet<SpLayoutMasivoUsuario> SpLayoutMasivoUsuarios { get; set; }
        public virtual DbSet<SpLayoutMasivoGetnetMit> SpLayoutMasivoGetnetMits { get; set; }
        public virtual DbSet<SpLayoutMasivoReingenieria> SpLayoutMasivoReingenierias { get; set; }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.HasAnnotation("Relational:Collation", "Modern_Spanish_CI_AS");

            modelBuilder.Entity<SpLayoutMasivo3>(entity => {
                entity.HasNoKey();
                entity.Property(e => e.DIAS_SLA_ADMIN)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_ADMIN");
                entity.Property(e => e.DIAS_SLA_GLOBAL)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_GLOBAL");
            });

            modelBuilder.Entity<SpLayoutMasivo>(entity => {
                entity.HasNoKey();
                entity.Property(e => e.DIAS_SLA_ADMIN)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_ADMIN");
                entity.Property(e => e.DIAS_SLA_GLOBAL)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_GLOBAL");
            });

            modelBuilder.Entity<SpLayoutMasivo2>(entity => {
                entity.HasNoKey();
                entity.Property(e => e.DIAS_SLA_ADMIN)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_ADMIN");
                entity.Property(e => e.DIAS_SLA_GLOBAL)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_GLOBAL");
            });

            modelBuilder.Entity<SpLayoutMasivo4>(entity => {
                entity.HasNoKey();
                entity.Property(e => e.DIAS_SLA_ADMIN)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_ADMIN");
                entity.Property(e => e.DIAS_SLA_GLOBAL)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_GLOBAL");
            });

            modelBuilder.Entity<SpLayoutMasivoUsuario>(entity =>
            {
                entity.HasNoKey();
                entity.Property(e => e.DIAS_SLA_ADMIN)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_ADMIN");
                entity.Property(e => e.DIAS_SLA_GLOBAL)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_GLOBAL");
            });

            modelBuilder.Entity<SpLayoutMasivoGetnetMit>(entity =>
            {
                entity.HasNoKey();
                entity.Property(e => e.DIAS_SLA_ADMIN)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_ADMIN");
                entity.Property(e => e.DIAS_SLA_GLOBAL)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_GLOBAL");
            modelBuilder.Entity<SpLayoutMasivoGetnetMit>(entity =>
            {
                entity.HasNoKey();
                entity.Property(e => e.DIAS_SLA_ADMIN)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_ADMIN");
                entity.Property(e => e.DIAS_SLA_GLOBAL)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_GLOBAL");
            });
            });
            
            modelBuilder.Entity<SpLayoutMasivoReingenieria>(entity => {
                entity.HasNoKey();
                entity.Property(e => e.DIAS_SLA_ADMIN)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_ADMIN");
                entity.Property(e => e.DIAS_SLA_GLOBAL)
                    .HasColumnType("decimal(18,2)")
                    .HasColumnName("DIAS_SLA_GLOBAL");
            });

            modelBuilder.Entity<BdAr>(entity =>
            {
                entity.HasKey(e => e.IdAr);

                entity.ToTable("BD_AR");

                entity.HasIndex(e => new { e.IdStatusAr, e.Status }, "missing_index_1096_1095");

                entity.HasIndex(e => new { e.NoAfiliacion, e.Status, e.IdProveedor, e.IdStatusAr }, "missing_index_1137_1136");

                entity.HasIndex(e => new { e.NoAr, e.Status, e.IdProveedor, e.IdStatusAr }, "missing_index_1139_1138");

                entity.HasIndex(e => new { e.Status, e.IdProveedor, e.IdStatusAr }, "missing_index_1145_1144");

                entity.HasIndex(e => new { e.IdStatusAr, e.Status, e.IdProveedor }, "missing_index_1157_1156");

                entity.HasIndex(e => new { e.Status, e.IdStatusAr }, "missing_index_118_117");

                entity.HasIndex(e => new { e.Status, e.IdStatusAr }, "missing_index_1224_1223");

                entity.HasIndex(e => new { e.IdCarga, e.NoAr, e.IdStatusAr }, "missing_index_122_121");

                entity.HasIndex(e => new { e.IdCarga, e.Status }, "missing_index_157_156");

                entity.HasIndex(e => new { e.IdCliente, e.NoAr, e.IdAr }, "missing_index_159_158");

                entity.HasIndex(e => e.NoAr, "missing_index_161_160");

                entity.HasIndex(e => e.IdCarga, "missing_index_163_162");

                entity.HasIndex(e => new { e.IdNegocio, e.IdProyecto, e.FecCierre, e.IdStatusAr }, "missing_index_165_164");

                entity.HasIndex(e => e.NoAr, "missing_index_167_166");

                entity.HasIndex(e => new { e.IdCarga, e.Status }, "missing_index_172_171");

                entity.HasIndex(e => e.NoAr, "missing_index_174_173");

                entity.HasIndex(e => e.NoAr, "missing_index_1760_1759");

                entity.HasIndex(e => e.NoAr, "missing_index_1762_1761");

                entity.HasIndex(e => e.Status, "missing_index_176_175");

                entity.HasIndex(e => new { e.IdCliente, e.NoAr, e.IdCarga }, "missing_index_178_177");

                entity.HasIndex(e => e.IdCarga, "missing_index_180_179");

                entity.HasIndex(e => new { e.IdCarga, e.Status }, "missing_index_184_183");

                entity.HasIndex(e => e.IdCarga, "missing_index_187_186");

                entity.HasIndex(e => new { e.IdCarga, e.Status }, "missing_index_189_188");

                entity.HasIndex(e => new { e.IdCarga, e.Status }, "missing_index_192_191");

                entity.HasIndex(e => new { e.IdCarga, e.Status }, "missing_index_198_197");

                entity.HasIndex(e => new { e.Status, e.IdStatusAr }, "missing_index_248_247");

                entity.HasIndex(e => new { e.NoAfiliacion, e.Status, e.IdStatusAr }, "missing_index_250_249");

                entity.HasIndex(e => new { e.NoAr, e.Status, e.IdStatusAr }, "missing_index_252_251");

                entity.HasIndex(e => new { e.IdCarga, e.Status }, "missing_index_2696_2695");

                entity.HasIndex(e => e.IdCarga, "missing_index_2708_2707");

                entity.HasIndex(e => new { e.IdCarga, e.Status }, "missing_index_2710_2709");

                entity.HasIndex(e => new { e.Status, e.IdProveedor, e.IdStatusAr }, "missing_index_2743_2742");

                entity.HasIndex(e => e.IdStatusAr, "missing_index_2815_2814");

                entity.HasIndex(e => new { e.NoAr, e.IdStatusAr }, "missing_index_2860_2859");

                entity.HasIndex(e => new { e.Status, e.IdStatusAr }, "missing_index_2932_2931");

                entity.HasIndex(e => new { e.Status, e.NoAr }, "missing_index_2_1");

                entity.HasIndex(e => e.NoAr, "missing_index_3318_3317");

                entity.HasIndex(e => new { e.IdStatusAr, e.Status }, "missing_index_3476_3475");

                entity.HasIndex(e => new { e.Status, e.IdStatusAr }, "missing_index_3478_3477");

                entity.HasIndex(e => new { e.Status, e.IdProveedor, e.IdStatusAr }, "missing_index_3487_3486");

                entity.HasIndex(e => e.NoAfiliacion, "missing_index_3491_3490");

                entity.HasIndex(e => e.IdTecnico, "missing_index_3618_3617");

                entity.HasIndex(e => new { e.IdStatusAr, e.Status }, "missing_index_3806_3805");

                entity.HasIndex(e => new { e.Status, e.IdProveedor, e.IdStatusAr }, "missing_index_4007_4006");

                entity.HasIndex(e => new { e.IdRegion, e.Status, e.IdStatusAr }, "missing_index_4014_4013");

                entity.HasIndex(e => new { e.Status, e.IdStatusAr }, "missing_index_4016_4015");

                entity.HasIndex(e => e.IdRegion, "missing_index_4018_4017");

                entity.HasIndex(e => new { e.Status, e.IdProveedor, e.IdStatusAr }, "missing_index_4098_4097");

                entity.HasIndex(e => new { e.IdStatusAr, e.Status, e.IdProveedor }, "missing_index_4101_4100");

                entity.HasIndex(e => e.IdProyecto, "missing_index_4130_4129");

                entity.HasIndex(e => new { e.IdProyecto, e.IdServicio, e.IdFalla }, "missing_index_4132_4131");

                entity.HasIndex(e => e.IdProyecto, "missing_index_4135_4134");

                entity.HasIndex(e => new { e.IdFalla, e.IdStatusAr, e.NoAfiliacion }, "missing_index_46_45");

                entity.HasIndex(e => new { e.IdFalla, e.IdStatusAr }, "missing_index_49_48");

                entity.HasIndex(e => new { e.NoAr, e.Status }, "missing_index_52_51");

                entity.HasIndex(e => e.IdZona, "missing_index_61_60");

                entity.HasIndex(e => new { e.Status, e.IdProveedor, e.IdStatusAr }, "missing_index_63_62");

                entity.HasIndex(e => e.Status, "missing_index_6_5");

                entity.HasIndex(e => e.Status, "missing_index_903_902");

                entity.Property(e => e.IdAr).HasColumnName("ID_AR");

                entity.Property(e => e.AgregarDiasAtencion).HasColumnName("AGREGAR_DIAS_ATENCION");

                entity.Property(e => e.Atiende)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("ATIENDE");

                entity.Property(e => e.AutorizadorRechazo)
                    .HasMaxLength(150)
                    .IsUnicode(false)
                    .HasColumnName("AUTORIZADOR_RECHAZO");

                entity.Property(e => e.Bitacora)
                    .HasMaxLength(4000)
                    .IsUnicode(false)
                    .HasColumnName("BITACORA");

                entity.Property(e => e.CadenaCierre)
                    .IsUnicode(false)
                    .HasColumnName("CADENA_CIERRE");

                entity.Property(e => e.CadenaCierreEscrita)
                    .HasMaxLength(3000)
                    .IsUnicode(false)
                    .HasColumnName("CADENA_CIERRE_ESCRITA");

                entity.Property(e => e.Caja)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("CAJA");

                entity.Property(e => e.CausaCancelacion)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("CAUSA_CANCELACION");

                entity.Property(e => e.CausaRechazo)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("CAUSA_RECHAZO");

                entity.Property(e => e.ClaveRechazo)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("CLAVE_RECHAZO");

                entity.Property(e => e.CodigoIntervencion)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("CODIGO_INTERVENCION");

                entity.Property(e => e.Colonia)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("COLONIA");

                entity.Property(e => e.Concepto)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("CONCEPTO");

                entity.Property(e => e.CorreoEjecutivo)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("CORREO_EJECUTIVO");

                entity.Property(e => e.Cp)
                    .HasMaxLength(5)
                    .IsUnicode(false)
                    .HasColumnName("CP");

                entity.Property(e => e.DescCorta)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("DESC_CORTA");

                entity.Property(e => e.DescEquipo)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("DESC_EQUIPO");

                entity.Property(e => e.DescNegocio)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("DESC_NEGOCIO");

                entity.Property(e => e.DescripcionTrabajo)
                    .HasMaxLength(3000)
                    .IsUnicode(false)
                    .HasColumnName("DESCRIPCION_TRABAJO");

                entity.Property(e => e.DigitoVerificador)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("DIGITO_VERIFICADOR");

                entity.Property(e => e.Direccion)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("DIRECCION");

                entity.Property(e => e.DireccionAlternaComercio)
                    .IsUnicode(false)
                    .HasColumnName("DIRECCION_ALTERNA_COMERCIO");

                entity.Property(e => e.DueBy)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("due_by");

                entity.Property(e => e.Duracion).HasColumnName("DURACION");

                entity.Property(e => e.Equipo)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("EQUIPO");

                entity.Property(e => e.Estado)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("ESTADO");

                entity.Property(e => e.FallaEncontrada)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("FALLA_ENCONTRADA");

                entity.Property(e => e.FecAlta)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_ALTA");

                entity.Property(e => e.FecAltaHorasAtencion)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_ALTA_HORAS_ATENCION");

                entity.Property(e => e.FecAltaReglaStatusAr)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_ALTA_REGLA_STATUS_AR");

                entity.Property(e => e.FecAtencion)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_ATENCION");

                entity.Property(e => e.FecAtencionOriginal)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_ATENCION_ORIGINAL");

                entity.Property(e => e.FecCarga)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_CARGA");

                entity.Property(e => e.FecCierre)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_CIERRE");

                entity.Property(e => e.FecConvenio)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_CONVENIO");

                entity.Property(e => e.FecFinIngeniero)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_FIN_INGENIERO");

                entity.Property(e => e.FecGarantia)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_GARANTIA");

                entity.Property(e => e.FecIniIngeniero)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_INI_INGENIERO");

                entity.Property(e => e.FecInicio)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_INICIO");

                entity.Property(e => e.FecIntento1)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_INTENTO_1");

                entity.Property(e => e.FecIntento2)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_INTENTO_2");

                entity.Property(e => e.FecIntento3)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_INTENTO_3");

                entity.Property(e => e.FecIntento4)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_INTENTO_4");

                entity.Property(e => e.FecLlegada)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_LLEGADA");

                entity.Property(e => e.FecLlegadaTerceros)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_LLEGADA_TERCEROS");

                entity.Property(e => e.FecStatusAr)
                    .HasColumnType("smalldatetime")
                    .HasColumnName("FEC_STATUS_AR");

                entity.Property(e => e.FolioServicio)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("FOLIO_SERVICIO");

                entity.Property(e => e.FolioTas)
                    .HasMaxLength(20)
                    .IsUnicode(false)
                    .HasColumnName("FOLIO_TAS");

                entity.Property(e => e.FolioTelecarga)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("FOLIO_TELECARGA");

                entity.Property(e => e.FolioTir)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("FOLIO_TIR");

                entity.Property(e => e.FolioValidacion)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("FOLIO_VALIDACION");

                entity.Property(e => e.HoraAtencionFin).HasColumnName("HORA_ATENCION_FIN");

                entity.Property(e => e.HoraAtencionIni).HasColumnName("HORA_ATENCION_INI");

                entity.Property(e => e.HorasAtencion).HasColumnName("HORAS_ATENCION");

                entity.Property(e => e.HorasAtencionWincor)
                    .HasColumnType("numeric(18, 0)")
                    .HasColumnName("HORAS_ATENCION_WINCOR");

                entity.Property(e => e.HorasGarantia).HasColumnName("HORAS_GARANTIA");

                entity.Property(e => e.HorasGarantiaWincor)
                    .HasColumnType("numeric(18, 0)")
                    .HasColumnName("HORAS_GARANTIA_WINCOR");

                entity.Property(e => e.IdAplicativo).HasColumnName("ID_APLICATIVO");

                entity.Property(e => e.IdArOriginal).HasColumnName("ID_AR_ORIGINAL");

                entity.Property(e => e.IdAttach1).HasColumnName("ID_ATTACH1");

                entity.Property(e => e.IdAttach2).HasColumnName("ID_ATTACH2");

                entity.Property(e => e.IdCalificaContacto).HasColumnName("ID_CALIFICA_CONTACTO");

                entity.Property(e => e.IdCalificaIntento1).HasColumnName("ID_CALIFICA_INTENTO_1");

                entity.Property(e => e.IdCalificaIntento2).HasColumnName("ID_CALIFICA_INTENTO_2");

                entity.Property(e => e.IdCalificaIntento3).HasColumnName("ID_CALIFICA_INTENTO_3");

                entity.Property(e => e.IdCalificaIntento4).HasColumnName("ID_CALIFICA_INTENTO_4");

                entity.Property(e => e.IdCarga).HasColumnName("ID_CARGA");

                entity.Property(e => e.IdCausa).HasColumnName("ID_CAUSA");

                entity.Property(e => e.IdCausaRechazo).HasColumnName("ID_CAUSA_RECHAZO");

                entity.Property(e => e.IdCliente).HasColumnName("ID_CLIENTE");

                entity.Property(e => e.IdConcepto).HasColumnName("ID_CONCEPTO");

                entity.Property(e => e.IdConectividad).HasColumnName("ID_CONECTIVIDAD");

                entity.Property(e => e.IdDescripcionTrabajo).HasColumnName("ID_DESCRIPCION_TRABAJO");

                entity.Property(e => e.IdDispatcher).HasColumnName("ID_DISPATCHER");

                entity.Property(e => e.IdEquipoCliente).HasColumnName("ID_EQUIPO_CLIENTE");

                entity.Property(e => e.IdEspecifTipoFalla).HasColumnName("ID_ESPECIF_TIPO_FALLA");

                entity.Property(e => e.IdEspecificaCausaRechazo).HasColumnName("ID_ESPECIFICA_CAUSA_RECHAZO");

                entity.Property(e => e.IdEstado).HasColumnName("ID_ESTADO");

                entity.Property(e => e.IdFalla).HasColumnName("ID_FALLA");

                entity.Property(e => e.IdModeloFalla).HasColumnName("ID_MODELO_FALLA");

                entity.Property(e => e.IdModeloReq).HasColumnName("ID_MODELO_REQ");

                entity.Property(e => e.IdMoneda).HasColumnName("ID_MONEDA");

                entity.Property(e => e.IdNegocio).HasColumnName("ID_NEGOCIO");

                entity.Property(e => e.IdPlaza).HasColumnName("ID_PLAZA");

                entity.Property(e => e.IdProducto).HasColumnName("ID_PRODUCTO");

                entity.Property(e => e.IdProveedor).HasColumnName("ID_PROVEEDOR");

                entity.Property(e => e.IdProyecto).HasColumnName("ID_PROYECTO");

                entity.Property(e => e.IdRegion).HasColumnName("ID_REGION");

                entity.Property(e => e.IdReglaStatusAr).HasColumnName("ID_REGLA_STATUS_AR");

                entity.Property(e => e.IdReporteCierre).HasColumnName("ID_REPORTE_CIERRE");

                entity.Property(e => e.IdResponsableCancelacionProgramado).HasColumnName("ID_RESPONSABLE_CANCELACION_PROGRAMADO");

                entity.Property(e => e.IdSegmento).HasColumnName("ID_SEGMENTO");

                entity.Property(e => e.IdServicio).HasColumnName("ID_SERVICIO");

                entity.Property(e => e.IdSolucion).HasColumnName("ID_SOLUCION");

                entity.Property(e => e.IdStatusAr).HasColumnName("ID_STATUS_AR");

                entity.Property(e => e.IdStatusReasonCodes).HasColumnName("ID_STATUS_REASON_CODES");

                entity.Property(e => e.IdStatusValidacionPrefacturacion).HasColumnName("ID_STATUS_VALIDACION_PREFACTURACION");

                entity.Property(e => e.IdTecnico).HasColumnName("ID_TECNICO");

                entity.Property(e => e.IdTipoCobro).HasColumnName("ID_TIPO_COBRO");

                entity.Property(e => e.IdTipoEquipo).HasColumnName("ID_TIPO_EQUIPO");

                entity.Property(e => e.IdTipoFallaEncontrada).HasColumnName("ID_TIPO_FALLA_ENCONTRADA");

                entity.Property(e => e.IdTipoPlaza).HasColumnName("ID_TIPO_PLAZA");

                entity.Property(e => e.IdTipoPrecio).HasColumnName("ID_TIPO_PRECIO");

                entity.Property(e => e.IdTipoServicio).HasColumnName("ID_TIPO_SERVICIO");

                entity.Property(e => e.IdUnidadAtendida).HasColumnName("ID_UNIDAD_ATENDIDA");

                entity.Property(e => e.IdUsuarioCierre).HasColumnName("ID_USUARIO_CIERRE");

                entity.Property(e => e.IdZona).HasColumnName("ID_ZONA");

                entity.Property(e => e.Insumos).HasColumnName("INSUMOS");

                entity.Property(e => e.IntensidadSenial)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("INTENSIDAD_SENIAL");

                entity.Property(e => e.IntentoContacto).HasColumnName("INTENTO_CONTACTO");

                entity.Property(e => e.IsActualizacion).HasColumnName("IS_ACTUALIZACION");

                entity.Property(e => e.IsBoletin).HasColumnName("IS_BOLETIN");

                entity.Property(e => e.IsCobrable).HasColumnName("IS_COBRABLE");

                entity.Property(e => e.IsDuplicada).HasColumnName("IS_DUPLICADA");

                entity.Property(e => e.IsExito).HasColumnName("IS_EXITO");

                entity.Property(e => e.IsFollowDispatch).HasColumnName("IS_FOLLOW_DISPATCH");

                entity.Property(e => e.IsGarantia).HasColumnName("IS_GARANTIA");

                entity.Property(e => e.IsIngresoManual).HasColumnName("IS_INGRESO_MANUAL");

                entity.Property(e => e.IsInstalacion).HasColumnName("IS_INSTALACION");

                entity.Property(e => e.IsInterfazBancomer).HasColumnName("IS_INTERFAZ_BANCOMER");

                entity.Property(e => e.IsLocal).HasColumnName("IS_LOCAL");

                entity.Property(e => e.IsPdf).HasColumnName("IS_PDF");

                entity.Property(e => e.IsProgramado).HasColumnName("IS_PROGRAMADO");

                entity.Property(e => e.IsRetipificado).HasColumnName("IS_RETIPIFICADO");

                entity.Property(e => e.IsRetiro).HasColumnName("IS_RETIRO");

                entity.Property(e => e.IsSimRemplazada)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("IS_SIM_REMPLAZADA");

                entity.Property(e => e.IsSoporteCliente).HasColumnName("IS_SOPORTE_CLIENTE");

                entity.Property(e => e.IsSustitucion).HasColumnName("IS_SUSTITUCION");

                entity.Property(e => e.IsTecnicoForaneo).HasColumnName("IS_TECNICO_FORANEO");

                entity.Property(e => e.MiComercio)
                    .HasMaxLength(100)
                    .HasColumnName("MI_COMERCIO")
                    .IsFixedLength(true);

                entity.Property(e => e.MinsDowntime)
                    .HasColumnName("MINS_DOWNTIME")
                    .HasDefaultValueSql("((0))");

                entity.Property(e => e.MotivoCobro)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("MOTIVO_COBRO");

                entity.Property(e => e.MotivoRetipificado)
                    .HasMaxLength(100)
                    .IsUnicode(false)
                    .HasColumnName("MOTIVO_RETIPIFICADO");

                entity.Property(e => e.NoAfiliacion)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NO_AFILIACION");

                entity.Property(e => e.NoAr)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NO_AR");

                entity.Property(e => e.NoDiasLiberacion).HasColumnName("NO_DIAS_LIBERACION");

                entity.Property(e => e.NoEquipo)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NO_EQUIPO");

                entity.Property(e => e.NoInventario)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NO_INVENTARIO");

                entity.Property(e => e.NoInventarioFalla)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NO_INVENTARIO_FALLA");

                entity.Property(e => e.NoReincidencias).HasColumnName("NO_REINCIDENCIAS");

                entity.Property(e => e.NoSerie)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NO_SERIE");

                entity.Property(e => e.NoSerieFalla)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("NO_SERIE_FALLA");

                entity.Property(e => e.NoSim)
                    .HasMaxLength(300)
                    .IsUnicode(false)
                    .HasColumnName("NO_SIM");

                entity.Property(e => e.NotViaticos).HasColumnName("NOT_VIATICOS");

                entity.Property(e => e.NotasRemedy)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("NOTAS_REMEDY");

                entity.Property(e => e.OtorganteSoporteCliente)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("OTORGANTE_SOPORTE_CLIENTE");

                entity.Property(e => e.OtorganteTas)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("OTORGANTE_TAS");

                entity.Property(e => e.OtorganteVobo)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("OTORGANTE_VOBO");

                entity.Property(e => e.OtorganteVoboCliente)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("OTORGANTE_VOBO_CLIENTE");

                entity.Property(e => e.OtorganteVoboTerceros)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("OTORGANTE_VOBO_TERCEROS");

                entity.Property(e => e.PersonaAtenderaComercio)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("PERSONA_ATENDERA_COMERCIO");

                entity.Property(e => e.Poblacion)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("POBLACION");

                entity.Property(e => e.Precio)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("PRECIO");

                entity.Property(e => e.PrecioExito)
                    .HasColumnType("numeric(18, 2)")
                    .HasColumnName("PRECIO_EXITO");

                entity.Property(e => e.ProveedorAtenderaComercio)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("PROVEEDOR_ATENDERA_COMERCIO");

                entity.Property(e => e.Responsable)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("RESPONSABLE");

                entity.Property(e => e.Rp).HasColumnName("RP");

                entity.Property(e => e.Rs).HasColumnName("RS");

                entity.Property(e => e.Segmento).HasColumnName("SEGMENTO");

                entity.Property(e => e.Sintoma)
                    .HasMaxLength(4000)
                    .IsUnicode(false)
                    .HasColumnName("SINTOMA");

                entity.Property(e => e.Status)
                    .HasMaxLength(10)
                    .IsUnicode(false)
                    .HasColumnName("STATUS");

                entity.Property(e => e.Telefono)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("TELEFONO");

                entity.Property(e => e.TelefonoComercio)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("TELEFONO_COMERCIO");

                entity.Property(e => e.TerminalAmex).HasColumnName("TERMINAL_AMEX");

                entity.Property(e => e.TipoFalla).HasColumnName("TIPO_FALLA");

                entity.Property(e => e.TipoServicio).HasColumnName("TIPO_SERVICIO");

                entity.Property(e => e.Traslado).HasColumnName("TRASLADO");

                entity.Property(e => e.VoltajeNeutro)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("VOLTAJE_NEUTRO");

                entity.Property(e => e.VoltajeTierra)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("VOLTAJE_TIERRA");

                entity.Property(e => e.VoltajeTierraNeutro)
                    .HasMaxLength(50)
                    .IsUnicode(false)
                    .HasColumnName("VOLTAJE_TIERRA_NEUTRO");
            });
            
            modelBuilder.Entity<BdColumnsLayout>(entity =>
            {
                entity.HasKey(e => e.IdColumnLayout)
                    .HasName("PK__BD_COLUM__15908F6A27388CC5");

                entity.ToTable("BD_COLUMNS_LAYOUTS");

                entity.Property(e => e.IdColumnLayout).HasColumnName("ID_COLUMN_LAYOUT");

                entity.Property(e => e.DescColumnLayout)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("DESC_COLUMN_LAYOUT");

                entity.Property(e => e.FecAlta)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_ALTA");

                entity.Property(e => e.IdLayout).HasColumnName("ID_LAYOUT");

                entity.Property(e => e.IdUsuarioAlta).HasColumnName("ID_USUARIO_ALTA");

                entity.Property(e => e.Orden).HasColumnName("ORDEN");

                entity.Property(e => e.Status)
                    .HasColumnName("STATUS")
                    .HasDefaultValueSql("((1))");
            });

            modelBuilder.Entity<CLayout>(entity =>
            {
                entity.HasKey(e => e.IdLayout)
                    .HasName("PK__C_LAYOUT__5B142612336110C6");

                entity.ToTable("C_LAYOUTS");

                entity.Property(e => e.IdLayout).HasColumnName("ID_LAYOUT");

                entity.Property(e => e.DescLayout)
                    .HasMaxLength(255)
                    .IsUnicode(false)
                    .HasColumnName("DESC_LAYOUT");

                entity.Property(e => e.FecAlta)
                    .HasColumnType("datetime")
                    .HasColumnName("FEC_ALTA");

                entity.Property(e => e.IdUsuarioAlta).HasColumnName("ID_USUARIO_ALTA");

                entity.Property(e => e.Status)
                    .HasColumnName("STATUS")
                    .HasDefaultValueSql("((1))");
            });

        }
    }
}
