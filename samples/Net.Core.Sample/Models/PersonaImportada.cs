using System;
using System.ComponentModel;
using System.Globalization;

namespace Net.Core.Sample.Models;

public class PersonaImportada
{
    public string Accion { get; set; }
    public string DocumentoIdentidad { get; set; }
    public string TipoDocumento { get; set; }

    [DisplayName("TipoDocumentoId")]
    public string TipoDocumentoIdStr { get; set; }
    public Guid TipoDocumentoId { get; set; }
    public string UserName { get; set; }
    public string Nombre { get; set; }
    public string Apellido1 { get; set; }
    public string Apellido2 { get; set; }
    public string Email { get; set; }
    public string EmailDescripcion { get; set; }
    public string Email2 { get; set; }
    public string Email2Descripcion { get; set; }
    public string Email3 { get; set; }
    public string Email3Descripcion { get; set; }
    public string Telefono1 { get; set; }
    public string Telefono1Descripcion { get; set; }
    public string Telefono2 { get; set; }
    public string Telefono2Descripcion { get; set; }
    public string Telefono3 { get; set; }
    public string Telefono3Descripcion { get; set; }
    public string Telefono4 { get; set; }
    public string Telefono4Descripcion { get; set; }
    public string TelefonoMovil { get; set; }
    public string TelefonoMovilDescripcion { get; set; }
    public bool TelefonoMovilAdmiteSms { get; set; }
    public string Fax { get; set; }
    public string FaxDescripcion { get; set; }
    public string Notas { get; set; }
    public bool RecibirBoletin { get; set; }
    public bool AceptoPoliticas { get; set; }
    public string TipoVia { get; set; }
    public Guid TipoViaId { get; set; }
    public string Direccion_DireccionPostal { get; set; }
    public string Direccion_Bloque { get; set; }
    public string Direccion_Numero { get; set; }
    public string Direccion_Escalera { get; set; }
    public string Direccion_Piso { get; set; }
    public string Direccion_Puerta { get; set; }
    public string Direccion_CodigoPostal { get; set; }
    public string Direccion_Poblacion { get; set; }
    public string Direccion_CodigoMunicipio { get; set; }
    public string Direccion_Municipio { get; set; }
    public string Direccion_CodigoProvincia { get; set; }
    public string Direccion_Provincia { get; set; }
    public string Direccion_CodigoPais { get; set; }
    public string Direccion_Pais { get; set; }
    public string Direccion_Observaciones { get; set; }
    public bool Bloqueado { get; set; }
    public string OrigenCarga { get; set; }
    public string OrigenIdentificador { get; set; }
    public Guid IdFuac { get; set; }

    [DisplayName("IdFuac")]
    public string FuacStr { get; set; }


    [DisplayName("FechaNacimiento")]
    public DateTime? FechaNacimientoStr { get; set; }
    [DisplayName("Sexo")]
    public int? SexoStr { get; set; }
    [DisplayName("FechaAceptacionPoliticas")]
    public string FechaAceptacionPoliticasStr { get; set; }
    [DisplayName("FechaBajaPoliticas")]
    public string FechaBajaPoliticasStr { get; set; }

    //  internal DateTime? FechaNacimiento => (!string.IsNullOrEmpty(FechaNacimientoStr)) ? Convert.ToDateTime(FechaNacimientoStr, CultureInfo.CurrentCulture) : (DateTime?)null;
    internal DateTime? FechaAceptacionPoliticas => (!string.IsNullOrEmpty(FechaAceptacionPoliticasStr)) ? Convert.ToDateTime(FechaAceptacionPoliticasStr, CultureInfo.CurrentCulture) : (DateTime?)null;
    internal DateTime? FechaBajaPoliticas => (!string.IsNullOrEmpty(FechaBajaPoliticasStr)) ? Convert.ToDateTime(FechaBajaPoliticasStr, CultureInfo.CurrentCulture) : (DateTime?)null;
    //internal int? Sexo => SexoStr!=null && SexoStr.Trim().Equals("Masculino", StringComparison.InvariantCultureIgnoreCase) ? 0 :
    //    SexoStr.Trim().Equals("Femenino", StringComparison.InvariantCultureIgnoreCase) ? (int?)1 : null;
}
