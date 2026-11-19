namespace TesoreriaCobranza
{
    enum MetodoIdentificacion
    {
        Posible =1 ,
        Compensada= 2,
        FolioFactura=4,
        ImporteAprox =8,
        SumaFacturas=16,
        NoIdentificado=32,
        FacturaMesCorriente=64,
        SumaRazonSocialFacturas =128,
        RazonSocial =256
    }

    public  enum TipoPolizas
    {
        Comisiones,
        Cobranza
    }
}