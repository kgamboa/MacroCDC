# Macro CDC PTC

Herramienta de conciliación y clasificación de movimientos bancarios (BBVA, SPEI, OpenPay, pólizas, CFDI).

## Qué hace la macro

Toma los estados de cuenta bancarios y los prepara para la contabilidad y la cobranza:

- **Genera archivo de bancos:** arma el Excel de movimientos por banco a partir de los estados de cuenta.
- **Categoriza bancos:** asigna a cada movimiento su tipo (Cobranza, Tesorería, Pago a proveedor, RCE, CALIFICATE, etc.) según el concepto y el importe.
- **Identifica facturas:** cruza los movimientos con los CFDI para encontrar la factura que corresponde a cada pago.
- **Ingresos pequeños:** procesa los ingresos de pequeña cuantía (RCE, CALIFICATE, pasarelas).
- **Genera pólizas:** crea las pólizas contables de bancos, cobranza, OpenPay, ingresos y cierre de mes.
- **Póliza semanal:** resume los movimientos de la semana para el registro contable.

La clasificación de rangos y el tipo de cada movimiento están en `ClasificacionBancos.cs`. Las pólizas se generan en `ProcesoPoliza*.cs`.

## Instalar (versión 4.2.0.105)

1. Recibe el archivo `CDC_Macro PTC 105.zip`.
2. Haz clic derecho sobre el zip y elige **Extraer todo**. Pon como destino una carpeta fácil de encontrar, por ejemplo `C:\CDC_Macro PTC`.
3. Entra a la carpeta extraída y haz doble clic en **setup.exe**.
4. Sigue el asistente de instalación. Si Windows muestra una advertencia de seguridad, revisa que el archivo venga de la fuente correcta y acepta solo si es así.
5. Al terminar, abre **Macro CDC PTC** desde el menú Inicio.

### Requisitos

- Windows 10 o 11.
- .NET Framework 4.8 (normalmente ya viene en Windows).
- Microsoft Excel instalado, porque la app lee y escribe archivos de Excel.
- Permiso para instalar programas en la PC.

### Si algo falla

- Si el instalador no abre, vuelve a extraer el zip completo antes de ejecutar `setup.exe`.
- Si pide actualizar, acepta la actualización.

## Despliegue en cliente

- **Instalación en la PC de Maricela (cliente):** se instaló como administrador de la máquina, desde una carpeta compartida en `Z:` (ClickOnce), y quedó en `C:`.
- **Versión instalada:** 4.2.0.105 (revisión ClickOnce), ensamblado 2.7.2.11. Pendiente confirmar en el equipo del cliente.
- **Actualizaciones:** ClickOnce busca la nueva versión en la ruta de `Z:` desde la que se instaló. Si `Z:` no está conectada, la app no se actualiza.
- **Error al crear acceso directo:** `Acceso denegado` en el escritorio de la cuenta del cliente. No afecta la app instalada.

## Compilar desde el código

1. Clona el repo: `https://github.com/kgamboa/MacroCDC.git`.
2. Copia `Clases\CuentasBancos.json` dentro de `Bancos\CategoriaBancos\Clases\`. No está en GitHub porque tiene cuentas bancarias reales.
3. Abre `CirculoCredito.sln` en Visual Studio y compila, o usa MSBuild:

```powershell
& 'C:\Program Files\Microsoft Visual Studio\18\Community\MSBuild\Current\Bin\MSBuild.exe' 'C:\Users\karlo\source\repos\MacroCDC\CirculoCredito.sln' /t:Build /p:Configuration=Debug /nologo
```

El ejecutable queda en `Bancos\CategoriaBancos\bin\Debug\MacroCDC_PTC.exe`.

## Reglas de clasificación

Los movimientos en pesos de **Santander**, **BBVA 0798** y **BAZ 450** se clasifican como RCE o CALIFICATE según el importe neto, con la tabla `RangosClasificacion` en `ClasificacionBancos.cs`. Cada rango va desde su inicio hasta el inicio del siguiente, con tope en 310.99. Importes de 311 o más no tienen regla y quedan como Cobranza.
