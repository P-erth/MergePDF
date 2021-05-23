using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Drawing;
using ZXing;

namespace MergePDF
{
    class Program
    {

        static void Main(string[] args)
        {
            
            string stdIn = "9800000000";
            foreach (string value in args) stdIn = value;
            string ruta = stdIn.Substring(0,2);
            string secuencia = stdIn.Substring(2,8);
            ruta = "02";
            secuencia = "0";
            bool isNoventaYOcho;

            ruta = "98";
            var path = "";
            if (ruta == "98")
            {
                path = Path.Combine(@"i:/_pdf_col/" , "ppdd-t2.txt");
                isNoventaYOcho = true;
            }
            else
            {
                path = @"I:/_pdf_col/ppdd.txt";
                isNoventaYOcho = false;
            }

            string textToParse = System.IO.File.ReadAllText(path);
            int paginas = textToParse.Length / 9010;
            int cont = 0;
            SortedDictionary<string, string> paginasDeRuta = new SortedDictionary<string, string>();

            //Recorrer todo el txt y quedarme con las que coinciden con la ruta
            for (int i = 0; i < paginas; i++)
            {
                //las rutas que coinciden las agregamos dentro de una lista
                if (i != 0) cont += 9010;
                string pagina = textToParse.Substring(cont, 9010);
                string lspSecuencia = pagina.Substring(661, 10);
                string rutaAEvaluar = lspSecuencia.Substring(0, 2);

                if (isNoventaYOcho)
                {
                    paginasDeRuta.Add(lspSecuencia, pagina);
                }
                else { if (String.Compare(rutaAEvaluar, ruta) == 0) paginasDeRuta.Add(lspSecuencia, pagina); }
            }



            //Creamos un documento unico
            PdfDocument document = new PdfDocument();
            document.Info.Title = "Cooperativa Electrica Colón - Buenos Aires";



            if(secuencia == "0")
            {
                foreach(string pagina in paginasDeRuta.Values)
                {
                    pdfGenerator(pagina, document);
                }
            }
            else
            {
                bool band = false;
                foreach (var pagina in paginasDeRuta)
                {
                    if(pagina.Key.Substring(2,8).Equals(secuencia))
                    {
                        band = true;
                    }
                    if (band) pdfGenerator(pagina.Value, document);
                }
            }

            document.Options.FlateEncodeMode = PdfFlateEncodeMode.BestCompression;
            string filename = "PdfPrueba.pdf";
            document.Save(filename);
            System.Diagnostics.Process.Start(filename);

            Console.ReadLine();
        }

        private static void pdfGenerator(string pagina, PdfDocument document)
        {
            int pivote = 0;
            // Create an empty page
            PdfPage page = document.AddPage();
            page.Size = PdfSharp.PageSize.A4;
            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //armado de variables
            String nis = pagina.Substring(pivote, 8);
            String nombre = pagina.Substring(pivote += 8, 30);
            String domiReal = pagina.Substring(pivote += 30, 30);
            String postal = pagina.Substring(pivote += 30, 30);
            String localidad = pagina.Substring(pivote += 30, 30);
            String socio = pagina.Substring(pivote += 30, 8);
            String socioDesde = pagina.Substring(pivote += 8, 8);
            String socioActa = pagina.Substring(pivote += 8, 8);
            String socioTipo = pagina.Substring(pivote += 8, 3);
            String socioDoc = pagina.Substring(pivote += 3, 11);
            //
            List<string> informaciones = new List<string>();//Tabla Inf
            informaciones.Add(pagina.Substring(pivote += 11, 52));
            for (int i = 0; i < 5; i++) informaciones.Add(pagina.Substring(pivote += 52, 52));
            //
            String cuit = pagina.Substring(pivote += 52, 11);
            String condiva = pagina.Substring(pivote += 11, 20);
            String cbu = pagina.Substring(pivote += 20, 22);
            String cuFecha = pagina.Substring(pivote += 22, 2) + "/" + pagina.Substring(pivote += 2, 2) + "/" + pagina.Substring(pivote += 2, 4);
            String cuHora = pagina.Substring(pivote += 4, 2) + ":" + pagina.Substring(pivote += 2, 2);
            String diaVto = pagina.Substring(pivote += 2, 2);
            String mesVto = pagina.Substring(pivote += 2, 2);
            String añoVto = pagina.Substring(pivote += 2, 4);
            //
            if (mesVto == "12")
            {
                mesVto = "01";
                int año = Int32.Parse(añoVto);
                añoVto = año.ToString();
            }
            else
            {
                int mes = Int32.Parse(mesVto);
                mes++;
                mesVto = mes.ToString();
            }
            //
            String proxVto = diaVto + "/" + mesVto + "/" + añoVto;
            String lspParte1 = pagina.Substring(pivote += 4, 1);
            String lspParte2 = pagina.Substring(pivote += 1, 4);
            String lsp = lspParte1 + "-" + lspParte2 + "-" + pagina.Substring(pivote += 4, 8);
            String codCom = pagina.Substring(pivote += 8, 2);
            String cesp = pagina.Substring(pivote += 2, 14);
            String cespEmis = pagina.Substring(pivote += 14, 2) + "/" + pagina.Substring(pivote += 2, 2) + "/" + pagina.Substring(pivote += 2, 4);
            String cespVto = pagina.Substring(pivote += 4, 2) + "/" + pagina.Substring(pivote += 2, 2) + "/" + pagina.Substring(pivote += 2, 4);
            String lspTarifa = pagina.Substring(pivote += 4, 6);
            String lspSocial = pagina.Substring(pivote += 6, 1);
            String lspMedidor = pagina.Substring(pivote += 1, 10);
            String lspEstadoAnt = pagina.Substring(pivote += 10, 10);
            String lspEstadoAnFec = pagina.Substring(pivote += 10, 2) + "/" + pagina.Substring(pivote += 2, 2) + "/" + pagina.Substring(pivote += 2, 4);
            String lspEstadoAc = pagina.Substring(pivote += 4, 10);
            String lspEstadoAcFec = pagina.Substring(pivote += 10, 2) + "/" + pagina.Substring(pivote += 2, 2) + "/" + pagina.Substring(pivote += 2, 4);
            String lspFactor = pagina.Substring(pivote += 4, 4);
            String lspConsumo = pagina.Substring(pivote += 4, 8);
            String lspSecuencia = pagina.Substring(pivote += 8, 10);
            //
            List<string> cuerpos = new List<string>();//Tabla Cuerpo
            cuerpos.Add(pagina.Substring(pivote += 10, 85));
            for (int i = 0; i < 39; i++) cuerpos.Add(pagina.Substring(pivote += 85, 85));
            List<string> deudas = new List<string>();//Tabla Deuda
            deudas.Add(pagina.Substring(pivote += 85, 50));
            for (int i = 0; i < 11; i++) deudas.Add(pagina.Substring(pivote += 50, 50));
            List<string> estadisticos = new List<string>();//Tabla estadisticos
            for (int i = 0; i < 7; i++) estadisticos.Add(pagina.Substring(pivote += 50, 50));
            List<string> recargos = new List<string>();//Tabla recargos
            for (int i = 0; i < 11; i++) recargos.Add(pagina.Substring(pivote += 50, 50));
            //
            String promedio = long.Parse(pagina.Substring(pivote += 50, 8)).ToString(); ///CONVERTIRLO A ENTERO Y DESPUES A STRING
            String totControl = pagina.Substring(pivote += 8, 8);
            if (totControl == "00000000") totControl = "0";
            //saco los ceros de la parte entera del importe
            String totImporteEntero = pagina.Substring(pivote += 8, 10);
            long parteEntera = long.Parse(totImporteEntero);
            totImporteEntero = parteEntera.ToString();
            String totImporteDecimal = pagina.Substring(pivote += 10, 2);
            // y guardo todo el importe entero
            String cod1 = pagina.Substring(pivote += 2, 28);
            //
            String lsp2 = lspParte1 + "-" + lspParte2 + "-" + pagina.Substring(pivote += 28, 8);
            List<string> cuerposTabla2 = new List<string>();
            cuerposTabla2.Add(pagina.Substring(pivote += 8, 85));
            for (int i = 0; i < 24; i++) cuerposTabla2.Add(pagina.Substring(pivote += 85, 85));//Tabla cuerpos hoja 2
            List<string> recargosTabla2 = new List<string>();
            recargosTabla2.Add(pagina.Substring(pivote += 85, 50));
            for (int i = 0; i < 11; i++) recargosTabla2.Add(pagina.Substring(pivote += 50, 50));//tabla recargos hoja 2
            List<string> sepelios = new List<string>();
            for (int i = 0; i < 12; i++) sepelios.Add(pagina.Substring(pivote += 50, 50));//tabla sepelios hoja 2
            String totControl2 = pagina.Substring(pivote += 50, 8);
            if (totControl2 == "00000000") totControl2 = "0";
            String totImporte2 = long.Parse(pagina.Substring(pivote += 8, 10)).ToString() + "." + pagina.Substring(pivote += 10, 2);
            String totImporte = totImporteEntero + "." + totImporteDecimal;
            String cod2 = pagina.Substring(pivote += 2, 28);

            //////////////////DATOS PARA QR1///////////////////
            String qr1 = "";
            qr1 += "000201";//Formato
            qr1 += "010212";//Metodo de inicio
            qr1 += "40230010com.yacare0105Y1103";//Datos Yacare
            qr1 += "5015001130545748831";//Cuil Empresa
            qr1 += "52044900"; //Codigo comercio
            qr1 += "5303032"; //Moneda de la TX
            String importe = Int32.Parse(cod1.Substring(18, 7)).ToString();
            String dec = cod1.Substring(25, 2);
            String cant = (importe.Length + 3).ToString();
            qr1 += "540" + cant + importe + "." + dec; //Importe
            qr1 += "5802AR"; //Codigo de pais
            qr1 += "5925COOPERATIVAELECTRICACOLON"; //Nombre de empresa
            qr1 += "6005Colon"; //Ciudad empresa
            qr1 += "61043280"; //Codigo postal
            String idCliente = cod1.Substring(4, 7);
            qr1 += "62230108" + totControl + "0607" + idCliente; //Datos cliente
            qr1 += "6304"; //digitos de mierda de verificacion de Yacaré
            String crc = CalculaCRC(qr1, Encoding.UTF8);
            qr1 += crc;
            ///////////////////////////////////////////////////
            //////////////////DATOS PARA QR2///////////////////
            String qr2 = "";
            qr2 += "000201";//Formato
            qr2 += "010212";//Metodo de inicio
            qr2 += "40230010com.yacare0105Y1103";//Datos Yacare
            qr2 += "5015001130545748831";//Cuil Empresa
            qr2 += "52044900"; //Codigo comercio
            qr2 += "5303032"; //Moneda de la TX
            importe = Int32.Parse(cod2.Substring(18, 7)).ToString();
            dec = cod2.Substring(25, 2);
            cant = (importe.Length + 3).ToString();
            qr2 += "540" + cant + importe + "." + dec; //Importe
            qr2 += "5802AR"; //Codigo de pais
            qr2 += "5925COOPERATIVAELECTRICACOLON"; //Nombre de empresa
            qr2 += "6005Colon"; //Ciudad empresa
            qr2 += "61043280"; //Codigo postal
            idCliente = cod2.Substring(4, 7);
            qr2 += "62230108" + totControl2 + "0607" + idCliente; //Datos cliente
            qr2 += "6304"; //digitos de mierda de verificacion de Yacaré
            String crc2 = CalculaCRC(qr1, Encoding.UTF8);
            qr2 += crc2;
            //////////////////////////////////////////////////////////
            DrawQR(gfx, qr1, qr2);
            DrawBarCode(gfx, cod1, cod2);
            // Declaracion de fuentes
            XFont fontCourierBold15 = new XFont("Courier New", 15, XFontStyle.Bold);
            XFont fontCourierBold14 = new XFont("Courier New", 14, XFontStyle.Bold);
            XFont fontCourierBold13 = new XFont("Courier New", 13, XFontStyle.Bold);
            XFont fontCourierBold7 = new XFont("Courier New", 7, XFontStyle.Bold);
            XFont fontCourier7 = new XFont("Courier New", 7, XFontStyle.Regular);
            XFont fontCourier6 = new XFont("Courier New", 6, XFontStyle.Regular);
            XFont fontCourierBold10 = new XFont("Courier New", 10, XFontStyle.Bold);
            XFont fontCourierBold9 = new XFont("Courier New", 9, XFontStyle.Bold);
            ///////////////////////////HOJA1//////////////////////////

            gfx.DrawString("NIS:  " + (long.Parse(nis)).ToString(), fontCourierBold15, XBrushes.Black, 410, 90);
            gfx.DrawString(nombre, fontCourierBold14, XBrushes.Black, 25, 92);
            gfx.DrawString(domiReal, fontCourierBold14, XBrushes.Black, 25, 105);
            gfx.DrawString(postal, fontCourierBold14, XBrushes.Black, 25, 118);
            gfx.DrawString(localidad, fontCourierBold14, XBrushes.Black, 25, 131);
            gfx.DrawString("Cond.Iva:" + condiva, fontCourierBold7, XBrushes.Black, 25, 142);
            gfx.DrawString("CUIT: " + cuit, fontCourierBold7, XBrushes.Black, 155, 142);
            gfx.DrawString("*" + cod1 + "*", fontCourier6, XBrushes.Black, 70, 168);
            //

            int posy = 142;
            if ((cbu == "0000000000000000000000") || (cbu == "                      ") || (cbu.Trim() == ""))
            {
                gfx.DrawString("CODIGO DE PAGO ELECTRONICO: " + (long.Parse(nis)).ToString(), fontCourierBold10, XBrushes.Black, 305, 150);
            }
            else
            {
                gfx.DrawString("El importe de esta factura sera debitado a su vencimiento de acuerdo al CBU", fontCourierBold7, XBrushes.Black, 245, posy);
                gfx.DrawString("número: " + cbu + " por lo tanto recomendamos verificar para tal", fontCourierBold7, XBrushes.Black, 245, posy += 7);
                gfx.DrawString("fecha tener el saldo disponible en su cuenta bancaria", fontCourierBold7, XBrushes.Black, 245, posy += 7);
            }
            if (lspSocial == "1") gfx.DrawString("**TARIFA SOCIAL**", fontCourierBold9, XBrushes.Black, 355, 165);

            posy = 185;
            //cuerpo
            foreach (string cuerpo in cuerpos) gfx.DrawString(cuerpo, fontCourier7, XBrushes.Black, 243, posy += 7);
            posy = 183;
            //deudas
            foreach (string deuda in deudas) gfx.DrawString(deuda, fontCourier6, XBrushes.Black, 35, posy += 7);
            posy = 255;
            //estadisticos
            foreach (string estadistico in estadisticos) gfx.DrawString(estadistico, fontCourier6, XBrushes.Black, 43, posy += 7);
            gfx.DrawString("Promedio : " + promedio, fontCourier6, XBrushes.Black, 105, posy += 7);
            //recargos
            posy = 326;
            foreach (string recargo in recargos) gfx.DrawString(recargo, fontCourier6, XBrushes.Black, 30, posy += 7);
            posy = 404;
            //informaciones al pie se la hoja 1
            foreach (string info in informaciones) gfx.DrawString(info, fontCourierBold7, XBrushes.Black, 30, posy += 7);
            gfx.DrawString("Numeración enitida como gran contribuyente el " + cuFecha + " a las " + cuHora, fontCourierBold7, XBrushes.Black, 34, posy += 7);
            //importa a pagar y vencimiento
            gfx.DrawString(proxVto, fontCourierBold14, XBrushes.Black, 382, 451);
            gfx.DrawString(totImporte, fontCourierBold14, XBrushes.Black, 495, 451);

            posy = 22;
            gfx.DrawString("Liq. Serv. Públicos", fontCourierBold7, XBrushes.Black, 400, posy);
            gfx.DrawString(lsp, fontCourierBold7, XBrushes.Black, 500, posy);
            gfx.DrawString("Código Comprobante", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(codCom, fontCourierBold7, XBrushes.Black, 555, posy);
            gfx.DrawString("C.E.S.P. Número", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(cesp, fontCourierBold7, XBrushes.Black, 505, posy);
            gfx.DrawString("Vencimiento CESP", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(cespVto, fontCourierBold7, XBrushes.Black, 522, posy);
            gfx.DrawString("Control de pago", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(totControl, fontCourierBold7, XBrushes.Black, 530, posy);
            gfx.DrawString("Fecha emisión", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(cespEmis, fontCourierBold7, XBrushes.Black, 522, posy);
            posy = 102;
            gfx.DrawString("Medidor Nro.", fontCourierBold7, XBrushes.Black, 275, posy);
            gfx.DrawString(lspMedidor, fontCourierBold7, XBrushes.Black, 352, posy);
            gfx.DrawString("Estado Anter.", fontCourierBold7, XBrushes.Black, 400, posy);
            gfx.DrawString((long.Parse(lspEstadoAnt)).ToString(), fontCourierBold7, XBrushes.Black, 470, posy);
            gfx.DrawString(lspEstadoAnFec, fontCourierBold7, XBrushes.Black, 505, posy);
            gfx.DrawString("Socio", fontCourierBold7, XBrushes.Black, 275, posy += 7);
            gfx.DrawString(socio, fontCourierBold7, XBrushes.Black, 360, posy);
            gfx.DrawString("Estado Actual", fontCourierBold7, XBrushes.Black, 400, posy);
            gfx.DrawString((long.Parse(lspEstadoAc)).ToString(), fontCourierBold7, XBrushes.Black, 470, posy);
            gfx.DrawString(lspEstadoAcFec, fontCourierBold7, XBrushes.Black, 505, posy);
            gfx.DrawString("Tarifa", fontCourierBold7, XBrushes.Black, 275, posy += 7);
            gfx.DrawString("Factor", fontCourierBold7, XBrushes.Black, 400, posy);
            gfx.DrawString((long.Parse(lspFactor)).ToString(), fontCourierBold7, XBrushes.Black, 480, posy);
            gfx.DrawString(lspTarifa, fontCourierBold7, XBrushes.Black, 368, posy);
            gfx.DrawString("Consumo", fontCourierBold7, XBrushes.Black, 400, posy += 7);
            gfx.DrawString((long.Parse(lspConsumo)).ToString(), fontCourierBold7, XBrushes.Black, 470, posy);
            gfx.DrawString("Pxmo. vto. desde", fontCourierBold7, XBrushes.Black, 275, posy += 7);
            gfx.DrawString(proxVto, fontCourierBold7, XBrushes.Black, 356, posy);
            gfx.DrawString("Secuencia", fontCourierBold7, XBrushes.Black, 400, posy);
            gfx.DrawString(lspSecuencia, fontCourierBold7, XBrushes.Black, 456, posy);
            ////////////////////////FINHOJA1////////////////////////////////////////////
            /////////////////////////HOJA2 //////////////////////////////

            gfx.DrawString("NIS:  " + (long.Parse(nis)).ToString(), fontCourierBold15, XBrushes.Black, 410, 548);
            posy = 536;

            gfx.DrawString(nombre, fontCourierBold13, XBrushes.Black, 25, posy);
            gfx.DrawString(domiReal, fontCourierBold13, XBrushes.Black, 25, posy += 10);
            gfx.DrawString(postal, fontCourierBold13, XBrushes.Black, 25, posy += 10);
            gfx.DrawString(localidad, fontCourierBold13, XBrushes.Black, 25, posy += 10);
            gfx.DrawString("Cond.Iva:" + condiva, fontCourierBold7, XBrushes.Black, 25, posy += 7);
            gfx.DrawString("CUIT: " + cuit, fontCourierBold7, XBrushes.Black, 155, posy);
            gfx.DrawString("*" + cod2 + "*", fontCourier6, XBrushes.Black, 70, 599);
            posy += 8;
            if ((cbu == "0000000000000000000000") || (cbu == "                      ") || (cbu.Trim() == ""))
            {
                gfx.DrawString("CODIGO DE PAGO ELECTRONICO: " + (long.Parse(nis)).ToString(), fontCourierBold10, XBrushes.Black, 305, posy);
            }
            else
            {
                posy -= 7;
                gfx.DrawString("El importe de esta factura sera debitado a su vencimiento de acuerdo al CBU", fontCourierBold7, XBrushes.Black, 245, posy);
                gfx.DrawString("número: " + cbu + " por lo tanto recomendamos verificar para tal", fontCourierBold7, XBrushes.Black, 245, posy += 7);
                gfx.DrawString("fecha tener el saldo disponible en su cuenta bancaria", fontCourierBold7, XBrushes.Black, 245, posy += 7);
            }
            if (lspSocial == "1") gfx.DrawString("**TARIFA SOCIAL**", fontCourierBold9, XBrushes.Black, 355, posy += 16);
            posy = 470;
            gfx.DrawString("Liq. Serv. Públicos", fontCourierBold7, XBrushes.Black, 400, posy);
            gfx.DrawString(lsp2, fontCourierBold7, XBrushes.Black, 500, posy);
            gfx.DrawString("Código Comprobante", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(codCom, fontCourierBold7, XBrushes.Black, 555, posy);
            gfx.DrawString("C.E.S.P. Número", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(cesp, fontCourierBold7, XBrushes.Black, 505, posy);
            gfx.DrawString("Vencimiento CESP", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(cespVto, fontCourierBold7, XBrushes.Black, 522, posy);
            gfx.DrawString("Fecha emisión", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(cespEmis, fontCourierBold7, XBrushes.Black, 522, posy);
            gfx.DrawString("Control de pago", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(totControl2, fontCourierBold7, XBrushes.Black, 530, posy);
            gfx.DrawString("Socio", fontCourierBold7, XBrushes.Black, 400, posy += 10);
            gfx.DrawString(socio, fontCourierBold7, XBrushes.Black, 530, posy);
            posy = 612;
            foreach (string recargo in recargosTabla2) gfx.DrawString(recargo, fontCourier6, XBrushes.Black, 30, posy += 6);
            posy = 692;
            foreach (string sepelio in sepelios) gfx.DrawString(sepelio, fontCourier6, XBrushes.Black, 30, posy += 6);

            gfx.DrawString(proxVto, fontCourierBold14, XBrushes.Black, 382, 824);
            gfx.DrawString(totImporte2, fontCourierBold14, XBrushes.Black, 504, 824);
            gfx.DrawString("Numeración enitida como gran contribuyente el " + cuFecha + " a las " + cuHora, fontCourierBold7, XBrushes.Black, 34, 827);
            posy = 612;
            foreach (string cuerpo in cuerposTabla2) gfx.DrawString(cuerpo, fontCourier7, XBrushes.Black, 242, posy += 6);

            ////////////////////FINHOJA2////////////////////
        }

        private static void DrawQR(XGraphics gfx, string qr1, string qr2)
        {
            XImage img = XImage.FromFile("template2.jpg");
            gfx.DrawImage(img, 0, 0);

            //Draw QR1
            var bcWriter = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new ZXing.Common.EncodingOptions
                {
                    Height = 100,
                    Width = 100,
                    Margin = 0
                },
            };
            Bitmap bm = bcWriter.Write(qr1);
            XImage img2 = XImage.FromGdiPlusImage((Image)bm);
            img2.Interpolate = false;
            gfx.DrawImage(img2, 495, 335);
            //Draw QR2
            bm = bcWriter.Write(qr2);
            img2 = XImage.FromGdiPlusImage((Image)bm);
            img2.Interpolate = false;
            gfx.DrawImage(img2, 495, 710);
        }
        private static void DrawBarCode(XGraphics gfx, string code1, string code2)
        {
            var bcWriter = new BarcodeWriter
            {
                Format = BarcodeFormat.CODE_128,
                Options = new ZXing.Common.EncodingOptions
                {
                    Height = 20,
                    Width = 253,
                    Margin = 0
                },
            };
            bcWriter.Options.PureBarcode = true;
            //Drawing BarCode1
            Bitmap bm = bcWriter.Write(code1);
            XImage img = XImage.FromGdiPlusImage((Image)bm);
            img.Interpolate = false;
            gfx.DrawImage(img, 30, 147);
            //Drawing BarCode2
            bm = bcWriter.Write(code2);
            img = XImage.FromGdiPlusImage((Image)bm);
            img.Interpolate = false;
            gfx.DrawImage(img, 30, 578);
        }

        static String CalculaCRC(String value, Encoding enc)
        {
            byte[] data = enc.GetBytes(value);
            int crcValue = 0xFFFF;
            for (int b = 0; b < data.Length; b++)
            {
                for (int i = 0; i < 8; i++)
                {
                    Boolean bit = (((int)((uint)data[b] >> (7 - i))) & 1) == 1;
                    Boolean c15 = (crcValue >> 15 & 1) == 1;
                    crcValue <<= 1;
                    if (c15 ^ bit)
                    {
                        crcValue ^= 0x1021;
                    }

                }

            }
            crcValue &= 0xffff;
            int hexString = (crcValue & 0xFFFF);
            hexString = (crcValue & 0xFFFF);
            string retu = Convert.ToString(hexString, 16);



            return retu;
        }


    }
}