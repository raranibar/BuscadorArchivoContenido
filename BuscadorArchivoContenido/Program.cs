using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Configuration;
using System.Text;
using System.Threading.Tasks;
using TCP.Conexión;
using TCP.FrameWork.RN;

namespace BuscadorArchivoContenido
{
    internal class Program
    {
        static void Main(string[] args)
        {
            const string DirectorioPdf = @"d:\ContenidoPdf\";
            const string DirectorioWord = @"d:\ContenidoWord\";

            Console.WriteLine("Inicio => {0}", DateTime.Now.ToString());
            bool abre = AbrirConexión();
            int cont = 0;

            //new Parámetro("Texto", "traslativa") ,
            //new Parámetro("Texto", "\"legitimación pasiva\"") ,
            //exec[dbo].[PaBuscarResoluciónJurisprudenciaAvanzada]  '\"legitimacion pasiva\"',0,null,null,16,null,null,39
            DataTable dtListaArchivos = Usuario.Sesión.EjecutarTabla("PaBuscarResoluciónJurisprudenciaAvanzada", new List<Parámetro>() {
                new Parámetro("Texto", "\"legitimación pasiva\"") ,
                new Parámetro("IdSección", 0) ,
                new Parámetro("FechaInicio", null) ,
                new Parámetro("FechaFin", null) ,
                new Parámetro("Recurso", 16) ,
                new Parámetro("Distrito", null) ,
                new Parámetro("Institución", null),
                new Parámetro("Magistrado", 39)
            });

            foreach (DataRow row in dtListaArchivos.Rows)
            {
                cont++;
                string IdResolucion = row[0].ToString();
                string TipoResolucion = row[4].ToString();
                string NumeroResolucion = row[5].ToString().Replace('/','-');

                switch ( TipoResolucion)
                {
                    case "VOTO ACLARATORIO": NumeroResolucion = NumeroResolucion + "_VA"; break;
                    case "VOTO DISIDENTE": NumeroResolucion = NumeroResolucion + "_VD"; break;
                    default:
                        break;
                }

                Byte[] Pdf = null;
                Byte[] Word = null;
                DataTable dt = Usuario.Sesión.EjecutarTabla("PaSeleccionarResoluciónArchivos", new List<Parámetro>() {
                 new Parámetro("IdResolución", IdResolucion)});
                //Genera en Word
                if (dt.Rows[0]["Word"] != System.DBNull.Value)
                {
                    Word = (Byte[])dt.Rows[0]["Word"];
                    GuardaWord(Word, DirectorioWord, NumeroResolucion.ToString());
                }
                //Genera en Pdf
                if (dt.Rows[0]["Pdf"] != System.DBNull.Value)
                {
                    Pdf = (Byte[])dt.Rows[0]["Pdf"];
                    GuardaPdf(Pdf, DirectorioPdf, NumeroResolucion.ToString());
                }                
            }

            Usuario.Sesión.Desconectar();
            Console.WriteLine("Fin => {0}", DateTime.Now.ToString());
            Console.WriteLine("Total de Expedientes => {0}", cont.ToString());
            Console.ReadLine();
        }

        static void GuardaWord(byte[] Word, string DirectoryName, string FileNameWord)
        {

            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            System.IO.FileStream fs = new System.IO.FileStream(DirectoryName + FileNameWord + ".docx", System.IO.FileMode.Create);
            fs.Write(Word, 0, Convert.ToInt32(Word.Length));
            fs.Close();

        }

        static void GuardaPdf(byte[] Pdf, string DirectoryName, string FileNamePdf)
        {
            if (!Directory.Exists(DirectoryName))
            {
                Directory.CreateDirectory(DirectoryName);
            }
            System.IO.FileStream fspdf = new System.IO.FileStream(DirectoryName + FileNamePdf + ".pdf", System.IO.FileMode.Create);
            fspdf.Write(Pdf, 0, Convert.ToInt32(Pdf.Length));
            fspdf.Close();

        }
        static bool AbrirConexión()
        {
            try
            {

                Usuario.Login = ConfigurationManager.AppSettings["Login"];
                Usuario.Password = ConfigurationManager.AppSettings["Password"];
                Usuario.Server = ConfigurationManager.AppSettings["Server"];
                Usuario.Base = ConfigurationManager.AppSettings["Base"];
                Usuario.ProveedorDB = ConfigurationManager.AppSettings["Proveedor"];
                Usuario.Autenticar();
            }
            catch (Exception ex)
            {
                if (ex.Message.StartsWith("I/O error for file CreateFile (open)"))
                    Console.WriteLine("Error al abrir la Conexión. Comuniquese con el Administrador", "Error en inicio de sesión");
                if (ex.Message.StartsWith("Your user name and password are not defined"))
                    Console.WriteLine("Error en nombre de usuario o contraseña, por favor verifique", "Error en inicio de sesión");
                if (ex.Message.StartsWith("Unable to complete network request to host"))
                    Console.WriteLine("Error al conectar al servidor. Comuniquese con el Administrador", "Error en inicio de sesión");
                return false;
            }
            return Usuario.Autenticado;
        }
    }
}