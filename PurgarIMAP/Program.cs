using ImapX;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Security.Authentication;
using System.Text;
using System.Threading;

namespace PurgarIMAP {
    class Program {
        static void Main(string[] args)
        {
            Thread.CurrentThread.CurrentCulture = new CultureInfo("en-US");
            string ayudastr = "Parámetros de entrada\n"+
             "Param     Descripción\n"+
             " -h       Ayuda. Este texto\n"+
             " -r       Solo mostrar (No eliminar)\n" +
             " -v       Verbose, salida a consola del detalle del proceso\n"+
             " -q       No pedir confirmaciones.\n" +
             " -e       Pedir confirmación para eliminar.\n" +
             " -u '' *  Usuario del cliente\n"+
             " -p '' *  Contraseña del cliente\n"+
             " -S '' *  Host IMAP\n"+
             " -P ''    Puerto del host\n"+
             " -s       Usar SSL\n" +
             " -T       Usar TLS\n" +
             " -d '' *  Días atrás a eliminar\n"+
             " -f '' *  Carpeta donde buscar mensajes. Case sensitive."; 
            foreach (string s in args)
                Console.Write(s + " ");
            Console.WriteLine("");
            int iUsr=-1, iPass=-1, iServer=-1, iDias=-1, iFolder=-1, iPuerto=-1;
            for (int i = 0; i < args.Count(); i++)
                switch (args[i])
                {
                    case "-u":
                        iUsr = i + 1;
                        break;
                    case "-p":
                        iPass = i + 1;
                        break;
                    case "-P":
                        iPuerto = i + 1;
                        break;
                    case "-S":
                        iServer = i + 1;
                        break;
                    case "-d":
                        iDias = i + 1;
                        break;
                    case "-f":
                        iFolder = i + 1;
                        break;
                }
            if (args.Contains("-h"))
            {
                Console.Write(ayudastr);
                Console.ReadLine();
                return;
            }
            string user = "";
            string pass = "";
            string server = "";
            double dias = 0;
            string fechastr = "";
            string folderemail = "";
            try
            {
                user = args[iUsr];
                pass = args[iPass];
                server = args[iServer];
                dias = -1 * Convert.ToDouble(args[iDias]);
                fechastr = DateTime.Now.AddDays(dias).ToString("dd-MMM-yyyy");
                folderemail = args[iFolder];
            }
            catch (Exception ex)
            {
                Console.WriteLine("Faltan parámetros necesarios");
                Console.ReadLine();
                return;
            }
            string puerto = "993";//IMAP default
            bool verbose = false;
            bool useSSL = false;
            bool useTLS = false;
            bool validarCert = false;
            bool pedirConfirmacionEliminar = false;
            bool pedirConfirmacion = false;
            bool soloMostrar = false;
            if (iPuerto != -1)
                puerto = args[iPuerto];
            if (args.Contains("-s"))
                useSSL = true;
            if (args.Contains("-T"))
                useTLS = true;
            if (args.Contains("-r"))
                soloMostrar = true;
            if (args.Contains("-e"))
                pedirConfirmacionEliminar = true;
            if (args.Contains("-q"))
                pedirConfirmacion= true;
            if (args.Contains("-v"))
            {
                verbose=true;
                Console.WriteLine(" ");
                Console.WriteLine("Conectando...");
            }
            else
            {
                Console.WriteLine("Procesando...");
            }
            ImapClient client;
            if (!useTLS)
                client = new ImapClient(server, Convert.ToInt32(puerto), useSSL, validarCert);
            else
                client = new ImapClient(server, Convert.ToInt32(puerto), SslProtocols.Tls, validarCert);
            //if (false)
            if (client.Connect())
            {
                if (verbose)
                {
                    Console.WriteLine("Conectado!");
                    Console.WriteLine("Iniciando Sesión...");
                }
                if (client.Login(user, pass))
                {
                    if (verbose)
                    {
                        Console.WriteLine("Sesión Iniciada!");
                        Console.WriteLine("Folders:");

                        foreach (Folder f in client.Folders)
                        {
                            Console.WriteLine("[F] " + f.Name);
                        }
                    }
                    var folder=client.Folders[folderemail];
                    if (verbose)
                    {
                        Console.WriteLine("--------------------------------");
                        Console.WriteLine("Descargando emails en ["+folderemail+"]");
                        Console.WriteLine("Desde: " + fechastr );
                    }
                    //fechastr = "12-apr-2018";
                    folder.Messages.Download("BEFORE " + fechastr);
                    if (verbose)
                    {
                        Console.WriteLine("Emails descargados");
                        Console.WriteLine("Llenando la lista de mensajes");
                    }
                    var messages = folder.Search();
                    if (verbose)
                    {
                        Console.WriteLine("--------------------------------");
                        Console.WriteLine("Emails [" + messages.LongLength + "]");
                        Console.WriteLine("--------------------------------");
                        foreach (Message m in messages)
                        {
                            Console.WriteLine("[" + m.Date + "] " + m.Subject);
                        }
                    }
                    var aEliminar = folder.Search("BEFORE "+fechastr);
                    if (verbose)
                    {
                        Console.WriteLine("--------------------------------");
                        Console.WriteLine("Eliminar anteriores a :"+fechastr);
                        Console.WriteLine("A eliminar:" + aEliminar.LongLength);
                        Console.WriteLine("--------------------------------");
                        foreach (Message m in aEliminar)
                        {
                            Console.WriteLine("[" + m.Date + "] " + m.Subject);
                        }
                        Console.WriteLine("--------------------------------");
                        Console.WriteLine(" ");
                    }
                    if(pedirConfirmacionEliminar || !pedirConfirmacion){
                        string confirmarElim="N";
                        if (pedirConfirmacionEliminar)
                        {
                            Console.WriteLine("Preciona Y para Eliminar [" + aEliminar.LongLength + "] mensajes");
                            confirmarElim = Console.ReadLine();
                        }
                    if (!pedirConfirmacion || confirmarElim.ToUpper().Equals("Y"))
                    {
                        try
                        {
                            int contador = 0;
                            foreach (Message m in aEliminar)
                            {
                                contador++;
                                if (verbose || !pedirConfirmacion)
                                { 
                                //Console.WriteLine("Eliminando >>  " + m.Subject);
                                drawTextProgressBar(contador, (int)aEliminar.LongLength);
                                }
                                if (!soloMostrar)
                                {
                                    m.Remove();
                                };
                                
                            }
                            drawTextProgressBar((int)aEliminar.LongLength, (int)aEliminar.LongLength);
                        }catch (Exception ex)
                        {
                            Console.WriteLine("(E!)\t"+ex.Message);
                        }
                        
                    }
                    }else{
                        try
                        {
                            int contador = 0;
                            foreach (Message m in aEliminar)
                            {
                                contador++;
                                if (verbose || !pedirConfirmacion)
                                {
                                    //Console.WriteLine("Eliminando >>  " + m.Subject);
                                    drawTextProgressBar(contador, (int)aEliminar.LongLength);
                                }
                                if (!soloMostrar) {
                                    m.Remove();
                                };
                            }
                            if (verbose || !pedirConfirmacion)
                            {
                                drawTextProgressBar((int)aEliminar.LongLength, (int)aEliminar.LongLength);
                            }
                        }catch (Exception ex)
                        {
                            Console.WriteLine("(E!)\t"+ex.Message);
                        }
                    }
                }
                else
                {
                    Console.Write("No se pudo iniciar Sesión");
                }
            }
            else
            {
                Console.Write("Conexión Fallida");
                
            }
            Console.WriteLine(" ");
            if (!pedirConfirmacion) { 
                Console.WriteLine("Preciona cuaquer tecla para cerrar");
                var confirmar = Console.ReadLine();
                if (confirmar.ToUpper().Equals("X"))
                {
                    return;
                }
            }
            else
            {
                return;
            }
        }


        private static void drawTextProgressBar(int progress, int total)
        {
            //draw empty progress bar
            Console.CursorLeft = 0;
            Console.Write("["); //start
            Console.CursorLeft = 32;
            Console.Write("]"); //end
            Console.CursorLeft = 1;
            float onechunk = 30.0f / total;

            //draw filled part
            int position = 1;
            for (int i = 0; i < onechunk * progress; i++)
            {
                Console.BackgroundColor = ConsoleColor.Gray;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw unfilled part
            for (int i = position; i <= 31; i++)
            {
                Console.BackgroundColor = ConsoleColor.Black;
                Console.CursorLeft = position++;
                Console.Write(" ");
            }

            //draw totals
            Console.CursorLeft = 35;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.Write("Eliminando "+progress.ToString() + " de " + total.ToString() + "    "); //blanks at the end remove any excess
        }
    }
}
