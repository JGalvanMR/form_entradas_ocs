using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;

namespace form_entradas_ocs.Data
{
    /// <summary>
    /// Envía correos desde un buzón fijo mediante Microsoft Graph (POST /users/{buzón}/sendMail).
    /// Autenticación de aplicación (credenciales de cliente) contra Microsoft Entra ID.
    /// Implementado sobre HttpWebRequest para ser compatible con .NET Framework 4.0 (sin MSAL ni System.Text.Json).
    /// El secreto se lee de la variable de entorno MRLUCKY_GRAPH_SECRET.
    /// </summary>
    public static partial class GraphMailService
    {
        private const string ClientId = "c982ac22-b59a-441c-905a-9eb273ade26b";
        private const string TenantId = "d20460dd-675d-4b51-87cc-9d10f9175633";
        // Cuenta desde la que se envían todos los correos del sistema.
        private const string Remitente = "sistemasgab@mrlucky.com.mx";
        private const string VariableSecreto = "MRLUCKY_GRAPH_SECRET";

        // Secreto integrado en el ejecutable: lo implementa GraphSecretoLocal.cs (fuera de git). Si no existe, solo se usa la variable de entorno.
        static partial void SecretoIntegrado(ref string secreto);

        private static readonly object _candado = new object();
        private static string _token;
        private static DateTime _tokenExpira = DateTime.MinValue;

        /// <summary>Envía un correo. Lanza excepción si falta el secreto o Graph responde con error.</summary>
        public static void Enviar(IEnumerable<string> para, IEnumerable<string> cc, IEnumerable<string> cco,
            string asunto, string cuerpoHtml, IEnumerable<string> adjuntos)
        {
            string token = ObtenerToken();

            StringBuilder sb = new StringBuilder();
            sb.Append("{\"message\":{");
            sb.Append("\"subject\":").Append(Json(asunto)).Append(",");
            sb.Append("\"body\":{\"contentType\":\"HTML\",\"content\":").Append(Json(cuerpoHtml)).Append("},");
            sb.Append("\"toRecipients\":").Append(Destinatarios(para));
            sb.Append(",\"ccRecipients\":").Append(Destinatarios(cc));
            sb.Append(",\"bccRecipients\":").Append(Destinatarios(cco));

            if (adjuntos != null)
            {
                List<string> items = new List<string>();
                foreach (string ruta in adjuntos)
                {
                    items.Add("{\"@odata.type\":\"#microsoft.graph.fileAttachment\",\"name\":" + Json(Path.GetFileName(ruta)) +
                        ",\"contentBytes\":\"" + Convert.ToBase64String(File.ReadAllBytes(ruta)) + "\"}");
                }
                if (items.Count > 0)
                    sb.Append(",\"attachments\":[").Append(string.Join(",", items.ToArray())).Append("]");
            }

            sb.Append("},\"saveToSentItems\":true}");

            string url = "https://graph.microsoft.com/v1.0/users/" + Uri.EscapeDataString(Remitente) + "/sendMail";
            Post(url, "application/json", sb.ToString(), token);
        }

        /// <summary>Atajo para el caso común: un solo destinatario directo y adjuntos opcionales.</summary>
        public static void Enviar(string para, string asunto, string cuerpoHtml, params string[] adjuntos)
        {
            Enviar(new string[] { para }, null, null, asunto, cuerpoHtml, adjuntos);
        }

        /// <summary>Notifica un error/evento interno a sistemas. Nunca lanza excepción.</summary>
        public static void NotificarError(string detalle)
        {
            try
            {
                Enviar("aescamilla@mrlucky.com.mx", "SISEMP - Entradas OC: aviso del sistema",
                    "<pre>" + WebUtility.HtmlEncode(detalle) + "</pre>");
            }
            catch (Exception)
            {
            }
        }

        private static string ObtenerToken()
        {
            lock (_candado)
            {
                if (_token != null && DateTime.UtcNow < _tokenExpira)
                    return _token;

                string secreto = Environment.GetEnvironmentVariable(VariableSecreto);
                if (string.IsNullOrEmpty(secreto) || secreto.Trim().Length == 0)
                    SecretoIntegrado(ref secreto);
                if (string.IsNullOrEmpty(secreto) || secreto.Trim().Length == 0)
                    throw new InvalidOperationException("Falta el secreto de correo (variable de entorno " + VariableSecreto + " o GraphSecretoLocal.cs).");

                string form = "client_id=" + Uri.EscapeDataString(ClientId) +
                    "&client_secret=" + Uri.EscapeDataString(secreto) +
                    "&scope=" + Uri.EscapeDataString("https://graph.microsoft.com/.default") +
                    "&grant_type=client_credentials";

                string respuesta = Post("https://login.microsoftonline.com/" + TenantId + "/oauth2/v2.0/token",
                    "application/x-www-form-urlencoded", form, null);

                _token = LeerCampo(respuesta, "access_token");
                int segundos;
                if (!int.TryParse(LeerCampo(respuesta, "expires_in"), out segundos)) segundos = 3000;
                // Se renueva 5 minutos antes de que expire.
                _tokenExpira = DateTime.UtcNow.AddSeconds(Math.Max(segundos - 300, 60));
                return _token;
            }
        }

        private static string Post(string url, string contentType, string contenido, string bearer)
        {
            // Microsoft exige TLS 1.2; en .NET 4.0 no existe el valor en el enum.
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | ServicePointManager.SecurityProtocol;

            byte[] datos = Encoding.UTF8.GetBytes(contenido);
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = contentType;
            req.ContentLength = datos.Length;
            if (bearer != null) req.Headers["Authorization"] = "Bearer " + bearer;

            using (Stream s = req.GetRequestStream())
                s.Write(datos, 0, datos.Length);

            try
            {
                using (HttpWebResponse resp = (HttpWebResponse)req.GetResponse())
                using (StreamReader sr = new StreamReader(resp.GetResponseStream(), Encoding.UTF8))
                    return sr.ReadToEnd();
            }
            catch (WebException ex)
            {
                string detalle = ex.Message;
                HttpWebResponse r = ex.Response as HttpWebResponse;
                if (r != null)
                {
                    using (StreamReader sr = new StreamReader(r.GetResponseStream(), Encoding.UTF8))
                        detalle = (int)r.StatusCode + ": " + sr.ReadToEnd();
                }
                throw new InvalidOperationException("Microsoft respondió " + detalle, ex);
            }
        }

        private static string Destinatarios(IEnumerable<string> direcciones)
        {
            List<string> items = new List<string>();
            if (direcciones != null)
                foreach (string d in direcciones)
                    if (!string.IsNullOrEmpty(d) && d.Trim().Length > 0)
                        items.Add("{\"emailAddress\":{\"address\":" + Json(d.Trim()) + "}}");
            return "[" + string.Join(",", items.ToArray()) + "]";
        }

        /// <summary>Extrae un valor simple (cadena o número) de un JSON plano como el de respuesta del token.</summary>
        private static string LeerCampo(string json, string campo)
        {
            string clave = "\"" + campo + "\"";
            int i = json.IndexOf(clave, StringComparison.Ordinal);
            if (i < 0) throw new InvalidOperationException("Respuesta de token inesperada: falta " + campo + ".");
            i = json.IndexOf(':', i + clave.Length) + 1;
            while (i < json.Length && json[i] == ' ') i++;

            if (json[i] == '"')
            {
                int fin = json.IndexOf('"', i + 1);
                return json.Substring(i + 1, fin - i - 1);
            }
            int f = i;
            while (f < json.Length && json[f] != ',' && json[f] != '}') f++;
            return json.Substring(i, f - i).Trim();
        }

        private static string Json(string texto)
        {
            if (texto == null) return "\"\"";
            StringBuilder sb = new StringBuilder("\"");
            foreach (char c in texto)
            {
                switch (c)
                {
                    case '"': sb.Append("\\\""); break;
                    case '\\': sb.Append("\\\\"); break;
                    case '\n': sb.Append("\\n"); break;
                    case '\r': sb.Append("\\r"); break;
                    case '\t': sb.Append("\\t"); break;
                    default:
                        if (c < ' ') sb.Append("\\u").Append(((int)c).ToString("x4"));
                        else sb.Append(c);
                        break;
                }
            }
            return sb.Append("\"").ToString();
        }
    }
}
